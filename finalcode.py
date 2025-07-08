import pandas as pd
import re
import time
import os
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

def get_current_page_number(page):
    try:
        lblpaging = page.locator("#lblpaging").inner_text(timeout=5000)
        match = re.search(r'Page\s+(\d+)\s+of\s+(\d+)', lblpaging)
        if match:
            return int(match.group(1)), int(match.group(2))
    except Exception:
        pass
    return None, None

def safe_extract_text(page, selector, by='css'):
    try:
        if by == 'css':
            return page.locator(selector).inner_text(timeout=5000).strip()
        else:
            return page.locator(f"xpath={selector}").inner_text(timeout=5000).strip()
    except Exception:
        return ""

def force_close_modal(page):
    try:
        page.locator("#LinkButton5").click(timeout=2000)
    except Exception:
        pass
    try:
        page.evaluate("$('#ProductCompany').modal('hide');")
    except Exception:
        pass
    try:
        page.evaluate("document.querySelector('#ProductCompany').style.display = 'none';")
    except Exception:
        pass

def extract_products_on_current_page(page):
    page.wait_for_selector("//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']", timeout=60000)
    data = []
    product_detail_buttons = page.locator("//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']")
    num_products = product_detail_buttons.count()
    print(f"Found {num_products} products to process on this page")

    for idx in range(num_products):
        print(f"Processing product {idx+1}/{num_products}")

        # Only ONE attempt per product; if fails, skip
        try:
            product_detail_buttons = page.locator("//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']")
            if idx >= product_detail_buttons.count():
                print(f"Product {idx+1} not found, skipping...")
                continue

            product_detail_btn = product_detail_buttons.nth(idx)
            product_detail_btn.scroll_into_view_if_needed(timeout=5000)
            time.sleep(0.1)

            clicked = False
            try:
                product_detail_btn.click(timeout=15000)
                clicked = True
            except Exception:
                try:
                    product_detail_btn.evaluate("el => el.click()")
                    clicked = True
                except Exception as e:
                    print(f"Click failed for product {idx+1}: {e}")
                    clicked = False

            if not clicked:
                print(f"Could not click product {idx+1}, skipping...")
                continue

            # Try to wait for modal, if fails, skip product
            modal_xpath = "//div[@id='ProductCompany' and contains(@class, 'show')]"
            try:
                page.wait_for_selector(modal_xpath, timeout=60000)
            except PlaywrightTimeoutError:
                print(f"Modal did not appear for product {idx+1}, skipping...")
                force_close_modal(page)
                continue

            # Expand "Item Description" accordion if not open
            item_desc_accordion_xpath = "//a[contains(@data-bs-toggle, 'collapse') and contains(text(), 'Item Description')]"
            item_desc_collapse_xpath = "//div[@id='shoes']"
            item_desc_accordion = page.locator(item_desc_accordion_xpath)
            item_desc_collapse = page.locator(item_desc_collapse_xpath)
            if "show" not in item_desc_collapse.get_attribute("class"):
                item_desc_accordion.click(timeout=5000)
                time.sleep(0.5)

            dpsu_shq = safe_extract_text(page, "#lblcompname")
            item_id = safe_extract_text(page, "#lblrefnoview")
            item_name = safe_extract_text(page, "#lblitemname1")
            oem_name = safe_extract_text(page, "#lbloemname")
            oem_country = safe_extract_text(page, "#lbloemcountry")
            oem_name_country = f"{oem_name}: {oem_country}" if oem_name or oem_country else ""
            nato_supply_class = safe_extract_text(page, "//th[contains(text(), 'NATO Supply Class')]/following-sibling::td", by='xpath')
            item_name_code = safe_extract_text(page, "//th[contains(text(), 'Item Name Code')]/following-sibling::td", by='xpath')

            quantity = ""
            value = ""
            try:
                import_value_accordion_xpath = "//a[contains(@data-bs-toggle, 'collapse') and contains(text(), 'Import Value, Quantity')]"
                import_value_collapse_xpath = "//div[@id='Estimated']"
                import_value_accordion = page.locator(import_value_accordion_xpath)
                import_value_collapse = page.locator(import_value_collapse_xpath)
                if "show" not in import_value_collapse.get_attribute("class"):
                    try:
                        import_value_accordion.click(timeout=5000)
                    except Exception:
                        import_value_accordion.evaluate("element => element.click()")
                    time.sleep(0.5)
                import_table_xpath = "//div[@id='Estimated']//table[contains(@class, 'table')][1]"
                import_table = page.locator(import_table_xpath)
                rows = import_table.locator("xpath=.//tr")
                for i in range(rows.count()):
                    row = rows.nth(i)
                    cells = row.locator("xpath=.//td")
                    if cells.count() >= 4:
                        year_text = cells.nth(0).inner_text().strip()
                        if year_text and year_text[0].isdigit():
                            quantity = cells.nth(1).inner_text().strip()
                            value = cells.nth(3).inner_text().strip()
                            break
            except Exception as e:
                print(f"Could not extract import values: {e}")

            data.append({
                "DPSU/SHQ": dpsu_shq,
                "Item Id (Portal)": item_id,
                "Item Name": item_name,
                "OEM Name:Country": oem_name_country,
                "NATO Supply Class": nato_supply_class,
                "Item Name Code": item_name_code,
                "Quantity": quantity,
                "Item value in Lakh(s) Rs (Qty*Price)": value
            })

            print(f"Successfully processed product {idx+1}")

            close_button = page.locator("#LinkButton5")
            close_button.click(timeout=15000)
            page.wait_for_selector(modal_xpath, state="hidden", timeout=60000)
            time.sleep(0.1)

        except Exception as e:
            print(f"Error processing product {idx+1}: {e}")
            force_close_modal(page)
            time.sleep(0.1)
            continue

    return data

def click_bottom_next_button_and_wait_for_page_change(page):
    current_page, total_pages = get_current_page_number(page)
    if current_page is None or total_pages is None:
        print("Could not determine current page number")
        return False
    for attempt in range(5):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2) # More robust wait before clicking Next, helps with slow loading and rate-limiting
            next_buttons = page.locator("#lnkbtnPgNext")
            count = next_buttons.count()
            if count == 0:
                print("No Next buttons found")
                return False
            next_btn = next_buttons.nth(count - 1)
            if not (next_btn.is_visible() and next_btn.is_enabled()):
                print("Next button not visible or enabled")
                return False
            next_btn.scroll_into_view_if_needed(timeout=5000)
            time.sleep(0.5)
            try:
                next_btn.click(timeout=15000)
            except Exception:
                next_btn.evaluate("element => element.click()")
            for _ in range(40):
                new_page, _ = get_current_page_number(page)
                if new_page and new_page > current_page:
                    print(f"Navigated to page {new_page}")
                    time.sleep(2) # Extra sleep after successful page navigation to avoid rate-limiting
                    return True
                time.sleep(0.5)
        except Exception as e:
            print(f"Error clicking next button (attempt {attempt+1}): {str(e)[:100]}")
            time.sleep(2) # Wait longer before retrying
    print("Failed to click Next button after retries")
    return False

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    def block_resources(route):
        if route.request.resource_type in ["image", "stylesheet", "font"]:
            route.abort()
        else:
            route.continue_()
    page.route("**/*", block_resources)

    page.goto("https://srijandefence.gov.in/")
    all_data = []
                                    #***going foward please replace the page number (eg: here they are 120) by the page number you want to stract extraction from
    # Navigate to page 120 before starting extraction
    print("Navigating to page 120...")   #replace here
    current_page, total_pages = get_current_page_number(page)
    if current_page is None or total_pages is None:
        print("Could not determine current page number. Exiting.")
        browser.close()
        exit()
    page_counter = 0
    while current_page < 120: #replace here
        print(f"Current page: {current_page}, navigating to next page...")
        next_clicked = click_bottom_next_button_and_wait_for_page_change(page)
        if not next_clicked:
            print("Could not navigate to next page. Exiting.")
            browser.close()
            exit()
        current_page, _ = get_current_page_number(page)
        print(f"Now on page {current_page}")

    print(f"Starting scraping from page {current_page}")
    while True:
        print(f"Processing page {current_page} of {total_pages}...")
        page_data = extract_products_on_current_page(page)
        all_data.extend(page_data)
        print(f"Extracted {len(page_data)} products from page {current_page}")

        page_counter += 1
        # Write to Excel every 10 pages and clear all_data
        if page_counter % 10 == 0 and all_data:
            df = pd.DataFrame(all_data)
            file_exists = os.path.exists("srijan_products_from_120.xlsx") #replace here
            if file_exists:
                with pd.ExcelWriter("srijan_products_from_120.xlsx", mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer: #replace 120 here
                    df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
            else:
                with pd.ExcelWriter("srijan_products_from_120.xlsx", mode='w', engine='openpyxl') as writer: #replace here
                    df.to_excel(writer, index=False, header=True)
            print(f"Saved {len(all_data)} products to srijan_products_from_120.xlsx (batch at page {current_page})") #replace here
            all_data = []

        if current_page >= total_pages:
            print("Reached last page.")
            break

        next_clicked = click_bottom_next_button_and_wait_for_page_change(page)
        if not next_clicked:
            print("Could not navigate to next page, stopping.")
            break

        page.evaluate("window.scrollTo(0, 0);")
        time.sleep(3) # Larger sleep between page navigations
        current_page += 1

    # Write any remaining data
    if all_data:
        df = pd.DataFrame(all_data)
        file_exists = os.path.exists("srijan_products_from_120.xlsx")
        if file_exists:
            with pd.ExcelWriter("srijan_products_from_120.xlsx", mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
        else:
            with pd.ExcelWriter("srijan_products_from_120.xlsx", mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=True)
        print(f"Saved final {len(all_data)} products to srijan_products_from_120.xlsx")

    browser.close()
