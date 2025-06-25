import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
import time
import re

options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-extensions')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

def get_current_page_number():
    try:
        lblpaging = driver.find_element(By.ID, "lblpaging").text
        match = re.search(r'Page\s+(\d+)\s+of\s+(\d+)', lblpaging)
        if match:
            return int(match.group(1)), int(match.group(2))
    except Exception:
        pass
    return None, None

def safe_extract_text(by, locator):
    try:
        return wait.until(EC.visibility_of_element_located((by, locator))).text.strip()
    except Exception:
        return ""

def force_close_modal():
    try:
        close_button = driver.find_element(By.ID, "LinkButton5")
        driver.execute_script("arguments[0].click();", close_button)
    except Exception:
        pass
    try:
        driver.execute_script("$('#ProductCompany').modal('hide');")
    except Exception:
        pass
    try:
        driver.execute_script("document.querySelector('#ProductCompany').style.display = 'none';")
    except Exception:
        pass

def extract_products_on_current_page():
    wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']")
        )
    )
    data = []
    product_detail_buttons = driver.find_elements(
        By.XPATH, "//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']"
    )
    num_products = len(product_detail_buttons)
    print(f"Found {num_products} products to process on this page")

    for idx in range(num_products):
        print(f"Processing product {idx+1}/{num_products}")

        retry_count = 0
        max_retries = 3

        while retry_count < max_retries:
            try:
                # Re-find buttons to avoid staleness
                product_detail_buttons = driver.find_elements(
                    By.XPATH, "//a[contains(@class, 'btn') and contains(@class, 'purple') and normalize-space()='Product Detail']"
                )
                if idx >= len(product_detail_buttons):
                    print(f"Product {idx+1} not found, skipping...")
                    break

                product_detail_btn = product_detail_buttons[idx]

                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_detail_btn)
                time.sleep(0.5)
                try:
                    product_detail_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", product_detail_btn)

                modal_xpath = "//div[@id='ProductCompany' and contains(@class, 'show')]"
                wait.until(EC.visibility_of_element_located((By.XPATH, modal_xpath)))

                # Expand "Item Description" accordion if not open
                item_desc_accordion_xpath = "//a[contains(@data-bs-toggle, 'collapse') and contains(text(), 'Item Description')]"
                item_desc_collapse_xpath = "//div[@id='shoes']"
                item_desc_accordion = wait.until(EC.presence_of_element_located((By.XPATH, item_desc_accordion_xpath)))
                item_desc_collapse = driver.find_element(By.XPATH, item_desc_collapse_xpath)
                if "show" not in item_desc_collapse.get_attribute("class"):
                    item_desc_accordion.click()
                    time.sleep(0.5)

                dpsu_shq = safe_extract_text(By.ID, "lblcompname")
                item_id = safe_extract_text(By.ID, "lblrefnoview")
                item_name = safe_extract_text(By.ID, "lblitemname1")
                oem_name = safe_extract_text(By.ID, "lbloemname")
                oem_country = safe_extract_text(By.ID, "lbloemcountry")
                oem_name_country = f"{oem_name}: {oem_country}" if oem_name or oem_country else ""
                nato_supply_class = safe_extract_text(By.XPATH, "//th[contains(text(), 'NATO Supply Class')]/following-sibling::td")
                item_name_code = safe_extract_text(By.XPATH, "//th[contains(text(), 'Item Name Code')]/following-sibling::td")

                quantity = ""
                value = ""
                try:
                    import_value_accordion_xpath = "//a[contains(@data-bs-toggle, 'collapse') and contains(text(), 'Import Value, Quantity')]"
                    import_value_collapse_xpath = "//div[@id='Estimated']"
                    import_value_accordion = wait.until(EC.presence_of_element_located((By.XPATH, import_value_accordion_xpath)))
                    import_value_collapse = driver.find_element(By.XPATH, import_value_collapse_xpath)
                    if "show" not in import_value_collapse.get_attribute("class"):
                        try:
                            import_value_accordion.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", import_value_accordion)
                        time.sleep(0.5)
                    import_table_xpath = "//div[@id='Estimated']//table[contains(@class, 'table')][1]"
                    import_table = wait.until(EC.presence_of_element_located((By.XPATH, import_table_xpath)))
                    rows = import_table.find_elements(By.XPATH, ".//tr")
                    for row in rows:
                        cells = row.find_elements(By.XPATH, ".//td")
                        if len(cells) >= 4:
                            year_text = cells[0].text.strip()
                            if year_text and year_text[0].isdigit():
                                quantity = cells[1].text.strip()
                                value = cells[3].text.strip()
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

                close_button = wait.until(EC.element_to_be_clickable((By.ID, "LinkButton5")))
                driver.execute_script("arguments[0].click();", close_button)
                wait.until(EC.invisibility_of_element_located((By.XPATH, modal_xpath)))
                time.sleep(1)
                break  # Success, exit retry loop

            except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException) as e:
                retry_count += 1
                print(f"Retry {retry_count}/{max_retries} for product {idx+1}: {str(e)[:100]}")
                force_close_modal()
                time.sleep(1)
                if retry_count >= max_retries:
                    print(f"Failed to process product {idx+1} after {max_retries} retries")
                    break

            except Exception as e:
                print(f"Error processing product {idx+1}: {e}")
                force_close_modal()
                time.sleep(1)
                break

    return data

def click_bottom_next_button_and_wait_for_page_change():
    current_page, total_pages = get_current_page_number()
    if current_page is None or total_pages is None:
        print("Could not determine current page number")
        return False
    for attempt in range(5):
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            next_buttons = driver.find_elements(By.ID, "lnkbtnPgNext")
            if not next_buttons:
                print("No Next buttons found")
                return False
            next_btn = next_buttons[-1]
            if next_btn.is_displayed() and next_btn.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
                time.sleep(0.5)
                try:
                    next_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", next_btn)
                for _ in range(40):
                    new_page, _ = get_current_page_number()
                    if new_page and new_page > current_page:
                        print(f"Navigated to page {new_page}")
                        return True
                    time.sleep(0.5)
        except Exception as e:
            print(f"Error clicking next button (attempt {attempt+1}): {str(e)[:100]}")
            time.sleep(1)
    print("Failed to click Next button after retries")
    return False

try:
    driver.get("https://srijandefence.gov.in/")
    all_data = []

    # Get total number of pages
    _, total_pages = get_current_page_number()
    if total_pages is None:
        print("Could not determine total number of pages. Exiting.")
        driver.quit()
        exit()

    current_page = 1
    while True:
        print(f"Processing page {current_page} of {total_pages}...")
        page_data = extract_products_on_current_page()
        all_data.extend(page_data)
        print(f"Extracted {len(page_data)} products from page {current_page}")

        # Check if last page
        if current_page >= total_pages:
            print("Reached last page.")
            break

        # Go to next page
        next_clicked = click_bottom_next_button_and_wait_for_page_change()
        if not next_clicked:
            print("Could not navigate to next page, stopping.")
            break

        # Scroll to top for fresh start
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        current_page += 1

    # Save all data to Excel
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel("srijan_products_all_pages.xlsx", index=False)
        print(f"Data saved to srijan_products_all_pages.xlsx - {len(all_data)} products total")
    else:
        print("No data was collected")

finally:
    driver.quit()
