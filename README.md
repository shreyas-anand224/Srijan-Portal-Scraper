# Srijan-Portal-Scraper
An automated web scraper built using python and playwright


**Required Libraries** - **Please install these prerequisites in order to run this code**


selenium (for browser automation)

pandas (for saving data to Excel)

openpyxl (for writing Excel .xlsx files with pandas)

chromedriver (for controlling Chrome; not a Python package, but must be installed and in your PATH)

**Installation Instructions**

1. **Install Python** (if not already installed)
Download and install Python 3.8+ from python.org.

2. Upgrade pip (optional but recommended)
bash
**python -m pip install --upgrade pip**
3. Install the required Python packages
Open your terminal or command prompt and run:

**bash
pip install selenium pandas openpyxl**
selenium: For browser automation

pandas: For data manipulation and saving to Excel

openpyxl: For writing Excel files (.xlsx) with pandas

4. Download ChromeDriver
Go to **https://sites.google.com/chromium.org/driver/**

Download the version matching your installed Chrome browser.

Unzip and place the chromedriver (or chromedriver.exe on Windows) somewhere in your PATH, or specify its path in your script, e.g.:

python
driver = webdriver.Chrome(executable_path="C:/path/to/chromedriver.exe", options=options)
5. (Optional) If running headless
Uncomment the --headless line in your options if you want to run without opening a browser window:

python
options.add_argument('--headless')
