from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from dotenv import load_dotenv
from data import *
from threading import Thread
import chromedriver_autoinstaller
from get_for_ex_trans import get_foreign_exchange_transactions
from add_rows_to_csv import parse_amount
import shutil
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains
from utils import *

# Load environment variables from .env file
load_dotenv(".env", override=True)

# Retrieve Max password from .env file
max_username = os.environ.get("MAX_USERNAME")
max_password = os.environ.get("MAX_PASSWORD")
id = os.environ.get("ID")
bank_password = os.environ.get("BANK_PASSWORD")
bank_id_code = os.environ.get("BANK_ID_CODE")


headless = False


def toggleHeadless(headlessState: int):  # should be 0 (off) or 1 (on)
    global headless
    if headlessState == 1:
        headless = True
    else:
        headless = False

def getExcelFileThreaded(year, month):
    Thread(target=getExcelFile, args=(year, month), daemon=True).start()

def initializeDriver():
    print("Initializing Chrome driver")
    # Ensure chromedriver_autoinstaller will provide the matching driver for the current Chrome.
    # chromedriver_autoinstaller.install() returns the path to the chromedriver executable
    # and will only download when the matching driver isn't already present.
    try:
        chromeDriverPath = chromedriver_autoinstaller.install()
        print(f"chromedriver_autoinstaller returned: {chromeDriverPath}")
    except Exception as e:
        print("chromedriver_autoinstaller.install() failed:", e)
        # Fallback to any chromedriver on PATH (useful if autoinstaller can't run)
        chromeDriverPath = shutil.which("chromedriver")
        if chromeDriverPath:
            print(f"Falling back to chromedriver on PATH: {chromeDriverPath}")
        else:
            raise Exception("No chromedriver available (autoinstaller failed and no chromedriver on PATH)") from e

    service = Service(executable_path=str(chromeDriverPath))
    print("Chrome driver service created")
    options = webdriver.ChromeOptions()
    # Use the new headless flag for recent Chrome versions
    if headless:
        options.add_argument("--headless=new")
    # keep default (non-headless) otherwise; remove invalid flag
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")

    try:
        driver = webdriver.Chrome(service=service, options=options)
        print("driver initialized successfully")
        return driver
    except Exception as e:
        print("Error initializing Chrome driver: ", e)
        return None
    
def connectToMax(driver):
    # wait until the pop up window shows
    try:
        print("Waiting for the pop-up window to appear")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "כניסה עם סיסמה"))
        )
        print("Pop-up window appeared")
        login_with_password_button = driver.find_element(
            By.PARTIAL_LINK_TEXT, "כניסה עם סיסמה"
        )
        login_with_password_button.click()
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("pop_up_window_screenshot.png")
        raise Exception("Failed to click on the pop-up window")

    # input email
    username_input = driver.find_element(By.ID, "user-name")
    username_input.send_keys(max_username)

    # input password
    password_input = driver.find_element(By.ID, "password")
    password_input.send_keys(max_password)

    # press enter to complete login
    password_input.send_keys(Keys.ENTER)

    # Wait for the ID input field to appear, if it exists
    try:
        print("Checking if ID input field appears")
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "idInput"))
        )
        print("ID input field appeared")
        try:
            id_input = driver.find_element(By.XPATH, "//div[@id='idInput']/input[1]")
            id_input.send_keys(id)
            time.sleep(3)
            # press enter to complete login
            password_input.send_keys(Keys.ENTER)
            time.sleep(3)
        except Exception as e:
            print("Error: ", e)
    except Exception:
        print("ID input field did not appear, continuing without it")

def goToMaxTransactionDetails(driver, year, month):
    # go to transaction details
    print("Going to transaction details")
    if month == "12":
        monthInt = "1"
        yearInt = str(int(year) + 1)
    else:
        monthInt = str(int(month) + 1)
        yearInt = year
    driver.get(
        "https://www.max.co.il/transaction-details/personal?filter=-1_-1_1_"
        + yearInt
        + "-"
        + monthInt
        + "-01_0_0_-1&sort=1a_1a_1a_1a_1a_1a"
    )

def downloadMaxExcelFile(driver, year, month):
    # wait until the download button comes up
    try:
        print("Waiting for the download button to appear")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "download-excel"))
        )
        print("Download button appeared")
        download_excel = driver.find_element(By.CLASS_NAME, "download-excel")
        download_excel.click()
        print("Done downloading the excel file")
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("screenshot.png")
        raise Exception("Failed to download the excel file")
    
def getDealTable(driver):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 
                                            "app-table.ng-star-inserted:nth-child(6) > div:nth-child(1)"))
        )
        print("Deal table appeared")
        dealTable = driver.find_element(
            By.CSS_SELECTOR,
            "app-table.ng-star-inserted:nth-child(6) > div:nth-child(1)",
        )
        if dealTable:
            print("Deal table found")
            html = dealTable.get_attribute("outerHTML")
            with open("foreign_exchange_transactions.html", "w", encoding="utf-8") as f:
                f.write(html)
            print("Deal table HTML saved to foreign_exchange_transactions.html")
        else:
            raise Exception("Deal table not found")
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("deal_table_screenshot.png")
        raise Exception("Deal table not found, check the screenshot")

def closeDriver(driver):
    try:
        driver.quit()
        print("Driver closed successfully")
    except Exception as e:
        print("Error closing the driver: ", e)

def saveExcelFile(year, month):
    downloads_path = str(Path.home() / "Downloads")
    destination_path = ""

    files = [
        os.path.join(downloads_path, f)
        for f in os.listdir(downloads_path)
        if f.endswith(".xlsx")
    ]
    latest_file = max(files, key=os.path.getctime) if files else None

    target_file = f"transaction-details_export_{year}_{month}.xlsx"
    shutil.move(
        latest_file, os.path.join(destination_path, os.path.basename(target_file))
    )

    get_foreign_exchange_transactions(
        "foreign_exchange_transactions.html", "transactions.csv"
    )
    parse_amount("transactions.csv", target_file)

def goToMax(driver):
    # go to url
    try:
        driver.get("https://www.max.co.il/login")
    except Exception as e:
        print("Error navigating to Max login page: ", e)
        driver.quit()
        return
    print("driver initialized, navigating to Max login page")
    
    connectToMax(driver)

def getExcelFile(year, month):
    year, month = checkDate(year, month)
    print(f"date = {year}-{month}")

    driver = initializeDriver()

    goToMax(driver)

    goToMaxTransactionDetails(driver, year, month)

    downloadMaxExcelFile(driver, year, month)

    getDealTable(driver)

    time.sleep(5)

    # goToDiscount(driver)

    # goToDiscountTransactionDetails(driver)
    # Get 3 months back transactions excel
    # Get the relevant timeframe - 10.{month} - 9.{month+1}
    # Check for the next keywords: משיכת שיק, החזר דיסקונט, הפקדת שיק, העברה ל, העברה מ, עמלת פעולה, אלטשולר שח, טפחות-משכנ, טפחות ס.בי, הע. ל, ביטוח לאומי - ילדים, מופ"ת, נאנומושן, הו"ק למיטב, עמלת סמס, 
    #

    # downloadDiscountExcelFile(driver)

    # manipulateDiscountExcelFile(year, month)

    closeDriver(driver)

    saveExcelFile(year, month)
