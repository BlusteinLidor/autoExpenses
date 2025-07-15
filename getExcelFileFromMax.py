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


def getExcelFile(year, month):
    # if year is not in the combobox options, set year to default value
    if year not in years:
        year = defaultYear
    # if month is not in the combobox options, set month to default value
    if month not in months:
        month = defaultMonth

    chromeDriverPath = chromedriver_autoinstaller.install()
    service = Service(executable_path=f"{chromeDriverPath}")
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    else:
        options.add_argument("--disable-headless-mode")
    # if the driver version is not up to date, download the latest version from the link below
    # https://googlechromelabs.github.io/chrome-for-testing/
    try:
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        print("Error initializing Chrome driver: ", e)

    # go to url
    driver.get("https://www.max.co.il/login")

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

    # wait until the pop up window came up
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

    try:
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

    time.sleep(5)

    driver.get("https://start.telebank.co.il/login/#/LOGIN_PAGE")

    # wait for the form to load
    try:
        print("Waiting for the Discount Bank login form to appear")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "loginForm"))
        )
        print("Discount Bank login form appeared")
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("discount_bank_login_form_screenshot.png")
        raise Exception("Failed to load Discount Bank login form")

    # input id
    bank_id_input = driver.find_element(By.ID, "tzId")
    bank_id_input.send_keys(id)

    # input password
    bank_password_input = driver.find_element(By.ID, "tzPassword")
    bank_password_input.send_keys(bank_password)

    # bank id code
    bank_id_code_input = driver.find_element(By.ID, "aidnum")
    bank_id_code_input.send_keys(bank_id_code)

    driver.find_element(By.CSS_SELECTOR, ".sendBtn").click()

    time.sleep(5)
    # wait for the transactions page to load

    driver.get("https://start.telebank.co.il/apollo/retail/#/OSH_LENTRIES_ALTAMIRA")

    try:
        print("Waiting for the transactions table to appear")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".advanced-search-btn-icon"))
        )
        print("Transactions table appeared")
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("transactions_table_screenshot.png")
        raise Exception("Failed to load transactions table")
    
    # click on the advanced search button
    try:
        advanced_search_button = driver.find_element(By.CSS_SELECTOR, ".advanced-search-btn-icon")
        advanced_search_button.click()
        print("Advanced search button clicked")
    except Exception as e:
        print("Error clicking advanced search button: ", e)
        driver.save_screenshot("advanced_search_button_screenshot.png")
        raise Exception("Failed to click on the advanced search button")

    try:
        print("Waiting for the pop-up to appear")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "fromDate"))
        )
        print("Pop-up appeared")
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("popup_screenshot.png")
        raise Exception("Failed to load the pop-up")
    
    # input the date range
    from_date_input = driver.find_element(By.ID, "fromDate")
    from_date_input.click()
    time.sleep(0.5)
    year_input = driver.find_element(By.CSS_SELECTOR, "button.current:nth-child(3) > span:nth-child(1)")
    if year_input.value_of_css_property("innerText") == year:
        year_input.click()
        driver.find_element(By.CSS_SELECTOR, "button.current:nth-child(2) > span:nth-child(1)").click()
        if month % 3 == 0:
            col = 3
        elif month % 3 == 1:
            col = 1
        else:
            col = 2
        if month <= 3:
            row = 1
        elif month <= 6:
            row = 2
        elif month <= 9:
            row = 3
        else:
            row = 4
        driver.find_element(By.CSS_SELECTOR, f"tr.ng-star-inserted:nth-child({row}) > td:nth-child({col}) > span:nth-child(1)").click()
        time.sleep(0.5)
        driver.find_element(By.CSS_SELECTOR, "tr.ng-star-inserted:nth-child(2) > td:nth-child(3) > span:nth-child(1)").click()
        time.sleep(0.5)
        driver.find_element(By.ID, "oshTransfersAdvancedSearchDateTO").click()
        year_input = driver.find_element(By.CSS_SELECTOR, "button.current:nth-child(3) > span:nth-child(1)")
        if year_input.value_of_css_property("innerText") == year:
            year_input.click()
            driver.find_element(By.CSS_SELECTOR, "button.current:nth-child(2) > span:nth-child(1)").click()
            col += 1
            if col > 3 and row < 4:
                col = 1
                row += 1
            driver.find_element(By.CSS_SELECTOR, f"tr.ng-star-inserted:nth-child({row}) > td:nth-child({col}) > span:nth-child(1)").click()
            time.sleep(0.5)
            driver.find_element(By.CSS_SELECTOR, "tr.ng-star-inserted:nth-child(2) > td:nth-child(3) > span:nth-child(1)").click()
            time.sleep(0.5) 
            driver.find_element(By.CSS_SELECTOR, "button.advanced-search-btn").click()
        
        

    time.sleep(5)

    try:
        driver.quit()
        print("Driver closed successfully")
    except Exception as e:
        print("Error closing the driver: ", e)

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
