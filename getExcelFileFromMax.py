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

# Load environment variables from .env file
load_dotenv()

# Retrieve Max password from .env file
username = os.environ.get("MAX_USERNAME")
password = os.environ.get("MAX_PASSWORD")
id = os.environ.get("MAX_ID")

headless = False

def toggleHeadless(headlessState: int): # should be 0 (off) or 1 (on)
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

    service = Service(executable_path="chromedriver.exe")
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    else:
        options.add_argument("--disable-headless-mode")
    # if the driver version is not up to date, download the latest version from the link below
    # https://googlechromelabs.github.io/chrome-for-testing/
    driver = webdriver.Chrome(service=service, options=options)

    # go to url
    driver.get("https://www.max.co.il/")

    # press on main login button
    main_login_button = driver.find_element(By.CLASS_NAME, "personal-text")
    main_login_button.click()

    # wait until the pop up window came 
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
    username_input.send_keys(username)

    # input password
    password_input = driver.find_element(By.ID, "password")
    password_input.send_keys(password)

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
            id_input = driver.find_element(
            By.XPATH, "//div[@id='idInput']/input[1]")
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
    except Exception as e:
        print("Error: ", e)
        driver.save_screenshot("screenshot.png")
        raise Exception("Failed to download the excel file")

    time.sleep(5)

    driver.quit()
    print("Done downloading the excel file")
