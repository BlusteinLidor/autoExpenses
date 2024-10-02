from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve Max password from .env file
password=os.environ.get("MAX_PASSWORD")

# months list
months=["1","2","3","4","5","6","7","8","9","10","11","12"]
# year list
years=["2022","2023","2024"]
# default year value
defaultYear = years[-1]
# default month value
defaultMonth = months[0]

def getExcelFile(year, month):
    #if year is not in the combobox options, set year to default value
    if year not in years:
        year = defaultYear
    # if month is not in the combobox options, set month to default value
    if month not in months:
        month = defaultMonth

    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome()

    # go to url
    driver.get("https://www.max.co.il/")

    # press on main login button
    main_login_button = driver.find_element(By.CLASS_NAME, "personal-text")
    main_login_button.click()

    # wait until the pop up window came up
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "כניסה עם סיסמה"))
    )
    # then click on the option - "login with password"
    login_with_password_button = driver.find_element(By.PARTIAL_LINK_TEXT, "כניסה עם סיסמה")
    login_with_password_button.click()

    # input email
    username_input = driver.find_element(By.ID, "user-name")
    username_input.send_keys("liidorr@gmail.com")

    # input password
    password_input = driver.find_element(By.ID, "password")
    password_input.send_keys(password)

    # press enter to complete login
    password_input.send_keys(Keys.ENTER)

    time.sleep(5)

    # go to transaction details
    driver.get("https://www.max.co.il/transaction-details/personal?filter=-1_-1_1_" + year + "-" + month + "-01_0_0_-1&sort=1a_1a_1a_1a_1a_1a")

    # wait until the pop up window came up
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, "download-excel"))
    )

    # open the months tab
    download_excel = driver.find_element(By.CLASS_NAME, "download-excel")
    download_excel.click()

    time.sleep(5)

    driver.quit()
