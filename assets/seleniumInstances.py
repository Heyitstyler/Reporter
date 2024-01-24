import os
import time
import subprocess
import pandas as pd

from selector import workingDir
from selector import barSelect
from selector import passwd
#from selector import dlspeed

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

def dlSummary():
    keyword = 'Summary'
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", workingDir)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
    options.add_argument("--headless")

    summary_driver = webdriver.Firefox(options=options)

    summary_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")

    username_field = summary_driver.find_element(By.NAME, 'session_username')
    username_field.send_keys(barSelect)
    password_field = summary_driver.find_element(By.NAME, 'session_password')
    password_field.send_keys(passwd)
    login_button = summary_driver.find_element(By.NAME, 'login')
    login_button.click()

    navigate_summary = summary_driver.find_element(By.ID, 'inventoriesButton')
    navigate_summary.click()
    time.sleep(1)
    full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
    full_summary.click()
    time.sleep(1)
    dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
    dropdown_summary.click()
    time.sleep(1)
    download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
    download_summary.click()

    while True:
    # List all files in the specified directory
        files = os.listdir(workingDir)

    # Check if any file contains the keyword
        for file in files:
            if file.startswith(keyword):
                print(f"Found file: {file}")
                time.sleep(0.5)
                summary_driver.quit()
                return
                
            

def dlUsage():
    keyword = 'Usage'
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", workingDir)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
    options.add_argument("--headless")

    usage_driver = webdriver.Firefox(options=options)

    usage_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")

    username_field = usage_driver.find_element(By.NAME, 'session_username')
    username_field.send_keys(barSelect)
    password_field = usage_driver.find_element(By.NAME, 'session_password')
    password_field.send_keys(passwd)
    login_button = usage_driver.find_element(By.NAME, 'login')
    login_button.click()

    navigate_reports = usage_driver.find_element(By.ID, 'reportsButton')
    navigate_reports.click()
    navigate_Usage = usage_driver.find_element(By.ID, "usageReportButton")
    navigate_Usage.click()

    use_start_date_drop = usage_driver.find_element(By.ID, "startInventoryId")
    use_start_date_drop.click()
    time.sleep(1)
    use_start_date_select = usage_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
    use_start_date_select.click()
    use_end_date_drop = usage_driver.find_element(By.ID, "endInventoryId")
    use_end_date_drop.click()
    time.sleep(1)
    use_end_date_select = usage_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
    use_end_date_select.click()
    time.sleep(1)

    run_js = 'runReport()'
    usage_driver.execute_script(run_js)
    time.sleep(4)
    download_js = 'downloadReport()'
    usage_driver.execute_script(download_js)

    while True:
    # List all files in the specified directory
        files = os.listdir(workingDir)

    # Check if any file contains the keyword
        for file in files:
            if file.startswith(keyword):
                print(f"Found file: {file}")
                time.sleep(0.5)
                usage_driver.quit()
                return



def dlVar():
    keyword = 'Variance'
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", workingDir)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
    options.add_argument("--headless")

    variance_driver = webdriver.Firefox(options=options)

    variance_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")

    username_field = variance_driver.find_element(By.NAME, 'session_username')
    username_field.send_keys(barSelect)
    password_field = variance_driver.find_element(By.NAME, 'session_password')
    password_field.send_keys(passwd)
    login_button = variance_driver.find_element(By.NAME, 'login')
    login_button.click()

    navigate_reports = variance_driver.find_element(By.ID, 'reportsButton')
    navigate_reports.click()
    navigate_Variance = variance_driver.find_element(By.ID, "varianceReportButton")
    navigate_Variance.click()

    var_start_date_drop = variance_driver.find_element(By.ID, "startInventoryId")
    var_start_date_drop.click()
    time.sleep(.5)
    var_start_date_select = variance_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
    var_start_date_select.click()
    var_end_date_drop = variance_driver.find_element(By.ID, "endInventoryId")
    var_end_date_drop.click()
    time.sleep(.5)
    var_end_date_select = variance_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
    var_end_date_select.click()
    time.sleep(.5)

    run_js = 'runReport()'
    variance_driver.execute_script(run_js)
    time.sleep(4)
    download_js = 'downloadReport()'
    variance_driver.execute_script(download_js)

    while True:
    # List all files in the specified directory
        files = os.listdir(workingDir)

    # Check if any file contains the keyword
        for file in files:
            if file.startswith(keyword):
                print(f"Found file: {file}")
                time.sleep(0.5)
                variance_driver.quit()
                return