import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By


#print(workingDir)

#print(workingDir)
#time.sleep(5)


# load bar database
bars = pd.read_csv("bardb.csv")

# user bar select
barSelect = input("What bar are we working with: ")

# pull pass from db
userRow = bars[bars["user"] == barSelect]
passwd = userRow["pass"]

os.chdir('downloads')
workingDir = os.getcwd()

# create a new instance of the Firefox driver
options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", workingDir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")

# options.add_argument("--headless")
variance_driver = webdriver.Firefox(options=options)
usage_driver = webdriver.Firefox(options=options)
summary_driver = webdriver.Firefox(options=options)

# go to the website
variance_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")
usage_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")
summary_driver.get("https://www.barkeepapp.com/BarkeepOnline/login.php")

# find the username field and type in the username
username_field = variance_driver.find_element(By.NAME, 'session_username')
username_field.send_keys(barSelect)
username_field = usage_driver.find_element(By.NAME, 'session_username')
username_field.send_keys(barSelect)
username_field = summary_driver.find_element(By.NAME, 'session_username')
username_field.send_keys(barSelect)

# find the password field and type in the password
password_field = variance_driver.find_element(By.NAME, 'session_password')
password_field.send_keys(passwd)
password_field = usage_driver.find_element(By.NAME, 'session_password')
password_field.send_keys(passwd)
password_field = summary_driver.find_element(By.NAME, 'session_password')
password_field.send_keys(passwd)

# click the login button
login_button = variance_driver.find_element(By.NAME, 'login')
login_button.click()
login_button = usage_driver.find_element(By.NAME, 'login')
login_button.click()
login_button = summary_driver.find_element(By.NAME, 'login')
login_button.click()

# open reports menu
navigate_reports = usage_driver.find_element(By.ID, 'reportsButton')
navigate_reports.click()
navigate_reports = variance_driver.find_element(By.ID, 'reportsButton')
navigate_reports.click()

# Navigate to reports
navigate_Variance = variance_driver.find_element(By.ID, "varianceReportButton")
navigate_Variance.click()
navigate_Usage = usage_driver.find_element(By.ID, "usageReportButton")
navigate_Usage.click()

# navigate to inventories
navigate_summary = summary_driver.find_element(By.ID, 'inventoriesButton')
navigate_summary.click()

# Choose full Inventory
time.sleep(3)
full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
full_summary.click()

# open Download menu
time.sleep(3)
dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
dropdown_summary.click()

# download summary
time.sleep(3)
download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
download_summary.click()

# run report
run_js = 'runReport()'
variance_driver.execute_script(run_js)
usage_driver.execute_script(run_js)

time.sleep(8)

# download report
download_js = 'downloadReport()'
variance_driver.execute_script(download_js)
usage_driver.execute_script(download_js)