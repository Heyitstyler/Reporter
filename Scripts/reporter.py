import os
import time
import subprocess
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

# Directory
dir_Scripts = os.getcwd()
os.chdir("..")
dir_Root = os.getcwd()
dir_Downloads = dir_Root + r"\_downloads"
dir_DB = dir_Root + r"\DB"



# load bar database
os.chdir(dir_DB)
bars = pd.read_csv("bardb.csv")

# user bar select
barSelect = input("What bar are we working with: ")


# pull pass from db
userRow = bars[bars["user"] == barSelect]

if userRow.empty:
    print("Username not found. Exiting.")
    time.sleep(5)
    exit()

passwd = userRow["pass"]
proper = userRow["proper"]

#print (passwd)
#print (proper)

# Select Download Speed
speed_Input = input("How fast is your internet? 1 - Fast, 2 - Average, 3 - Slow: ")

if speed_Input == '1':
    dlspeed = 20
elif speed_Input == '2':
    dlspeed = 25
elif speed_Input == '3':
    dlspeed = 30
else:
    print("Invalid Entry")
    time.sleep(5)
    exit()
    

# Make the bar folder
os.chdir(dir_Downloads)
exists = os.path.exists(barSelect)
if not exists:
    os.makedirs(barSelect)
os.chdir(barSelect)


dir_BarFolder = os.path.join(dir_Downloads, barSelect)
print (dir_BarFolder)
workingDir = os.getcwd()


# create a new instance of the Firefox driver
options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", workingDir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
options.add_argument("--headless")

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
time.sleep(2)
full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
full_summary.click()

# open Download menu
time.sleep(2)
dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
dropdown_summary.click()

# download summary
time.sleep(2)
download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
download_summary.click()

# select inventories for usage report
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

# select inventories for variance report
var_start_date_drop = variance_driver.find_element(By.ID, "startInventoryId")
var_start_date_drop.click()
time.sleep(1)
var_start_date_select = variance_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
var_start_date_select.click()
var_end_date_drop = variance_driver.find_element(By.ID, "endInventoryId")
var_end_date_drop.click()
time.sleep(1)
var_end_date_select = variance_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
var_end_date_select.click()
time.sleep(3)

# run report
run_js = 'runReport()'
variance_driver.execute_script(run_js)
usage_driver.execute_script(run_js)

# let the report run
time.sleep(4)

# download report
download_js = 'downloadReport()'
variance_driver.execute_script(download_js)
usage_driver.execute_script(download_js)
time.sleep(dlspeed)

# close drivers
variance_driver.quit()
usage_driver.quit()
summary_driver.quit()

# Finished
print ("Downloads Complete! Adjusting Variance.")
os.chdir(dir_Scripts)