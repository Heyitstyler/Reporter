import os
import time
import threading
import pandas as pd
import glob
import xlwings
from directory import *
from selector import *
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

# # Directory
# dir_Assets = os.getcwd()
# os.chdir("..")
# dir_Root = os.getcwd()
# dir_Downloads = dir_Root + r"\_downloads"
# dir_DB = dir_Root + r"\DB"

# # load bar database
# os.chdir(dir_DB)
# bars = pd.read_csv("bardb.csv")

# # user bar select
# barSelect = input("What bar are we working with: ")

# # pull pass from db
# userRow = bars[bars["user"] == barSelect]

# if userRow.empty:
#     print("Username not found. Exiting.")
#     time.sleep(5)
#     exit()

# passwd = userRow["pass"]
# proper = userRow["proper"]

# # Select Download Speed
# speed_Input = input("How fast is your internet? 1 - Fast, 2 - Average, 3 - Slow: ")

# if speed_Input == '1':
#     dlspeed = 20
# elif speed_Input == '2':
#     dlspeed = 25
# elif speed_Input == '3':
#     dlspeed = 30
# else:
#     print("Invalid Entry")
#     time.sleep(5)
#     exit()


# # Make the bar folder
# os.chdir(dir_Downloads)
# exists = os.path.exists(barSelect)
# if not exists:
#     os.makedirs(barSelect)
# os.chdir(barSelect)

# dir_BarFolder = os.path.join(dir_Downloads, barSelect)
# print (dir_BarFolder)
# workingDir = os.getcwd()

from seleniumInstances import dlSummary
t1 = threading.Thread(target=dlSummary)
t1.start()

from seleniumInstances import dlUsage
t2 = threading.Thread(target=dlUsage)
t2.start()

from seleniumInstances import dlVar
t3 = threading.Thread(target=dlVar)
t3.start()

t1.join()
t2.join()
t3.join()

from adjuster import adjust
t4 = threading.Thread(target=adjust)
t4.start()
t4.join()



from adjuster import namer
t5 = threading.Thread(target=namer)
t5.start()
t5.join()


print ("Done!")

restart = input("Would you like to run another bar? (y/n)")
if restart == "y":
    os.chdir(dir_Root)
    os.system(dir_Root + r"/Run.bat")
else:
    quit()