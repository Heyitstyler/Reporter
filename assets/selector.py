import os
import time
import pandas as pd
from directory import *

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

# Select Download Speed
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


# Make the bar folder
os.chdir(dir_Downloads)
exists = os.path.exists(barSelect)
if not exists:
    os.makedirs(barSelect)
os.chdir(barSelect)

dir_BarFolder = os.path.join(dir_Downloads, barSelect)
print (dir_BarFolder)
workingDir = os.getcwd()


