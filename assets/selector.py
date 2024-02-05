import os
import pandas as pd
import datetime
from directory import *

# load bar database
os.chdir(dir_DB)
bars = pd.read_csv("bardb.csv")

# user bar select
while True:
    barSelect = input("What bar are we working with: ")

    # Pull user row from the database
    userRow = bars[bars["user"] == barSelect]

    if userRow.empty:
        print("Username not found. Please try again.")
        # Optional: You might want to add a condition to break the loop after several attempts
        continue  # This will cause the loop to start over again
    else:
        # Code to proceed with the operation for the found username
        break  # Exit the loop when a valid username is found

passwd = userRow["pass"]
proper = userRow["proper"]


# Make the bar folder
    
current_date = datetime.datetime.now()
formatted_date = current_date.strftime(' %Y-%m-%d')

os.chdir(dir_Downloads)
exists = os.path.exists(barSelect + formatted_date)
if not exists:
    os.makedirs(barSelect + formatted_date)
os.chdir(barSelect + formatted_date)

dir_BarFolder = os.path.join(dir_Downloads, barSelect + formatted_date)
print (dir_BarFolder)
workingDir = os.getcwd()


