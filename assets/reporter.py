import os
import time
import threading
import pandas as pd
from directory import *
from selector import *
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

time1 = time.perf_counter()

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

time2 = time.perf_counter()

print (f"Done! Ran reports in {time2 - time1:0.2f} seconds.")

restart = input("Would you like to run another bar? (y/n)")
try:
    if restart == "y":
        os.chdir(dir_Root)
        os.system('Run.bat')
    else:
        quit()
except Exception as e:
    print(f"An error occurred: {e}")
    with open("error_log.txt", "a") as log_file:
        log_file.write(f"An error occurred: {e}\n")