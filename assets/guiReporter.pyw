import os
import sys
import time
import threading
import glob
import queue
import datetime
import signal
import pandas as pd
# from directory import *
from barlist import *
from tkinter import *
from PIL import ImageTk, Image
import xlwings as xw

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

#Directory
dir_Assets = os.getcwd()
os.chdir("..")
dir_Root = os.getcwd()
dir_Downloads = dir_Root + r"\_downloads"
dir_DB = dir_Root + r"\DB"

#Root
root = Tk()
root.geometry("800x500")
root.title("Reporter")
root.resizable(False, False)
message_queue = queue.Queue()

#Top Labels
comp_Label = Label(root, text="Companies", background="light blue", width=10, pady=10, font=('Arial', 24))
comp_Label.grid(row=0, column=0, pady=10)

bars_Label = Label(root, text="Bars", background="light blue", width=10, pady=10, font=('Arial', 24))
bars_Label.grid(row=0, column=1, pady=10)

report_Label = Label(root, text="Reporter", background="light blue", width=10, pady=10, font=('Arial', 24))
report_Label.grid(row=0, column=2, pady=10)


#Company Frame
comp_Frame = LabelFrame(root)
comp_Frame.pack_propagate(False)
comp_Frame.config(height=390, width=200)
comp_Frame.grid(row=1, column=0, padx=30, pady=5)


# Bars Frame
bars_Frame = LabelFrame(root)
bars_Frame.pack_propagate(False)
bars_Frame.config(height=390, width=225)
bars_Frame.grid(row=1, column=1, padx=15, pady=25)


# Report Frame
report_Frame = LabelFrame(root)
report_Frame.grid_propagate(False)
report_Frame.config(height=390, width=225)
report_Frame.grid(row=1, column=2, padx=30, pady=5)


# Sub-Report Frames
report_button = Button(report_Frame, text="Run Report", bg="Red", activebackground="yellow", font=("Arial", 16), pady=5, state=DISABLED, anchor=N)
report_button.grid(row=0, column=0, pady=15)


status = Label(report_Frame, text="Status: Ready")
status.grid(row=1, column=0)

# History Frame
hist_Frame = LabelFrame(report_Frame)
hist_Frame.config(height=20, width=200, text="History", labelanchor=N, font=('Arial', 12))
hist_Frame.grid(row=2, column=0, padx=10, ipady=120)
hist_Frame.grid_propagate(False)
hist_Frame.pack_propagate(False)



# Define Button Click Functions
def resetbg(button):
    button.config(bg="light grey")


def adjust(mode):
    print(f"{mode} adjust")
    try:
        # For Windows
        os.system("taskkill /f /im excel.exe")
        print("Excel has been closed.")
    except Exception as e:
        print(f"An error occurred: {e}")
    try:
        matching_files = glob.glob(os.path.join(dir_BarFolder, 'VarianceReport*.xlsx'))
        if matching_files:
            # Iterate through matching files
            for excel_file_path in matching_files:
                # Open the Excel file without displaying the Excel application window
                app = xw.App(visible=False)
                workbook = app.books.open(excel_file_path)

                # Specify the VBA macro name
                macro_name = 'varianceFix'

                # Specify the path to the VBA script or Personal Macro Workbook
                vba_script_path = os.path.join(dir_Assets, 'macroBook.xlsm')

                # Run the VBA macro from the specified script file
                workbook.api.Application.Run("'" + vba_script_path + "'!Module1.varianceFix")

                # Save changes and close the workbook
                workbook.save()
                workbook.close()

                # Close the Excel application
                app.quit()
        else:
            print("No Excel files starting with 'VarianceReport' found in the specified directory.")
    except Exception as e:
        print(str(e))
        input ("Press any button to continue")



def namer(mode):
    global proper
    print(f"{mode} namer")
    proper_str = proper.iloc[0] if isinstance(proper, pd.Series) else str(proper)  # Convert to string

    for filename in os.listdir(dir_BarFolder):
        if os.path.isfile(os.path.join(dir_BarFolder, filename)):
            if proper_str not in filename:
                # Splitting the filename from its extension
                file_base, file_extension = os.path.splitext(filename)
                new_filename = proper_str + "_" + file_base + file_extension

                # Check if the new filename already exists
                count = 1
                while os.path.exists(os.path.join(dir_BarFolder, new_filename)):
                    new_filename = f"{proper_str}_{file_base}_{count}{file_extension}"
                    count += 1
                
                os.rename(os.path.join(dir_BarFolder, filename), os.path.join(dir_BarFolder, new_filename))
                print(f"Renamed '{filename}' to '{new_filename}'")


def on_company_click(button, mode):
    # Define what happens when a button is clicked
    # For demonstration, you can print the mode or handle it as needed
    for widget in comp_Frame.winfo_children():
        widget.configure(bg="light grey")
    for widget in bars_Frame.winfo_children():
        widget.destroy()
    button.configure(bg="dark grey")
    if mode == "EEG": bars_EEG()
    if mode == "Pedal": bars_PEDAL()
    if mode == "Porch": bars_PORCH()
    if mode == "Babbos": bars_BABBOS()
    if mode == "Independant": bars_INDEPENDANT()
    print(f"Button for {mode} clicked")


def on_bar_click(button, mode):
    global hist_Frame, dir_BarFolder, proper, userRow, passwd, workingDir, barSelect, report1_button
    print(f"{mode} is selected")
    bars = pd.read_csv(dir_DB + "\\bardb.csv")

    while True:
        try:
            if mode is None:
                pass
            else:
                barSelect = mode
        except:
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
        
    
    for widget in bars_Frame.winfo_children():
        widget.configure(bg="light grey")
    # for widget in report_Frame.winfo_children():
    #     widget.destroy()
    button.configure(bg="dark grey")
    report_button.forget()
    report_Frame.update()
    report1_button = Button(report_Frame, text="Run Report", background="green", activebackground="yellow", font=("Arial", 16), pady=5)
    report1_button.config(bg="lime", state=NORMAL, command=lambda button=button, mode=mode: run_report(button, mode))
    report1_button.grid(row=0, column=0, pady=15)
    report_Frame.update()




def run_report(button, mode):
    global dir_BarFolder, workingDir
    status.config(text="Status: Running")
    report1_button.config(bg="yellow", state=DISABLED)
    root.update()
    print(f"{mode} is selected in reporter")

    current_date = datetime.datetime.now()
    formatted_date = current_date.strftime(' %Y-%m-%d')

    os.chdir(dir_Downloads)
    exists = os.path.exists(barSelect + formatted_date)
    if not exists:
        os.makedirs(barSelect + formatted_date)
    os.chdir(barSelect + formatted_date)

    dir_BarFolder = os.path.join(dir_Downloads, barSelect + formatted_date)
    os.chdir(dir_BarFolder)
    workingDir = os.getcwd()

    t1 = threading.Thread(target=dlSummary, kwargs={'mode': mode})
    cwd = os.getcwd()
    print(cwd)
    t1.start()


    t2 = threading.Thread(target=dlUsage, kwargs={'mode': mode})
    t2.start()


    t3 = threading.Thread(target=dlVar, kwargs={'mode': mode})
    t3.start()


    t1.join() 
    sum_hist = Label(hist_Frame, text=f"{sum_e}")
    sum_hist.pack() 
    root.update()

    t2.join()
    use_hist = Label(hist_Frame, text=f"{use_e}")
    use_hist.pack()
    root.update()

    t3.join()
    var_hist = Label(hist_Frame, text=f"{var_e}")
    var_hist.pack()
    root.update()







    t4 = threading.Thread(target=adjust, kwargs={'mode': mode})
    t4.start()
    t4.join()


    t5 = threading.Thread(target=namer, kwargs={'mode': mode})
    t5.start()
    t5.join()
    report1_button.config(bg="lime", state=NORMAL)
    status.config(text="Status: Ready")
    # subprocess.run(["init.bat", selected], shell=True)




#Company List
for text, mode in COMPANIES:
    button = Button(comp_Frame, text=text, bg="light grey", font=('Arial', 16))
    button.config(command=lambda button=button, mode=mode:[on_company_click(button, mode)])
    button.pack(pady=15)


#Bars Lists
def bars_EEG():
    for text, mode in EEGBARS:
        button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
        button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
        button.pack(pady=5)

def bars_PEDAL():
    for text, mode in PEDALBARS:
        button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
        button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
        button.pack(pady=5)

def bars_PORCH():
    for text, mode in PORCHBARS:
        button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
        button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
        button.pack(pady=5)

def bars_BABBOS():
    for text, mode in BABBOSBARS:
        button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
        button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
        button.pack(pady=3)
   
def bars_INDEPENDANT():
    for text, mode in INDEPENDANTBARS:
        button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
        button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
        button.pack(pady=5)

def dlSummary(mode):
    global sum_e
    try:
        os.chdir(dir_BarFolder)
        keyword = 'Summary'
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.download.dir", dir_BarFolder)
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

        time.sleep(0.5)

        navigate_summary = summary_driver.find_element(By.ID, 'inventoriesButton')
        navigate_summary.click()
        time.sleep(0.5)
        full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
        full_summary.click()
        time.sleep(1)
        dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
        dropdown_summary.click()
        time.sleep(0.5)
        download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
        download_summary.click()

        while True:
        # List all files in the specified directory
            files = os.listdir(workingDir)

        # Check if any file contains the keyword
            for file in files:
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    sum_e = (f"{mode} Summary report")
                    time.sleep(1)
                    summary_driver.close()
                    time.sleep(0.5)
                    summary_driver.quit()
                    return
    except Exception as sum_e:
        sum_e = ("Error Collecting Summary Report")
        summary_driver.close()
        time.sleep(1)
        summary_driver.quit()
        return
                
            

def dlUsage(mode):
    global use_e
    try:
        dl = f"{dir_Root} + \\_downloads"
        keyword = 'Usage'
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.download.dir", dir_BarFolder)
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

        time.sleep(0.5)

        navigate_reports = usage_driver.find_element(By.ID, 'reportsButton')
        navigate_reports.click()
        navigate_Usage = usage_driver.find_element(By.ID, "usageReportButton")
        navigate_Usage.click()

        time.sleep(1)

        use_start_date_drop = usage_driver.find_element(By.ID, "startInventoryId")
        use_start_date_drop.click()
        time.sleep(0.5)
        use_start_date_select = usage_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
        use_start_date_select.click()
        use_end_date_drop = usage_driver.find_element(By.ID, "endInventoryId")
        use_end_date_drop.click()
        time.sleep(0.5)
        use_end_date_select = usage_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
        use_end_date_select.click()
        time.sleep(0.5)

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
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    use_e = (f"{mode} Usage Report")
                    time.sleep(1)
                    usage_driver.close()
                    time.sleep(0.5)
                    usage_driver.quit()
                    return
    except:
        use_e = ("Error Collecting Usage Report")
        usage_driver.close()
        time.sleep(1)
        usage_driver.quit()



def dlVar(mode):
    global var_e
    try:
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

        time.sleep(0.5)

        navigate_reports = variance_driver.find_element(By.ID, 'reportsButton')
        navigate_reports.click()
        navigate_Variance = variance_driver.find_element(By.ID, "varianceReportButton")
        navigate_Variance.click()

        time.sleep(1)

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
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    var_e = (f"{mode} Variance Report")
                    time.sleep(1)
                    variance_driver.close()
                    time.sleep(0.5)
                    variance_driver.quit()
                    return
    except:
        var_e = ("Error Collecting Variance Report")
        variance_driver.close()
        time.sleep(1)
        variance_driver.quit()


root.mainloop()