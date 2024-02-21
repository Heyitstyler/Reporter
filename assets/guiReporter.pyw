import os
import sys
import time
import threading
import glob
import queue
import datetime
import signal
import pandas as pd
import requests
# from directory import *
from barlist import *
from tkinter import *
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from PIL import ImageTk, Image
import xlwings as xw

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

#Version
version = "Reporter 1.7"
hist_Track = 0

#Directory
dir_Assets = os.getcwd()
os.chdir("..")
dir_Root = os.getcwd()
dir_Downloads = dir_Root + r"\_downloads"
dir_DB = dir_Root + r"\DB"

try:
    checkint = requests.get("https://www.google.com", timeout=3)
except:
    print("Failed to determine download speed. do you have an internet")


def download_file(url):

    local_filename = url.split('/')[-1]
    start = time.time()
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    end = time.time()
    return end - start, local_filename


def calculate_speed(download_time, file_size_mb):
    # Speed in Mbps
    speed_mbps = file_size_mb / download_time
    os.remove("5MB.zip")
    return speed_mbps * 8  # Convert MBps to Mbps

# URL of the 10MB file
file_url = 'http://ipv4.download.thinkbroadband.com/5MB.zip'
file_size_mb = 10  # Size of the file in MB

download_time, _ = download_file(file_url)
download_speed_mbps = calculate_speed(download_time, file_size_mb)

print(f"Download completed in {download_time:.2f} seconds.")
print(f"Download speed: {download_speed_mbps:.2f} Mbps")

download_time = round(download_time, 2)
download_speed_mbps = round(download_speed_mbps, 2)
loadTime = 1
if download_time > 50:
    loadTime = loadTime * 6
elif download_time > 45:
    loadTime = loadTime * 5
elif download_time > 40:
    loadTime = loadTime * 4
elif download_time > 35:
    loadTime = loadTime * 3
elif download_time > 30:
    loadTime = loadTime * 2

os.chdir(dir_Root)
log = open("dllog.txt", "a")
L = [f"Download Time: {download_time} ", f" Download Speed: {download_speed_mbps}\n"]
log.writelines(L)
log.close()


#Update DB
def updateDB():
    global hist_Track
    listpyURL = "https://raw.githubusercontent.com/Heyitstyler/Reporter/main/assets/barlist.py"
    csvURL = "https://raw.githubusercontent.com/Heyitstyler/Reporter/main/DB/bardb.csv"
    try:
        os.chdir(dir_DB)
        requests.get(csvURL, timeout=5)
        try:
            os.remove("bardb.backup.csv")
        except:
            print("no backup bardb")
        os.rename("bardb.csv", "bardb.backup.csv")
        download_file(csvURL)
        print("Downloaded New bardb.csv")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Downloaded new bardb.csv").pack()
        hist_Track = hist_Track + 1
    except:
        print("error downloading new bardb.csv")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Error downloading new bardb.csv").pack()
        hist_Track = hist_Track + 1
    
    try:
        os.chdir(dir_Assets)
        requests.get(listpyURL, timeout=5)
        try:
            os.remove("barlist.backup.py")
        except:
            print("no backup barlist")
        os.rename("barlist.py", "barlist.backup.py")
        download_file(listpyURL)
        print("Downloaded New barlist.py")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Downloaded New barlist.py").pack()
        hist_Track = hist_Track + 1
    except:
        print("error downloading new barlist.py")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Error downloading new barlist.py").pack()
        hist_Track = hist_Track + 1

#Root
root = Tk()
root.geometry("800x525")
root.title("Reporter")
root.resizable(False, False)
message_queue = queue.Queue()

#top menu
topMenu = Menu(root)
root.config(menu=topMenu)

update_menu = Menu(topMenu)
topMenu.add_cascade(label="Update", menu=update_menu)
update_menu.add_command(label="Update Bar Database", command = lambda:updateDB())

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
comp_Frame.config(height=425, width=200)
comp_Frame.grid(row=1, column=0, padx=30, pady=5)


# Bars Frame
bars_Frame = LabelFrame(root)
bars_Frame.pack_propagate(False)
bars_Frame.config(height=425, width=225, pady=5)
bars_Frame.grid(row=1, column=1, padx=15, pady=5)


# Report Frame
report_Frame = LabelFrame(root)
report_Frame.grid_propagate(False)
report_Frame.config(height=425, width=226)
report_Frame.grid(row=1, column=2, padx=30, pady=5)


# Sub-Report Frames
report_button = Button(report_Frame, text="Run Report", bg="Red", activebackground="yellow", font=("Arial", 16), pady=5, state=DISABLED, anchor=N)
report_button.grid(row=0, column=0, pady=15, columnspan=3)

rpt = IntVar()
rpt.set(1)

reportTypeFull = Radiobutton(report_Frame, text="Full", variable=rpt, value=1)
reportTypeReport = Radiobutton(report_Frame, text="Just Report", variable=rpt, value=2)
reportTypeInvoice = Radiobutton(report_Frame, text="Just Invoice", variable=rpt, value=3)

reportTypeFull.grid(row=1, column=0)
reportTypeReport.grid(row=1, column=1)
reportTypeInvoice.grid(row=1, column=2)

status = Label(report_Frame, text="Status: Ready")
status.grid(row=2, column=0, columnspan=3)

# History Frame
def history():
    global hist_Frame
    hist_Frame = LabelFrame(report_Frame)
    hist_Frame.config(height=20, width=200, text="History", labelanchor=N, font=('Arial', 12))
    hist_Frame.grid(row=3, column=0, padx=10, ipady=130, columnspan=3)
    hist_Frame.grid_propagate(False)
    hist_Frame.pack_propagate(False)
history()




# Define Button Click Functions
def resetbg(button):
    button.config(bg="light grey")


def adjust(mode):
    try:
        # For Windows
        os.system("taskkill /f /im excel.exe")
        print("Excel has been closed.")
    except:
        print(f"No Excel files to close.")
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
    global hist_Frame, dir_BarFolder, proper, userRow, passwd, workingDir, barSelect, report1_button, street, city, inv, price
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
            continue
        else:
            break

    passwd = userRow["pass"].iloc[0]
    proper = userRow["proper"].iloc[0]
    street = userRow["street"].iloc[0]
    city = userRow["city"].iloc[0]
    inv = userRow["invoicename"].iloc[0]
    price = userRow["price"].iloc[0]
        
    
    for widget in bars_Frame.winfo_children():
        widget.configure(bg="light grey")
    button.configure(bg="dark grey")
    report_button.config(bg="lime", state=NORMAL, command=lambda button=button, mode=mode: run_report(button, mode))
    report_Frame.update()




def run_report(button, mode):
    global dir_BarFolder, workingDir, loadTime, hist_Track
    time1 = time.perf_counter()

    status.config(text="Status: Running")
    report_button.config(bg="yellow", state=DISABLED)
    root.update()
    print(f"Running reports for {mode}")

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

    rptOption = rpt.get()

    if rptOption == 1:

        if hist_Track >= 9:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
            

        hist_Track = hist_Track + 4
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

        t6 = threading.Thread(target=Invoice)
        t6.start()
        t6.join()
        inv_hist = Label(hist_Frame, text=f"Generated {proper} Invoice")
        inv_hist.pack()

        time2 = time.perf_counter()

        print(f"Ran Reporter in {time2 - time1:0.2f} seconds.")

        report_button.config(bg="lime", state=NORMAL)
        status.config(text="Status: Ready")

    elif rptOption == 2:

        if hist_Track >= 10:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()

        hist_Track = hist_Track + 3
    
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

        time2 = time.perf_counter()

        print(f"Downloaded reports in {time2 - time1:0.2f} seconds.")

        report_button.config(bg="lime", state=NORMAL)
        status.config(text="Status: Ready")

    elif rptOption == 3:

        if hist_Track >= 12:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()

        hist_Track = hist_Track + 1

        t6 = threading.Thread(target=Invoice)
        t6.start()
        t6.join()
        inv_hist = Label(hist_Frame, text=f"Generated {proper} Invoice")
        inv_hist.pack()

        time2 = time.perf_counter()

        print(f"Generated Invoice in {time2 - time1:0.2f} seconds.")

        report_button.config(bg="lime", state=NORMAL)
        status.config(text="Status: Ready")



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
        button.pack(pady=2)

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
        button.pack(pady=3)

def dlSummary(mode):
    global sum_e
    try:
        sum_e = "Failed"
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

        time.sleep(loadTime)

        navigate_summary = summary_driver.find_element(By.ID, 'inventoriesButton')
        navigate_summary.click()
        time.sleep(loadTime)
        full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
        full_summary.click()
        time.sleep(loadTime)
        dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
        dropdown_summary.click()
        time.sleep(loadTime/2)
        download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
        download_summary.click()

        while True:
        # List all files in the specified directory
            files = os.listdir(workingDir)

        # Check if any file contains the keyword
            for file in files:
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    sum_e = (f"{proper} Summary report")
                    time.sleep(1)
                    summary_driver.close()
                    time.sleep(0.5)
                    summary_driver.quit()
                    return
                
    except:
        sum_e = ("Error Collecting Summary Report")
        summary_driver.close()
        time.sleep(1)
        summary_driver.quit()
        os.chdir(dir_Root)
        log = open("dllog.txt", "a")
        L = [f"Failed Summary Report\n"]
        log.writelines(L)
        log.close()
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

        time.sleep(loadTime)

        navigate_reports = usage_driver.find_element(By.ID, 'reportsButton')
        navigate_reports.click()
        navigate_Usage = usage_driver.find_element(By.ID, "usageReportButton")
        navigate_Usage.click()

        time.sleep(loadTime)

        use_start_date_drop = usage_driver.find_element(By.ID, "startInventoryId")
        use_start_date_drop.click()
        time.sleep(loadTime/2)
        use_start_date_select = usage_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
        use_start_date_select.click()
        use_end_date_drop = usage_driver.find_element(By.ID, "endInventoryId")
        use_end_date_drop.click()
        time.sleep(loadTime/2)
        use_end_date_select = usage_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
        use_end_date_select.click()
        time.sleep(loadTime/2)

        run_js = 'runReport()'
        usage_driver.execute_script(run_js)
        time.sleep(loadTime*4)
        download_js = 'downloadReport()'
        usage_driver.execute_script(download_js)

        while True:
        # List all files in the specified directory
            files = os.listdir(workingDir)

        # Check if any file contains the keyword
            for file in files:
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    use_e = (f"{proper} Usage Report")
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
        os.chdir(dir_Root)
        log = open("dllog.txt", "a")
        L = [f"Failed Usage Report\n"]
        log.writelines(L)
        log.close()



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

        time.sleep(loadTime)

        navigate_reports = variance_driver.find_element(By.ID, 'reportsButton')
        navigate_reports.click()
        navigate_Variance = variance_driver.find_element(By.ID, "varianceReportButton")
        navigate_Variance.click()

        time.sleep(loadTime)

        var_start_date_drop = variance_driver.find_element(By.ID, "startInventoryId")
        var_start_date_drop.click()
        time.sleep(loadTime/2)
        var_start_date_select = variance_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
        var_start_date_select.click()
        var_end_date_drop = variance_driver.find_element(By.ID, "endInventoryId")
        var_end_date_drop.click()
        time.sleep(loadTime/2)
        var_end_date_select = variance_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
        var_end_date_select.click()
        time.sleep(loadTime/2)

        run_js = 'runReport()'
        variance_driver.execute_script(run_js)
        time.sleep(loadTime*4)
        download_js = 'downloadReport()'
        variance_driver.execute_script(download_js)

        while True:
        # List all files in the specified directory
            files = os.listdir(workingDir)

        # Check if any file contains the keyword
            for file in files:
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    var_e = (f"{proper} Variance Report")
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
        os.chdir(dir_Root)
        log = open("dllog.txt", "a")
        L = [f"Failed Variance Report\n"]
        log.writelines(L)
        log.close()

def Invoice():
    os.chdir(dir_BarFolder)
    today = datetime.datetime.now()
    c = canvas.Canvas(f"{proper} invoice {today:%Y-%m-%d}.pdf", pagesize=letter)

    # Set font size, color, and spacing
    font_size = 12
    marginleft = 70
    spacey = 17

    # Service Provider Information
    c.setFont("Helvetica-Bold", 18)
    c.setFillColorRGB(7/255, 55/255, 99/255)  # Setting text color to black
    c.drawString(marginleft, 720 - spacey, "GDS Consulting LLC")
    c.setFont("Helvetica", font_size)
    c.setFillColor("black")  # Setting text color to black
    c.drawString(marginleft, 720 - spacey*2, "3650 South Joshua Tree Lane")
    c.drawString(marginleft, 720 - spacey*3, "Gilbert, Arizona 85297")
    c.drawString(marginleft, 720 - spacey*4, "Phone - (480) 593-0573")
    c.drawString(marginleft, 720 - spacey*5, "Email - GDSConsultingllc@gmail.com")

    # Client Information
    c.setFillColorRGB(7/255, 55/255, 99/255)  # Setting text color to black
    c.drawString(marginleft, 610 - spacey*2, "BILLED TO")
    c.setFont("Helvetica-Bold", font_size)
    c.setFillColor("black")  # Setting text color to black
    c.drawString(marginleft, 610 - spacey*3, f"{proper}")
    c.drawString(marginleft, 610 - spacey*4, f"{street}")
    c.drawString(marginleft, 610 - spacey*5, f"{city}")

    # Invoice Details
    c.setFont("Helvetica-Bold", 24)
    c.setFillColorRGB(7/255, 55/255, 99/255)  # Setting text color to blue
    c.drawString(marginleft, 570 - spacey*7, "Invoice")
    c.setFont("Helvetica", font_size)
    c.setFillColor("black")  # Setting text color to black
    c.drawString(marginleft, 570 - spacey*8.2, f"Invoice #{inv}:{today:%m%d%Y}")
    c.drawString(marginleft, 570 - spacey*9.2, f"{today:%m/%d/%Y}")
    c.drawString(marginleft + 10, 570 - spacey*12, "Description")
    c.drawString(marginleft + 440, 570 - spacey*12, "Total")
    c.setFont("Helvetica-Bold", font_size)
    c.drawString(marginleft + 10, 570 - spacey*13.5, "Audit & Consultation")
    c.drawString(marginleft + 425, 570 - spacey*13.5, f"${price}.00")
    c.line(marginleft, 570 - spacey*15, marginleft + 480, 570 - spacey*15)
    c.drawString(marginleft + 280, 570 - spacey*17, "Total")
    c.drawString(marginleft + 430, 570 - spacey*17, f"${price}.00")
    c.setFillColorRGB(207/255, 226/255, 243/255)
    c.rect(marginleft + 275, 566 - spacey*19, 200, 16, stroke=0, fill=1)
    c.setFillColor("black")
    c.drawString(marginleft + 280, 570 - spacey*19, "Amount Due")
    c.drawString(marginleft + 430, 570 - spacey*19, f"${price}.00")

    c.save()

root.mainloop()
