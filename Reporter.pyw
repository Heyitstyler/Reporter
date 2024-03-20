import os
import csv
import sys
import time
import threading
import glob
import queue
import datetime
import signal
import pandas as pd
import requests
import xlwings as xw
import subprocess
from shutil import copyfile
from collections import deque
from tkinter import *
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from PIL import ImageTk, Image
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

#Version
version = "2.4"
hist_Track = 0
barQueue = []
update_bat_cont = r"""set arg1=%1
timeout 3
copy "Reporter.exe" "%~1"
copy "Reporter.Console.exe" "%~1"
pause
"""


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
    speed_mbps = file_size_mb / download_time
    os.remove("5MB.zip")
    return speed_mbps * 8

# Update DB
def initialDB():
    csvURL = "https://raw.githubusercontent.com/Heyitstyler/Reporter/main/assets/bardb.csv"
    try:
        os.chdir(dir_Assets)
        requests.get(csvURL, timeout=5)
        download_file(csvURL)
    except:
        print("Error downloading Bar Database")

def initialMacro():
    MacroURL = "https://github.com/Heyitstyler/Reporter/raw/main/assets/macroBook.xlsm"
    try:
        os.chdir(dir_Assets)
        requests.get(MacroURL, timeout=5)
        download_file(MacroURL)
    except:
        print("Error downloading Macro Book")

def updateDB():
    global hist_Track
    csvURL = "https://raw.githubusercontent.com/Heyitstyler/Reporter/main/assets/bardb.csv"
    try:
        os.chdir(dir_DB)
        requests.get(csvURL, timeout=5)
        try:
            os.remove("bardb.backup.csv")
        except:
            print("No backup bardb")
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
    except Exception as e:
        print(f"Error downloading new bardb.csv {e}")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Error downloading new bardb.csv").pack()
        hist_Track = hist_Track + 1


def updateRep():
                # Specify the path and name of the batch file you want to create


    global hist_Track
    repURL = "https://raw.githubusercontent.com/Heyitstyler/Reporter/main/Reporter.pyw"
    exeURL = "https://github.com/Heyitstyler/Reporter/releases/latest/download/Reporter.exe"
    exeConURL = "https://github.com/Heyitstyler/Reporter/releases/latest/download/Reporter.Console.exe"
    if installType == "SOURCE":
        try:
            os.chdir(dir_Root)
            requests.get(repURL, timeout=5)
            try:
                os.remove("Reporter.backup.pyw")
            except:
                print("No backup Reporter")
            os.rename("Reporter.pyw", "Reporter.backup.pyw")
            download_file(repURL)
            print("Downloaded New Reporter")
            if hist_Track >= 11:
                hist_Frame.forget()
                hist_Track = 0
                history()
                root.update()
            Label(hist_Frame, text="Downloaded new Reporter").pack()
            hist_Track = hist_Track + 1
        except Exception as e:
            print(f"Error downloading new Reporter {e}")
            if hist_Track >= 11:
                hist_Frame.forget()
                hist_Track = 0
                history()
                root.update()
            Label(hist_Frame, text="Error downloading new Reporter").pack()
            hist_Track = hist_Track + 1

    elif installType == "EXE":
        try:
            batch_file_path = os.path.join(dir_Update, "update.bat")

            # Use 'with' statement to open a file and ensure proper closure
            with open(batch_file_path, 'w') as batch_file:
                # Write the content to the batch file
                batch_file.write(update_bat_cont)

            print(f"Batch file created at {batch_file_path}")
            os.chdir(dir_Update)
            requests.get(exeURL, timeout=5)
            try:
                download_file(exeURL)
                download_file(exeConURL)
                print("Downloaded New Reporter")
                if os.path.exists(update_Bat):
                    subprocess.Popen([update_Bat, f"{dir_Root}"], stdout=None, stdin=None, stderr=None, creationflags=subprocess.CREATE_NEW_CONSOLE)
                    sys.exit()
                else:
                    # Specify the path and name of the batch file you want to create
                    batch_file_path = os.path.join(dir_Update, "update.bat")

                # Use 'with' statement to open a file and ensure proper closure
                with open(batch_file_path, 'w') as batch_file:
                    # Write the content to the batch file
                    batch_file.write(update_bat_cont)

                print(f"Batch file created at {batch_file_path}")
            except Exception as e:
                print ({e})
                return
            
            if hist_Track >= 11:
                hist_Frame.forget()
                hist_Track = 0
                history()
                root.update()
            Label(hist_Frame, text="Downloaded new Reporter").pack()
            hist_Track = hist_Track + 1
        except Exception as e:
            print(f"Error downloading new Reporter {e}")
            if hist_Track >= 11:
                hist_Frame.forget()
                hist_Track = 0
                history()
                root.update()
            Label(hist_Frame, text="Error downloading new Reporter").pack()
            hist_Track = hist_Track + 1


def updateMacro():
    global hist_Track
    MacroURL = "https://github.com/Heyitstyler/Reporter/raw/main/assets/macroBook.xlsm"
    try:
        os.chdir(dir_Assets)
        requests.get(MacroURL, timeout=5)
        try:
            os.remove("macroBook.backup.xlsm")
        except:
            print("No backup Macro Book")
        os.rename("macroBook.xlsm", "macroBook.backup.xlsm")
        download_file(MacroURL)
        print("Downloaded New Macro Book")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Downloaded new Macro Book").pack()
        hist_Track = hist_Track + 1
    except Exception as e:
        print(f"Error downloading new Macro Book {e}")
        if hist_Track >= 11:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()
        Label(hist_Frame, text="Error downloading new Macro Book").pack()
        hist_Track = hist_Track + 1

#Directory
appdata = os.getenv('APPDATA')
appdataPATH = os.path.join(appdata, "Reporter")
if not os.path.exists("Reporter Downloads"):
    os.makedirs("Reporter Downloads")

if os.path.exists("Scripts") and os.path.isfile("pyvenv.cfg"):
    installType = "SOURCE"
else:
    installType = "EXE"

if installType == "SOURCE":
    if os.path.exists("assets"):
        dir_Root = os.getcwd()
        dir_Assets = os.path.join(dir_Root, "assets")
        dir_Downloads = os.path.join(dir_Root, "Reporter Downloads")
        dir_DB = dir_Assets
        bardbloc = os.path.join(dir_Assets, "bardb.csv")

    else:
        os.makedirs("assets")
        dir_Root = os.getcwd()
        dir_Assets = os.path.join(dir_Root, "assets")
        dir_Downloads = os.path.join(dir_Root, "Reporter Downloads")
        dir_DB = dir_Assets
        bardbloc = os.path.join(dir_Assets, "bardb.csv")
        initialDB()
        initialMacro()

elif installType == "EXE":
    if os.path.exists(appdataPATH + "\\assets") and os.path.exists(appdataPATH + "\\update"):
        dir_Root = os.getcwd()
        dir_Assets = os.path.join(appdataPATH, "assets")
        dir_Downloads = os.path.join(dir_Root, "Reporter Downloads")
        dir_DB = dir_Assets
        bardbloc = os.path.join(dir_Assets, "bardb.csv")
        dir_Update = os.path.join(appdataPATH, "update")
        update_Bat = os.path.join(dir_Update, "update.bat")

    else:
        os.makedirs(appdataPATH)
        os.makedirs(appdataPATH + "\\assets")
        os.makedirs(appdataPATH + "\\update")
        dir_Root = os.getcwd()
        dir_Assets = os.path.join(appdataPATH, "assets")
        dir_Downloads = os.path.join(dir_Root, "Reporter Downloads")
        dir_DB = dir_Assets
        bardbloc = os.path.join(dir_Assets, "bardb.csv")
        dir_Update = os.path.join(appdataPATH, "update")
        update_Bat = os.path.join(dir_Update, "update.bat")
        initialDB()
        initialMacro()
    

# Initial internet check
try:
    checkint = requests.get("https://www.google.com", timeout=3)
except:
    print("Can't contact Google. Are you connected to the internet?")


# Determine internet speed
    
# URL of the 5MB file
file_url = 'http://ipv4.download.thinkbroadband.com/1MB.zip'
file_size_mb = 1


def add_to_list(group_name, data_tuple):
    # Check if the list already exists in globals; if not, initialize it
    if group_name not in globals():
        globals()[group_name] = []
    globals()[group_name].append(data_tuple)


# Set to keep track of unique group names
unique_groups = set()
company_names = set()

# Open the CSV file
with open(bardbloc, mode='r') as csv_file:
    # Create a CSV reader
    csv_reader = csv.DictReader(csv_file)
    
    # Iterate through each row in the CSV
    for row in csv_reader:
        if row['group'] == 'Canceled':
            continue
        # Extract and process the group name
        group_name = row['group'].upper() + 'BARS'
        company_names.add(group_name)
        
        # Add the original group name to the set of unique groups
        original_group_name = row['group']
        unique_groups.add(original_group_name)
        
        # Create a tuple from the appropriate column and the 'user' column
        data_tuple = (row['proper'], row['user'])
        
        # Add the tuple to the correct list
        add_to_list(group_name, data_tuple)

# Now, create the COMPANIES list from the unique groups set
COMPANIES = [(group, group) for group in unique_groups]

# If you want to sort the COMPANIES list alphabetically by the first element of the tuples
COMPANIES.sort(key=lambda x: x[0])
        


# Root
root = Tk()
root.geometry("800x530")
root.title(f"Reporter {version}")
root.resizable(False, False)

# Top Labels
comp_Label = Label(root, text="Companies", background="light blue", width=10, pady=10, font=('Arial', 24))
comp_Label.grid(row=0, column=0, pady=8)

bars_Label = Label(root, text="Bars", background="light blue", width=10, pady=10, font=('Arial', 24))
bars_Label.grid(row=0, column=1, pady=8)

report_Label = Label(root, text="Reporter", background="light blue", width=10, pady=10, font=('Arial', 24))
report_Label.grid(row=0, column=2, pady=8)


# Company Frame
comp_Frame = LabelFrame(root)
comp_Frame.pack_propagate(False)
comp_Frame.config(height=440, width=200)
comp_Frame.grid(row=1, column=0, padx=30, pady=5)


# Bars Frame
bars_Frame = LabelFrame(root)
bars_Frame.pack_propagate(False)
bars_Frame.config(height=440, width=225, pady=5)
bars_Frame.grid(row=1, column=1, padx=15, pady=5)


# Report Frame
report_Frame = LabelFrame(root)
report_Frame.grid_propagate(False)
report_Frame.config(height=440, width=226)
report_Frame.grid(row=1, column=2, padx=30, pady=5)

rpt_Queue = IntVar()
rpt_Queue.set(0)
queue_Select = rpt_Queue.get()

# Sub-Report Frames
report_button = Button(report_Frame, text="Run Report", bg="Red", activebackground="yellow", font=("Arial", 16), pady=8, state=DISABLED)
add_Queue_Button = Button(report_Frame, text="Add", bg="#c4abd2", activebackground="#94819f", font=("Arial", 14), pady=8, width=5, command=lambda: add_Queue(user))
run_Queue_Button = Button(report_Frame, text="Run Queue", bg="lime", activebackground="yellow", font=("Arial", 14), pady=8, command=lambda: run_Queue())


folder_Button = Button(report_Frame, text="Open Downloads Folder", bg="light grey", font=("Arial", 12), pady=3, anchor=N, command=lambda:os.startfile(dir_Downloads))
folder_Button.grid(row=4, column=0, columnspan=3)

def change_Modes(value):
    if value == 0:
        add_Queue_Button.grid_remove()
        run_Queue_Button.grid_remove()
        report_Frame.update()
        root.update()
        report_button.grid(row=0, column=0, pady=8, columnspan=3)
        report_Frame.update()
        root.update()
    elif value == 1:
        report_button.grid_remove()
        report_Frame.update()
        root.update()
        add_Queue_Button.grid(row=0, column=0, pady=10, padx=10, columnspan=2, sticky=W)
        run_Queue_Button.grid(row=0, column=1, pady=10, padx=10, columnspan=2, sticky=E)
        report_Frame.update()
        root.update()
    elif value == 2:
        report_button.grid(row=0, column=0, pady=8, columnspan=3)
        report_Frame.update()

change_Modes(2)

# Download type selector
rpt = IntVar()
rpt.set(1)

reportTypeFull = Radiobutton(report_Frame, text="Full", variable=rpt, value=1)
reportTypeReport = Radiobutton(report_Frame, text="Just Report", variable=rpt, value=2)
reportTypeInvoice = Radiobutton(report_Frame, text="Just Invoice", variable=rpt, value=3)

reportTypeFull.grid(row=1, column=0)
reportTypeReport.grid(row=1, column=1)
reportTypeInvoice.grid(row=1, column=2)

# Status
status = Label(report_Frame, text="Status: Ready")
status.grid(row=2, column=0, columnspan=3)

# Top menu
topMenu = Menu(root)
root.config(menu=topMenu)

update_menu = Menu(topMenu)
topMenu.add_cascade(label="Update", menu=update_menu)
update_menu.add_command(label="Update Bar Database", command = lambda:updateDB())
update_menu.add_command(label="Update Reporter", command = lambda:updateRep())
update_menu.add_command(label="Update Macro Book", command = lambda:updateMacro())

genOrder = IntVar()
genOrder.set(0)
order_Option = genOrder.get()

optional_Menu = Menu(topMenu)
topMenu.add_cascade(label='Optional', menu=optional_Menu)
optional_Menu.add_checkbutton(label="Generate Order Report", variable=genOrder, onvalue=1, offvalue=0)

edit_Menu = Menu(topMenu)
topMenu.add_cascade(label="Edit", menu=edit_Menu)
edit_Menu.add_command(label="Edit Bar Database", command = lambda:os.startfile(bardbloc))

mode_Menu = Menu(topMenu)
topMenu.add_cascade(label="Mode", menu=mode_Menu)
mode_Menu.add_command(label="Single Bar Mode", command = lambda:change_Modes(0))
mode_Menu.add_command(label="Queue Mode", command = lambda:change_Modes(1))


# History Frame
def history():
    global hist_Frame
    hist_Frame = LabelFrame(report_Frame)
    hist_Frame.config(height=20, width=210, text="History", labelanchor=N, font=('Arial', 12))
    hist_Frame.grid(row=3, column=0, padx=5, ipady=130, columnspan=3)
    hist_Frame.grid_propagate(False)
    hist_Frame.pack_propagate(False)
history()


for text, mode in COMPANIES:
    compbutton = Button(comp_Frame, text=text, bg="light grey", font=('Arial', 16))
    compbutton.config(command=lambda button=compbutton, mode=mode:[on_company_click(button, mode)])
    compbutton.pack(pady=15)


def on_company_click(button, mode):
    # Define what happens when a button is clicked
    # For demonstration, you can print the mode or handle it as needed
    for widget in comp_Frame.winfo_children():
        widget.configure(bg="light grey")
    for widget in bars_Frame.winfo_children():
        widget.destroy()
    button.configure(bg="dark grey")
    bars_for_group(f'{mode}')
    print(f"Button for {mode} clicked")


def on_bar_click(button, mode):
    global hist_Frame, dir_BarFolder, proper, userRow, passwd, workingDir, barSelect, street, city, inv, price, user, bargroup
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

    user = userRow["user"].iloc[0]
    passwd = userRow["pass"].iloc[0]
    proper = userRow["proper"].iloc[0]
    street = userRow["street"].iloc[0]
    city = userRow["city"].iloc[0]
    inv = userRow["invoicename"].iloc[0]
    price = userRow["price"].iloc[0]
    bargroup = userRow["group"].iloc[0]
        
    
    for widget in bars_Frame.winfo_children():
        widget.configure(bg="light grey")
    button.configure(bg="dark grey")
    report_button.config(bg="lime", state=NORMAL, command=lambda button=button, mode=mode: run_report(mode))
    report_Frame.update()


def run_report(mode):
    global dir_BarFolder, workingDir, hist_Track, order_Option, proper, audit_date
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
    order_Option = genOrder.get()

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

        summary_File = glob.glob(os.path.join(dir_BarFolder, 'Summary_report*.xlsx'))[0]
        summary_noExtention = summary_File.rsplit('.', 1)[0]
        parts = summary_noExtention.split("_")
        date_str = "-".join(parts[-3:]).strip()  # Join the last three parts to get the date string in 'MM-DD-YYYY' format

        # Parse the date string into a datetime object
        date_obj = datetime.datetime.strptime(date_str, "%m-%d-%Y")

        # Format the datetime object as desired (e.g., 'YYYY-MM-DD')
        audit_date = date_obj.strftime("%Y-%m-%d").strip()
        print(f"Audit was performed on {audit_date}")

        t4 = threading.Thread(target=adjust, kwargs={'mode': mode})
        t4.start()
        t4.join()
        if order_Option == 1:
            inv_hist = Label(hist_Frame, text=f"{proper} Order Report")
            inv_hist.pack()
            hist_Track = hist_Track + 1


        t5 = threading.Thread(target=namer, kwargs={'mode': mode})
        t5.start()
        t5.join()
        

        t6 = threading.Thread(target=Invoice)
        t6.start()
        t6.join()
        inv_hist = Label(hist_Frame, text=f"Generated {proper} Invoice")
        inv_hist.pack()

        emailGenReport()
        emailGenInvoice()

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

        summary_File = glob.glob(os.path.join(dir_BarFolder, 'Summary_report*.xlsx'))[0]
        summary_noExtention = summary_File.rsplit('.', 1)[0]
        parts = summary_noExtention.split("_")
        date_str = "-".join(parts[-3:]).strip()  # Join the last three parts to get the date string in 'MM-DD-YYYY' format

        # Parse the date string into a datetime object
        date_obj = datetime.datetime.strptime(date_str, "%m-%d-%Y")

        # Format the datetime object as desired (e.g., 'YYYY-MM-DD')
        audit_date = date_obj.strftime("%Y-%m-%d").strip()
        print(f"Audit was performed on {audit_date}")

        t4 = threading.Thread(target=adjust, kwargs={'mode': mode})
        t4.start()
        t4.join()
        if order_Option == 1:
            inv_hist = Label(hist_Frame, text=f"{proper} Order Report")
            inv_hist.pack()
            hist_Track = hist_Track + 1


        t5 = threading.Thread(target=namer, kwargs={'mode': mode})
        t5.start()
        t5.join()

        emailGenReport()
        emailGenInvoice()

        time2 = time.perf_counter()

        print(f"Downloaded reports in {time2 - time1:0.2f} seconds.")
        print(formatted_date)
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

        emailGenInvoice()

        time2 = time.perf_counter()

        print(f"Generated Invoice in {time2 - time1:0.2f} seconds.")

        report_button.config(bg="lime", state=NORMAL)
        status.config(text="Status: Ready")
        
    os.chdir(dir_Root)


def add_Queue(user):
    global barQueue, hist_Track
    # if 'barQueue' not in globals():  # Initialize barQueue if it's not already
    #     barQueue = deque()
    barQueue.append(user)
    Queue_hist = Label(hist_Frame, text=f"Queued {user}")

    if hist_Track >= 12:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()

    hist_Track = hist_Track + 1
    
    Queue_hist.pack()
    print(list(barQueue))
    print(len(barQueue))

def run_Queue():
    global barQueue, barSelect, hist_Track, hist_Frame, dir_BarFolder, proper, userRow, passwd, workingDir, street, city, inv, price, user, bargroup
    run_Queue_Button.config(state=DISABLED, bg="red")
    bars = pd.read_csv(dir_DB + "\\bardb.csv")
    qTime1 = time.perf_counter()
    print(len(barQueue))
    if len(barQueue) == 0:
        print("Queue is Empty")
    else:
        
        while len(barQueue) > 0:
            bars = pd.read_csv(dir_DB + "\\bardb.csv")
            barSelect = barQueue.pop()
            mode = barSelect
            userRow = bars[bars["user"] == barSelect]
            user = userRow["user"].iloc[0]
            passwd = userRow["pass"].iloc[0]
            proper = userRow["proper"].iloc[0]
            street = userRow["street"].iloc[0]
            city = userRow["city"].iloc[0]
            inv = userRow["invoicename"].iloc[0]
            price = userRow["price"].iloc[0]
            bargroup = userRow["group"].iloc[0]
            barQueue.pop
            run_report(mode)
            print(barQueue)
            Queue_hist = Label(hist_Frame, text=f"Queued {user}")
            
            
            
    qTime2 = time.perf_counter()
    qTime = qTime2 - qTime1
    print(f"Queue Completed in {qTime:0.2f}")
    if hist_Track >= 12:
            hist_Frame.forget()
            hist_Track = 0
            history()
            root.update()

    hist_Track = hist_Track + 1
    queue_Done = Label(hist_Frame, text=f"Queued Completed in {qTime:0.2f}!")
    queue_Done.pack()
    run_Queue_Button.config(bg="lime", state=NORMAL)


def bars_for_group(group_name):
    # Construct the name of the list variable for the specified group
    bars_list_name = group_name.upper() + 'BARS'
    
    # Access the list variable dynamically using globals()
    if bars_list_name in globals():
        bars_list = globals()[bars_list_name]
        for text, mode in bars_list:
            button = Button(bars_Frame, text=text, bg="light grey", font=('Arial', 16))
            button.config(command=lambda button=button, mode=mode: on_bar_click(button, mode))
            
            # Adjust padding based on the group name if needed
            pady_value = 2  # Default padding, can be adjusted as needed
            
            button.pack(pady=pady_value)

# Selenium Instances
def dlSummary(mode):
    global sum_e, audit_date
    found_Sum = "False"
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
        sumWait = WebDriverWait(summary_driver, 90)

        summary_driver.get("https://www.barkeepapp.com/BarkeepOnline/inventories.php")

        login_Loaded = sumWait.until(EC.presence_of_element_located((By.NAME, 'session_username')))
        username_field = summary_driver.find_element(By.NAME, 'session_username')
        username_field.send_keys(barSelect)
        password_field = summary_driver.find_element(By.NAME, 'session_password')
        password_field.send_keys(passwd)
        login_button = summary_driver.find_element(By.NAME, 'login')
        login_button.click()

        inventories_Loaded = sumWait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')))
        full_summary = summary_driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div[3]/div[2]/table/tbody/tr[1]/td[1]/a[1]')
        full_summary.click()

        full_Loaded = sumWait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="dropdownMenu1"]')))
        dropdown_summary = summary_driver.find_element(By.XPATH, '//*[@id="dropdownMenu1"]')
        dropdown_summary.click()

        full_Loaded = sumWait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')))
        download_summary = summary_driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/ul/li[1]/a')
        download_summary.click()

        while found_Sum =="False":
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
                    os.chdir(dir_Root)
                    found_Sum = "True"
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
    found_Use = "False"
    try:
        keyword = 'Usage'
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.download.dir", dir_BarFolder)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
        options.add_argument("--headless")

        usage_driver = webdriver.Firefox(options=options)

        usage_driver.get("https://www.barkeepapp.com/BarkeepOnline/usageReport.php")
        waitUse = WebDriverWait(usage_driver, 90)

        login_Loaded = waitUse.until(EC.presence_of_element_located((By.NAME, 'session_username')))
        username_field = usage_driver.find_element(By.NAME, 'session_username')
        username_field.send_keys(barSelect)
        password_field = usage_driver.find_element(By.NAME, 'session_password')
        password_field.send_keys(passwd)
        login_button = usage_driver.find_element(By.NAME, 'login')
        login_button.click()


        usage_Loaded = waitUse.until(EC.presence_of_element_located((By.ID, "startInventoryId")))
        use_start_date_drop = usage_driver.find_element(By.ID, "startInventoryId")
        use_start_date_drop.click()

        use_start_date_select = usage_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
        use_start_date_select.click()
        use_end_date_drop = usage_driver.find_element(By.ID, "endInventoryId")
        use_end_date_drop.click()

        use_end_date_select = usage_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
        use_end_date_select.click()


        run_js = 'runReport()'
        usage_driver.execute_script(run_js)
        report_Loaded = waitUse.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[4]/div/div[2]/div[2]/table/tbody/tr[2]")))
        download_js = 'downloadReport()'
        usage_driver.execute_script(download_js)

        while found_Use == "False":
        # List all files in the specified directory
            files = os.listdir(workingDir)

        # Check if any file contains the keyword
            for file in files:
                if file.startswith(keyword) and not file.endswith(".part"):
                    print(f"Found file: {file}")
                    use_e = (f"{proper} Usage Report")
                    time.sleep(1.5)
                    usage_driver.close()
                    time.sleep(0.5)
                    usage_driver.quit()
                    os.chdir(dir_Root)
                    found_Use = "True"
                    
    except Exception as e:
        use_e = ("Error Collecting Usage Report")
        error_e = (f"Error Collecting Usage Report {e}")
        print(error_e)
        usage_driver.close()
        time.sleep(1)
        usage_driver.quit()
        os.chdir(dir_Root)
        log = open("dllog.txt", "a")
        L = [f"Failed Usage Report\n{e}\n"]
        log.writelines(L)
        log.close()


def dlVar(mode):
    global var_e
    found_Var = "False"
    try:
        keyword = 'Variance'
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.download.dir", workingDir)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-gzip")
        options.add_argument("--headless")

        variance_driver = webdriver.Firefox(options=options)

        variance_driver.get("https://www.barkeepapp.com/BarkeepOnline/varianceReport.php")
        waitVar = WebDriverWait(variance_driver, 90)

        login_Loaded = waitVar.until(EC.presence_of_element_located((By.NAME, 'session_username')))
        username_field = variance_driver.find_element(By.NAME, 'session_username')
        username_field.send_keys(barSelect)
        password_field = variance_driver.find_element(By.NAME, 'session_password')
        password_field.send_keys(passwd)
        login_button = variance_driver.find_element(By.NAME, 'login')
        login_button.click()


        variance_Loaded = waitVar.until(EC.presence_of_element_located((By.ID, 'startInventoryId')))
        var_start_date_drop = variance_driver.find_element(By.ID, "startInventoryId")
        var_start_date_drop.click()
        # time.sleep(loadTime/2)
        var_start_date_select = variance_driver.find_element(By.XPATH, '//select[@id="startInventoryId"]/option[3]')
        var_start_date_select.click()
        var_end_date_drop = variance_driver.find_element(By.ID, "endInventoryId")
        var_end_date_drop.click()
        # time.sleep(loadTime/2)
        var_end_date_select = variance_driver.find_element(By.XPATH, '//select[@id="endInventoryId"]/option[2]')
        var_end_date_select.click()
        # time.sleep(loadTime/2)

        run_js = 'runReport()'
        variance_driver.execute_script(run_js)
        # time.sleep(loadTime*4)
        report_Loaded = waitVar.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[4]/div/div[1]/div/div[2]/div[2]/table/tbody/tr[2]")))
        download_js = 'downloadReport()'
        variance_driver.execute_script(download_js)

        while found_Var == "False":
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
                    os.chdir(dir_Root)
                    found_Var = "True"
                
    except Exception as e:
        var_e = ("Error Collecting Variance Report")
        error_e = (f"Error Collecting Variance Report {e}")
        print(error_e)
        variance_driver.close()
        time.sleep(1)
        variance_driver.quit()
        os.chdir(dir_Root)
        log = open("dllog.txt", "a")
        L = [f"Failed Variance Report\n"]
        log.writelines(L)
        log.close()

# Report Adjustments
def adjust(mode):
    today = datetime.datetime.now()
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
                os.chdir(dir_Root)
        else:
            print("No Excel files starting with 'VarianceReport' found in the specified directory.")
    except Exception as e:
        print(str(e))
        input ("Press any button to continue")
    

    
    print(order_Option)

    if order_Option == 1:
        rename = f"Order_Report_{today:%Y}_{today:%m}_{today:%d}.xlsx"
        
        usage_File = glob.glob(os.path.join(dir_BarFolder, 'Usage_Report*.xlsx'))[0]
        copyfile(usage_File, f"{dir_BarFolder}\\{rename}")
        order_File = glob.glob(os.path.join(dir_BarFolder, 'Order_Report*.xlsx'))[0]
        
        try:
            
            if os.path.exists(order_File):
                    # Open the Excel file without displaying the Excel application window
                    app = xw.App(visible=False)
                    workbook = app.books.open(order_File)

                    # Specify the VBA macro name
                    macro_name = 'varianceFix'

                    # Specify the path to the VBA script or Personal Macro Workbook
                    vba_script_path = os.path.join(dir_Assets, 'macroBook.xlsm')

                    # Run the VBA macro from the specified script file
                    workbook.api.Application.Run("'" + vba_script_path + "'!Module2.orderReport")

                    # Save changes and close the workbook
                    workbook.save()
                    workbook.close()

                    # Close the Excel application
                    app.quit()
                    os.chdir(dir_Root)
                    print("Generated Order Report")
            else:
                os.chdir(dir_Root)
                print("No Excel files starting with 'Order_Report' found in the specified directory.")
        except Exception as e:
            print(str(e))
            input ("Press any button to continue")
    else:
        print("No Extras Selected")
        os.chdir(dir_Root)
        return


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
    os.chdir(dir_Root)


def emailGenReport():
    global audit_date
    link_Proper = proper.replace(" ", "%20")
    os.chdir(dir_BarFolder)
    subject = f"GDS%20Consulting's%20Pour%20Cost%20Reports%20for%20{link_Proper}%20on%20{audit_date}"
    body = f"GDS%20Consulting's%20Pour%20Cost%20Reports%20for%20{link_Proper}%20on%20{audit_date}"
    mail_link = f"https://mail.google.com/mail/u/gdsconsultingllc@gmail.com/?fs=1&tf=cm&source=mailto&su={subject}&body={body}"
    
    with open('Report Email.url','w') as f:
        f.write(f"""[InternetShortcut]
    URL={mail_link}
    """)


def emailGenInvoice():
    global audit_date
    link_Proper = proper.replace(" ", "%20")
    os.chdir(dir_BarFolder)
    subject = f"GDS%20Consulting's%20Invoice%20for%20{link_Proper}%20on%20{audit_date}"
    body = f"GDS%20Consulting's%20Pour%20Cost%20Reports%20for%20{link_Proper}%20on%20{audit_date}%0aThanks%20so%20much%21"
    mail_link = f"https://mail.google.com/mail/u/gdsconsultingllc@gmail.com/?fs=1&tf=cm&source=mailto&su={subject}&body={body}"
    
    with open('Invoice Email.url','w') as f:
        f.write(f"""[InternetShortcut]
    URL={mail_link}
    """)

root.mainloop()
