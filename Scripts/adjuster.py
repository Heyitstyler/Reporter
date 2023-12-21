from reporter import dir_BarFolder
from reporter import dir_Scripts
from reporter import dir_Root
import os
import glob
import xlwings as xw
import time

# Directory where your Excel files are located
#directory = dir_BarFolder

# Path to the folder where your VBA script or Personal Macro Workbook is located
#vba_script_folder = dir_Scripts

# Search for Excel files starting with "VarianceReport"
matching_files = glob.glob(os.path.join(dir_BarFolder, 'VarianceReport*.xls'))

if matching_files:
    # Iterate through matching files
    for excel_file_path in matching_files:
        # Open the Excel file without displaying the Excel application window
        app = xw.App(visible=False)
        workbook = app.books.open(excel_file_path)

        # Specify the VBA macro name
        macro_name = 'varianceFix'

        # Specify the path to the VBA script or Personal Macro Workbook
        vba_script_path = os.path.join(dir_Scripts, 'macroBook.xlsm')

        # Run the VBA macro from the specified script file
        workbook.api.Application.Run("'" + vba_script_path + "'!Module1.varianceFix")

        # Save changes and close the workbook
        workbook.save()
        workbook.close()

        # Close the Excel application
        app.quit()
else:
    print("No Excel files starting with 'VarianceReport' found in the specified directory.")

# Restart?
restart = input("Would you like to run another bar? (y/n)")
if restart == "y":
    os.chdir(dir_Root)
    os.system(dir_Root + r"/reporter.bat")
else:
    quit()