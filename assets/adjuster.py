from selector import *
from directory import *
import pandas as pd
import os
import glob
import xlwings as xw

# Search for Excel files starting with "VarianceReport"
def adjust():
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



def namer():
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