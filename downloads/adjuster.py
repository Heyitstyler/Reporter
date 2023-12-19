import os
import glob
import xlwings as xw

# Directory where your Excel files are located
directory = os.getcwd()

# Search for Excel files starting with "VarianceReport"
matching_files = glob.glob(os.path.join(directory, 'VarianceReport*.xls'))

if matching_files:
    # Iterate through matching files
    for excel_file_path in matching_files:
        # Open the Excel file without displaying the Excel application window
        app = xw.App(visible=False)
        workbook = app.books.open(excel_file_path)

        # Get the folder where the Excel file is located
        excel_folder = os.path.dirname(excel_file_path)

        # Specify the VBA script file name (assuming it has a .xlsm extension)
        vba_script_file = os.path.join(excel_folder, 'macroBook.xlsm')

        # Run the VBA macro from the specified script file
        workbook.api.Application.Run("'" + vba_script_file + "'!varianceFix")

        # Save changes and close the workbook
        workbook.save()
        workbook.close()

        # Close the Excel application
        app.quit()
else:
    print("No Excel files starting with 'VarianceReport' found in the specified directory.")