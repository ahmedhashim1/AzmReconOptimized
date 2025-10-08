import pandas as pd
import xlwings as xw
from datetime import date
import time
import re


def modify_excel_file_final(daily_file_path, master_file_path):
    """
    Modifies an Excel file by rearranging columns and adding data from a master file
    while preserving original formatting.
    """
    app = None
    try:
        # --- PHASE 1: Use xlwings to perform the column movements first ---
        print("PHASE 1: Starting column movements.")

        # Create a new xlwings app instance and make it visible
        # app = xw.App(visible=True)
        app = xw.apps.active

        # Open the daily workbook
        daily_wb = app.books.open(daily_file_path)
        daily_sheet = daily_wb.sheets[0]

        last_row = daily_sheet.range('A1').end('down').row

        # 1. Immediately move 'رقم العقد' column (B) to Column S
        daily_sheet.range('B:B').api.Cut(Destination=daily_sheet.range('S1').api)

        # 2. Then move 'اسم المفوتر' column (now A) to Column B
        daily_sheet.range('A:A').api.Cut(Destination=daily_sheet.range('B1').api)

        # 3. Insert 1 Column Before Column B
        daily_sheet.range('B1').api.EntireColumn.Insert()

        # 4. Insert Headers 'Cust' and 'Index' into Column A and Column B
        daily_sheet.range('A1').value = 'Cust'
        daily_sheet.range('B1').value = 'Index'

        # Write the 'fdate' data to Column U immediately after column movements
        print("DEBUG: Writing 'fdate' to Column U.")
        daily_sheet.range('U1').value = 'fdate'
        today = date.today()
        daily_sheet.range('U2:U' + str(last_row)).value = today
        print("PHASE 1: Column movements and 'fdate' column completed successfully.")

        # --- PHASE 2: Perform pandas logic to get data from master file ---
        print("PHASE 2: Starting pandas data processing.")

        # Get the 'اسم المفوتر' column from the sheet, which is now in Column C
        # Start from row 2 to exclude the header
        print("DEBUG: Fetching 'اسم المفوتر' from Column C.")
        اسم_المفوتر_col = daily_sheet.range('C2:C' + str(last_row)).options(ndim=1).value
        اسم_المفوتر_df = pd.DataFrame(اسم_المفوتر_col, columns=['اسم المفوتر'])
        print("DEBUG: Created 'اسم المفوتر' DataFrame.")

        # Load only the specified columns from the master Excel file
        print("DEBUG: Loading data from master file.")
        master_df = pd.read_excel(master_file_path, usecols=['Arabic', 'Name', 'Index'])
        print("DEBUG: Master DataFrame created.")

        # Merge the dataframes
        print("DEBUG: Starting data merge.")
        master_df.rename(columns={'Arabic': 'اسم المفوتر', 'Name': 'Cust', 'Index': 'Index'}, inplace=True)
        merged_df = pd.merge(اسم_المفوتر_df, master_df, on='اسم المفوتر', how='left')
        print("DEBUG: Data merge completed. Merged DataFrame shape:", merged_df.shape)

        # Write the new 'Cust' and 'Index' values to the sheet using a bulk write
        print("DEBUG: Starting bulk write for 'Cust' and 'Index' values.")
        daily_sheet.range('A2').value = merged_df[['Cust', 'Index']].values
        print("DEBUG: Bulk write for 'Cust' and 'Index' completed.")

        # Save the changes to the workbook
        daily_wb.save()

        print(f"Successfully updated {daily_file_path}. All changes were made while preserving original formatting.")
        print("The file will remain open. Please close it manually when you are done.")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Do not close the app, as requested
        pass


def run_vba_sub(macro_name):
    """
    Runs a VBA Sub from PERSONAL.XLSB that does not require any arguments.
    """
    try:
        # Get a reference to the active Excel application
        app = xw.apps.active

        # Check if PERSONAL.XLSB is already open.
        personal_wb = None
        for wb in app.books:
            if wb.name.upper() == 'PERSONAL.XLSB':
                personal_wb = wb
                break

        # If PERSONAL.XLSB is not found, explicitly open it
        if personal_wb is None:
            personal_wb = xw.Book('PERSONAL.XLSB')

        # Run the VBA Sub directly on the Excel Application object.
        # Note: We do not try to capture a return value.
        personal_wb.app.api.Run(macro_name)

        print(f"VBA macro '{macro_name}' executed successfully.")
        return True
    except Exception as e:
        print(f"An error occurred while running VBA macro: {e}")
        return False




def debug_vba_functions(module_name):
    """
    Connects to the active Excel application and prints the names of all
    functions and subs in a specified VBA module within PERSONAL.XLSB.
    """
    try:
        # Connect to the active Excel application instance
        app = xw.apps.active

        # Check if PERSONAL.XLSB is already open
        personal_wb = None
        for wb in app.books:
            if wb.name.upper() == 'PERSONAL.XLSB':
                personal_wb = wb
                break

        # If PERSONAL.XLSB is not found, explicitly open it
        if personal_wb is None:
            print("DEBUG: PERSONAL.XLSB not found in open workbooks. Attempting to open it.")
            personal_wb = xw.Book('PERSONAL.XLSB')

        print("DEBUG: PERSONAL.XLSB workbook reference obtained.")

        # Get the specific VBA component (module) by its name
        try:
            vba_component = personal_wb.api.VBProject.VBComponents(module_name)
            print(f"DEBUG: Found module '{module_name}'.")
        except Exception as e:
            print(f"DEBUG: Error accessing module '{module_name}': {e}")
            return

        # Get the code module for the component
        code_module = vba_component.CodeModule

        print("DEBUG: Listing functions and subroutines:")

        # Regular expressions to find Sub and Function declarations
        function_pattern = re.compile(r'^\s*(?:Public|Private)?\s*Function\s+(\w+)\s*\(.*', re.IGNORECASE)
        sub_pattern = re.compile(r'^\s*(?:Public|Private)?\s*Sub\s+(\w+)\s*\(.*', re.IGNORECASE)

        # Loop through each line of the code module
        line_number = code_module.CountOfLines
        line = 1
        found_count = 0
        while line <= line_number:
            line_text = code_module.Lines(line, 1)

            # Search for function and sub definitions
            func_match = function_pattern.match(line_text)
            sub_match = sub_pattern.match(line_text)

            if func_match:
                print(f"    - Function: {func_match.group(1)}")
                found_count += 1
            elif sub_match:
                print(f"    - Sub: {sub_match.group(1)}")
                found_count += 1

            line += 1

        if found_count == 0:
            print(f"    No functions or subroutines found in module '{module_name}'.")

        print("\nScript completed successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")


# This is how you provide the parameters to the script
if __name__ == "__main__":
    daily_file = r"D:\Freelance\Azm\2025\Sep\21\Test\DailyFile_1.xlsx"  # Replace with the actual file name parameter
    master_file = r"D:\Freelance\Azm\2025\CustomerNamesLookUp.xlsx"  # Make sure this file exists in the same directory or provide the full path

    modify_excel_file_final(daily_file, master_file)
    # debug_vba_functions('SplitSheet2')
    # vba_function_name = 'SplitSheet2.SplitDataset()'
    # run_vba_sub(vba_function_name)