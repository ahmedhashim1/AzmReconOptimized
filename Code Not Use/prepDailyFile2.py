import pandas as pd
import xlwings as xw
from datetime import date
import time
import os
import numpy as np


def modify_excel_file_final(daily_file_path, master_file_path):
    """
    Modifies an Excel file by rearranging columns and adding data from a master file
    while preserving original formatting.
    """
    app = None
    try:
        # --- PHASE 1: Use xlwings to perform the column movements first ---
        print("PHASE 1: Starting column movements.")

        # Connect to the existing, visible Excel application instance
        app = xw.apps.active

        # Open the daily workbook
        daily_wb = app.books.open(daily_file_path)
        daily_sheet = daily_wb.sheets[0]

        last_row = daily_sheet.range('A1').end('down').row

        # Corrected error: Check for None or empty sheets
        if last_row is None or last_row < 2:
            print(
                "Error: The daily file is empty or contains only headers. Please ensure there is at least one row of data.")
            daily_wb.close()
            return

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
        print("DEBUG: Fetching 'اسم المفوتر' from Column C.")
        اسم_المفوتر_col = daily_sheet.range('C2:C' + str(last_row)).options(ndim=1).value
        اسم_المفوتر_df = pd.DataFrame(اسم_المفوتر_col, columns=['اسم المفوتر'])
        print("DEBUG: Created 'اسم المفوتر' DataFrame.")

        # Load only the specified columns from the master Excel file
        print("DEBUG: Loading data from master file.")
        master_df = pd.read_excel(master_file_path, usecols=['Arabic', 'Name', 'Index', 'Type', 'Transf Type'])
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
        pass


def add_helper_and_split_files(daily_file_path, master_file_path):
    """
    Adds a 'Helper' sheet to the daily file and splits the data into separate
    Excel files for each unique customer.
    """
    app = None
    try:
        # Connect to the existing, visible Excel application instance
        app = xw.apps.active

        # Read the now-fully-modified sheet into a DataFrame for splitting
        full_data = pd.read_excel(daily_file_path, header=None)

        # Get the actual headers from the first row and set them
        headers = full_data.iloc[0].tolist()
        modified_df = full_data[1:].copy()
        modified_df.columns = headers

        # Clean the 'Cust' column by converting any None/NaN to a string
        modified_df['Cust'] = modified_df['Cust'].astype(str)

        # --- DEBUG OUTPUT ---
        #print("DEBUG: First 5 rows of the DataFrame being used for splitting:")
        #print(modified_df.head().to_string())
        # --- END DEBUG OUTPUT ---

        # --- PART 1: Split the data into separate files for each customer ---
        print("\nPART 1: Splitting data by customer.")

        unique_customers = modified_df['Cust'].unique()
        print(f"Found {len(unique_customers)} unique customers. Creating a file for each.")

        # Get the directory of the daily file
        file_dir = os.path.dirname(daily_file_path)

        for customer in unique_customers:
            print(f"  -> Processing data for customer: {customer}")
            customer_df = modified_df[modified_df['Cust'] == customer].copy()
            new_file_name = f"{customer}.xlsx"
            new_file_path = os.path.join(file_dir, new_file_name)
            customer_df.to_excel(new_file_path, index=False)
            print(f"  -> Successfully created {new_file_name}")

        print("\nPART 1 completed successfully.")

        # --- PART 2: Add the 'Helper' sheet with a unique summary to the original file ---
        print("\nPART 2: Adding 'Helper' sheet to the original daily file.")

        # Create a new DataFrame for the Helper sheet based on the modified data
        master_df = pd.read_excel(master_file_path, usecols=['Arabic', 'Name', 'Index', 'Type', 'Transf Type'])
        master_df.rename(columns={'Arabic': 'اسم المفوتر', 'Name': 'Cust', 'Index': 'Index', 'Type': 'TransType',
                                  'Transf Type': 'BillerType'}, inplace=True)

        # Get unique records from the modified daily file and merge with master data for TransType and BillerType
        helper_df = pd.merge(modified_df.drop_duplicates(subset=['اسم المفوتر']), master_df, on=['اسم المفوتر'],
                             how='left', suffixes=('', '_master'))

        # Select and rename columns for the Helper sheet
        final_helper_df = pd.DataFrame()
        final_helper_df['CustomerName'] = helper_df['Cust']
        final_helper_df['Index'] = helper_df['Index']
        final_helper_df['ArabicName'] = helper_df['اسم المفوتر']
        final_helper_df['HyperLink'] = ''
        final_helper_df['TransType'] = helper_df['TransType']
        final_helper_df['BillerType'] = helper_df['BillerType']

        # Sort the data by the Index column before writing
        final_helper_df.sort_values(by='Index', inplace=True)

        # Re-open the daily workbook to work with the Helper sheet
        daily_wb = app.books.open(daily_file_path)

        # Check if a 'Helper' sheet already exists and add or clear it
        helper_sheet = None
        for sheet in daily_wb.sheets:
            if sheet.name == 'Helper':
                helper_sheet = sheet
                break

        if helper_sheet is None:
            helper_sheet = daily_wb.sheets.add(name='Helper')
        else:
            helper_sheet.clear()

        # Write the Helper DataFrame to the new sheet without the index
        helper_sheet.range('A1').options(index=False).value = final_helper_df

        daily_wb.save()

        print("PART 2 completed successfully. 'Helper' sheet added to original file.")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        pass


# This is how you provide the parameters to the script
if __name__ == "__main__":
    daily_file = r"D:\Freelance\Azm\2025\Sep\21\Test\DailyFile_1.xlsx"  # Replace with the actual file name parameter
    master_file = r"D:\Freelance\Azm\2025\CustomerNamesLookUp.xlsx"  # Make sure this file exists in the same directory or provide the full path

    # Run the initial file modification process
    modify_excel_file_final(daily_file, master_file)

    # Run the second part of the process
    add_helper_and_split_files(daily_file, master_file)