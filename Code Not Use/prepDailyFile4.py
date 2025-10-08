import pandas as pd
import xlwings as xw
from datetime import date
import time
import os
import numpy as np
import mysql.connector
from mysql.connector import Error
import concurrent.futures
from threading import Lock

mysql_config = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "azm"
}


def get_mysql_connection_pool():
    """Create a connection pool for better performance"""
    try:
        # Create connection pool
        from mysql.connector import pooling
        pool_config = mysql_config.copy()
        pool_config.update({
            'pool_name': 'azm_pool',
            'pool_size': 10,
            'pool_reset_session': True,
            'autocommit': True
        })

        pool = pooling.MySQLConnectionPool(**pool_config)
        print(f"Created connection pool with {pool.pool_size} connections")
        return pool
    except mysql.connector.Error as err:
        print(f"Error creating connection pool: {err}")
        return None



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

        # Use pandas only to get unique customers efficiently
        temp_df = pd.read_excel(daily_file_path)
        unique_customers = temp_df['Cust'].unique()

        # --- PART 1: Split the data into separate files for each customer ---
        print("\nPART 1: Splitting data by customer using filters for speed and formatting.")

        print(f"Found {len(unique_customers)} unique customers. Creating a file for each.")

        # Open the daily workbook to work with filters
        daily_wb = app.books.open(daily_file_path)
        daily_sheet = daily_wb.sheets[0]

        file_dir = os.path.dirname(daily_file_path)

        # Check and remove any existing filter
        if daily_sheet.api.AutoFilterMode:
            daily_sheet.api.AutoFilterMode = False

        # Get the full data range for filtering
        full_range = daily_sheet.used_range

        successful_splits = 0
        for i, customer in enumerate(unique_customers, 1):
            print(f"[{i}/{len(unique_customers)}] Processing: {customer}")

            try:
                # Apply filter to the full range
                full_range.api.AutoFilter(
                    Field=1,
                    Criteria1=str(customer)
                )

                # Copy visible cells only (headers and filtered data)
                visible_cells = full_range.api.SpecialCells(12)  # xlCellTypeVisible
                visible_cells.Copy()

                # Create new workbook and paste
                new_wb = app.books.add()
                new_ws = new_wb.sheets[0]

                # Paste with formatting
                new_ws.range('A1').api.PasteSpecial(-4104)  # xlPasteAll

                # Clear clipboard
                app.api.CutCopyMode = False

                # Auto-adjust columns
                new_ws.autofit()

                # Save file
                safe_filename = "".join(c for c in str(customer) if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
                if not safe_filename:
                    safe_filename = f"Customer_{i}"

                output_file = os.path.join(file_dir, f"{safe_filename}.xlsx")
                new_wb.save(output_file)
                new_wb.close()

                successful_splits += 1
                print(f"  ✓ Saved: {safe_filename}.xlsx")

            except Exception as e:
                print(f"  ✗ Error processing {customer}: {str(e)}")
                # Continue to the next customer in case of an error
                continue

        # Clean up
        if daily_sheet.api.AutoFilterMode:
            daily_sheet.api.AutoFilterMode = False
        daily_wb.save()
        daily_wb.close()

        print(f"\nCompleted! Successfully split {successful_splits}/{len(unique_customers)} customers")
        print(f"Files saved in: {os.path.abspath(file_dir)}")

        # --- PART 2: Add the 'Helper' sheet with a unique summary to the original file ---
        print("\nPART 2: Adding 'Helper' sheet to the original daily file.")

        # Create a new DataFrame for the Helper sheet based on the modified data
        master_df = pd.read_excel(master_file_path, usecols=['Arabic', 'Name', 'Index', 'Type', 'Transf Type'])
        master_df.rename(columns={'Arabic': 'اسم المفوتر', 'Name': 'Cust', 'Index': 'Index', 'Type': 'TransType',
                                  'Transf Type': 'BillerType'}, inplace=True)

        # Get unique records from the modified daily file and merge with master data for TransType and BillerType
        helper_df = pd.merge(temp_df.drop_duplicates(subset=['اسم المفوتر']), master_df, on=['اسم المفوتر'], how='left',
                             suffixes=('', '_master'))

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