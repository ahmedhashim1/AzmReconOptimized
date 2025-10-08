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

# MySQL configuration details
mysql_config = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "azm"
}

# Define a global connection pool
connection_pool = None
pool_lock = Lock()


def get_mysql_connection_pool():
    """Create or get a connection pool for better performance"""
    global connection_pool
    if connection_pool is None:
        with pool_lock:
            if connection_pool is None:
                try:
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
                    connection_pool = pool
                except mysql.connector.Error as err:
                    print(f"Error creating connection pool: {err}")
    return connection_pool


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
        today = date.today()
        daily_sheet.range('U1').value = 'fdate'
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


def create_customer_file(customer_name, query_date, output_dir):
    """
    Fetches data for a single customer from the database and saves it to a new, formatted Excel file.
    This function uses xlwings with robust error handling for multi-threading.
    """
    app = None
    conn = None
    try:
        pool = get_mysql_connection_pool()
        if not pool:
            return f"Failed to get connection for {customer_name}"

        conn = pool.get_connection()
        cursor = conn.cursor(dictionary=True)

        query = "SELECT * FROM dailyfiledto WHERE `Cust` = %s AND fdate = %s"
        cursor.execute(query, (customer_name, query_date))

        df = pd.DataFrame(cursor.fetchall())

        # Explicitly convert columns to string to preserve formatting
        for col in ['InvoiceNum', 'InternalCode', 'ContractNum']:
            if col in df.columns:
                df[col] = df[col].astype(str).apply(lambda x: f"'{x}")

        safe_filename = "".join(c for c in str(customer_name) if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
        output_file = os.path.join(output_dir, f"{safe_filename}.xlsx")

        if not df.empty:
            # Create a new, isolated xlwings app instance for this thread
            app = xw.App(visible=False)
            new_wb = app.books.add()
            new_ws = new_wb.sheets[0]

            # Write the DataFrame to the new sheet without the index
            new_ws.range('A1').options(index=False).value = df

            # Apply formatting
            new_ws.autofit()
            new_ws.range('A1').expand('right').api.Font.Bold = True

            new_wb.save(output_file)
            new_wb.close()

            return f"✓ Saved: {safe_filename}.xlsx"
        else:
            return f"No data for customer {customer_name}"

    except Exception as e:
        # If an error occurs, try to kill the app forcefully
        if app is not None:
            try:
                app.kill()
            except Exception:
                pass
        return f"✗ Error processing {customer_name}: {e}"
    finally:
        if app is not None:
            try:
                # Always attempt a clean quit
                app.quit()
            except Exception:
                pass
        if conn and conn.is_connected():
            conn.close()


def add_helper_and_split_files_from_db(daily_file_path, master_file_path):
    """
    Adds a 'Helper' sheet to the daily file using the Excel file and splits the data into separate
    Excel files for each unique customer using MySQL.
    """
    app = None
    conn = None
    try:
        today = date.today()

        # --- PART 1: Split the data into separate files for each customer using MySQL ---
        print("\nPART 1: Splitting data by customer using MySQL queries.")
        pool = get_mysql_connection_pool()
        if not pool:
            print("Database connection pool not available. Exiting split process.")
            return

        conn = pool.get_connection()
        cursor = conn.cursor()

        # Get unique customers from the daily data in the database for today's date
        query_customers = "SELECT DISTINCT `Cust` FROM dailyfiledto WHERE fdate = %s"
        cursor.execute(query_customers, (today,))
        unique_customers = [row[0] for row in cursor.fetchall()]
        print(f"Found {len(unique_customers)} unique customers for {today}. Creating a file for each.")

        file_dir = os.path.dirname(daily_file_path)

        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_customer = {executor.submit(create_customer_file, customer, today, file_dir): customer for
                                  customer in unique_customers}
            for future in concurrent.futures.as_completed(future_to_customer):
                customer_name = future_to_customer[future]
                try:
                    result = future.result()
                    print(result)
                except Exception as exc:
                    print(f"An error occurred while creating file for {customer_name}: {exc}")

        print("\nCompleted! File splitting process is finished.")

        # --- PART 2: Add the 'Helper' sheet with a unique summary to the original file ---
        print("\nPART 2: Adding 'Helper' sheet to the original daily file from Excel files.")

        # Read the now-fully-modified daily file into a DataFrame
        full_data = pd.read_excel(daily_file_path)

        # Create a new DataFrame for the Helper sheet based on the modified data
        master_df = pd.read_excel(master_file_path, usecols=['Arabic', 'Name', 'Index', 'Type', 'Transf Type'])
        master_df.rename(columns={'Arabic': 'اسم المفوتر', 'Name': 'Cust', 'Index': 'Index', 'Type': 'TransType',
                                  'Transf Type': 'BillerType'}, inplace=True)

        # Get unique records from the modified daily file and merge with master data for TransType and BillerType
        helper_df = pd.merge(full_data.drop_duplicates(subset=['اسم المفوتر']), master_df, on=['اسم المفوتر'],
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

        # Connect to Excel to add the Helper sheet
        app = xw.apps.active
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

        helper_sheet.range('A1').options(index=False).value = final_helper_df

        daily_wb.save()

        print("PART 2 completed successfully. 'Helper' sheet added to original file.")

    except Error as e:
        print(f"An error occurred: {e}")
    finally:
        if conn and conn.is_connected():
            conn.close()


if __name__ == "__main__":
    daily_file = r"D:\Freelance\Azm\2025\Sep\21\Test\DailyFile_1.xlsx"  # Replace with the actual file name parameter
    master_file = r"D:\Freelance\Azm\2025\CustomerNamesLookUp.xlsx"  # Make sure this file exists in the same directory or provide the full path

    # Run the initial file modification process
    modify_excel_file_final(daily_file, master_file)

    # Run the new, database-based process
    # NOTE: This assumes you have already uploaded the data from your Excel file
    # into a MySQL table named 'dailyfiledto'.
    add_helper_and_split_files_from_db(daily_file, master_file)