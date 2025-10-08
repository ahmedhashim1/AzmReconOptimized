import xlwings as xw
import pandas as pd
import os
import datetime
import config
import time
from concurrent.futures import ThreadPoolExecutor

# Configuration settings
m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
date = datetime.datetime(m_year, m_month, m_day)
trans_date = date.strftime("%Y/%m/%d")

curr_year = date.today().year
curr_month = date.today().strftime("%B")
curr_month_small = date.today().strftime("%b")
curr_day = date.today().strftime("%d")

path_year = date.strftime("%Y")
path_month_full = date.strftime("%B")
path_month_abbr = date.strftime("%b")
path_day = date.strftime("%d")

DAILY_FILE_BASE = config.config.dailyfile_base
file_name = config.config.dailyfile_name
customers_file = rf"{DAILY_FILE_BASE}\{path_year}\{path_month_abbr}\{path_day}\{file_name}"
today_sheet_name = f"{path_day}-{path_month_abbr}"
BILLER_REPORT_BASE = config.config.biller_base


def update_customer_report(customer_name, biller_type, amount):
    """
    Worker function to update a single customer's report file.
    This function will be executed by a separate thread.
    """
    try:
        customer_report_directory = BILLER_REPORT_BASE
        biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"
        customer_report_file = os.path.join(customer_report_directory, biller_report_file)

        # Open the customer report file
        customer_report_wb = xw.Book(customer_report_file)
        customer_report_sheet = customer_report_wb.sheets[0]

        # Determine the paste range based on biller type
        if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
            m_paste_range = "J5"
        elif biller_type in ("Biller With Sub-biller"):
            m_paste_range = "L5"
        else:
            print(f"Unknown biller type: {biller_type} for customer {customer_name}. Skipping.")
            customer_report_wb.close()
            return

        # Paste the amount and save
        customer_report_sheet.range(m_paste_range).value = amount
        customer_report_wb.save()
        customer_report_wb.close()
        print(f"Successfully updated report for {customer_name}.")

    except FileNotFoundError:
        print(f"Customer report file not found for: {customer_name}")
    except Exception as e:
        print(f"Error processing customer {customer_name}: {e}")


def process_customer_data(customer_list_file_path, customer_data_file_path):
    """
    Processes customer data using a thread pool to update report files concurrently.
    """
    start_time = time.time()  # Start the timer

    try:
        # 1. Open the customer list workbook and create a DataFrame
        customer_list_wb = xw.Book(customer_list_file_path)
        customer_list_sheet = customer_list_wb.sheets["Helper"]
        df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
        df = df.iloc[:, :6]

        # 2. Open the customer data workbook (only once)
        customer_data_wb = xw.Book(customer_data_file_path)
        customer_data_sheet = customer_data_wb.sheets[today_sheet_name]

        # 3. Create a thread pool to handle concurrent updates
        with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            # 4. Loop through the customer list and submit tasks to the pool
            for index, row in df.iterrows():
                customer_name = row["CustomerName"]
                biller_type = row["BillerType"]

                # Get the amount from the customer data file
                amount = None
                for cell in customer_data_sheet.range("B6:B" + str(customer_data_sheet.cells.last_cell.row)):
                    if cell.value == customer_name:
                        amount = cell.offset(0, 14).value
                        print(f"Found amount {amount} for {customer_name}.")
                        break

                if amount is None:
                    print(f"Customer name '{customer_name}' not found in customer data file.")
                    continue

                # Submit the update task to the thread pool
                executor.submit(update_customer_report, customer_name, biller_type, amount)

        # 5. Do not close workbooks as requested

        end_time = time.time()  # End the timer
        total_time = end_time - start_time
        print("Customer data processing complete.")
        print(f"Total time taken: {total_time:.2f} seconds.")

    except Exception as e:
        print(f"An error occurred: {e}")


# Script execution starts here
customer_list_file = config.config.dailyfile_name
customer_data_file = rf"All Billers Reconciliation Summary - {curr_month}.xlsm"
process_customer_data(customer_list_file, customer_data_file)