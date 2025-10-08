import xlwings as xw
import pandas as pd
import os
import datetime
import config
import time


m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
date = datetime.datetime(m_year, m_month, m_day)
trans_date = date.strftime("%Y/%m/%d")

curr_year = date.today().year
curr_month = date.today().strftime("%B")  # Full month name
curr_month_small = date.today().strftime("%b")  # Abbreviated month name
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

def process_customer_data(customer_list_file_path, customer_data_file_path):
    """
    Processes customer data, retrieves amounts, and pastes them into customer report files.

    Args:
        customer_list_file_path (str): Path to the customer list Excel file.
        customer_data_file_path (str): Path to the customer data Excel file.
        customer_report_directory (str): Path to the directory containing customer report files.
    """

    # try:
    # 1. Open the customer list workbook and create a DataFrame
    start_time = time.time()  # Start the timer
    customer_list_wb = xw.Book(customer_list_file_path)
    customer_list_sheet = customer_list_wb.sheets["Helper"]
    df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
    df = df.iloc[:, :6]  # Select only columns A to F
    today_sheet_name = f"{path_day}-{path_month_abbr}"
    # 2. Open the customer data workbook (only once)
    customer_data_wb = xw.Book(customer_data_file_path)
    customer_data_sheet = customer_data_wb.sheets[today_sheet_name]  # Assuming data is on the first sheet

    # 3 & 4. Loop through the customer list
    for index, row in df.iterrows():
        customer_name = row["CustomerName"]
        biller_type = row["BillerType"]

        # 3. Get the amount from the customer data file
        amount = None
        for cell in customer_data_sheet.range("B6:B" + str(customer_data_sheet.cells.last_cell.row)):
            if cell.value == customer_name:
                amount = cell.offset(0, 14).value  # Column P (offset 15)
                print(f"{customer_name} with {amount}")
                break
        if amount is None:
            print(f"Customer name '{customer_name}' not found in customer data file.")
            continue

        # 4. Open the customer report file and paste the amount
        # try:
        customer_report_directory = BILLER_REPORT_BASE
        biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"

        customer_report_file = os.path.join(customer_report_directory, biller_report_file) # Assuming file extension is xlsx. Change if needed.
        customer_report_wb = xw.Book(customer_report_file) # Open report file
        customer_report_sheet = customer_report_wb.sheets[0] # Assuming data is on the first sheet

        if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
            m_paste_range = "J5"

        elif biller_type in ("Biller With Sub-biller"):
            m_paste_range = "L5"

        else:
            print(f"Unknown biller type: {biller_type}")
            continue

        customer_report_sheet.range(m_paste_range).value = amount
        customer_report_wb.save()
        customer_report_wb.close()  # Close the report file

        # except FileNotFoundError:
        #     print(f"Customer report file not found for: {customer_name}")
        #     continue

        # except Exception as e:
        #     print(f"Error processing customer {customer_name}: {e}")
        #     continue

    # customer_data_wb.close()  # Close the customer data workbook (after processing all customers)
    customer_list_wb.save() # Save the changes to the customer list file.
    # customer_list_wb.close() # Close the customer list file.
    end_time = time.time()  # End the timer
    total_time = end_time - start_time

    print(f"Total time taken: {total_time:.2f} seconds.")
    print("Customer data processing complete.")

    # except Exception as e:
    #     print(f"An error occurred: {e}")


customer_list_file = config.config.dailyfile_name
customer_data_file = rf"All Billers Reconciliation Summary - {curr_month}.xlsm"


process_customer_data(customer_list_file, customer_data_file)