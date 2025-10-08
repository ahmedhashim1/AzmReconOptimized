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
path_month_abbr = date.today().strftime("%b")
path_day = date.today().strftime("%d")

DAILY_FILE_BASE = config.config.dailyfile_base
file_name = config.config.dailyfile_name
customers_file = rf"{DAILY_FILE_BASE}\{path_year}\{path_month_abbr}\{path_day}\{file_name}"
BILLER_REPORT_BASE = config.config.biller_base

# Path construction for the invoice base
INVOICE_BASE = config.config.invoice_base
invoice_base_folder = INVOICE_BASE
biller_summary_name = f"All Billers Reconciliation Summary - {path_month_full}.xlsm"
biller_summary_path_name = rf"Billers Summary\{path_year}\{path_month_abbr}\{biller_summary_name}"
biller_summary_path = os.path.join(invoice_base_folder, biller_summary_path_name)


def update_customer_report(customer_name, biller_type, amount):
    """
    Worker function to update a single customer's report file using pandas.
    """
    try:
        customer_report_directory = BILLER_REPORT_BASE
        biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"
        customer_report_file = os.path.join(customer_report_directory, biller_report_file)

        # Read the Excel file into a pandas DataFrame
        df_report = pd.read_excel(customer_report_file, header=None)

        # Determine the cell to update based on biller type
        if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
            row_index, col_index = 4, 9  # J5 is row 5, col 10 (0-indexed)
        elif biller_type in ("Biller With Sub-biller"):
            row_index, col_index = 4, 11  # L5 is row 5, col 12 (0-indexed)
        else:
            print(f"Unknown biller type: {biller_type} for customer {customer_name}. Skipping.")
            return

        # Update the DataFrame in memory
        df_report.iloc[row_index, col_index] = amount

        # Save the updated DataFrame back to the Excel file
        df_report.to_excel(customer_report_file, header=False, index=False)

        print(f"Successfully updated report for {customer_name}.")

    except FileNotFoundError:
        print(f"Customer report file not found for: {customer_name}")
    except Exception as e:
        print(f"Error processing customer {customer_name}: {e}")


def process_customer_data_optimized(customer_list_file_path, customer_data_file_path):
    """
    Processes customer data using pandas and a thread pool.
    """
    start_time = time.time()  # Start the timer

    try:
        # Read customer list and data files into DataFrames (one-time operation)
        df_customer_list = pd.read_excel(customer_list_file_path, sheet_name="Helper")
        df_customer_data = pd.read_excel(customer_data_file_path, sheet_name=f"{path_day}-{path_month_abbr}")

        # Create a thread pool to handle concurrent updates
        with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            for index, row in df_customer_list.iterrows():
                customer_name = row["CustomerName"]
                biller_type = row["BillerType"]

                # Find the amount for the customer from the data DataFrame
                amount_row = df_customer_data[df_customer_data.iloc[:, 1] == customer_name]
                if not amount_row.empty:
                    # Amount is in the 15th column (index 14)
                    amount = amount_row.iloc[0, 14]
                    print(f"Found amount {amount} for {customer_name}.")
                    # Submit the update task to the thread pool
                    executor.submit(update_customer_report, customer_name, biller_type, amount)
                else:
                    print(f"Customer name '{customer_name}' not found in customer data file.")
                    continue

        end_time = time.time()  # End the timer
        total_time = end_time - start_time
        print("Customer data processing complete.")
        print(f"Total time taken: {total_time:.2f} seconds.")

    except Exception as e:
        print(f"An error occurred: {e}")


# Script execution starts here
customer_list_file = customers_file
customer_data_file = biller_summary_path

process_customer_data_optimized(customer_list_file, customer_data_file)