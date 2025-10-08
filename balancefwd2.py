import xlwings as xw
import pandas as pd
import datetime
import config
import os
from ProcessRecon import assign_open_workbook


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

def update_biller_amounts(helper_file_name):
    """Updates biller amounts in individual Excel files based on a customer list.

    Args:
        helper_file_name: Name of the Excel file containing the customer list.
        curr_month: Full name of the current month (e.g., "October").
    """

    try:
        # 1. Read customer list into pandas DataFrame
        try:
            app = xw.apps.active
            app.display_alerts = False
            helper_wb = xw.books[helper_file_name]  # Assuming it's already open
            df = helper_wb.sheets["Helper"].range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
            df = df.iloc[:, :6] # Select only columns A to F

            required_columns = ["CustomerName", "BillerType"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(
                    f"Starting workbook is missing the following required columns: {missing_columns}. Available columns are: {df.columns.tolist()}")

        except KeyError:
            raise ValueError(f"Helper file '{helper_file_name}' not found. Make sure it's open.")
        except Exception as e:
            raise ValueError(f"Error reading data from helper file: {e}")


        # 2. Get amount to copy and loop through customers
        amount_to_copy = helper_wb.sheets["Helper"].range("G1").value
        last_copied_value = None  # Initialize to None

        for _, row in df.iterrows():  # Use _ for the index, as it's not needed
            customer_name = row["CustomerName"]
            biller_type = row["BillerType"]

            # 3. Determine target cell and copy cell based on biller type
            if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
                m_paste_range = "J6"
                m_copy_range = "L6"
            elif biller_type == "Biller With Sub-biller":
                m_paste_range = "B6"
                m_copy_range = "D6"
            else:
                print(f"Unknown biller type: {biller_type}. Skipping.")
                continue

            try:
                # 4. Open/Switch to customer's workbook (assuming it's already open)
                customer_wb_name = f"{customer_name} - {path_month_full} Internal Reconciliation Summary.xlsx"
                print(customer_name)
                try:
                    customer_wb = xw.books[customer_wb_name]
                except KeyError:
                    print(f"Warning: Customer workbook '{customer_wb_name}' not found. Skipping.")
                    continue

                # 5. Paste amount and copy new value
                target_cell = customer_wb.sheets[today_sheet_name].range(m_paste_range)  # Assumes sheet name is "Sheet1"
                if last_copied_value == None:
                    target_cell.value = amount_to_copy
                else:
                    target_cell.value = last_copied_value

                copy_cell = customer_wb.sheets[today_sheet_name].range(m_copy_range)
                last_copied_value = copy_cell.value  # Store the copied value for the next loop

            except Exception as e:
                print(f"Error processing {customer_name}: {e}")
                continue  # Go to the next customer

        print("Biller amounts updated successfully.")
        app.display_alerts = True

    except Exception as e:
        print(f"An error occurred: {e}")


# Example usage:
helper_file_name = config.config.dailyfile_name  # Replace with the actual name
update_biller_amounts(helper_file_name)