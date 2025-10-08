import xlwings as xw
import pandas as pd
import os
import datetime
import config
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from typing import Dict, Any, List, Tuple
import queue

# Global variables
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

# Thread-local storage for xlwings app instances
thread_local_data = threading.local()


def safe_excel_operation(func, *args, max_retries=2, **kwargs):
    """
    Safely execute an Excel operation with retries and error handling.

    Args:
        func: Function to execute
        *args: Arguments for the function
        max_retries: Maximum number of retries
        **kwargs: Keyword arguments for the function

    Returns:
        Result of the function or raises exception
    """
    for attempt in range(max_retries + 1):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < max_retries:
                print(f"Attempt {attempt + 1} failed, retrying: {str(e)}")
                # Cleanup and recreate app on error
                cleanup_xlwings_app()
                time.sleep(1)  # Wait before retry
            else:
                raise e


def get_or_open_workbook(app, file_path: str, file_description: str = "file"):
    """
    Get an existing workbook or open it if not already open with better error handling.

    Args:
        app: xlwings App instance
        file_path: Path to the Excel file
        file_description: Description for logging purposes

    Returns:
        xlwings Workbook object
    """

    def _open_workbook():
        try:
            # First try to reference by full path
            workbook = xw.Book(file_path)
            print(f"Found already open {file_description}: {os.path.basename(file_path)}")
            return workbook
        except:
            try:
                # Try to reference by filename only (in case path differs)
                filename = os.path.basename(file_path)
                workbook = xw.Book(filename)
                print(f"Found already open {file_description}: {filename}")
                return workbook
            except:
                # File not open, try to open it
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"{file_description} not found: {file_path}")

                workbook = app.books.open(file_path)
                print(f"Opened {file_description}: {os.path.basename(file_path)}")
                return workbook

    return safe_excel_operation(_open_workbook)


def get_xlwings_app():
    """Get or create an xlwings app instance for the current thread."""
    if not hasattr(thread_local_data, 'app'):
        # Make Excel visible since we're just opening files for review
        thread_local_data.app = xw.App(visible=True, add_book=False)
        thread_local_data.app.display_alerts = False
        thread_local_data.app.screen_updating = True  # Keep screen updating on for visibility
    return thread_local_data.app


def cleanup_xlwings_app():
    """Clean up the xlwings app instance for the current thread."""
    if hasattr(thread_local_data, 'app'):
        try:
            # Don't quit the app - leave it open with the files for user review
            # thread_local_data.app.quit()
            del thread_local_data.app
        except:
            pass


def open_single_customer_report(customer_data: Tuple[str, str]) -> Tuple[bool, str, str]:
    """
    Open a single customer's report file for review.

    Args:
        customer_data: Tuple of (customer_name, biller_type)

    Returns:
        Tuple of (success, customer_name, message)
    """
    customer_name, biller_type = customer_data

    def _open_customer_report():
        app = get_xlwings_app()

        # Construct file path
        customer_report_directory = BILLER_REPORT_BASE
        biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"
        customer_report_file = os.path.join(customer_report_directory, biller_report_file)

        # Use helper function to get or open the customer report workbook
        customer_report_wb = get_or_open_workbook(app, customer_report_file, f"report file for {customer_name}")

        try:
            # Just open the file - no modifications needed
            customer_report_sheet = customer_report_wb.sheets[0]

            # Determine the relevant cell based on biller type for informational purposes
            if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
                relevant_cell = "J5"
                current_value = customer_report_sheet.range(relevant_cell).value
            elif biller_type in ("Biller With Sub-biller"):
                relevant_cell = "L5"
                current_value = customer_report_sheet.range(relevant_cell).value
            else:
                relevant_cell = "N/A"
                current_value = "N/A"

            # Select the relevant cell to highlight it for the user
            if relevant_cell != "N/A":
                customer_report_sheet.range(relevant_cell).select()

            return True, customer_name, f"Opened successfully. Current amount cell ({relevant_cell}): {current_value}"

        except Exception as e:
            raise e

    try:
        return safe_excel_operation(_open_customer_report, max_retries=1)
    except Exception as e:
        return False, customer_name, f"Error opening report: {str(e)}"


def open_customer_reports_multithreaded(customer_list_file_path: str, max_workers: int = 3):
    """
    Opens customer report files using multi-threading for faster access.

    Args:
        customer_list_file_path: Path to the customer list Excel file
        max_workers: Maximum number of threads to use (default: 3)
    """
    start_time = time.time()

    try:
        print(f"Starting multi-threaded report opening with {max_workers} workers...")

        # 1. Load customer list
        print("Loading customer list...")
        main_app = xw.App(visible=True, add_book=False)  # Visible for user interaction
        main_app.display_alerts = False

        # Try to get existing workbook first, then open if not found
        customer_list_wb = get_or_open_workbook(main_app, customer_list_file_path, "customer list file")
        customer_list_sheet = customer_list_wb.sheets["Helper"]
        df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
        df = df.iloc[:, :6]  # Select only columns A to F

        print(f"Loaded {len(df)} customers from list.")

        # 2. Prepare data for opening reports
        customer_data_list = []
        for index, row in df.iterrows():
            customer_name = row["CustomerName"]
            biller_type = row["BillerType"]
            customer_data_list.append((customer_name, biller_type))

        print(f"Prepared {len(customer_data_list)} customer reports for opening.")
        print(f"Will open all {len(customer_data_list)} customer reports automatically.")

        # 3. Open customer reports using ThreadPoolExecutor
        successful_opens = 0
        failed_opens = 0
        failed_customers = []

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_customer = {
                executor.submit(open_single_customer_report, customer_data): customer_data[0]
                for customer_data in customer_data_list
            }

            # Process completed tasks
            for future in as_completed(future_to_customer):
                customer_name = future_to_customer[future]
                try:
                    success, returned_name, message = future.result()
                    if success:
                        successful_opens += 1
                        print(f"✓ {returned_name}: {message}")
                    else:
                        failed_opens += 1
                        failed_customers.append((returned_name, message))
                        print(f"✗ {returned_name}: {message}")
                except Exception as e:
                    failed_opens += 1
                    failed_customers.append((customer_name, f"Exception: {str(e)}"))
                    print(f"✗ {customer_name}: Exception occurred - {str(e)}")

        # 4. Retry failed customers sequentially if any
        if failed_customers:
            print(f"\n🔄 Retrying {len(failed_customers)} failed customers sequentially...")
            retry_successful = 0

            for customer_name, error_msg in failed_customers:
                # Find the customer data for retry
                customer_retry_data = None
                for customer_data in customer_data_list:
                    if customer_data[0] == customer_name:
                        customer_retry_data = customer_data
                        break

                if customer_retry_data:
                    print(f"🔄 Retrying {customer_name}...")
                    time.sleep(1)  # Wait between retries
                    success, returned_name, message = open_single_customer_report(customer_retry_data)

                    if success:
                        retry_successful += 1
                        successful_opens += 1
                        failed_opens -= 1
                        print(f"✓ RETRY SUCCESS - {returned_name}: {message}")
                    else:
                        print(f"✗ RETRY FAILED - {returned_name}: {message}")

            print(f"🔄 Retry results: {retry_successful} successful out of {len(failed_customers)} attempts")

        # 5. Report results
        end_time = time.time()
        total_time = end_time - start_time

        print("\n" + "=" * 60)
        print("REPORT OPENING SUMMARY")
        print("=" * 60)
        print(f"Total reports processed: {len(customer_data_list)}")
        print(f"Successfully opened: {successful_opens}")
        print(f"Failed to open: {failed_opens}")
        print(f"Total time taken: {total_time:.2f} seconds")
        if len(customer_data_list) > 0:
            print(f"Average time per report: {total_time / len(customer_data_list):.2f} seconds")
        print("Report opening complete.")
        print("\nNOTE: All opened Excel files are left open for your review.")
        print("You can now review and modify the reports as needed.")

    except Exception as e:
        print(f"An error occurred in main processing: {e}")


def open_customer_reports_sequential(customer_list_file_path: str):
    """
    Sequential version for opening customer reports one by one automatically.
    """
    start_time = time.time()

    try:
        print("Starting sequential report opening...")

        # Load customer list
        main_app = xw.App(visible=True, add_book=False)
        main_app.display_alerts = False

        customer_list_wb = get_or_open_workbook(main_app, customer_list_file_path, "customer list file")
        customer_list_sheet = customer_list_wb.sheets["Helper"]
        df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
        df = df.iloc[:, :6]

        print(f"Loaded {len(df)} customers from list.")

        successful_opens = 0
        failed_opens = 0

        for index, row in df.iterrows():
            customer_name = row["CustomerName"]
            biller_type = row["BillerType"]

            print(f"[{index + 1}/{len(df)}] Opening report for: {customer_name}")

            success, returned_name, message = open_single_customer_report((customer_name, biller_type))

            if success:
                successful_opens += 1
                print(f"✓ {returned_name}: {message}")
            else:
                failed_opens += 1
                print(f"✗ {returned_name}: {message}")

            # Small delay between operations
            time.sleep(0.5)

        end_time = time.time()
        total_time = end_time - start_time

        print(f"\nSEQUENTIAL OPENING SUMMARY:")
        print(f"Successfully opened: {successful_opens}")
        print(f"Failed to open: {failed_opens}")
        print(f"Processing time: {total_time:.2f} seconds")

    except Exception as e:
        print(f"An error occurred: {e}")


# Main execution
if __name__ == "__main__":
    # Start timing the entire script execution
    script_start_time = time.time()
    print("=" * 60)
    print("BILLER REPORTS OPENER - SCRIPT STARTED")
    print("=" * 60)
    print(f"Start time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Processing date: {trans_date}")
    print()

    customer_list_file = config.config.dailyfile_name

    try:
        print("Choose opening method:")
        print("1. Multi-threaded opening (faster, opens multiple reports simultaneously)")
        print("2. Sequential opening (one by one, with user control)")

        method_choice = input("Enter your choice (1 or 2): ").strip()

        if method_choice == "2":
            open_customer_reports_sequential(customer_list_file)
        else:
            # Default to multi-threaded
            max_workers = 3  # You can adjust this number
            try:
                workers_input = input(f"Enter number of workers (default {max_workers}): ").strip()
                if workers_input:
                    max_workers = int(workers_input)
            except:
                pass

            open_customer_reports_multithreaded(customer_list_file, max_workers=max_workers)

    except Exception as e:
        print(f"Critical error during script execution: {e}")

    finally:
        # Calculate and display total script execution time
        script_end_time = time.time()
        total_script_time = script_end_time - script_start_time

        print("\n" + "=" * 60)
        print("SCRIPT EXECUTION COMPLETED")
        print("=" * 60)
        print(f"End time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Total script execution time: {total_script_time:.2f} seconds")

        if total_script_time > 60:
            print(f"Total script execution time: {total_script_time / 60:.2f} minutes")
            minutes = int(total_script_time // 60)
            seconds = int(total_script_time % 60)
            print(f"Formatted time: {minutes} minutes and {seconds} seconds")

        print("\nAll reports are left open for your review and editing.")
        print("=" * 60)