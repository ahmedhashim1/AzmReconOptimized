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
    Simplified version without COM handler for better compatibility.

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
        thread_local_data.app = xw.App(visible=False, add_book=False)
        thread_local_data.app.display_alerts = False
        thread_local_data.app.screen_updating = False
    return thread_local_data.app


def cleanup_xlwings_app():
    """Clean up the xlwings app instance for the current thread."""
    if hasattr(thread_local_data, 'app'):
        try:
            thread_local_data.app.quit()
            del thread_local_data.app
        except:
            pass


def get_customer_amounts(customer_data_file_path: str, customer_names: List[str]) -> Dict[str, float]:
    """
    Pre-load all customer amounts from the data file to avoid repeated file access.

    Args:
        customer_data_file_path: Path to the customer data file
        customer_names: List of customer names to look for

    Returns:
        Dictionary mapping customer names to their amounts
    """
    print("Pre-loading customer amounts...")
    customer_amounts = {}

    try:
        app = get_xlwings_app()
        # Use helper function to get or open the customer data workbook
        customer_data_wb = get_or_open_workbook(app, customer_data_file_path, "customer data file")
        customer_data_sheet = customer_data_wb.sheets[today_sheet_name]

        # Get all data in column B and P at once
        last_row = customer_data_sheet.cells.last_cell.row
        names_range = customer_data_sheet.range(f"B6:B{last_row}")
        amounts_range = customer_data_sheet.range(f"P6:P{last_row}")

        names_values = names_range.value
        amounts_values = amounts_range.value

        # Handle single row case
        if not isinstance(names_values, list):
            names_values = [names_values]
            amounts_values = [amounts_values]

        # Create mapping
        for name, amount in zip(names_values, amounts_values):
            if name and name in customer_names:
                customer_amounts[name] = amount
                print(f"Loaded: {name} -> {amount}")

        # Don't close customer_data_wb - leave it open as requested
        print(f"Pre-loaded {len(customer_amounts)} customer amounts.")

    except Exception as e:
        print(f"Error pre-loading customer amounts: {e}")
        cleanup_xlwings_app()
        raise

    return customer_amounts


def process_single_customer(customer_data: Tuple[str, str, float]) -> Tuple[bool, str, str]:
    """
    Process a single customer's report file with enhanced error handling.

    Args:
        customer_data: Tuple of (customer_name, biller_type, amount)

    Returns:
        Tuple of (success, customer_name, message)
    """
    customer_name, biller_type, amount = customer_data

    def _process_customer():
        app = get_xlwings_app()

        # Construct file path
        customer_report_directory = BILLER_REPORT_BASE
        biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"
        customer_report_file = os.path.join(customer_report_directory, biller_report_file)

        # Use helper function to get or open the customer report workbook
        customer_report_wb = get_or_open_workbook(app, customer_report_file, f"report file for {customer_name}")

        try:
            customer_report_sheet = customer_report_wb.sheets[0]

            # Determine paste range based on biller type
            if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
                m_paste_range = "J5"
            elif biller_type in ("Biller With Sub-biller"):
                m_paste_range = "L5"
            else:
                return False, customer_name, f"Unknown biller type: {biller_type}"

            # Update amount and save with lock to prevent conflicts
            customer_report_sheet.range(m_paste_range).value = amount
            customer_report_wb.save()

            # Only close if we opened it (not if it was already open)
            try:
                if customer_report_wb.app.books.count > 1:  # Don't close if it's the only book
                    customer_report_wb.close()
            except:
                pass  # Ignore close errors

            return True, customer_name, f"Successfully updated with amount: {amount}"

        except Exception as e:
            try:
                customer_report_wb.close()
            except:
                pass
            raise e

    try:
        return safe_excel_operation(_process_customer, max_retries=1)
    except Exception as e:
        # Clean up on error
        cleanup_xlwings_app()
        return False, customer_name, f"Error after retries: {str(e)}"


def process_customer_data_multithreaded(customer_list_file_path: str, customer_data_file_path: str,
                                        max_workers: int = 3):
    """
    Processes customer data using multi-threading for improved performance.
    Reduced max_workers default to 2 for better stability with Excel.

    Args:
        customer_list_file_path: Path to the customer list Excel file
        customer_data_file_path: Path to the customer data Excel file
        max_workers: Maximum number of threads to use (default: 2 for stability)
    """
    start_time = time.time()

    try:
        print(f"Starting multi-threaded processing with {max_workers} workers...")
        print("Note: Reduced worker count for better Excel stability")

        # 1. Load customer list
        print("Loading customer list...")
        main_app = xw.App(visible=False, add_book=False)
        main_app.display_alerts = False
        main_app.screen_updating = False

        # Try to get existing workbook first, then open if not found
        customer_list_wb = get_or_open_workbook(main_app, customer_list_file_path, "customer list file")
        customer_list_sheet = customer_list_wb.sheets["Helper"]
        df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
        df = df.iloc[:, :6]  # Select only columns A to F

        print(f"Loaded {len(df)} customers from list.")

        # 2. Pre-load all customer amounts
        customer_names = df["CustomerName"].tolist()
        customer_amounts = get_customer_amounts(customer_data_file_path, customer_names)

        # 3. Prepare data for processing
        customer_data_list = []
        for index, row in df.iterrows():
            customer_name = row["CustomerName"]
            biller_type = row["BillerType"]

            if customer_name in customer_amounts:
                amount = customer_amounts[customer_name]
                customer_data_list.append((customer_name, biller_type, amount))
            else:
                print(f"Customer name '{customer_name}' not found in customer data file.")

        print(f"Prepared {len(customer_data_list)} customers for processing.")

        # 4. Process customers using ThreadPoolExecutor with reduced workers
        successful_updates = 0
        failed_updates = 0
        failed_customers = []

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_customer = {
                executor.submit(process_single_customer, customer_data): customer_data[0]
                for customer_data in customer_data_list
            }

            # Process completed tasks
            for future in as_completed(future_to_customer):
                customer_name = future_to_customer[future]
                try:
                    success, returned_name, message = future.result()
                    if success:
                        successful_updates += 1
                        print(f"✓ {returned_name}: {message}")
                    else:
                        failed_updates += 1
                        failed_customers.append((returned_name, message))
                        print(f"✗ {returned_name}: {message}")
                except Exception as e:
                    failed_updates += 1
                    failed_customers.append((customer_name, f"Exception: {str(e)}"))
                    print(f"✗ {customer_name}: Exception occurred - {str(e)}")

        # 5. Retry failed customers sequentially (single-threaded for stability)
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
                    time.sleep(2)  # Wait between retries
                    success, returned_name, message = process_single_customer(customer_retry_data)

                    if success:
                        retry_successful += 1
                        successful_updates += 1
                        failed_updates -= 1
                        print(f"✓ RETRY SUCCESS - {returned_name}: {message}")
                    else:
                        print(f"✗ RETRY FAILED - {returned_name}: {message}")

            print(f"🔄 Retry results: {retry_successful} successful out of {len(failed_customers)} attempts")

        # 6. Save customer list but don't close files as requested
        customer_list_wb.save()
        # Don't close customer_list_wb or main_app - leave them open as requested

        # 7. Report results
        end_time = time.time()
        total_time = end_time - start_time

        print("\n" + "=" * 60)
        print("PROCESSING SUMMARY")
        print("=" * 60)
        print(f"Total customers processed: {len(customer_data_list)}")
        print(f"Successful updates: {successful_updates}")
        print(f"Failed updates: {failed_updates}")
        print(f"Total time taken: {total_time:.2f} seconds")
        print(f"Average time per customer: {total_time / len(customer_data_list):.2f} seconds")
        print("Customer data processing complete.")

    except Exception as e:
        print(f"An error occurred in main processing: {e}")
        # Cleanup any remaining xlwings instances
        try:
            if 'main_app' in locals():
                main_app.quit()
        except:
            pass


def process_customer_data_original(customer_list_file_path: str, customer_data_file_path: str):
    """
    Original single-threaded version for comparison.
    """
    start_time = time.time()
    customer_list_wb = xw.Book(customer_list_file_path)
    customer_list_sheet = customer_list_wb.sheets["Helper"]
    df = customer_list_sheet.range("A1").expand('table').options(pd.DataFrame, header=1, index=False).value
    df = df.iloc[:, :6]

    customer_data_wb = xw.Book(customer_data_file_path)
    customer_data_sheet = customer_data_wb.sheets[today_sheet_name]

    successful_updates = 0
    failed_updates = 0

    for index, row in df.iterrows():
        customer_name = row["CustomerName"]
        biller_type = row["BillerType"]

        # Get amount
        amount = None
        for cell in customer_data_sheet.range("B6:B" + str(customer_data_sheet.cells.last_cell.row)):
            if cell.value == customer_name:
                amount = cell.offset(0, 14).value
                print(f"{customer_name} with {amount}")
                break

        if amount is None:
            print(f"Customer name '{customer_name}' not found in customer data file.")
            failed_updates += 1
            continue

        try:
            customer_report_directory = BILLER_REPORT_BASE
            biller_report_file = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx"
            customer_report_file = os.path.join(customer_report_directory, biller_report_file)

            customer_report_wb = xw.Book(customer_report_file)
            customer_report_sheet = customer_report_wb.sheets[0]

            if biller_type in ("Single Biller", "Single Biller with Adv Wallet"):
                m_paste_range = "J5"
            elif biller_type in ("Biller With Sub-biller"):
                m_paste_range = "L5"
            else:
                print(f"Unknown biller type: {biller_type}")
                failed_updates += 1
                continue

            customer_report_sheet.range(m_paste_range).value = amount
            customer_report_wb.save()
            customer_report_wb.close()
            successful_updates += 1

        except Exception as e:
            print(f"Error processing customer {customer_name}: {e}")
            failed_updates += 1

    customer_list_wb.save()
    end_time = time.time()
    total_time = end_time - start_time

    print(f"\nORIGINAL METHOD SUMMARY:")
    print(f"Successful: {successful_updates}, Failed: {failed_updates}")
    print(f"Success rate: {(successful_updates / (successful_updates + failed_updates) * 100):.1f}%")
    print(f"Processing time: {total_time:.2f} seconds")
    if (successful_updates + failed_updates) > 0:
        print(f"Average time per customer: {total_time / (successful_updates + failed_updates):.2f} seconds")
        print(f"Throughput: {(successful_updates + failed_updates) / total_time:.1f} customers/second")


# Main execution
if __name__ == "__main__":
    # Start timing the entire script execution
    script_start_time = time.time()
    print("=" * 60)
    print("SCRIPT EXECUTION STARTED")
    print("=" * 60)
    print(f"Start time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Processing date: {trans_date}")
    print()

    customer_list_file = config.config.dailyfile_name
    customer_data_file = rf"All Billers Reconciliation Summary - {curr_month}.xlsm"

    try:
        # Use multithreaded version with reduced workers for stability
        # Start with 2 workers - you can increase to 3-4 if stable
        process_customer_data_multithreaded(customer_list_file, customer_data_file, max_workers=2)

        # Uncomment below to run original version for comparison
        # print("\n" + "="*60)
        # print("RUNNING ORIGINAL VERSION FOR COMPARISON")
        # print("="*60)
        # process_customer_data_original(customer_list_file, customer_data_file)

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
        print(f"Total script execution time: {total_script_time / 60:.2f} minutes")

        # Additional time breakdown for better insights
        if total_script_time > 60:
            minutes = int(total_script_time // 60)
            seconds = int(total_script_time % 60)
            print(f"Formatted time: {minutes} minutes and {seconds} seconds")

        print("=" * 60)