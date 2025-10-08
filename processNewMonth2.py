import os
import shutil
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Date setup
date_str = "2025-09-30"
date = datetime(2025, 9, 30)
date_obj = datetime.strptime(date_str, "%Y-%m-%d")


def add_months_return_month_name(date_obj, months):
    """Adds months to a datetime object and returns the full month name."""
    new_date = date_obj + relativedelta(months=months)
    return new_date.strftime("%B")


def copy_and_rename_files(base_dir, old_month, new_month):
    """
    Copies and renames Excel files within a directory structure.

    Args:
        base_dir: The base directory of the invoice folder.
        old_month: The name of the old month (e.g., "January").
        new_month: The name of the new month (e.g., "February").
    """
    # File paths
    path_year = date.strftime("%Y")  # Year to search for (e.g., "2025")
    path_month_full = date.strftime("%B")
    path_month_full_new = add_months_return_month_name(date_obj, 1)
    path_month_abbr = date.strftime("%b")
    path_day = date.strftime("%d")

    with xw.App(visible=False) as app:
        for customer_folder in os.listdir(base_dir):
            customer_path = os.path.join(base_dir, customer_folder)
            if os.path.isdir(customer_path):
                for year_folder in os.listdir(customer_path):
                    # Only process the folder for the specified year
                    if year_folder == path_year:
                        year_path = os.path.join(customer_path, year_folder)
                        if os.path.isdir(year_path):
                            old_month_path = os.path.join(year_path, old_month)
                            new_month_path = os.path.join(year_path, new_month)

                            if os.path.exists(old_month_path):
                                if not os.path.exists(new_month_path):
                                    os.makedirs(new_month_path)

                                for file in os.listdir(old_month_path):
                                    if file.endswith(".xlsx"):
                                        old_file_path = os.path.join(old_month_path, file)
                                        new_file_name = f"{customer_folder} - {path_month_full_new} Internal Reconciliation Summary.xlsx"
                                        new_file_path = os.path.join(new_month_path, new_file_name)

                                        shutil.copy2(old_file_path, new_file_path)

                                        try:
                                            # Open and modify the copied file
                                            with app.books.open(new_file_path) as wb:
                                                # Get visible sheets
                                                visible_sheets = [
                                                    sheet for sheet in wb.sheets
                                                    if sheet.visible != xw.constants.SheetVisibility.xlSheetVeryHidden
                                                    and sheet.visible != xw.constants.SheetVisibility.xlSheetHidden
                                                ]

                                                # Delete unnecessary sheets
                                                for sheet in wb.sheets:
                                                    if sheet.name not in ["Template", "Summary", "DataForFilters"] and sheet.name != visible_sheets[-1].name:
                                                        try:
                                                            sheet.delete()
                                                        except Exception as e:
                                                            print(f"Failed to delete sheet '{sheet.name}' in file '{new_file_path}': {e}")

                                                # Select the "Summary" sheet
                                                summary_sheet = wb.sheets["Summary"]
                                                summary_sheet.activate()

                                                # Set date for the first day of the new month
                                                first_day_of_new_month = datetime.strptime(f"01-{new_month}-{path_year}", "%d-%b-%Y").date()
                                                summary_sheet.range('A5').value = first_day_of_new_month.strftime("%m/%d/%Y")

                                                # Save and close the workbook
                                                wb.save()
                                        except Exception as e:
                                            print(f"Error processing file '{new_file_path}': {e}")

        # Ensure the Excel application is properly closed
        # app.quit()


if __name__ == "__main__":
    base_dir = rf"D:\Freelance\Azm\OneDrive - AZM Saudi\Customers\Reconcilation Reports"  # Replace with the actual base directory path
    # Ensure to change 2 dates also in top of this module to the last working day of old month
    old_month = "Sep"  # Replace with the actual old month
    new_month = "Oct"  # Replace with the actual new month

    copy_and_rename_files(base_dir, old_month, new_month)