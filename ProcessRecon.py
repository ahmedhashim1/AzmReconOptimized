import concurrent
import os
import shutil
import pymysql
import openpyxl
import pandas
import config
import xlwings as xw
import mysql.connector
import win32com.client as win32
from mysql.connector import Error
import datetime
from xlwings import Book, Sheet, Range
from win32com.client import Dispatch
from typing import List, Any, Union


# date = datetime.now()
m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
date = datetime.datetime(m_year, m_month, m_day)
trans_date = date.strftime("%Y/%m/%d")

BILLER_REPORT_BASE = config.config.biller_base
# BILLER_REPORT_BASE = rf"E:\ReconTest\Biller Reports"

def get_first_listobject_name(workbook, sheet_name):
  """
  Gets the name of the first ListObject in a specified sheet.

  Args:
    workbook_path: Path to the Excel file.
    sheet_name: Name of the sheet containing the ListObject.

  Returns:
    The name of the first ListObject in the sheet, or None if no ListObjects are found.
  """

  try:
    wb = workbook
    wb.visible = True
    sheet = wb.sheets[sheet_name]

    # Check if any ListObjects exist in the sheet
    if len(sheet.tables) > 0:
      first_listobject = sheet.tables[0]  # Get the first ListObject
      return first_listobject.name
    else:
      return None

  except Exception as e:
    print(f"Error: {e}")
    return None


def get_range_address_from_named_range(workbook, named_range, sheet_name=None):
  """
  Resolves a named range or table to a range address.
  If the named range is scoped to a specific sheet, provide the sheet_name.
  """
  try:
    if sheet_name:
      # Access the named range scoped to a specific sheet
      sheet = workbook.sheets(sheet_name)
      named_item = sheet.names(named_range)
    else:
      # Access the named range scoped to the workbook
      named_item = workbook.Names(named_range)

    # Get the range address
    return named_item.RefersToRange.Address
  except Exception as e:
    print(f"Error resolving named range '{named_range}': {e}")
    return None

def change_pivot_data_source2(workbook_path, sheet_name, pivot_table_name, new_data_source):
  excel_app = Dispatch("Excel.Application")
  excel_app.DisplayAlerts = False

  try:
    # Open the Excel workbook
    wb = workbook_path
    wb.visible = True
    sheet = wb.sheets(sheet_name)

    # Debug: Print workbook and sheet info
    print(f"Workbook: {wb.name}, Sheet: {sheet.name}")

    # Resolve named range or table to a range address (if applicable)
    if "!" not in new_data_source:  # Check if it's not already a range address
      resolved_range = new_data_source
      if resolved_range:
        new_data_source = f"{sheet_name}!{resolved_range}"
        print(f"Resolved named range '{new_data_source}' to range address: {new_data_source}")
      else:
        print(f"Error: Could not resolve named range '{new_data_source}'.")
        return

    # Get the PivotTable
    try:
      pivot_table = sheet.PivotTables(pivot_table_name)  # Access the existing PivotTable by name
      print(f"PivotTable '{pivot_table_name}' found.")
    except Exception as e:
      print(f"Error: PivotTable '{pivot_table_name}' not found. Details: {e}")
      return

    # Update the PivotTable's data source
    try:
      # Get the PivotCache associated with the PivotTable
      pivot_cache = pivot_table.PivotCache()
      print("PivotCache accessed successfully.")

      # Debug: Print current source data
      print(f"Current SourceData: {pivot_cache.SourceData}")

      # Change the source data of the PivotCache
      pivot_cache.SourceData = new_data_source
      print(f"SourceData updated to: {new_data_source}")

      # Refresh the PivotTable
      pivot_table.RefreshTable()
      print(f"PivotTable '{pivot_table_name}' refreshed successfully.")
    except Exception as e:
      print(f"Error updating PivotTable: {e}")

  except Exception as e:
    print(f"Error: {e}")
  finally:
    # Quit the Excel application
    excel_app.DisplayAlerts = True
    # excel_app.Quit()


# Example usage:
# change_pivot_data_source("C:/path/to/workbook.xlsx", "Sheet1", "PivotTable1", "Sheet1!A1:D100")

def change_pivot_data_source(workbook_path, sheet_name, pivot_table_name, new_data_source):
  # excel_app = Dispatch("Excel.Application")
  # excel_app.DisplayAlerts = False
  app = xw.apps.active
  app.display_alerts = False
  # Open the Excel workbook
  wb = assign_open_workbook(workbook_path)
  sheet = wb.sheets[sheet_name]

  # Get the PivotTable
  try:

    pivot_table = sheet.api.PivotTables(pivot_table_name)  # Access the existing PivotTable by name
    print(f"PivotTable '{pivot_table_name}' found.")
  except Exception as e:
    print(f"Error: PivotTable '{pivot_table_name}' not found.")
    return



  # Update the PivotTable's data source
  try:

    # Update the PivotCache with the new data source
    pivot_table.ChangePivotCache(
      wb.api.PivotCaches().Create(
        SourceType=1,  # xlDatabase
        SourceData=new_data_source  # New data source range (e.g., "Sheet1!A1:D100")
      )
    )
    # # Get the PivotCache associated with the PivotTable
    # pivot_cache = pivot_table.PivotCache()
    # print("PivotCache accessed successfully.")
    # print(f"Current SourceData: {pivot_cache.SourceData}")
    # print(f"New SourceData name: {new_data_source}")
    #
    # # Change the source data of the PivotCache
    # pivot_cache.SourceData = new_data_source
    # Refresh the PivotTable
    pivot_table.RefreshTable()
    print(f"PivotTable '{pivot_table_name}' data source successfully updated to '{new_data_source}'.")
    # excel_app.DisplayAlerts = True
    app.display_alerts = True
  except Exception as e:
    print(f"Error updating PivotTable: {e}")

def copy_value_between_sheets(workbook, source_cell, target_cell):
  """
  Copies a value from a cell in the 2nd-to-last sheet to a cell in the last sheet.

  Args:
      workbook (xw.Book): The xlwings workbook object.
      source_cell (str): The cell in the 2nd-to-last sheet to copy from (e.g., "J15").
      target_cell (str): The cell in the last sheet to copy to (e.g., "J12").
  """
  # Get the source and target sheets
  source_sheet = workbook.sheets[-2]
  target_sheet = workbook.sheets[-1]

  # Get the value from the source cell
  value_to_copy = source_sheet.range(source_cell).value

  # Copy the value to the target cell
  target_sheet.range(target_cell).value = value_to_copy

  print(f"Value copied from {source_sheet.name} {source_cell} to {target_sheet.name} {target_cell}.")


def Add_report_Date(workbook, target_cell):
  """
  Copies a value from a cell in the 2nd-to-last sheet to a cell in the last sheet.

  Args:
      workbook (xw.Book): The xlwings workbook object.
      source_cell (str): The cell in the 2nd-to-last sheet to copy from (e.g., "J15").
      target_cell (str): The cell in the last sheet to copy to (e.g., "J12").
  """
  # Get the source and target sheets
  source_sheet = workbook.sheets[-2]
  target_sheet = workbook.sheets[-1]

  # Get the value from the source cell
  value_to_copy = source_sheet.range(source_cell).value

  # Copy the value to the target cell
  target_sheet.range(target_cell).value = value_to_copy

  print(f"Value copied from {source_sheet.name} {source_cell} to {target_sheet.name} {target_cell}.")
def ensure_folder_exists(folder_path):
  """
  Checks if the given folder path exists. If not, creates the folder.

  Args:
      folder_path (str): The path to the folder to check/create.
  """
  if not os.path.exists(folder_path):
    os.makedirs(folder_path)
    print(f"Folder created at: {folder_path}")
  else:
    print(f"Folder already exists at: {folder_path}")

def biller_report_create2(source_path, new_path):
  """
    Save an Excel workbook to a new path by copying it using the OS.

    Parameters:
        source_path (str): The full path of the source Excel file.
        new_path (str): The full path where the new Excel file will be saved.
    """
  try:
      if os.path.exists(new_path):
        os.remove(new_path)
      # Copy the file from source to destination
      shutil.copy(source_path, new_path)
      print(f"Workbook copied successfully to: {new_path}")
  except FileNotFoundError:
      print(f"Error: The source file '{source_path}' does not exist.")
  except PermissionError:
      print(f"Error: Permission denied for accessing '{new_path}'.")
  except Exception as e:
      print(f"An unexpected error occurred: {e}")


  # raise
def export_data_to_list_object_xlwings(workbook, sheet_name, list_object_name, data, columns, cust_name,cust_type ,exp_type):
  """
  Imports data into a specified ListObject in an Excel workbook using xlwings.

  Args:
    workbook: The xlwings Workbook object.
    sheet_name: The name of the sheet containing the ListObject.
    list_object_name: The name of the ListObject.
    data: A MySql Cursor
    columns: A list of column names.

  Raises:
    Exception: If the sheet or ListObject is not found.
  """
  # print(cust_name)
  try:
    # Get the sheet
    if exp_type == "Recon":
      sh_name = sheet_name
    else:
      sh_name = f"{cust_name} Report"

    sheet = workbook.sheets[sh_name]

    # Find the list object
    list_object = sheet.api.ListObjects(list_object_name)
    if not list_object:
      raise Exception(f"ListObject '{list_object_name}' not found on sheet '{sheet_name}'.")

    # Clear existing data
    list_object.DataBodyRange.ClearContents()

    # Get the start cell of the list object
    start_cell = list_object.Range.Cells(1, 1)

    # Calculate dimensions
    if len(data) > 1:
      num_rows = len(data) + 2
    else:
      num_rows = len(data) + 1

    table_range = list_object.Range
    second_column = table_range.Columns(2)
    second_column.NumberFormat = "@"

    if cust_type == "Single Biller" or cust_type == "Single Biller with Adv Wallet":
      int_code_column = table_range.Columns(8)
      int_code_column.NumberFormat = "@"
    else:
      int_code_column = table_range.Columns(10)
      int_code_column.NumberFormat = "@"


    num_cols = len(columns)
    end_cell = start_cell.Offset(num_rows + 1, num_cols + 1)  # Bottom-right corner of the new range

    new_range = sheet.range((start_cell.Row, start_cell.Column), (end_cell.Row, end_cell.Column)).api
    list_object.Resize(new_range)
    # Write data to the list object


    write_range = new_range
    write_start_cell = list_object.Range.Cells(2, 2)  # Start from the second column
    write_range = sheet.range((write_start_cell.Row, write_start_cell.Column),
                              (write_start_cell.Row + num_rows - 2, write_start_cell.Column + num_cols - 1)).api
    write_range.Value = data

    # Write serial numbers in the first column
    serial_range = list_object.Range.Columns(1)  # First column of the table
    header_row = serial_range.Cells(1, 1)
    footer_row = serial_range.Cells(num_rows + 2, 1)  # Assuming footer is the last row after data

    serial_range.Cells(2, 1).Formula = f"=ROW()-ROW({header_row.Address})"  # Set formula for second row

    return True

  except Exception as e:
    print(f"Error importing data: {e}")
    return False

def export_data_to_list_object_xlwings_claude(workbook, sheet_name, list_object_name,
                                                 data, columns, cust_name, cust_type, exp_type):
  """
  Optimized function to import data into a specified ListObject in an Excel workbook using xlwings.

  Args:
      workbook: The xlwings Workbook object.
      sheet_name: The name of the sheet containing the ListObject.
      list_object_name: The name of the ListObject.
      data: A MySQL Cursor object
      columns: A list of column names.
      cust_name: Customer name
      cust_type: Customer type
      exp_type: Export type

  Returns:
      bool: True if successful, False otherwise
  """
  try:
    # Get the sheet name
    sh_name = sheet_name if exp_type == "Recon" else f"{cust_name} Report"
    sheet = workbook.sheets[sh_name]

    # Find the list object
    list_object = sheet.api.ListObjects(list_object_name)
    if not list_object:
      raise Exception(f"ListObject '{list_object_name}' not found on sheet '{sh_name}'.")

    # Convert cursor data to list efficiently
    # This is crucial - fetch all data at once instead of iterating
    if hasattr(data, 'fetchall'):
      data_list = data.fetchall()
    elif hasattr(data, '__iter__'):
      data_list = list(data)
    else:
      raise ValueError("Data parameter must be a MySQL cursor or iterable")

    print(f"Fetched {len(data_list)} rows from cursor")  # Debug info

    # Clear existing data first
    try:
      if list_object.DataBodyRange is not None:
        list_object.DataBodyRange.ClearContents()
    except:
      pass  # Sometimes DataBodyRange doesn't exist if table is empty

    # Early return if no data
    if not data_list:
      return True

    # Calculate dimensions
    num_rows = len(data_list)
    num_cols = len(columns)

    print(f"Data dimensions: {num_rows} rows x {num_cols} columns")  # Debug info

    # Disable screen updating for better performance
    app = workbook.app
    screen_updating = app.screen_updating
    calculation = app.calculation

    app.screen_updating = False
    app.calculation = 'manual'

    try:
      # Resize the list object FIRST
      start_cell = list_object.Range.Cells(1, 1)
      # Calculate the new range: header + data rows, serial column + data columns
      total_rows = num_rows + 1  # +1 for header
      total_cols = num_cols + 1  # +1 for serial number column

      end_row = start_cell.Row + total_rows - 1
      end_col = start_cell.Column + total_cols - 1

      new_range = sheet.range((start_cell.Row, start_cell.Column),
                              (end_row, end_col)).api
      list_object.Resize(new_range)

      print(f"Resized table to: {start_cell.Row}:{end_row}, {start_cell.Column}:{end_col}")  # Debug

      # Convert data to proper format for xlwings
      if data_list and isinstance(data_list[0], tuple):
        data_array = [list(row) for row in data_list]
      else:
        data_array = data_list

      # Write data starting from row 2, column 2 (skip header row and serial column)
      data_start_row = start_cell.Row + 1  # Skip header
      data_start_col = start_cell.Column + 1  # Skip serial column

      # Use xlwings range notation
      data_range = sheet.range((data_start_row, data_start_col),
                               (data_start_row + num_rows - 1, data_start_col + num_cols - 1))

      print(
        f"Writing data to range: {data_start_row}:{data_start_row + num_rows - 1}, {data_start_col}:{data_start_col + num_cols - 1}")  # Debug

      # Write all data at once - this is the key optimization
      data_range.value = data_array

      # Add serial numbers in the first column
      serial_start_row = data_start_row
      serial_col = start_cell.Column
      serial_range = sheet.range((serial_start_row, serial_col),
                                 (serial_start_row + num_rows - 1, serial_col))

      # Create serial numbers array
      serial_numbers = [[i + 1] for i in range(num_rows)]
      serial_range.value = serial_numbers

      # Apply number formats after data is written
      table_range = list_object.Range

      # Second column as text (which is column 2 in the table, column data_start_col in the sheet)
      table_range.Columns(2).NumberFormat = "@"

      # Int code column based on customer type
      int_code_col = 8 if cust_type in ["Single Biller", "Single Biller with Adv Wallet"] else 10
      if int_code_col <= total_cols:
        table_range.Columns(int_code_col).NumberFormat = "@"

      print("Data export completed successfully")  # Debug
      return True

    finally:
      # Restore Excel settings
      app.screen_updating = screen_updating
      app.calculation = calculation

  except Exception as e:
    print(f"Error importing data: {e}")
    import traceback
    traceback.print_exc()  # This will help debug the exact issue
    return False

def export_data_to_list_object_xlwings_chatgpt(workbook, sheet_name, list_object_name, data, columns, cust_name,
                                                 cust_type, exp_type):
  """
  Imports data into a specified ListObject in an Excel workbook using xlwings.
  Optimized for performance by minimizing Excel API calls.

  Args:
      workbook: The xlwings Workbook object.
      sheet_name: The name of the sheet containing the ListObject.
      list_object_name: The name of the ListObject.
      data: A list of lists or tuples containing the data to export.
      columns: A list of column names.
      cust_name: Customer name for sheet naming.
      cust_type: Customer type for conditional formatting.
      exp_type: Export type for sheet naming.

  Returns:
      bool: True on success, False on error.
  """
  try:
    # Determine the sheet name
    if exp_type == "Recon":
      sh_name = sheet_name
    else:
      sh_name = f"{cust_name} Report"

    sheet = workbook.sheets[sh_name]

    # Get the ListObject
    try:
      list_object = sheet.api.ListObjects(list_object_name)
    except Exception:
      raise ValueError(f"ListObject '{list_object_name}' not found on sheet '{sh_name}'.")

    # Get the DataBodyRange of the list object
    # This is the most efficient way to get the data range without headers.
    data_body_range = sheet.range(list_object.DataBodyRange.Address)

    # Clear existing data in the DataBodyRange
    data_body_range.clear_contents()

    # Write data in a single, efficient operation.
    # xlwings will automatically resize the ListObject if the data has more rows.
    if data:
      data_body_range.resize(len(data), len(columns)).value = data

    # Apply number formatting in a single call for each column
    # Use a list of column numbers to apply formatting efficiently
    # Assuming the first column is for serial numbers, data starts from the second column.

    # Column 2 is always formatted as text
    col_2 = list_object.ListColumns(2).DataBodyRange
    col_2.number_format = "@"

    # Conditional formatting for the "int_code" column
    if cust_type in ["Single Biller", "Single Biller with Adv Wallet"]:
      int_code_col_num = 8
    else:
      int_code_col_num = 10

    # Get the specific column by its index within the ListObject
    int_code_col = list_object.ListColumns(int_code_col_num).DataBodyRange
    int_code_col.number_format = "@"

    # Write serial numbers in a single block
    # This is much faster than setting a formula for each cell
    if data:
      serial_numbers = [[i + 1] for i in range(len(data))]
      serial_range = list_object.ListColumns(1).DataBodyRange
      serial_range.value = serial_numbers

    return True

  except Exception as e:
    print(f"Error importing data: {e}")
    return False

def export_data_to_list_object_xlwings2(workbook, sheet_name, list_object_name, data, columns, cust_name, cust_type,
                                       exp_type):
  """
  Efficiently exports data into a specified ListObject in an Excel workbook using xlwings.
  Assumes Column 1 is for serial numbers; data starts from Column 2.
  """

  try:
    app = workbook.app
    # Turn off ScreenUpdating and DisplayAlerts for performance
    # app.screen_updating = False
    app.display_alerts = False

    if exp_type == "Recon":
      sh_name = sheet_name
    else:
      sh_name = f"{cust_name} Report"

    sheet = workbook.sheets[sh_name]

    # Find the list object
    list_object = sheet.api.ListObjects(list_object_name)
    if not list_object:
      raise Exception(f"ListObject '{list_object_name}' not found on sheet '{sh_name}'.")

    # Determine the number of rows needed for data (excluding header)
    num_data_rows = len(data)
    num_mysql_columns = len(columns)

    # Calculate the total number of columns needed in the ListObject:
    # One column for serial numbers + all MySQL data columns
    total_table_columns = 1 + num_mysql_columns

    # Clear existing data in the data body range
    if list_object.DataBodyRange and list_object.DataBodyRange.Rows.Count > 0:
      list_object.DataBodyRange.ClearContents()

    # Resize the ListObject to accommodate the new data + header row and the additional serial column
    # Get the top-left cell of the entire table range
    table_start_cell = list_object.Range.Cells(1, 1)

    # Calculate the new range for the table, including header, and the additional serial column
    new_table_end_cell_row = table_start_cell.Row + num_data_rows
    new_table_end_cell_col = table_start_cell.Column + total_table_columns - 1  # Adjust for 0-based index difference if thinking of count

    new_table_range_address = sheet.range(
      (table_start_cell.Row, table_start_cell.Column),
      (new_table_end_cell_row, new_table_end_cell_col)
    ).address

    # Resize the list object using its address
    list_object.Resize(sheet.range(new_table_range_address).api)

    # Write all MySQL data to the data body range, starting from the second column
    if num_data_rows > 0:
      # The data body range starts from the second row of the table (header is row 1)
      data_body_start_row = list_object.Range.Row + 1
      # Data should start from the second column of the table
      data_body_start_col_in_excel = list_object.Range.Column + 1

      # Define the range where ALL MySQL data will be written
      # This range will span from the second column of the table for 'num_mysql_columns' wide
      data_write_range = sheet.range(
        (data_body_start_row, data_body_start_col_in_excel),
        (data_body_start_row + num_data_rows - 1, data_body_start_col_in_excel + num_mysql_columns - 1)
      )
      data_write_range.value = data  # Write all fetched data as-is

    # Apply number format to specific columns (only once for the entire column)
    # These column numbers are 1-based relative to the start of the ListObject itself.
    table_range = list_object.Range

    # Second column of the ListObject (where the first MySQL data column is now)
    second_table_column = table_range.Columns(2)
    second_table_column.NumberFormat = "@"

    if cust_type == "Single Biller" or cust_type == "Single Biller with Adv Wallet":
      # Assuming 'int_code_column' refers to the 8th column of the *table*
      # This will now be the (8-1)th column of the MySQL data, starting from table column 2.
      # So, table column 8 is the 7th MySQL column.
      int_code_column = table_range.Columns(8)
      int_code_column.NumberFormat = "@"
    else:
      # Assuming 'int_code_column' refers to the 10th column of the *table*
      # This will now be the (10-1)th column of the MySQL data, starting from table column 2.
      # So, table column 10 is the 9th MySQL column.
      int_code_column = table_range.Columns(10)
      int_code_column.NumberFormat = "@"

    # Generate and write serial numbers in the first column of the ListObject
    if num_data_rows > 0:
      serial_numbers = [[i + 1] for i in range(num_data_rows)]
      # Get the data body range of the first column of the ListObject
      serial_column_range = list_object.ListColumns(1).DataBodyRange
      serial_column_range.value = serial_numbers

    return True

  except Exception as e:
    print(f"Error importing data: {e}")
    return False
  finally:
    # Always re-enable ScreenUpdating and DisplayAlerts
    app = workbook.app
    # app.screen_updating = True
    app.display_alerts = True


def delete_blank_or_na_rows(workbook, sheet_name, list_object_name):
  """
  Deletes rows from a ListObject (Table) in an Excel sheet if they contain blank cells or "#N/A" using filters.

  Args:
    workbook (xw.Book): The Excel workbook object.
    sheet_name (str): The name of the sheet containing the ListObject.
    list_object_name (str): The name of the ListObject (Table) to process.
  """

  try:
    # Get the specified sheet
    sheet = workbook.sheets[sheet_name]
    if not sheet:
      raise Exception(f"Sheet '{sheet_name}' not found.")

    # Find the ListObject in the specified sheet
    list_object = sheet.api.ListObjects(str(list_object_name))
    if not list_object:
      raise Exception(f"ListObject '{list_object_name}' not found on sheet '{sheet_name}'.")

    # Get the DataBodyRange of the ListObject
    data_body_range = list_object.DataBodyRange
    if not data_body_range:
      print(f"ListObject '{list_object_name}' has no data.")
      return

    # Apply filter for blank cells in the first column (adjust Field as needed)
      # Apply filter for blank cells in the first column
    data_body_range.AutoFilter(Field=2, Criteria1="=#N/A")  # Filter for blank cells

    # Get the first row and last row of the DataBodyRange
    first_row = data_body_range.Rows(1).Row
    last_row = data_body_range.Rows.Count + first_row - 1
    # print(first_row)
    # print(last_row)


    # Clear the entire range
    # sheet.range(f"{first_row}:{last_row}").api.ClearContents()
    data_body_range.ClearContents()

    # Clear filters
    # Check if AutoFilter is applied and remove it
    if list_object.AutoFilter.FilterMode:
      list_object.AutoFilter.ShowAllData()  # Show all data and remove the filter
      # print(f"AutoFilter has been removed from ListObject '{list_object_name}'.")
    else:
      print(f"No active AutoFilter found on ListObject '{list_object_name}'.")

    # Determine the last row of the DataBodyRange
    last_row = data_body_range.Rows.Count

    # Iterate through rows in reverse order
    num = 0
    for row_index in range(data_body_range.Rows.Count, 0, -1):
      row_range = data_body_range.Rows(row_index)
      # Check if the row is blank
      if all(cell.Value in [None, ""] for cell in row_range.Cells):
        row_range.Delete()  # Delete the entire row
        num = num + 1
        # print(f"Deleted blank row {row_index}.")
      else:
        # Stop when a row with data is encountered
        # print(f"Encountered a non-blank row at {row_index}. Stopping.")
        break

        # list_object.Refresh()

    print(f"Total {num} blank rows in ListObject '{list_object_name}' on sheet '{sheet_name}' have been deleted.")

  except Exception as e:
    print(f"An error occurred: {e}")



###############################################################
#### MODIFIED New Main Function to Import data from MySql #####
###############################################################
def import_mysql_to_excel_xlwings_mod(mysql_con, query, wb, sheet_name, list_object_name, cname, ctype):
  """
  Imports data from a MySQL database into a specified Excel sheet's list object.

  :param mysql_config: Dictionary containing MySQL connection parameters (host, user, password, database)
  :param query: SQL query to fetch the data
  :param sheet_name: Name of the sheet in the active workbook
  :param list_object_name: Name of the list object in the specified sheet
  """
  path_year = date.strftime("%Y")
  path_month_full = date.strftime("%B")
  path_month_abbr = date.strftime("%b")
  path_day = date.strftime("%d")

  biller_report_folder = BILLER_REPORT_BASE

  biller_report_template = rf"{cname}\{cname} Report xx-month.xlsx"
  biller_report_temp_path = os.path.join(biller_report_folder, biller_report_template)

  biller_report_file_folder = rf"{cname}\{path_year}\{path_month_abbr}"
  biller_report_folder_path = os.path.join(biller_report_folder, biller_report_file_folder)

  biller_report_file = rf"{cname}\{path_year}\{path_month_abbr}\{cname} Report {path_day}-{path_month_full}.xlsx"
  biller_report_path = os.path.join(biller_report_folder, biller_report_file)

  try:
    # Connect to the MySQL database

    # connection = mysql.connector.connect(**mysql_config)
    connection = mysql_con
    if connection.is_connected():
    # if connection:
      print("Connection Acquired")
      cursor = connection.cursor()
      cursor.execute(query)
      data = cursor.fetchall()

      # print(len(data))
    # print(data)
      columns = [desc[0] for desc in cursor.description]  # Get column names

    # Connect to the active Excel application
      app = xw.apps.active
      if not app:
        raise Exception("No active Excel application found.")

    # Get the active workbook
      workbook = assign_open_workbook(wb)
      if not workbook:
        raise Exception("No active workbook found.")

    # Get the specified sheet
      sheet = workbook.sheets[sheet_name]
      wb_rec_lo_name = get_first_listobject_name(workbook, sheet_name)
      if not sheet:
        raise Exception(f"Sheet '{sheet_name}' not found.")

      # print(len(data))
      export_data_to_list_object_xlwings(workbook, sheet_name, list_object_name, data, columns, cname,ctype,"Recon")
      # export_data_to_list_object_xlwings_chatgpt(workbook, sheet_name, list_object_name, data, columns, cname, ctype, "Recon")
      sheet.range("G2").value = date

      if len(data) > 1:
        delete_blank_or_na_rows(workbook, sheet_name, list_object_name)

      if (ctype == 'Biller With Sub-biller'):
        change_pivot_data_source(wb, sheet_name, "PivotSummary", list_object_name)
        # update_pivot_data_source(wb, sheet_name, "PivotSummary", list_object_name)

      if (ctype == 'Single Biller with Adv Wallet'):
          copy_value_between_sheets(wb, 'J15', 'J12')

      ensure_folder_exists(biller_report_folder_path)
      # print(biller_report_path)
      # biller_report_create(biller_report_temp_path,biller_report_path)
      biller_report_create2(biller_report_temp_path, biller_report_path)

      if os.path.exists(biller_report_path):
        # excel_app = Dispatch("Excel.Application")
        # excel_app.DisplayAlerts = False

        print(f"Processing {cname} Biller Report")
        wb_br = xw.Book(biller_report_path)
        # wb_br_sh = wb.sheets[f"{customer_name} Report"]
        wb_br_shname = f"{cname} Report"
        wb_br_lo_name = get_first_listobject_name(wb_br, wb_br_shname)

        export_data_to_list_object_xlwings(wb_br, wb_br_shname, wb_br_lo_name, data, columns, cname, ctype, "BillerReport")
        wb_br.sheets[wb_br_shname].range("G2").value = date


        if len(data) > 1:
          delete_blank_or_na_rows(wb_br, wb_br_shname, wb_br_lo_name)

        if (ctype == 'Biller With Sub-biller'):
          change_pivot_data_source(wb_br, wb_br_shname, "SummaryTable", wb_br_lo_name)
          # update_pivot_data_source(wb_br, wb_br_shname, "SummaryTable", wb_br_lo_name)

        wb_br.save()
        # excel_app.DisplayAlerts = True
        wb_br.close()



      print(f"Data successfully imported into sheet '{sheet_name}', list object '{list_object_name}'.")

  except Error as e:
     print(f"MySQL error: {e}")
  except Exception as ex:
     print(f"Error: {ex}")
  finally:
  #   if connection.is_connected():
      cursor.close()
  #     connection.close()
  #     print("MySQL connection closed.")


def assign_open_workbook(workbook):
  """
  Assigns an already open Excel workbook to an xlwings Workbook object.

  Args:
      workbook_name: The name of the open workbook (e.g., "Book1.xlsx").

  Returns:
      The xlwings Workbook object if found, otherwise raises an error.
  """
  #try:
  # print(os.path.basename(workbook.fullname))
  app = xw.apps.active  # Get the active app instance
  if not app:
    raise RuntimeError("No active Excel application found.")

        # Find the workbook in the list of open workbooks
  workbook_name = os.path.basename(workbook.fullname)
  for book in app.books:
    if book.name == workbook_name:
      # print(workbook_name)
      return book


  #except Exception as e:
  #  print(f"An error occurred: {e}")
  #  return None


def sheet_exists_in_open_workbook(op_workbook, sheet_name):
  """
  Checks if a sheet with the given name exists in an already opened Excel workbook.

  Args:
    workbook: The xlwings Book object representing the open workbook.
    sheet_name: The name of the sheet to check.

  Returns:
    True if the sheet exists, False otherwise.
  """

  try:
    workbook = op_workbook
    if not workbook:
      raise Exception("No active workbook found.")

    # Check if the sheet exists
    sheet_names = [sheet.name for sheet in workbook.sheets]
    # print(sheet_names)
    # return sheet_name in sheet_names # Attempt to access the sheet by name
    return sheet_name in sheet_names
  except KeyError:
    return False



########################################################
#### New Function to clear Listobject and add rows #####
########################################################
def clear_and_add_rows_to_listobject(op_workbook, sheet_name, listobject_name, num_rows):
  """
  Clears the contents of an existing Excel ListObject and adds the specified number of rows.

  Args:
    workbook_path: Path to the Excel workbook.
    sheet_name: Name of the sheet containing the ListObject.
    listobject_name: Name of the ListObject.
    num_rows: Number of rows to add.

  Raises:
    ValueError: If num_rows is not a positive integer.
    ValueError: If the ListObject is not found.
  """

  if not isinstance(num_rows, int) or num_rows <= 0:
    raise ValueError("num_rows must be a positive integer.")

  try:
    # Open the workbook
    # app = xw.App(visible=False)
    wb = op_workbook
    sheet = wb[sheet_name]

    # Find the ListObject by name
    listobject = None
    for table in sheet.range("A1").tables:
      if table.name == listobject_name:
        listobject = table
        break

    if not listobject:
      raise ValueError(f"ListObject '{listobject_name}' not found in sheet '{sheet_name}'.")

    # Clear the contents of the ListObject
    listobject.clear_contents()

    # Add the specified number of rows
    listobject.range.api.ListRows.Add(num_rows)

    # Save and close the workbook
    wb.save()

  except Exception as e:
    print(f"An error occurred: {e}")
    # Close the workbook in case of an error




# Unhide and copy template with new name
###############################################################################################################
def copy_and_rename_sheet(workbook_path, source_sheet_name, new_sheet_name):
  """
  Copies a sheet in an Excel workbook, renames it, and moves it to the last position.

  Args:
    workbook_path: Path to the Excel workbook.
    source_sheet_name: Name of the sheet to copy.
    new_sheet_name: New name for the copied sheet.
  """
  #try
  wb = assign_open_workbook(workbook_path)
  source_sheet = wb.sheets[source_sheet_name]

  # Unhide the source sheet
  source_sheet.api.Visible = True

  # Copy the sheet
  new_sheet = source_sheet.copy()

  # Rename the copied sheet
  new_sheet.name = new_sheet_name

  # Move the new sheet to the last position
  new_sheet.api.Move(After=wb.sheets[-1].api)
  source_sheet.api.Visible = False

  wb.save()
    #wb.close()

  # except Exception as e:
  #   print(f"An error occurred: {e}")


import xlwings as xw





def copy_full_pivot_table(source_workbook_names, target_workbook_path, sheet_name, start_columns):
  """
  Copies the full PivotTable (including headers, data, and footer) from a list of *open* workbooks to a target workbook,
  with dynamic starting columns and a shared sheet name using the xlwings API.
  Only processes the FIRST pivot table in the source sheet.
  """
  try:
    app = xw.apps.active
    workbooks = app.books

    for wb in workbooks:
      if wb.name in source_workbook_names:
        ws = wb.sheets[sheet_name]
        try:
          print(f"Processing workbook: {wb.name}")
          start_column = start_columns.get(wb.name, 1)

          # Access PivotTables directly
          pivot_tables = ws.api.PivotTables()
          print(f"Found {pivot_tables.Count} PivotTable(s) in '{sheet_name}' of workbook '{wb.name}'")

          if pivot_tables.Count > 0:
            pt = pivot_tables.Item(1)  # First pivot table
            print(f"Processing FIRST PivotTable: {pt.Name}")

            # Access the full TableRange2 (including headers, data, and footer)
            pivot_table_range = pt.TableRange2
            print(f"Pivot Table Range Address: {pivot_table_range.Address}")

            # Check if pivot table range is valid
            if pivot_table_range is None:
              print("TableRange2 is None, no data in PivotTable range.")
              continue

            # Get the full pivot table data (headers, data, footer)
            pivot_table_data = pivot_table_range.value

            # Open the target workbook (either from the open books or by opening it)
            try:
              target_wb = app.books[target_workbook_path.split("\\")[-1]]
              print("Using opened file")
            except:
              target_wb = app.books.open(target_workbook_path)
              print("Target workbook not opened, opening...")

            target_sht = target_wb.sheets[sheet_name]

            # Get the last row with data in the target sheet starting from the specified column
            last_row = target_sht.api.Cells(target_sht.api.Rows.Count, start_column).End(-4162).Row
            print(f"Last row in target sheet: {last_row}")

            # Paste the entire PivotTable data into the target sheet
            target_sht.range((last_row + 2, start_column)).value = pivot_table_data
            print(f"Data successfully written to target workbook '{target_workbook_path}' at column {start_column}")

            # Save the target workbook
            if target_wb != wb:
              target_wb.save()
              print(f"Target workbook '{target_workbook_path}' saved.")
            else:
              print(f"Target workbook '{target_workbook_path}' already saved.")
          else:
            print(f"No PivotTables found in '{sheet_name}' of workbook '{wb.name}'")

        except Exception as e:
          print(f"Error processing workbook {wb.name}: {e}")
  except Exception as e:
    print(f"An error occurred: {e}")


def copy_full_pivot_table2(source_workbook_names, target_workbook_path, sheet_name, start_columns):
  """
  Copies the full PivotTable (including headers, data, and footer) from a list of *open* workbooks to a target workbook,
  with dynamic starting columns and a shared sheet name using the xlwings API.
  Only processes the FIRST pivot table in the source sheet.
  """
  try:
    app = xw.apps.active
    workbooks = app.books

    for wb in workbooks:
      if wb.name in source_workbook_names:
        ws = wb.sheets[sheet_name]
        try:
          print(f"Processing workbook: {wb.name}")
          start_column = start_columns.get(wb.name, 1)

          # Access PivotTables directly
          pivot_tables = ws.api.PivotTables()
          print(f"Found {pivot_tables.Count} PivotTable(s) in '{sheet_name}' of workbook '{wb.name}'")

          if pivot_tables.Count > 0:
            pt = pivot_tables.Item(1)  # First pivot table
            print(f"Processing FIRST PivotTable: {pt.Name}")

            # Get the full pivot table range using TableRange2
            pivot_table_range = pt.TableRange1  # This includes headers, data, and footer

            # Check if pivot_table_range is a valid range object
            if pivot_table_range:
              print(f"Pivot Table Range Address: {pivot_table_range.address}")

              # Get the data from the pivot table range
              pivot_table_data = pivot_table_range.value

              if not pivot_table_data:
                print("Pivot table data is None or empty.")
                continue

              print(f"Pivot table data (range): {pivot_table_data}")

              # Open the target workbook (either from the open books or by opening it)
              try:
                target_wb = app.books[target_workbook_path.split("\\")[-1]]
                print("Using opened file")
              except:
                target_wb = app.books.open(target_workbook_path)
                print("Target workbook not opened, opening...")

              target_sht = target_wb.sheets[sheet_name]

              # Get the last row with data in the target sheet starting from the specified column
              last_row = target_sht.api.Cells(target_sht.api.Rows.Count, start_column).End(-4162).Row
              print(f"Last row in target sheet: {last_row}")

              # Paste the entire PivotTable data into the target sheet
              target_sht.range((last_row + 2, start_column)).value = pivot_table_data
              print(f"Data successfully written to target workbook '{target_workbook_path}' at column {start_column}")

              # Save the target workbook
              if target_wb != wb:
                target_wb.save()
                print(f"Target workbook '{target_workbook_path}' saved.")
              else:
                print(f"Target workbook '{target_workbook_path}' already saved.")
            else:
              print(f"Pivot table range is not valid.")
          else:
            print(f"No PivotTables found in '{sheet_name}' of workbook '{wb.name}'")

        except Exception as e:
          print(f"Error processing workbook {wb.name}: {e}")
  except Exception as e:
    print(f"An error occurred: {e}")


# from deepseek
def copy_pivot_data_from_open_workbooks_dynamic_columnDS(
        source_workbook_names,
        target_workbook_path,
        sheet_name,
        start_columns,
        table_names
):
  """
  Copies PivotTable data (as values with formatting) from open workbooks into a target workbook.
  After pasting:
    1. Clears the 3rd column of the pasted block (including header).
    2. Inserts INDEX/MATCH formulas in the cleared column using the given table name
       (table already exists in the summary workbook).

  Parameters:
      source_workbook_names (list): List of source workbook names to pull from.
      target_workbook_path (str): Path to target workbook.
      sheet_name (str): Sheet name in both source and target.
      start_columns (dict): Mapping workbook_name -> start column in summary sheet.
      table_name (str): Name of the Excel table in the summary workbook to reference in formulas.
  """
  try:
    app = xw.apps.active
    workbooks = app.books

    for wb in workbooks:
      if wb.name in source_workbook_names:
        ws = wb.sheets[sheet_name]
        try:
          print(f"Processing workbook: {wb.name}")
          start_column = start_columns.get(wb.name, 1)
          table_name = table_names.get(wb.name, 1)

          # Access PivotTables directly
          pivot_tables = ws.api.PivotTables()
          if pivot_tables.Count == 0:
            print(f"No PivotTables found in '{sheet_name}' of workbook '{wb.name}'")
            continue

          pt = pivot_tables.Item(1)  # First pivot table
          pivot_table_range = pt.TableRange2
          if pivot_table_range is None:
            print("TableRange2 is None, no data in PivotTable range.")
            continue

          # Open or attach target workbook
          try:
            target_wb = app.books[target_workbook_path.split("\\")[-1]]
            print("Using opened target workbook")
          except:
            target_wb = app.books.open(target_workbook_path)
            print("Target workbook not opened, opening...")

          target_sht = target_wb.sheets[sheet_name]

          # Get the last row with data in the target sheet starting from start_column
          last_row = target_sht.api.Cells(target_sht.api.Rows.Count, start_column).End(-4162).Row

          # Destination range
          dest_start_cell = target_sht.range((last_row + 2, start_column))
          dest_range = dest_start_cell.resize(
            pivot_table_range.Rows.Count,
            pivot_table_range.Columns.Count
          )

          # Copy pivot table data WITH formatting
          pivot_table_range.Copy()
          dest_range.api.PasteSpecial(Paste=-4122)  # Values + number formats
          dest_range.api.PasteSpecial(Paste=12)  # Formats
          app.api.CutCopyMode = False
          print(f"Pasted pivot at row {last_row + 2}, col {start_column}")

          # ---- STEP 1: Clear 3rd column (header + data) ----
          col3_range = dest_range.columns(3)
          col3_range.clear_contents()
          print("Cleared 3rd column of pasted table")

          # ---- STEP 2: Insert INDEX/MATCH formula ----
          first_row = dest_range.row
          last_row_new = dest_range.row + dest_range.rows.count - 1

          # First col of pasted block → lookup value
          col1_letter = target_sht.range((1, dest_range.column)).get_address(0, 0)[0]

          # Template: row ref will be auto-filled
          base_formula = (
            f"=INDEX({table_name}[IBAN],"
            f"MATCH({col1_letter}{first_row + 1},{table_name}[Name],0),1)"
          )

          # Write once and Excel will auto-fill down
          formula_range = target_sht.range(
            (first_row + 1, dest_range.column + 2),  # 3rd col
            (last_row_new, dest_range.column + 2)
          )
          formula_range.formula = base_formula
          print(f"Inserted INDEX/MATCH formula into 3rd column using {table_name}")

          # Save target workbook
          if target_wb != wb:
            target_wb.save()
            print(f"Target workbook '{target_workbook_path}' saved.")

        except Exception as e:
          print(f"Error processing workbook {wb.name}: {e}")

  except Exception as e:
    print(f"An error occurred: {e}")

def filter_and_delete_zero_amount_rows(workbook_name, sheet_name):
  """
  Filters a ListObject in an Excel sheet by the 2nd column (amount) for zero values,
  and then deletes the entire rows containing those zero values.

  Args:
      workbook_name (str): The name of the already open Excel workbook.
      sheet_name (str): The name of the sheet containing the ListObject.
  """
  try:
    # Connect to the already open workbook
    wb = xw.books(workbook_name)
    sheet = wb.sheets(sheet_name)

    # Access the first ListObject
    table = sheet.tables[0]

    # Determine the amount column index (2nd column, so index 1)
    amount_column_index = 1

    # Apply the filter
    table.range.api.AutoFilter(Field=amount_column_index + 1, Criteria1=0)  # field parameter is 1 based.

    # Get the filtered range (excluding headers)
    filtered_range = table.range.api.SpecialCells(12)  # 12 represents xlCellTypeVisible

    # Check if any filtered rows exist
    if filtered_range is not None:
      # Get the row numbers of the filtered rows
      filtered_rows = [area.Row for area in filtered_range.Areas]

      # Turn off filter to allow deletion
      table.range.api.AutoFilter(Field=amount_column_index + 1)

      # Delete the rows in reverse order to avoid shifting issues
      for row_num in sorted(filtered_rows, reverse=True):
        sheet.range(row_num, 1).api.EntireRow.Delete()

    else:
      print("No zero amount rows found.")

  except Exception as e:
    print(f"An error occurred: {e}")
    if 'wb' in locals():
      try:
        table.range.api.AutoFilter(Field=amount_column_index + 1)  # ensure filter is off even on error.
      except:
        pass  # if no filter was ever set, this will fail.


# Add these optimized functions to ProcessRecon.py

def batch_mysql_query(connection_pool, customers_df, trans_date):
  """
  Execute batch MySQL queries to fetch all biller data efficiently
  """
  try:
    connection = connection_pool.get_connection()
    cursor = connection.cursor()

    # Group customers by biller type to minimize queries
    single_billers = []
    multi_billers = []

    for _, row in customers_df.iterrows():
      if row['BillerType'] in ['Single Biller', 'Single Biller with Adv Wallet']:
        single_billers.append(row['CustomerName'])
      else:
        multi_billers.append(row['CustomerName'])

    all_data = {}

    # Single query for all single billers
    if single_billers:
      placeholders = ','.join(['%s'] * len(single_billers))
      single_query = f"""
                SELECT Cust, InvoiceNum, InvAmount, AmountPaid, PayDate, OpFee, PostPaidShare, InternalCode
                FROM dailyfiledto
                WHERE Cust IN ({placeholders}) AND fdate = %s
                ORDER BY Cust
            """
      cursor.execute(single_query, single_billers + [trans_date])

      # Group by customer
      current_customer = None
      for row in cursor.fetchall():
        if row[0] != current_customer:
          current_customer = row[0]
          all_data[current_customer] = {
            'data': [],
            'type': 'Single Biller',
            'columns': ['InvoiceNum', 'InvAmount', 'AmountPaid', 'PayDate', 'OpFee', 'PostPaidShare', 'InternalCode']
          }
        all_data[current_customer]['data'].append(row[1:])

    # Single query for all multi billers
    if multi_billers:
      placeholders = ','.join(['%s'] * len(multi_billers))
      multi_query = f"""
                SELECT Cust, InvoiceNum, InvAmount, AmountPaid, PayDate, OpFee, PostPaidShare, SubBillerShare, SubBillerName, InternalCode
                FROM dailyfiledto
                WHERE Cust IN ({placeholders}) AND fdate = %s
                ORDER BY Cust
            """
      cursor.execute(multi_query, multi_billers + [trans_date])

      # Group by customer
      current_customer = None
      for row in cursor.fetchall():
        if row[0] != current_customer:
          current_customer = row[0]
          all_data[current_customer] = {
            'data': [],
            'type': 'Multi Biller',
            'columns': ['InvoiceNum', 'InvAmount', 'AmountPaid', 'PayDate', 'OpFee', 'PostPaidShare', 'SubBillerShare',
                        'SubBillerName', 'InternalCode']
          }
        all_data[current_customer]['data'].append(row[1:])

    cursor.close()
    connection.close()

    print(f"Fetched data for {len(all_data)} billers in batch operation")
    return all_data

  except Exception as e:
    print(f"Error in batch MySQL query: {e}")
    return {}


def bulk_excel_operations(workbook, operations):
  """
  Perform multiple Excel operations in a single batch with minimal API calls
  """
  try:
    app = workbook.app

    # Disable all updates
    original_settings = {
      'screen_updating': app.screen_updating,
      'calculation': app.calculation,
      'display_alerts': app.display_alerts
    }

    app.screen_updating = False
    app.calculation = 'manual'
    app.display_alerts = False

    try:
      # Execute all operations
      for operation in operations:
        operation()

    finally:
      # Restore settings
      for setting, value in original_settings.items():
        setattr(app, setting, value)

  except Exception as e:
    print(f"Error in bulk Excel operations: {e}")


def optimized_table_resize_and_populate(sheet, table_name, data, serial_numbers=True):
  """
  Highly optimized table resize and population
  """
  try:
    table = sheet.api.ListObjects(table_name)

    if not data:
      return True

    num_rows = len(data)
    num_cols = len(data[0]) if data else 0

    # Calculate total columns (data + serial if needed)
    total_cols = num_cols + (1 if serial_numbers else 0)

    # Resize table in one operation
    start_cell = table.Range.Cells(1, 1)
    new_range = sheet.range(
      (start_cell.Row, start_cell.Column),
      (start_cell.Row + num_rows, start_cell.Column + total_cols - 1)
    ).api

    table.Resize(new_range)

    # Write all data at once
    data_start_col = 2 if serial_numbers else 1
    data_range = sheet.range(
      (start_cell.Row + 1, start_cell.Column + data_start_col - 1),
      (start_cell.Row + num_rows, start_cell.Column + data_start_col + num_cols - 2)
    )
    data_range.value = data

    # Add serial numbers if needed
    if serial_numbers:
      serial_range = sheet.range(
        (start_cell.Row + 1, start_cell.Column),
        (start_cell.Row + num_rows, start_cell.Column)
      )
      serial_range.value = [[i + 1] for i in range(num_rows)]

    return True

  except Exception as e:
    print(f"Error in optimized table operations: {e}")
    return False


def parallel_biller_processing(biller_data_list, max_workers=3):
  """
  Process multiple billers in parallel with controlled concurrency
  """
  results = []

  with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
    # Submit all tasks
    future_to_biller = {
      executor.submit(process_single_biller_optimized, biller_data): biller_data[0]
      for biller_data in biller_data_list
    }

    # Collect results
    for future in concurrent.futures.as_completed(future_to_biller):
      biller_name = future_to_biller[future]
      try:
        result = future.result(timeout=300)  # 5 minute timeout per biller
        results.append((biller_name, result))
        print(f"Completed processing for {biller_name}")
      except Exception as exc:
        print(f"Biller {biller_name} generated an exception: {exc}")
        results.append((biller_name, None))

  return results


def process_single_biller_optimized(biller_data):
  """
  Optimized single biller processing with minimal Excel API calls
  """
  customer_name, row_data, all_biller_data = biller_data
  biller_type = row_data['BillerType']

  try:
    path_year = date.strftime("%Y")
    path_month_full = date.strftime("%B")
    path_month_abbr = date.strftime("%b")
    path_day = date.strftime("%d")

    invoice_file_path_name = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} - {path_month_full} Internal Reconciliation Summary.xlsx"
    invoice_path = os.path.join(config.config.invoice_base, invoice_file_path_name)

    if not os.path.exists(invoice_path):
      print(f"Invoice for {customer_name} not found")
      return None

    # Use Excel application lock for thread safety
    with excel_lock:
      wb = xw.Book(invoice_path)
      today_sheet_name = f"{path_day}-{path_month_abbr}"

      if sheet_exists_in_open_workbook(wb, today_sheet_name):
        return wb

      # Batch all Excel operations
      operations = []

      # Add sheet creation operation
      operations.append(lambda: copy_and_rename_sheet(wb, "Template", today_sheet_name))

      # Get pre-fetched data
      if customer_name in all_biller_data:
        biller_info = all_biller_data[customer_name]
        data = biller_info['data']
        columns = biller_info['columns']

        # Add data export operation
        ws_lo = get_first_listobject_name(wb, today_sheet_name)
        operations.append(lambda: optimized_table_resize_and_populate(
          wb.sheets[today_sheet_name], ws_lo, data, serial_numbers=True
        ))

        # Add date setting operation
        operations.append(lambda: wb.sheets[today_sheet_name].range("G2").__setattr__('value', date))

        # Execute all operations in batch
        bulk_excel_operations(wb, operations)

        # Handle post-processing
        if len(data) > 1:
          delete_blank_or_na_rows_optimized(wb, today_sheet_name, ws_lo)

        if biller_type == 'Biller With Sub-biller':
          change_pivot_data_source_optimized(wb, today_sheet_name, "PivotSummary", ws_lo)

        if biller_type == 'Single Biller with Adv Wallet':
          copy_value_between_sheets(wb, 'J15', 'J12')

        # Process biller report asynchronously if possible
        process_biller_report_async(customer_name, biller_type, data, columns)

      return wb

  except Exception as e:
    print(f"Error processing {customer_name}: {e}")
    return None


def process_biller_report_async(customer_name, biller_type, data, columns):
  """
  Asynchronous biller report processing to avoid blocking main thread
  """
  try:
    path_year = date.strftime("%Y")
    path_month_full = date.strftime("%B")
    path_month_abbr = date.strftime("%b")
    path_day = date.strftime("%d")

    BILLER_REPORT_BASE = config.config.biller_base

    # Prepare paths
    biller_report_template = rf"{customer_name}\{customer_name} Report xx-month.xlsx"
    biller_report_temp_path = os.path.join(BILLER_REPORT_BASE, biller_report_template)

    biller_report_folder_path = os.path.join(BILLER_REPORT_BASE, rf"{customer_name}\{path_year}\{path_month_abbr}")
    biller_report_path = os.path.join(BILLER_REPORT_BASE,
                                      rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx")

    # Create folder and copy template
    ensure_folder_exists(biller_report_folder_path)
    biller_report_create2(biller_report_temp_path, biller_report_path)

    if os.path.exists(biller_report_path):
      wb_br = xw.Book(biller_report_path)
      wb_br_shname = f"{customer_name} Report"
      wb_br_lo_name = get_first_listobject_name(wb_br, wb_br_shname)

      # Batch operations for biller report
      operations = [
        lambda: optimized_table_resize_and_populate(wb_br.sheets[wb_br_shname], wb_br_lo_name, data),
        lambda: wb_br.sheets[wb_br_shname].range("G2").__setattr__('value', date)
      ]

      bulk_excel_operations(wb_br, operations)

      if len(data) > 1:
        delete_blank_or_na_rows_optimized(wb_br, wb_br_shname, wb_br_lo_name)

      if biller_type == 'Biller With Sub-biller':
        change_pivot_data_source_optimized(wb_br, wb_br_shname, "SummaryTable", wb_br_lo_name)

      wb_br.save()
      wb_br.close()

  except Exception as e:
    print(f"Error in async biller report for {customer_name}: {e}")


# Memory management and caching utilities
class ExcelOperationCache:
  """Cache frequently used Excel objects and operations"""

  def __init__(self):
    self._workbook_cache = {}
    self._sheet_cache = {}
    self._table_cache = {}

  def get_workbook(self, path):
    if path not in self._workbook_cache:
      self._workbook_cache[path] = xw.Book(path)
    return self._workbook_cache[path]

  def get_sheet(self, workbook, sheet_name):
    key = f"{workbook.name}_{sheet_name}"
    if key not in self._sheet_cache:
      self._sheet_cache[key] = workbook.sheets[sheet_name]
    return self._sheet_cache[key]

  def clear_cache(self):
    self._workbook_cache.clear()
    self._sheet_cache.clear()
    self._table_cache.clear()


# Global cache instance
excel_cache = ExcelOperationCache()


def memory_efficient_data_processing(data_chunks, chunk_size=1000):
  """
  Process data in chunks to avoid memory issues with large datasets
  """
  for i in range(0, len(data_chunks), chunk_size):
    chunk = data_chunks[i:i + chunk_size]
    yield chunk


def optimize_excel_application_settings():
  """
  Configure Excel application for maximum performance
  """
  try:
    app = xw.apps.active
    if app:
      # Performance optimizations
      app.screen_updating = False
      app.calculation = 'manual'
      app.display_alerts = False
      app.enable_events = False

      # Memory optimizations
      app.api.Application.CutCopyMode = False

      return app
  except Exception as e:
    print(f"Error optimizing Excel settings: {e}")
    return None


def restore_excel_application_settings(app, original_settings=None):
  """
  Restore Excel application to normal settings
  """
  try:
    if app:
      if original_settings:
        for setting, value in original_settings.items():
          setattr(app, setting, value)
      else:
        # Default restoration
        app.screen_updating = True
        app.calculation = 'automatic'
        app.display_alerts = True
        app.enable_events = True

  except Exception as e:
    print(f"Error restoring Excel settings: {e}")


# Database optimization utilities
def create_optimized_mysql_connection():
  """
  Create MySQL connection with performance optimizations
  """
  try:
    mysql_config_optimized = {
      "host": "localhost",
      "user": "root",
      "password": "root",
      "database": "azm",
      "autocommit": True,
      "sql_mode": "TRADITIONAL",
      "charset": "utf8mb4",
      "use_unicode": True,
      "buffered": True,
      "raw": False,
      "consume_results": True
    }

    connection = mysql.connector.connect(**mysql_config_optimized)

    # Optimize connection settings
    cursor = connection.cursor()
    cursor.execute("SET SESSION query_cache_type = ON")
    cursor.execute("SET SESSION query_cache_size = 67108864")  # 64MB
    cursor.close()

    return connection

  except Exception as e:
    print(f"Error creating optimized MySQL connection: {e}")
    return None


def delete_blank_or_na_rows_optimized(workbook, sheet_name, list_object_name):
  """
  Optimized version of delete_blank_or_na_rows with better performance
  """
  try:
    sheet = workbook.sheets[sheet_name]
    list_object = sheet.api.ListObjects(str(list_object_name))

    if not list_object:
      print(f"ListObject '{list_object_name}' not found")
      return

    data_body_range = list_object.DataBodyRange
    if not data_body_range:
      print(f"ListObject '{list_object_name}' has no data")
      return

    # Disable screen updating for performance
    app = workbook.app
    screen_updating = app.screen_updating
    calculation = app.calculation

    app.screen_updating = False
    app.calculation = 'manual'

    try:
      # Use Excel's built-in filtering for better performance
      # Filter for #N/A values in the second column (typically the amount column)
      data_body_range.AutoFilter(Field=2, Criteria1="=#N/A")

      # Get visible cells (filtered results)
      try:
        visible_cells = data_body_range.SpecialCells(12)  # xlCellTypeVisible
        if visible_cells:
          # Clear content of filtered rows
          visible_cells.ClearContents()
          print(f"Cleared #N/A rows in ListObject '{list_object_name}'")
      except:
        # No filtered rows found
        pass

      # Remove the filter
      if list_object.AutoFilter.FilterMode:
        list_object.AutoFilter.ShowAllData()

      # Alternative approach: Remove completely empty rows
      last_row = data_body_range.Rows.Count
      rows_deleted = 0

      # Iterate from bottom to top to avoid index shifting
      for row_index in range(last_row, 0, -1):
        row_range = data_body_range.Rows(row_index)

        # Check if entire row is empty
        is_empty = True
        for cell in row_range.Cells:
          if cell.Value not in [None, "", "#N/A"]:
            is_empty = False
            break

        if is_empty:
          row_range.Delete()
          rows_deleted += 1
        else:
          # Stop when we find a non-empty row (optimization)
          break

      if rows_deleted > 0:
        print(f"Deleted {rows_deleted} empty rows from ListObject '{list_object_name}'")

    finally:
      # Restore Excel settings
      app.screen_updating = screen_updating
      app.calculation = calculation

  except Exception as e:
    print(f"Error in optimized delete blank rows: {e}")


def change_pivot_data_source_optimized(workbook, sheet_name, pivot_table_name, new_data_source):
  """
  Optimized version of change_pivot_data_source with better error handling and performance
  """
  try:
    # Get the workbook object (handle both xlwings Book and string path)
    if isinstance(workbook, str):
      wb = assign_open_workbook(workbook)
    else:
      wb = workbook

    sheet = wb.sheets[sheet_name]

    # Disable updates for performance
    app = wb.app
    screen_updating = app.screen_updating
    calculation = app.calculation
    display_alerts = app.display_alerts

    app.screen_updating = False
    app.calculation = 'manual'
    app.display_alerts = False

    try:
      # Access the pivot table
      pivot_table = sheet.api.PivotTables(pivot_table_name)
      print(f"Found PivotTable '{pivot_table_name}' in sheet '{sheet_name}'")

      # Create new pivot cache with optimized settings
      new_pivot_cache = wb.api.PivotCaches().Create(
        SourceType=1,  # xlDatabase
        SourceData=new_data_source
      )

      # Update the pivot table's cache
      pivot_table.ChangePivotCache(new_pivot_cache)

      # Refresh the pivot table
      pivot_table.RefreshTable()

      print(f"Successfully updated PivotTable '{pivot_table_name}' data source to '{new_data_source}'")
      return True

    except Exception as e:
      print(f"Error updating PivotTable '{pivot_table_name}': {e}")

      # Fallback method: try direct source data update
      try:
        pivot_cache = pivot_table.PivotCache()
        pivot_cache.SourceData = new_data_source
        pivot_table.RefreshTable()
        print(f"Updated PivotTable '{pivot_table_name}' using fallback method")
        return True
      except Exception as e2:
        print(f"Fallback method also failed: {e2}")
        return False

    finally:
      # Restore Excel settings
      app.screen_updating = screen_updating
      app.calculation = calculation
      app.display_alerts = display_alerts

  except Exception as e:
    print(f"Error in optimized pivot data source change: {e}")
    return False


# Additional helper function for better pivot table handling
def refresh_all_pivot_tables_optimized(workbook, sheet_name):
  """
  Refresh all pivot tables in a sheet with optimized performance
  """
  try:
    sheet = workbook.sheets[sheet_name]

    # Disable updates
    app = workbook.app
    screen_updating = app.screen_updating
    app.screen_updating = False

    try:
      pivot_tables = sheet.api.PivotTables()

      if pivot_tables.Count > 0:
        print(f"Refreshing {pivot_tables.Count} pivot table(s) in sheet '{sheet_name}'")

        for i in range(1, pivot_tables.Count + 1):
          try:
            pt = pivot_tables.Item(i)
            pt.RefreshTable()
            print(f"Refreshed pivot table: {pt.Name}")
          except Exception as e:
            print(f"Error refreshing pivot table {i}: {e}")
      else:
        print(f"No pivot tables found in sheet '{sheet_name}'")

    finally:
      app.screen_updating = screen_updating

  except Exception as e:
    print(f"Error refreshing pivot tables: {e}")


# Enhanced error handling wrapper
def excel_operation_with_retry(operation, max_retries=3, delay=1):
  """
  Execute Excel operations with retry logic for better reliability
  """
  import time

  for attempt in range(max_retries):
    try:
      return operation()
    except Exception as e:
      print(f"Excel operation failed (attempt {attempt + 1}): {e}")
      if attempt < max_retries - 1:
        time.sleep(delay)
        print(f"Retrying in {delay} seconds...")
      else:
        print(f"Excel operation failed after {max_retries} attempts")
        raise e