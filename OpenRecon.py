import os
import pandas as pd
import config
import ProcessRecon
import xlwings as xw
import time
import win32com.client as win32
import mysql.connector
from ProcessRecon import (ensure_folder_exists, copy_full_pivot_table, copy_full_pivot_table2,
                          copy_pivot_data_from_open_workbooks_dynamic_columnDS, biller_report_create2,
                          copy_value_between_sheets,
                          get_first_listobject_name, sheet_exists_in_open_workbook, copy_and_rename_sheet,
                          assign_open_workbook,
                          import_mysql_to_excel_xlwings_mod, export_data_to_list_object_xlwings,
                          filter_and_delete_zero_amount_rows)
import datetime
from xlwings import Book, Sheet, Range
from pathlib import Path
from win32com.client import Dispatch
from mysql.connector import Error
import concurrent.futures
from threading import Lock

# Configuration
INVOICE_BASE = config.config.invoice_base
m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
date = datetime.datetime(m_year, m_month, m_day)
trans_date = date.strftime("%Y/%m/%d")

mysql_config = {
    "host": "localhost",
    "user": "root",
    "password": "root",
    "database": "azm"
}

# Global lock for Excel operations
excel_lock = Lock()


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

def safe_sheet_name(name: str) -> str:
    """Make Excel sheet name safe and unique (<=31 chars, no invalid chars)."""
    invalid = ['\\', '/', '?', '*', ':', '[', ']']
    for ch in invalid:
        name = name.replace(ch, '_')
    return name[:31]

def fetch_all_biller_data(customers_df, connection_pool):
    """Fetch all biller data in a single database operation"""
    try:
        connection = connection_pool.get_connection()
        cursor = connection.cursor()

        # Build a single query for all billers
        all_data = {}

        # Group billers by type for efficient querying
        single_billers = []
        multi_billers = []

        for _, row in customers_df.iterrows():
            customer_name = row['CustomerName']
            biller_type = row['BillerType']

            if biller_type in ['Single Biller', 'Single Biller with Adv Wallet']:
                single_billers.append(customer_name)
            else:
                multi_billers.append(customer_name)

        # Execute batch queries
        if single_billers:
            placeholders = ','.join(['%s'] * len(single_billers))
            sql_single = f"""
                SELECT
                      Cust,
                      CAST(InvoiceNum AS CHAR) AS InvoiceNum,
                      InvAmount,
                      AmountPaid,
                      PayDate,
                      OpFee,
                      PostPaidShare,
                      CAST(InternalCode AS CHAR) AS InternalCode
                    FROM dailyfiledto
                    WHERE Cust IN ({placeholders}) AND fdate = %s
            """
            cursor.execute(sql_single, single_billers + [trans_date])
            # single_results = cursor.fetchall()
            data = cursor.fetchall()
            columns = ["Cust","InvoiceNum", "InvAmount", "AmountPaid", "PayDate", "OpFee", "PostPaidShare", "InternalCode"]
            single_results = normalize_dataframe(data, columns, cust_type="Single Biller")

            # Group results by customer
            for row in single_results:
                cust = row[0]
                if cust not in all_data:
                    all_data[cust] = {'data': [], 'type': 'Single Biller'}
                all_data[cust]['data'].append(row[1:])  # Exclude customer name

        if multi_billers:
            placeholders = ','.join(['%s'] * len(multi_billers))
            sql_multi = f"""
                SELECT
                      Cust,
                      CAST(InvoiceNum AS CHAR) AS InvoiceNum,
                      InvAmount,
                      AmountPaid,
                      PayDate,
                      OpFee,
                      PostPaidShare,
                      SubBillerShare,
                      SubBillerName,
                      CAST(InternalCode AS CHAR) AS InternalCode
                    FROM dailyfiledto
                    WHERE Cust IN ({placeholders}) AND fdate = %s
            """
            cursor.execute(sql_multi, multi_billers + [trans_date])
            # multi_results = cursor.fetchall()
            data = cursor.fetchall()
            columns = ["Cust","InvoiceNum", "InvAmount", "AmountPaid", "PayDate", "OpFee", "PostPaidShare",
                       "SubBillerShare", "SubBillerName", "InternalCode"]
            multi_results = normalize_dataframe(data, columns, cust_type="Biller With Sub-biller")

            # Group results by customer
            for row in multi_results:
                cust = row[0]
                if cust not in all_data:
                    all_data[cust] = {'data': [], 'type': 'Multi Biller'}
                all_data[cust]['data'].append(row[1:])  # Exclude customer name

        ### FOR DEBUGGING SQL WITH EXCEL ROWS
        # print(f"[DEBUG] Thiqah collected rows: {len(all_data.get('Thiqah', {}).get('data', []))}")
        # print(f"[DEBUG] ThiqaNafith collected rows: {len(all_data.get('ThiqaNafith', {}).get('data', []))}")

        cursor.close()
        connection.close()

        return all_data

    except Exception as e:
        print(f"Error fetching batch data: {e}")
        return {}


def normalize_dataframe(rows, columns, cust_type):
    """
    Convert fetched MySQL rows into a DataFrame and force specific columns to string.

    Args:
        rows: list of tuples (from cursor.fetchall())
        columns: list of column names
        cust_type: type of customer (decides InternalCode position: 8 vs 10)
    Returns:
        list-of-lists ready for Excel
    """
    if not rows:
        return []

    # Build DataFrame
    df = pd.DataFrame(rows, columns=columns)

    # Always force InvoiceNum to string
    if "InvoiceNum" in df.columns:
        df["InvoiceNum"] = df["InvoiceNum"].astype(str)

    # Force InternalCode (only if present)
    if "InternalCode" in df.columns:
        df["InternalCode"] = df["InternalCode"].astype(str)

    # Convert NaNs to empty string
    df = df.fillna("")

    # Return as list-of-lists
    return df.values.tolist()

def process_single_biller(biller_data, connection_pool):
    """Process a single biller's data - designed for parallel execution"""
    customer_name, row_data, all_biller_data = biller_data
    biller_type = row_data['BillerType']

    try:
        path_year = date.strftime("%Y")
        path_month_full = date.strftime("%B")
        path_month_abbr = date.strftime("%b")
        path_day = date.strftime("%d")

        invoice_file_path_name = rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} - {path_month_full} Internal Reconciliation Summary.xlsx"
        invoice_path = os.path.join(INVOICE_BASE, invoice_file_path_name)
        file_path = Path(invoice_path)

        if not os.path.exists(file_path):
            print(f"Invoice for {customer_name} not found at {invoice_path}")
            return None

        print(f"Processing Reconciliation for {customer_name}")

        # Excel operations need to be synchronized
        with excel_lock:
            wb = xw.Book(file_path)
            wb.visible = True
            wb.activate()
            today_sheet_name = f"{path_day}-{path_month_abbr}"

            if sheet_exists_in_open_workbook(wb, today_sheet_name):
                print(f"Sheet '{today_sheet_name}' already exists for {customer_name}")
                return wb

            copy_and_rename_sheet(wb, "Template", today_sheet_name)

            # Get data from pre-fetched results
            if customer_name in all_biller_data:
                data = all_biller_data[customer_name]['data']
                biller_type_from_data = all_biller_data[customer_name]['type']

                # Define columns based on biller type
                if biller_type_from_data == 'Single Biller':
                    columns = ['InvoiceNum', 'InvAmount', 'AmountPaid', 'PayDate', 'OpFee', 'PostPaidShare',
                               'InternalCode']
                else:
                    columns = ['InvoiceNum', 'InvAmount', 'AmountPaid', 'PayDate', 'OpFee', 'PostPaidShare',
                               'SubBillerShare', 'SubBillerName', 'InternalCode']

                ws_lo = get_first_listobject_name(wb, today_sheet_name)

                # Use optimized data export function
                success = export_data_to_list_object_xlwings_optimized(
                    wb, today_sheet_name, ws_lo, data, columns, customer_name, biller_type, "Recon"
                )

                if success:
                    wb.sheets[today_sheet_name].range("G2").value = date

                    if len(data) > 1:
                        delete_blank_or_na_rows_optimized(wb, today_sheet_name, ws_lo)


                    if biller_type == 'Biller With Sub-biller':
                        change_pivot_data_source_optimized(wb, today_sheet_name, "PivotSummary", ws_lo)

                    if biller_type == 'Single Biller with Adv Wallet':
                        copy_value_between_sheets(wb, 'J15', 'J12')

                    # Process biller report
                    process_biller_report_optimized(customer_name, biller_type, data, columns, connection_pool)

            return wb

    except Exception as e:
        print(f"Error processing {customer_name}: {e}")
        return None


def export_data_to_list_object_xlwings_optimized(workbook, sheet_name, list_object_name,
                                                 data, columns, cust_name, cust_type, exp_type):
    """
    Fast exporter for Excel ListObject:
      - Expands/shrinks table body using block Range.Insert/Delete
      - Preserves totals row exactly
      - Copies formatting from an existing body row (not header/totals)
      - Clears and rewrites only DataBodyRange
      - Pre-converts required columns (2, 8/10) to text before writing
      - Applies text number format to those columns in DataBody
    """
    import win32com.client

    try:
        # --- Normalize data ---
        if hasattr(data, "fetchall"):
            data = data.fetchall()
        if data and isinstance(data[0], tuple):
            data = [list(r) for r in data]
        if data is None:
            data = []

        num_rows = len(data)
        num_data_cols = len(columns) if columns else (len(data[0]) if data else 0)

        # --- Get sheet + COM ListObject ---
        sh_name = sheet_name if exp_type == "Recon" else f"{cust_name} Report"
        sheet = workbook.sheets[sh_name]
        ws = sheet.api
        table = ws.ListObjects(list_object_name)
        if not table:
            raise Exception(f"ListObject '{list_object_name}' not found on sheet '{sh_name}'.")

        # --- Detect totals row ---
        totals_visible = False
        try:
            totals_visible = bool(table.ShowTotals)
        except Exception:
            pass

        # --- Current body rows ---
        header_row = int(table.HeaderRowRange.Row)
        start_col = int(table.Range.Column)
        last_col = int(table.Range.Column + table.Range.Columns.Count - 1)

        if totals_visible:
            try:
                totals_row = int(table.TotalsRowRange.Row)
                body_end_row = totals_row - 1
            except Exception:
                body_end_row = int(table.Range.Row + table.Range.Rows.Count - 1) - 1
        else:
            body_end_row = int(table.Range.Row + table.Range.Rows.Count - 1)

        body_start_row = header_row + 1
        current_body_rows = max(0, body_end_row - body_start_row + 1)

        # --- Expand rows if needed ---
        if num_rows > current_body_rows:
            need = num_rows - current_body_rows
            print(f"[export-fast] Inserting {need} rows into '{list_object_name}'")

            # Format source: first body row if exists, else header
            if current_body_rows > 0 and table.DataBodyRange is not None:
                format_source = table.DataBodyRange.Rows(1)
            else:
                format_source = table.HeaderRowRange
                print(f"[export-fast] WARNING: no body row found, using header format as fallback.")

            insert_at_row = body_end_row + 1
            insert_range = ws.Range(
                ws.Cells(insert_at_row, start_col),
                ws.Cells(insert_at_row + need - 1, last_col)
            )
            insert_range.Insert(Shift=-4121)  # xlShiftDown

            # Resize table
            new_range = ws.Range(
                table.Range.Cells(1, 1),
                ws.Cells(body_end_row + need, last_col)
            )
            table.Resize(new_range)

            # Reapply formatting
            try:
                new_rows_range = ws.Range(
                    ws.Cells(insert_at_row, start_col),
                    ws.Cells(insert_at_row + need - 1, last_col)
                )
                format_source.Copy()
                new_rows_range.PasteSpecial(Paste=-4122)  # xlPasteFormats
                ws.Application.CutCopyMode = False
            except Exception as e:
                print(f"[export-fast] Warning: failed to reapply formats: {e}")

        # --- Shrink rows if needed ---
        elif num_rows < current_body_rows:
            to_remove = current_body_rows - num_rows
            if to_remove > 0:
                print(f"[export-fast] Deleting {to_remove} rows from '{list_object_name}'")
                delete_start = body_start_row + num_rows
                delete_end = body_end_row
                delete_range = ws.Range(ws.Cells(delete_start, start_col),
                                        ws.Cells(delete_end, last_col))
                delete_range.Delete(Shift=-4162)  # xlShiftUp
                table = ws.ListObjects(list_object_name)  # refresh ref

        # --- Clear DataBody contents ---
        if table.DataBodyRange is not None:
            sheet.range(table.DataBodyRange.Address).clear_contents()

        # --- Apply text format BEFORE writing ---
        if num_rows > 0 and num_data_cols > 0:
            body_range = table.DataBodyRange
            db_start_row = int(body_range.Row)
            db_start_col = int(body_range.Column)

            text_cols = [2]  # InvoiceNum always 2nd
            if cust_type in ["Single Biller", "Single Biller with Adv Wallet"]:
                text_cols.append(8)
            else:  # Biller With Sub-biller
                text_cols.append(10)

            for col_index in text_cols:
                if col_index <= (1 + num_data_cols):
                    sheet_col = db_start_col + (col_index - 1)
                    rng = ws.Range(ws.Cells(db_start_row, sheet_col),
                                   ws.Cells(db_start_row + num_rows - 1, sheet_col))
                    rng.NumberFormat = "@"

        # --- Write serials + data ---
        if num_rows > 0 and num_data_cols > 0:
            body_range = table.DataBodyRange
            db_start_row = int(body_range.Row)
            db_start_col = int(body_range.Column)

            # Serial numbers
            serial_range = sheet.range((db_start_row, db_start_col),
                                       (db_start_row + num_rows - 1, db_start_col))
            serial_range.value = [[i + 1] for i in range(num_rows)]

            # Data values
            data_range = sheet.range((db_start_row, db_start_col + 1),
                                     (db_start_row + num_rows - 1, db_start_col + num_data_cols))

            print(f"[DEBUG] Exporting {num_rows} rows for {cust_name} ({list_object_name})")
            data_range.value = data

        print(f"[export-fast] Exported {num_rows} rows to '{list_object_name}' on '{sh_name}'")
        return True

    except Exception as exc:
        import traceback
        print(f"❌ export_data_to_listobject_fast failed: {exc}")
        traceback.print_exc()
        return False


def process_biller_report_optimized(customer_name, biller_type, data, columns, connection_pool):
    """Optimized biller report processing"""
    try:
        path_year = date.strftime("%Y")
        path_month_full = date.strftime("%B")
        path_month_abbr = date.strftime("%b")
        path_day = date.strftime("%d")

        BILLER_REPORT_BASE = config.config.biller_base

        biller_report_template = rf"{customer_name}\{customer_name} Report xx-month.xlsx"
        biller_report_temp_path = os.path.join(BILLER_REPORT_BASE, biller_report_template)

        biller_report_folder_path = os.path.join(BILLER_REPORT_BASE, rf"{customer_name}\{path_year}\{path_month_abbr}")
        biller_report_path = os.path.join(BILLER_REPORT_BASE,
                                          rf"{customer_name}\{path_year}\{path_month_abbr}\{customer_name} Report {path_day}-{path_month_full}.xlsx")

        ensure_folder_exists(biller_report_folder_path)
        biller_report_create2(biller_report_temp_path, biller_report_path)

        if os.path.exists(biller_report_path):
            wb_br = xw.Book(biller_report_path)
            wb_br_shname = f"{customer_name} Report"
            wb_br_lo_name = get_first_listobject_name(wb_br, wb_br_shname)

            export_data_to_list_object_xlwings_optimized(
                wb_br, wb_br_shname, wb_br_lo_name, data, columns, customer_name, biller_type, "BillerReport"
            )

            wb_br.sheets[wb_br_shname].range("G2").value = date

            if len(data) > 1:
                delete_blank_or_na_rows_optimized(wb_br, wb_br_shname, wb_br_lo_name)


            if biller_type == 'Biller With Sub-biller':
                change_pivot_data_source_optimized(wb_br, wb_br_shname, "SummaryTable", wb_br_lo_name)

            wb_br.save()
            wb_br.close()

    except Exception as e:
        print(f"Error processing biller report for {customer_name}: {e}")


def delete_blank_or_na_rows_optimized(workbook, sheet_name, list_object_name):
    """Deletes only rows in a ListObject where ALL cells are blank or '#N/A'."""
    try:
        sheet = workbook.sheets[sheet_name]
        list_object = sheet.api.ListObjects(str(list_object_name))
        data_body = list_object.DataBodyRange
        if not data_body:
            print(f"ListObject '{list_object_name}' has no data.")
            return

        # Read values as a 2D list via xlwings (not .api)
        values = sheet.range(data_body.Address).value
        if not isinstance(values, list):  # Single-row tables can return a flat list
            values = [values]

        rows_to_delete = []
        for i, row in enumerate(values):
            row_list = row if isinstance(row, list) else [row]
            # ✅ delete only if all cells are empty or '#N/A'
            if all(v in (None, "", "#N/A") for v in row_list):
                rows_to_delete.append(i + 1)

        for idx in reversed(rows_to_delete):
            list_object.ListRows(idx).Delete()

        print(f"Deleted {len(rows_to_delete)} completely blank rows from '{list_object_name}'.")
    except Exception as e:
        print(f"An error occurred (delete_blank_or_na_rows_optimized): {e}")


def change_pivot_data_source_optimized(workbook, sheet_name, pivot_table_name, new_data_source):
    """Optimized pivot table data source change"""
    try:
        wb = assign_open_workbook(workbook)
        sheet = wb.sheets[sheet_name]

        pivot_table = sheet.api.PivotTables(pivot_table_name)

        # Update pivot cache efficiently
        pivot_table.ChangePivotCache(
            wb.api.PivotCaches().Create(
                SourceType=1,  # xlDatabase
                SourceData=new_data_source
            )
        )

        pivot_table.RefreshTable()
        print(f"PivotTable '{pivot_table_name}' updated successfully")

    except Exception as e:
        print(f"Error updating pivot table: {e}")


def OpenReconFiles():
    """Main optimized function"""
    excel_app = Dispatch("Excel.Application")
    excel_app.DisplayAlerts = False

    # Disable screen updating globally
    excel_app.ScreenUpdating = False
    excel_app.Calculation = -4135  # xlCalculationManual

    try:
        # File paths
        path_year = date.strftime("%Y")
        path_month_full = date.strftime("%B")
        path_month_abbr = date.strftime("%b")
        path_day = date.strftime("%d")

        DAILY_FILE_BASE = config.config.dailyfile_base
        file_name = config.config.dailyfile_name
        customers_file = rf"{DAILY_FILE_BASE}\{path_year}\{path_month_abbr}\{path_day}\{file_name}"

        # Create connection pool
        connection_pool = get_mysql_connection_pool()
        if not connection_pool:
            print("Failed to create connection pool")
            return

        # Read customer data
        try:
            customers_df = pd.read_excel(customers_file)
            print(f"Processing {len(customers_df)} customers")
        except Exception as e:
            print(f"Error reading customer list: {e}")
            return

        # Fetch all data at once
        print("Fetching all biller data...")
        all_biller_data = fetch_all_biller_data(customers_df, connection_pool)

        # Process billers in parallel (limited concurrency for Excel stability)
        print("Processing billers...")
        biller_tasks = []

        for index, row in customers_df.iterrows():
            customer_name = row['CustomerName']
            biller_tasks.append((customer_name, row, all_biller_data))

        # Use ThreadPoolExecutor with limited workers for Excel stability
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            future_to_biller = {
                executor.submit(process_single_biller, task, connection_pool): task[0]
                for task in biller_tasks
            }

            processed_workbooks = []
            for future in concurrent.futures.as_completed(future_to_biller):
                biller_name = future_to_biller[future]
                try:
                    wb = future.result()
                    if wb:
                        processed_workbooks.append(wb)
                except Exception as exc:
                    print(f"Biller {biller_name} generated an exception: {exc}")

        # Process biller summary
        process_biller_summary_optimized(path_year, path_month_full, path_month_abbr, path_day)

    finally:
        # Re-enable Excel features
        excel_app.ScreenUpdating = True
        excel_app.Calculation = -4105  # xlCalculationAutomatic
        excel_app.DisplayAlerts = True


def delete_blank_or_zero_from_listobject_open(file_ref, sheet_name, table_name):
    """
    Deletes entire worksheet rows where the 2nd column of a ListObject
    is blank, zero (numeric or text), or shows '-' due to Accounting formatting.
    Works on already open workbook.
    """
    try:
        # --- Normalize workbook reference ---
        book = None
        if isinstance(file_ref, xw.main.Book):
            book = file_ref
        else:
            fname = str(file_ref).strip()
            for app in xw.apps:
                for b in app.books:
                    if b.name == fname or b.fullname.endswith(fname):
                        book = b
                        break
                if book:
                    break

        if book is None:
            raise RuntimeError(f"Workbook '{file_ref}' is not open.")

        ws = book.sheets[sheet_name]
        lo = ws.api.ListObjects(table_name)

        if lo.DataBodyRange is None:
            print(f"No data in ListObject '{table_name}' on '{sheet_name}'")
            return

        # Get DataBodyRange values as 2D list
        rng = ws.range(lo.DataBodyRange.Address)
        values = rng.value

        # Ensure always list-of-lists
        if not isinstance(values[0], list):
            values = [values]

        rows_to_delete = []
        for i, row in enumerate(values, start=1):
            val = row[1]  # second column
            if (
                val is None or val == "" or val == "0" or
                (isinstance(val, (int, float)) and abs(val) < 1e-9)
            ):
                # mark actual row number in sheet
                rows_to_delete.append(lo.DataBodyRange.Row + i - 1)

        # Delete bottom-to-top to avoid shifting
        for row_idx in sorted(rows_to_delete, reverse=True):
            ws.api.Rows(row_idx).Delete()

        try:
            if ws.api.FilterMode:
                ws.api.ShowAllData()
        except Exception:
            try:
                lo.AutoFilter.ShowAllData()
            except Exception:
                pass

        print(f"✅ Deleted {len(rows_to_delete)} rows from '{table_name}' in '{sheet_name}' of '{book.name}'")

    except Exception as e:
        print(f"❌ Error cleaning rows: {e}")




def process_biller_summary_optimized(path_year, path_month_full, path_month_abbr, path_day):
    """Optimized biller summary processing"""
    try:
        invoice_base_folder = INVOICE_BASE
        biller_summary_name = f"All Billers Reconciliation Summary - {path_month_full}.xlsm"
        today_sheet_name = f"{path_day}-{path_month_abbr}"

        biller_summary_path_name = rf"Billers Summary\{path_year}\{path_month_abbr}\{biller_summary_name}"
        biller_summary_path = os.path.join(invoice_base_folder, biller_summary_path_name)

        if not os.path.exists(biller_summary_path):
            print("Biller Summary not available")
            return

        wb_bs = xw.Book(biller_summary_path)

        if sheet_exists_in_open_workbook(wb_bs, today_sheet_name):
            print(f"Sheet '{today_sheet_name}' already exists in summary")
            return

        copy_and_rename_sheet(wb_bs, "Template", today_sheet_name)
        wb_bs.sheets[today_sheet_name].range("A7").value = date

        # Define source workbooks and columns
        source_workbook_names = [
            rf"Bcare - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"Damin - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"Thiqah - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"ThiqaNafith - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"Asnad - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"Tatbeeq - {path_month_full} Internal Reconciliation Summary.xlsx",
            rf"TameeniElectronic - {path_month_full} Internal Reconciliation Summary.xlsx",
        ]

        start_columns = {
            rf"Bcare - {path_month_full} Internal Reconciliation Summary.xlsx": 1,
            rf"Damin - {path_month_full} Internal Reconciliation Summary.xlsx": 4,
            rf"Thiqah - {path_month_full} Internal Reconciliation Summary.xlsx": 4,
            rf"ThiqaNafith - {path_month_full} Internal Reconciliation Summary.xlsx": 8,
            rf"Asnad - {path_month_full} Internal Reconciliation Summary.xlsx": 8,
            rf"Tatbeeq - {path_month_full} Internal Reconciliation Summary.xlsx": 8,
            rf"TameeniElectronic - {path_month_full} Internal Reconciliation Summary.xlsx": 11,
        }

        table_names = {
            rf"Bcare - {path_month_full} Internal Reconciliation Summary.xlsx": "Table7",
            rf"Damin - {path_month_full} Internal Reconciliation Summary.xlsx": "Table8",
            rf"Thiqah - {path_month_full} Internal Reconciliation Summary.xlsx": "Table15",
            rf"Tatbeeq - {path_month_full} Internal Reconciliation Summary.xlsx": "Table21",
            rf"TameeniElectronic - {path_month_full} Internal Reconciliation Summary.xlsx": "Table16",
        }

        copy_pivot_data_from_open_workbooks_dynamic_columnDS(
            source_workbook_names, wb_bs.name, today_sheet_name, start_columns,table_names
        )

        curr_lo = get_first_listobject_name(wb_bs, today_sheet_name)
        delete_blank_or_zero_from_listobject_open(wb_bs,today_sheet_name,curr_lo)

    except Exception as e:
        print(f"Error processing biller summary: {e}")

def measure_execution_time(func):
  """
  Decorator to measure the execution time of a function.

  Args:
    func: The function to be timed.

  Returns:
    The decorated function.
  """
  def wrapper(*args, **kwargs):
    start_time = time.time()
    result = func(*args, **kwargs)
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Execution time of {func.__name__}: {execution_time:.4f} seconds")
    return result
  return wrapper

@measure_execution_time
def main():
    OpenReconFiles()
    for wb in xw.apps.active.books:
        if not wb.name.lower().endswith("personal.xlsb"):
            wb.app.api.Windows(wb.name).Activate()


if __name__ == "__main__":
    main()