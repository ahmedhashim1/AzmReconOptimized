import pandas as pd
from datetime import date, datetime
from win32com.client import Dispatch
import os
import config

# Configuration
INVOICE_BASE = config.config.invoice_base
m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
date = datetime(m_year, m_month, m_day)
trans_date = date.strftime("%Y/%m/%d")

path_year = date.strftime("%Y")
path_month_full = date.strftime("%B")
path_month_abbr = date.strftime("%b")
path_day = date.strftime("%d")

biller_summary_name = f"All Billers Reconciliation Summary - {path_month_full}.xlsm"
today_sheet_name = f"{path_day}-{path_month_abbr}"

invoice_base_folder = INVOICE_BASE
biller_summary_path_name = rf"Billers Summary\{path_year}\{path_month_abbr}\{biller_summary_name}"
biller_summary_path = os.path.join(invoice_base_folder, biller_summary_path_name)

DAILY_FILE_BASE = config.config.dailyfile_base
file_name = config.config.dailyfile_name
bank_rep_path = rf"{DAILY_FILE_BASE}\{path_year}\{path_month_abbr}\{path_day}"


def generate_bank_report(input_file_path, sheet_name, output_file_path):
    """
    Generates a bank report from an already opened Excel file, from a specified sheet.

    Args:
        input_file_path (str): The full path to the Excel file to process.
        sheet_name (str): The name of the worksheet to process.
        output_file_path (str): The full path where the output file will be saved.
    """
    try:
        # Connect to an existing Excel application instance
        excel_app = Dispatch("Excel.Application")
        excel_app.Visible = True

        # Suppress Excel's warning dialog boxes
        excel_app.DisplayAlerts = False

        # Assume the file is already opened and get the workbook and worksheet
        input_file_name = os.path.basename(input_file_path)
        workbook = excel_app.Workbooks(input_file_name)
        ws = workbook.Worksheets(sheet_name)

        # --- Extract data from the main ListObject ---
        list_object = ws.ListObjects(1)
        table_range = list_object.DataBodyRange
        df_main = pd.DataFrame(table_range.Value)

        list_object_rows = list_object.DataBodyRange.Rows.Count

        biller_names = df_main.iloc[:, 0].values
        main_amounts = df_main.iloc[:, 6].values

        # --- Extract data from the parallel 'Total' column (Column P) ---
        start_row = table_range.Row
        col_p_index = 16

        total_range = ws.Range(ws.Cells(start_row, col_p_index),
                               ws.Cells(start_row + list_object_rows - 1, col_p_index))
        total_amounts = [row[0] for row in total_range.Value]

        # --- Combine and process the dataframes ---
        data_dict = {
            "Billers": biller_names,
            "Main Billers Amount": main_amounts,
            "Sub Billers Amount": [''] * len(main_amounts),
            "Total": total_amounts
        }

        df_report = pd.DataFrame(data_dict)

        df_report.loc[len(df_report)] = ['', '', '', '']
        df_report.loc[len(df_report)] = ['Total', '', '', '']

        # --- Finalizing the new sheet and writing data ---
        new_workbook = excel_app.Workbooks.Add()
        new_sheet = new_workbook.Worksheets("Sheet1")

        rows, cols = df_report.shape
        new_sheet.Range(new_sheet.Cells(2, 1), new_sheet.Cells(1 + rows, cols)).Value = df_report.values.tolist()

        for i, header in enumerate(df_report.columns):
            new_sheet.Cells(1, i + 1).Value = header

        last_row = new_sheet.UsedRange.Rows.Count
        data_rows_end = last_row - 2

        if last_row > 1:
            for i in range(2, data_rows_end + 1):
                new_sheet.Cells(i, 3).Formula = f"=D{i}-B{i}"
                # Highlight rows where Column C has a value
                if new_sheet.Cells(i, 3).Value is not None and new_sheet.Cells(i, 3).Value != 0:
                    new_sheet.Range(f"A{i}:D{i}").Interior.Color = 13434879  # Light Yellow (RGB: 255, 255, 204)
                    # For light green, use 13434876 (RGB: 204, 255, 204)

            total_row_index = last_row
            new_sheet.Cells(total_row_index, 2).Formula = f"=SUM(B2:B{data_rows_end})"
            new_sheet.Cells(total_row_index, 3).Formula = f"=SUM(C2:C{data_rows_end})"
            new_sheet.Cells(total_row_index, 4).Formula = f"=SUM(D2:D{data_rows_end})"

            new_sheet.Range(f"A{total_row_index}:D{total_row_index}").Font.Bold = True

        # --- Apply formatting ---
        # Make headers bold and centered
        header_range = new_sheet.Range(f"A1:D1")
        header_range.Font.Bold = True
        header_range.HorizontalAlignment = -4108  # xlCenter

        new_sheet.Range("A:D").Columns.ColumnWidth = 20

        # Apply borders to the main data table (excluding the blank row and total row)
        data_range = new_sheet.Range(f"A1:D{data_rows_end}")
        data_range.Borders.LineStyle = 1

        # Apply the same borders to the total row
        total_row_range = new_sheet.Range(f"A{total_row_index}:D{total_row_index}")
        total_row_range.Borders.LineStyle = 1

        new_sheet.Range("B:D").NumberFormat = '_("SAR"* #,##0.00_);_("SAR"* (#,##0.00);_("SAR"* "-"??_);_(@_)'

        print("Bank report successfully generated!")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if 'new_workbook' in locals():
            new_workbook.SaveAs(output_file_path)
            new_workbook.Close(SaveChanges=True)

        excel_app.DisplayAlerts = True

        print("Original file remains open.")


# Example usage:
if __name__ == "__main__":
    file_path = rf"{biller_summary_path}"
    sheet_to_use = rf"{today_sheet_name}"

    os.makedirs(bank_rep_path, exist_ok=True)

    output_file = os.path.join(bank_rep_path, f"B2B_Transfers_{date.today().strftime('%Y%m%d')}.xlsx")

    generate_bank_report(file_path, sheet_to_use, output_file)