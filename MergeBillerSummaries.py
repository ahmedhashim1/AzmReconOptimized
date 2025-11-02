import xlwings as xw
import pandas as pd
from pathlib import Path
import time


def connect_excel_workbook(file_name: str):
    """Connect to existing workbook or open new one."""
    file_path = Path(file_name).resolve()

    # Try to connect to already-open workbook
    for app_instance in xw.apps:
        for wb in app_instance.books:
            if wb.name.lower() == file_path.name.lower():
                print(f"✅ Connected to already-open workbook: {wb.name}")
                return app_instance, wb

    # Open new workbook if not found
    print(f"📂 Opening workbook: {file_path}")
    app = xw.App(visible=False)
    wb = app.books.open(str(file_path))
    return app, wb


def make_unique_headers(headers):
    """Make column headers unique by appending numbers to duplicates."""
    seen = {}
    unique_headers = []

    for idx, header in enumerate(headers):
        # Special handling for column P (index 15) - always name it "Amount_rec"
        if idx == 15:
            unique_headers.append("Amount_rec")
            continue

        # Handle None or empty headers
        if header is None or str(header).strip() == "":
            header = "Unnamed"
        else:
            header = str(header).strip()

        # Make unique if duplicate
        if header in seen:
            seen[header] += 1
            unique_header = f"{header}_{seen[header]}"
        else:
            seen[header] = 0
            unique_header = header

        unique_headers.append(unique_header)

    return unique_headers


def extract_data_from_sheet(sh, excluded_sheets):
    """
    Extract data from A6:P[last_row] where last_row is from used range.
    Headers are in row 6, data starts from row 7.
    """
    try:
        # Skip excluded sheets
        if sh.name in excluded_sheets:
            print(f"⏭️  Skipping excluded sheet: {sh.name}")
            return None

        # Get the FIRST used range only (stops at first empty row)
        # This prevents including non-relevant data after gaps
        used_range = sh.used_range
        if used_range is None or used_range.last_cell is None:
            print(f"ℹ️  Sheet '{sh.name}' is empty")
            return None

        # Find the actual last row by checking for first empty row or "Total" in column B
        # Start from row 7 (first data row) and find first empty cell or Total row
        last_row = 7
        for row_num in range(7, used_range.last_cell.row + 1):
            cell_value = sh.range(f"B{row_num}").value
            if cell_value is None or str(cell_value).strip() == "":
                # Found first empty row, stop here
                last_row = row_num - 1
                break
            # Check if cell contains "Total" anywhere in the text (case-insensitive)
            cell_str = str(cell_value).strip().lower()
            if "total" in cell_str:
                # Found Total row, stop before this row
                last_row = row_num - 1
                print(f"   ⚠️  Stopped at row {row_num} (found 'Total': {cell_value})")
                break
        else:
            # No empty row found, use the used range last row
            last_row = used_range.last_cell.row

        # Check if there's enough data (at least header row 6 and one data row)
        if last_row < 7:
            print(f"ℹ️  Sheet '{sh.name}' has no data (last row: {last_row})")
            return None

        # Read headers from row 6 (A6:P6)
        header_range = sh.range(f"A6:P6")
        headers = header_range.value

        # Check if headers exist
        if not headers or all(h is None for h in headers):
            print(f"⚠️  Sheet '{sh.name}' has no headers in row 6")
            return None

        # Make headers unique to avoid pandas concat error
        headers = make_unique_headers(headers)

        # Read data from row 7 to last_row (A7:P[last_row])
        data_range = sh.range(f"A7:P{last_row}")
        data = data_range.value

        # Handle single row case (xlwings returns list instead of list of lists)
        if last_row == 7:
            data = [data]

        # Check if there's actual data
        if not data or all(row is None or all(cell is None for cell in row) for row in data):
            print(f"ℹ️  Sheet '{sh.name}' has no data rows")
            return None

        # Create DataFrame
        df = pd.DataFrame(data, columns=headers)

        # Remove completely empty rows
        df = df.dropna(how='all')

        if len(df) == 0:
            print(f"ℹ️  Sheet '{sh.name}' has no valid data after removing empty rows")
            return None

        # Add sheet name as first column
        df.insert(0, "SheetName", sh.name)

        print(f"✅ Extracted {len(df)} rows from '{sh.name}' (range: A6:P{last_row})")
        return df

    except Exception as e:
        print(f"⚠️  Error processing sheet '{sh.name}': {e}")
        import traceback
        traceback.print_exc()
        return None


def make_summary_all_sequential(file_name: str):
    """
    Merge data from A6:P[last_row] from all sheets into Overall Summary.
    Headers from row 6, data from row 7 onwards.
    """
    start = time.time()

    # Connect to workbook
    app, wb = connect_excel_workbook(file_name)

    # Define excluded sheets (case-sensitive set)
    excluded = {"Template", "List Entries", "Monthly Summary", "IBANS","IBAN", "Overall Summary"}

    # Prepare destination sheet
    if "Overall Summary" not in [s.name for s in wb.sheets]:
        dest_sh = wb.sheets.add("Overall Summary")
        print("📄 Created new 'Overall Summary' sheet")
    else:
        dest_sh = wb.sheets["Overall Summary"]
        dest_sh.clear()
        print("🧹 Cleared existing 'Overall Summary' sheet")

    # Get ALL sheets first, then filter (don't filter during comprehension)
    all_sheets = list(wb.sheets)
    sheet_objs = []

    print(f"\n🔍 Filtering sheets...")
    for sh in all_sheets:
        if sh.name not in excluded:
            sheet_objs.append(sh)
        else:
            print(f"   ⏭️  Excluding: {sh.name}")

    total_sheets = len(sheet_objs)
    print(f"\n📊 Processing {total_sheets} sheets (extracting A6:P[last_row])...\n")

    # Process sheets sequentially
    all_data = []
    successful_sheets = []

    for idx, sh in enumerate(sheet_objs, 1):
        print(f"[{idx}/{total_sheets}] Processing: {sh.name}")
        df = extract_data_from_sheet(sh, excluded)
        if df is not None:
            all_data.append(df)
            successful_sheets.append(sh.name)

    # Combine and write results
    if all_data:
        print(f"\n🔄 Combining data from {len(all_data)} sheets...")

        try:
            # Concat with ignore_index to avoid index conflicts
            combined_df = pd.concat(all_data, ignore_index=True, sort=False)

            # Write to Excel starting from A1
            dest_sh.range("A1").options(index=False, header=True).value = combined_df

            # Auto-fit columns for better visibility
            try:
                dest_sh.autofit()
            except:
                pass

            print(f"\n✅ Successfully merged {len(all_data)} sheets into 'Overall Summary'")
            print(f"📈 Total rows: {len(combined_df)}")
            print(f"📊 Total columns: {len(combined_df.columns)}")
            print(f"📋 Sheets included: {', '.join(successful_sheets[:5])}" +
                  (f"... and {len(successful_sheets) - 5} more" if len(successful_sheets) > 5 else ""))

        except Exception as e:
            print(f"\n❌ Error combining data: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("\n⚠️  No data found in any sheets.")

    # Save and cleanup
    elapsed = time.time() - start
    print(f"\n⏱️  Completed in {elapsed:.2f}s")

    wb.save()
    print("💾 Workbook saved")

    # Only quit if we opened the app
    if app.visible is False:
        app.quit()
        print("🔒 Excel application closed")


if __name__ == "__main__":
    make_summary_all_sequential(
        r"All Billers Reconciliation Summary - October.xlsm"
    )