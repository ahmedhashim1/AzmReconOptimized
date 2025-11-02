import xlwings as xw
import pandas as pd
import numpy as np

MAIN_COLUMNS = [
    "Date", "Biller Name", "Total Amount Paid", "Total Amount received (bank)",
    "Total Amount (paid-Sadad fees)", "Difference (C-D)", "Bank Transfer Charge", "Amount transfer to BILLER",
    "Sadad Fees", "Azm Fees", "Total Fees", "Number of Bills", "Matched", "Status"
]


def find_main_table_end_column(df, header_row_idx):
    """Find where the main table ends by looking for the last main table column"""
    row_values = df.iloc[header_row_idx].values
    last_main_col = 0

    for i, val in enumerate(row_values[:25]):  # Check first 25 columns
        if pd.notna(val) and str(val).strip():
            val_str = str(val).strip()
            # Check if this looks like a main table column
            if any(main_col.lower() in val_str.lower() for main_col in MAIN_COLUMNS):
                last_main_col = i

    return last_main_col


def find_parallel_table_start(df, header_row_idx, main_table_end):
    """Find where parallel table starts - look for content after gap"""
    row_values = df.iloc[header_row_idx].values

    # Look for content after the main table with at least 1-2 column gap
    for i in range(main_table_end + 2, min(len(row_values), main_table_end + 10)):
        if pd.notna(row_values[i]) and str(row_values[i]).strip():
            return i

    return None


def find_data_end_row(df, header_row_idx):
    """Find where the actual data ends (before company summaries)"""
    end_row = header_row_idx + 1
    max_rows_to_check = min(len(df), header_row_idx + 200)

    for i in range(header_row_idx + 1, max_rows_to_check):
        # Check first few columns for stop patterns
        first_cells = [str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                       for j in range(min(5, len(df.columns)))]

        # Stop if we find company summary or other irrelevant data
        if any(pattern in cell for cell in first_cells
               for pattern in ["Company Name", "Sum of حصة", "Sum of", "Allied Cooperative"]):
            break

        # Check if row has any meaningful data
        row_data = df.iloc[i, :20]  # Check first 20 columns
        non_empty = sum(1 for val in row_data if pd.notna(val) and str(val).strip() != "")

        if non_empty > 0:
            end_row = i + 1
        elif non_empty == 0:
            # If we hit 2 consecutive empty rows, probably end of data
            next_row_empty = True
            if i + 1 < len(df):
                next_row_data = df.iloc[i + 1, :20]
                next_row_non_empty = sum(1 for val in next_row_data if pd.notna(val) and str(val).strip() != "")
                if next_row_non_empty > 0:
                    next_row_empty = False

            if next_row_empty:
                break

    return end_row


def extract_main_table(df, header_row_idx, main_table_end, data_end_row):
    """Extract main table data"""
    print(f"   📊 Extracting main table: columns 0-{main_table_end}, rows {header_row_idx + 1}-{data_end_row}")

    # Get headers
    main_headers = [str(h).replace("\xa0", " ").strip() if pd.notna(h) else f"COL_{i}"
                    for i, h in enumerate(df.iloc[header_row_idx, :main_table_end + 1])]

    # Extract main table data
    main_df = df.iloc[header_row_idx + 1:data_end_row, :main_table_end + 1].copy()
    main_df.columns = main_headers

    # Keep only known main columns that exist
    existing_main_cols = [col for col in MAIN_COLUMNS if col in main_df.columns]
    main_df = main_df[existing_main_cols]

    # Clean data
    main_df = main_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Handle merged date cells
    if "Date" in main_df.columns:
        main_df["Date"] = main_df["Date"].ffill()

    # Remove total rows
    if "Biller Name" in main_df.columns:
        main_df = main_df[~main_df["Biller Name"].astype(str).str.strip().str.lower().isin(["total", ""])]

    # Remove rows where key columns are all empty
    key_cols = ["Date", "Biller Name"]
    existing_key_cols = [col for col in key_cols if col in main_df.columns]
    if existing_key_cols:
        main_df = main_df.dropna(subset=existing_key_cols, how='all')

    print(f"   ✅ Main table extracted: {len(main_df)} rows, {len(main_df.columns)} columns")
    return main_df


def extract_parallel_table(df, header_row_idx, parallel_start, data_end_row):
    """Extract parallel table data and return as separate columns"""
    print(f"   📊 Extracting parallel table: starting from column {parallel_start}")

    # Find how many columns the parallel table has
    header_row = df.iloc[header_row_idx, parallel_start:]
    parallel_cols = 0
    for i, val in enumerate(header_row):
        if pd.notna(val) and str(val).strip():
            parallel_cols = i + 1
        elif parallel_cols > 0 and i > parallel_cols + 2:  # Allow 2 empty columns gap
            break

    if parallel_cols == 0:
        return pd.DataFrame()

    parallel_end = parallel_start + parallel_cols
    print(f"   📊 Parallel table spans columns {parallel_start}-{parallel_end}")

    # Extract parallel data
    parallel_df = df.iloc[header_row_idx + 1:data_end_row, parallel_start:parallel_end].copy()

    # Create generic column names for parallel data
    parallel_col_names = []
    for i in range(parallel_cols):
        header_val = df.iloc[header_row_idx, parallel_start + i]
        if pd.notna(header_val) and str(header_val).strip():
            col_name = f"Parallel_{str(header_val).strip().replace(' ', '_')}"
        else:
            col_name = f"Parallel_Col_{i + 1}"
        parallel_col_names.append(col_name)

    parallel_df.columns = parallel_col_names

    # Clean data
    parallel_df = parallel_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Remove completely empty rows
    parallel_df = parallel_df.dropna(how='all')

    print(f"   ✅ Parallel table extracted: {len(parallel_df)} rows, {len(parallel_df.columns)} columns")
    print(f"   📊 Parallel columns: {list(parallel_df.columns)}")
    return parallel_df


def combine_main_and_parallel(main_df, parallel_df):
    """Combine main and parallel tables by aligning rows"""
    if parallel_df.empty:
        return main_df

    print(f"   🔄 Combining main ({len(main_df)} rows) with parallel ({len(parallel_df)} rows)")

    # Handle case where parallel table might have +1 row than main table
    if len(parallel_df) > len(main_df):
        print(f"   📊 Parallel table has {len(parallel_df) - len(main_df)} extra row(s)")
        # Extend main_df to match parallel_df length
        extra_rows = len(parallel_df) - len(main_df)
        empty_rows = pd.DataFrame(index=range(len(main_df), len(parallel_df)), columns=main_df.columns)
        main_df_aligned = pd.concat([main_df, empty_rows], ignore_index=True)
        parallel_df_aligned = parallel_df.reset_index(drop=True)
    elif len(main_df) > len(parallel_df):
        # Extend parallel_df to match main_df length
        main_df_aligned = main_df.reset_index(drop=True)
        extra_rows = pd.DataFrame(index=range(len(parallel_df), len(main_df)), columns=parallel_df.columns)
        parallel_df_aligned = pd.concat([parallel_df, extra_rows], ignore_index=True)
    else:
        # Same length, just reset indices
        main_df_aligned = main_df.reset_index(drop=True)
        parallel_df_aligned = parallel_df.reset_index(drop=True)

    # Combine the dataframes horizontally
    result_df = pd.concat([main_df_aligned, parallel_df_aligned], axis=1)

    print(f"   ✅ Combined table: {len(result_df)} rows, {len(result_df.columns)} columns")
    return result_df


def extract_tables_from_sheet(df, header_row_idx):
    """Extract both main and parallel tables from a sheet"""
    print(f"   🔎 Processing sheet data starting from header row {header_row_idx}")

    # Find boundaries
    main_table_end = find_main_table_end_column(df, header_row_idx)
    parallel_table_start = find_parallel_table_start(df, header_row_idx, main_table_end)
    data_end_row = find_data_end_row(df, header_row_idx)

    print(f"   📊 Main table ends at column {main_table_end}")
    print(f"   📊 Data ends at row {data_end_row}")
    if parallel_table_start:
        print(f"   📊 Parallel table starts at column {parallel_table_start}")
    else:
        print(f"   📊 No parallel table found")

    # Extract main table
    main_df = extract_main_table(df, header_row_idx, main_table_end, data_end_row)

    # Extract parallel table if it exists
    parallel_df = pd.DataFrame()
    if parallel_table_start is not None:
        parallel_df = extract_parallel_table(df, header_row_idx, parallel_table_start, data_end_row)

    # Combine tables
    result_df = combine_main_and_parallel(main_df, parallel_df)

    return result_df


def merge_excel_sheets_opened(file_name: str):
    """Main function to merge all Excel sheets"""
    # Connect to active Excel app
    app = xw.apps.active
    if app is None:
        raise RuntimeError("No active Excel instance found. Open the file first.")

    # Find workbook
    wb = None
    for book in app.books:
        if book.name == file_name:
            wb = book
            break
    if wb is None:
        raise FileNotFoundError(f"Workbook '{file_name}' is not opened in Excel.")

    print(f"✅ Found workbook: {wb.name}\n")

    all_data = []

    # Process all visible sheets except Overall Summary
    visible_sheets = [s for s in wb.sheets if s.visible and s.name != "Overall Summary"]

    for sheet in visible_sheets:
        print(f"🔎 Processing sheet: {sheet.name}")
        used_range = sheet.used_range
        if used_range is None or used_range.value is None:
            print(f"⚠️ Sheet {sheet.name} is empty, skipping\n")
            continue

        df = pd.DataFrame(used_range.value)
        df.dropna(how="all", inplace=True)
        if df.empty:
            print(f"⚠️ Sheet {sheet.name} has no data, skipping\n")
            continue

        # Find header row with 'Date'
        header_row_idx = None
        for i in range(min(10, len(df))):
            row_values = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[i]]
            if "Date" in row_values:
                header_row_idx = i
                break

        if header_row_idx is None:
            print(f"⚠️ No header row with 'Date' found in {sheet.name}, skipping\n")
            continue

        # Extract tables from this sheet
        try:
            sheet_df = extract_tables_from_sheet(df, header_row_idx)
            if not sheet_df.empty:
                sheet_df['Sheet_Name'] = sheet.name  # Add source sheet identifier

                # Debug: Print columns from this sheet
                print(f"   🔍 DEBUG: Sheet {sheet.name} columns: {list(sheet_df.columns)}")

                all_data.append(sheet_df)
                print(f"   ✅ Sheet processed successfully: {len(sheet_df)} rows")
            else:
                print(f"   ⚠️ No data extracted from sheet")
        except Exception as e:
            print(f"   ❌ Error processing sheet: {e}")
            import traceback
            traceback.print_exc()

        print()  # Empty line for readability

    # Combine all data
    if not all_data:
        raise ValueError("No valid data found in any sheets to merge.")

    print(f"🔄 Combining data from {len(all_data)} sheets...")

    # Combine all dataframes
    final_df = pd.concat(all_data, ignore_index=True, sort=False)

    # Move Sheet_Name column to the end if it exists
    if 'Sheet_Name' in final_df.columns:
        sheet_name_col = final_df.pop('Sheet_Name')
        final_df['Sheet_Name'] = sheet_name_col

    print(f"✅ Data combined: {len(final_df)} total rows, {len(final_df.columns)} columns")

    # Display preview
    print(f"\n🔎 Final Merged Data Preview:")
    print(f"Columns: {list(final_df.columns)}")
    print(final_df.head(10))

    # Create or clear Overall Summary sheet
    summary_sheet_name = "Overall Summary"
    if summary_sheet_name in [s.name for s in wb.sheets]:
        print(f"🗑️ Deleting existing '{summary_sheet_name}' sheet...")
        wb.sheets[summary_sheet_name].delete()

    print(f"📄 Creating new '{summary_sheet_name}' sheet...")
    summary_sheet = wb.sheets.add(summary_sheet_name, before=wb.sheets[0])

    # Write data to sheet
    values_to_write = [final_df.columns.tolist()] + final_df.values.tolist()
    summary_sheet.range("A1").value = values_to_write

    # Format headers - with safety check for column count
    num_cols = len(final_df.columns)
    if num_cols > 0 and num_cols <= 16384:  # Excel's maximum columns
        try:
            # For more than 26 columns, we need to handle column letters differently
            if num_cols <= 26:
                end_col = chr(65 + num_cols - 1)
            else:
                # For columns beyond Z, use AA, AB, etc.
                if num_cols <= 702:  # Up to ZZ
                    first_letter = chr(65 + (num_cols - 27) // 26)
                    second_letter = chr(65 + (num_cols - 27) % 26)
                    end_col = first_letter + second_letter
                else:
                    # For very large numbers, just format first 26 columns
                    end_col = "Z"
                    num_cols = 26

            header_range = summary_sheet.range(f"A1:{end_col}1")
            header_range.color = (79, 129, 189)  # Blue background
            header_range.api.Font.Color = 16777215  # White text
            header_range.api.Font.Bold = True
        except Exception as e:
            print(f"⚠️ Warning: Could not format headers due to too many columns: {e}")
    else:
        print(f"⚠️ Warning: Cannot format headers - too many columns ({num_cols})")

    print(
        f"✅ Merged data written to '{summary_sheet_name}' sheet: {len(final_df)} rows, {len(final_df.columns)} columns")
    print(f"\n🎉 Process completed successfully!")


if __name__ == "__main__":
    file_name = "All Billers Reconciliation Summary - October.xlsm"
    merge_excel_sheets_opened(file_name)