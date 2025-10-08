import xlwings as xw
import pandas as pd
import numpy as np

# A global list to be populated dynamically
DYNAMIC_MAIN_COLUMNS = []


def find_all_headers(df, header_row_idx):
    """Dynamically find all headers in a given sheet and handle duplicates."""
    headers = []
    header_counts = {}
    header_row_values = df.iloc[header_row_idx].values
    for i, val in enumerate(header_row_values):
        if pd.notna(val) and str(val).strip():
            header_val = str(val).replace("\xa0", " ").strip()

            # Handle duplicate headers
            if header_val in header_counts:
                header_counts[header_val] += 1
                headers.append(f"{header_val}_{header_counts[header_val]}")
            else:
                header_counts[header_val] = 0
                headers.append(header_val)

    return headers


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


def extract_tables_from_sheet(df, header_row_idx):
    """Extract all data from a sheet and align to the master column list."""
    print(f"   🔎 Processing sheet data starting from header row {header_row_idx}")

    # Get the headers for the current sheet. We do this before slicing to ensure we have the correct number of headers.
    sheet_columns = find_all_headers(df, header_row_idx)

    data_end_row = find_data_end_row(df, header_row_idx)

    print(f"   📊 Data ends at row {data_end_row}")

    # Extract all data from the sheet from the header row down using column indexes
    sheet_data = df.iloc[header_row_idx + 1:data_end_row, :len(sheet_columns)].copy()

    # Assign the correct headers to the DataFrame
    sheet_data.columns = sheet_columns

    # Reindex the sheet data based on the master column list
    reindexed_df = sheet_data.reindex(columns=DYNAMIC_MAIN_COLUMNS)

    # Clean data
    reindexed_df = reindexed_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Handle merged date cells
    if "Date" in reindexed_df.columns:
        reindexed_df["Date"] = reindexed_df["Date"].ffill()

    # Remove total rows
    if "Biller Name" in reindexed_df.columns:
        reindexed_df = reindexed_df[
            ~reindexed_df["Biller Name"].astype(str).str.strip().str.lower().isin(["total", ""])]

    # Remove rows where key columns are all empty
    key_cols = ["Date", "Biller Name"]
    existing_key_cols = [col for col in key_cols if col in reindexed_df.columns]
    if existing_key_cols:
        reindexed_df = reindexed_df.dropna(subset=existing_key_cols, how='all')

    print(f"   ✅ Main table extracted: {len(reindexed_df)} rows, {len(reindexed_df.columns)} columns")
    return reindexed_df


def combine_dataframes(all_data):
    """Combine a list of dataframes into a single one with a consistent structure."""
    if not all_data:
        return pd.DataFrame()

    final_df = pd.concat(all_data, ignore_index=True, sort=False)

    # Drop rows where 'Date' is blank - a key step for a clean dataset
    if 'Date' in final_df.columns:
        initial_rows = len(final_df)
        final_df = final_df.dropna(subset=['Date'])
        final_rows = len(final_df)
        print(f"🧹 Cleaned data: Dropped {initial_rows - final_rows} rows with blank 'Date' values.")

    # Move Sheet_Name column to the end
    if 'Sheet_Name' in final_df.columns:
        sheet_name_col = final_df.pop('Sheet_Name')
        final_df['Sheet_Name'] = sheet_name_col

    # Drop rows where 'Biller Name' is blank
    if 'Biller Name' in final_df.columns:
        initial_rows = len(final_df)
        final_df = final_df.dropna(subset=['Biller Name'])
        final_rows = len(final_df)
        print(f"🧹 Cleaned data: Dropped {initial_rows - final_rows} rows with blank 'Biller Name' values.")

    return final_df


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

    all_data_frames = []
    global DYNAMIC_MAIN_COLUMNS

    # Get a list of all sheets to find the one with the most headers
    sheets_to_process = [s for s in wb.sheets if s.visible and s.name.lower() not in ["overall summary", "template",
                                                                                      "ibans"] and s.api.Visible != xw.constants.SheetVisibility.xlSheetVeryHidden]

    if not sheets_to_process:
        raise ValueError("No sheets found to process. Please check sheet visibility and names.")

    # --- PHASE 1: Build the master column list from the LAST sheet processed ---
    reference_sheet = sheets_to_process[-1] if sheets_to_process else None
    if reference_sheet is None:
        raise RuntimeError("⚠️ No sheets to process. Cannot build master column list.")

    df_ref = pd.DataFrame(reference_sheet.used_range.value)
    header_row_idx_ref = None
    for i in range(min(10, len(df_ref))):
        if "Date" in [str(x).strip() for x in df_ref.iloc[i].values if pd.notna(x)]:
            header_row_idx_ref = i
            break

    if header_row_idx_ref is None:
        raise RuntimeError(f"⚠️ 'Date' header not found in the last sheet '{reference_sheet.name}'. Cannot proceed.")

    DYNAMIC_MAIN_COLUMNS = find_all_headers(df_ref, header_row_idx_ref)
    print(f"🧠 Dynamically built master column list from '{reference_sheet.name}':")
    print(DYNAMIC_MAIN_COLUMNS)
    print("-" * 50)

    # --- PHASE 2: Process and merge all valid sheets ---
    for sheet in sheets_to_process:
        print(f"🔎 Processing sheet: {sheet.name}")
        used_range = sheet.used_range
        df = pd.DataFrame(used_range.value)

        # Find header row with 'Date' on the full DataFrame
        header_row_idx = None
        for i in range(min(10, len(df))):
            row_values = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[i]]

            if "Date" in row_values and "Biller Name" in row_values:
                header_row_idx = i
                break

        # If a header is found, proceed with processing
        if header_row_idx is not None:
            # Extract tables from this sheet
            try:
                sheet_df = extract_tables_from_sheet(df, header_row_idx)
                if not sheet_df.empty:
                    sheet_df['Sheet_Name'] = sheet.name  # Add source sheet identifier
                    all_data_frames.append(sheet_df)
                    print(f"   ✅ Sheet processed successfully: {len(sheet_df)} rows")
                else:
                    print(f"   ⚠️ No data extracted from sheet")
            except Exception as e:
                print(f"   ❌ Error processing sheet: {e}")
                import traceback
                traceback.print_exc()
        else:
            print(f"⚠️ No header row with 'Date' and 'Biller Name' found in {sheet.name}, skipping\n")

        print()  # Empty line for readability

    # Combine all data
    if not all_data_frames:
        raise ValueError("No valid data found in any sheets to merge.")

    print(f"🔄 Combining data from {len(all_data_frames)} sheets...")

    final_df = combine_dataframes(all_data_frames)

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
                if num_cols <= 702:
                    first_letter = chr(65 + (num_cols - 1) // 26 - 1)
                    second_letter = chr(65 + (num_cols - 1) % 26)
                    end_col = first_letter + second_letter
                else:
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
    file_name = "All Billers Reconciliation Summary - December.xlsm"
    merge_excel_sheets_opened(file_name)