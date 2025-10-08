import xlwings as xw
import pandas as pd

MAIN_COLUMNS = [
    "Date", "Biller Name", "Total Amount Paid", "Total Amount received (bank)",
    "Total Amount (paid-Sadad fees)", "Difference (C-D)","Bank Transfer Charge","Amount transfer to BILLER",
    "Sadad Fees", "Azm Fees", "Total Fees", "Number of Bills", "Matched", "Status"
]

def merge_excel_sheets_opened(file_name: str):
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

    # Process all visible sheets (removed [:4] limit)
    visible_sheets = [s for s in wb.sheets if s.visible]

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

        # --- Detect header dynamically ---
        header_row_idx = None
        for i in range(min(10, len(df))):
            row_values = [str(x).strip() if x is not None else "" for x in df.iloc[i]]
            if "Date" in row_values:
                header_row_idx = i
                break

        if header_row_idx is None:
            print(f"⚠️ No header row with 'Date' found in {sheet.name}, skipping\n")
            continue

        headers = [str(h).replace("\xa0", " ").strip() if h is not None else f"IGNORE_{i}"
                   for i, h in enumerate(df.iloc[header_row_idx])]
        df = df.iloc[header_row_idx + 1:]  # keep rows below header
        df.columns = headers

        # Keep only main columns that exist in this sheet
        existing_cols = [col for col in MAIN_COLUMNS if col in df.columns]
        df = df[existing_cols]

        # Strip string values
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        # --- Handle merged date cells ---
        if "Date" in df.columns:
            df["Date"] = df["Date"].ffill()  # forward fill merged cells

        # # Drop rows where Date contains 'Total'
        # df = df[~df["Biller Name"].astype(str).str.contains("Total", na=False)]

        # Drop rows where Biller Name = 'Total'
        if "Biller Name" in df.columns:
            df = df[df["Biller Name"].astype(str).str.strip().str.lower() != "total"]

        # Cutoff irrelevant tables after main table
        cutoff_idx = None
        for idx, first_cell in enumerate(df.iloc[:, 0]):
            if isinstance(first_cell, str) and (
                "Company Name" in first_cell or "Sum of حصة المفوتر الفرعي" in first_cell
            ):
                cutoff_idx = idx
                break
        if cutoff_idx is not None:
            df = df.iloc[:cutoff_idx]

        if df.empty:
            print(f"⚠️ Table empty after cleaning in {sheet.name}, skipping\n")
            continue

        print(f"   ✅ {len(df)} rows kept from {sheet.name}\n")
        all_data.append(df)

    if not all_data:
        raise ValueError("No valid data found in visible sheets to merge.")

    # Concatenate all sheets
    merged_df = pd.concat(all_data, ignore_index=True)

    # # Optional: sort by Date → Biller Name for better structure
    # if "Date" in merged_df.columns and "Biller Name" in merged_df.columns:
    #     merged_df.sort_values(by=["Date", "Biller Name"], inplace=True, ignore_index=True)

    print(f"🔎 Merged Data Preview (first 10 rows):")
    print(merged_df.head(10))

    # Write to Overall Summary
    if "Overall Summary" in [s.name for s in wb.sheets]:
        summary_sheet = wb.sheets["Overall Summary"]
        summary_sheet.clear()
    else:
        summary_sheet = wb.sheets.add("Overall Summary")

    summary_sheet.range("A1").value = [merged_df.columns.tolist()] + merged_df.values.tolist()
    print("✅ Merged data written to 'Overall Summary' sheet.")


if __name__ == "__main__":
    file_name = rf"All Billers Reconciliation Summary - April.xlsx"
    merge_excel_sheets_opened(file_name)