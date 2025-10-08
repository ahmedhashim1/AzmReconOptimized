import pandas as pd
import os
from datetime import datetime
import numpy as np
import re

# ⚠️ MODIFIED REGEX FOR EXPLICIT INCLUSION OF \t, \n, \r ⚠️
# This regex matches all control characters (C0, including the requested \t, \n, \r) and DEL (\x7F).
NON_PRINTABLE_REGEX = re.compile(r'[\t\n\r\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')

# --------------------------------------------------------------------------------------
# --- Configuration ---
# --------------------------------------------------------------------------------------

# ⚠️ UPDATE THESE FILE PATHS FOR YOUR EXCEL FILES ⚠️
SOURCE_FILE_LIVE = rf"D:\Freelance\Azm\2025\Sep\28\AllCustomersDailyFile_28.xlsx"
SOURCE_FILE_TEST = rf"E:\ReconTest\DailyFiles\2025\Sep\28\AllCustomersDailyFile_28_test.xlsx"
SHEET_NAME = "DailyFileDTO"

# This must be the EXACT header string for your custom unique identifier column (e.g., 'U_ID').
KEY_COLUMN_HEADER = "U_ID"

# This list must contain the EXACT header strings for ALL columns that hold numeric/currency data.
NUMERIC_COLUMNS_LIST = [
    # Add your Arabic column names here:
    "قيمة الفاتورة",
    "المبلغ المدفوع",
    "رسوم العمليات",
    "حصة المفوتر",
    "حصة المفوتر الفرعي",
    # Add other numeric columns as needed...
]

# Numeric Tolerance: This is the max absolute difference allowed between two numbers
# before they are flagged as different (e.g., 0.000001).
NUMERIC_TOLERANCE = 1e-6

# Specify the desired output path and file name. Use {timestamp} for a unique name.
REPORT_FILE_PATH_TEMPLATE = r"D:\Freelance\Azm\DB_Comparison_Report_{timestamp}.xlsx"

# --------------------------------------------------------------------------------------
# --- Derived Configuration and Data Retrieval (STRICTLY ORIGINAL) ---
# --------------------------------------------------------------------------------------

REPORT_COLUMNS = [KEY_COLUMN_HEADER, 'Difference_Column', 'Value (Live)', 'Value (Test)']
REPORT_TYPES = {col: 'object' for col in REPORT_COLUMNS}
REPORT_TYPES[KEY_COLUMN_HEADER] = 'object'


def fetch_excel_data(file_path, sheet_name):
    """Reads data up to the first blank row. ONLY coerces the Key column to string."""
    print(f"Loading data from: {file_path} - Sheet: '{sheet_name}'...")
    try:
        # 1. Quick Read to find the end of clean data
        temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, usecols=[KEY_COLUMN_HEADER])
        last_clean_row_index = temp_df[temp_df[KEY_COLUMN_HEADER].isna()].index.min()

        num_rows_to_read = len(temp_df) if pd.isna(last_clean_row_index) else last_clean_row_index

        if KEY_COLUMN_HEADER not in temp_df.columns:
            raise KeyError(f"Key column '{KEY_COLUMN_HEADER}' not found in file: {os.path.basename(file_path)}")

        numeric_dtypes = {col: float for col in NUMERIC_COLUMNS_LIST if col in temp_df.columns}

        # 2. Final Read with Row Count Limit and Type Enforcement
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            dtype=numeric_dtypes,
            header=0,
            nrows=num_rows_to_read
        )

        # 3. Handle Key Column for MERGE (Convert to string but DO NOT strip whitespace)
        # This preserves all whitespace in the key for strict comparison.
        df[KEY_COLUMN_HEADER] = df[KEY_COLUMN_HEADER].astype(str).fillna('NO_ID_FOUND')

        # 4. Keep All other Columns STRICTLY ORIGINAL

        print(f"Successfully loaded {len(df)} transactional rows from {os.path.basename(file_path)}.")
        return df
    except FileNotFoundError:
        print(f"\n❌ ERROR: File not found at path: {file_path}")
        return None
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {e}")
        return None


# --------------------------------------------------------------------------------------
# --- Main Comparison Logic (WITH NON-PRINTABLE DETECTION) ---
# --------------------------------------------------------------------------------------

def compare_excel_files_to_excel(df_live, df_test, key_column_header, report_path_template):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = report_path_template.format(timestamp=timestamp)
    results = {}

    if df_live is None or df_test is None:
        print("\nComparison aborted due to data loading error.")
        return

    print("\n--- Starting Row-by-Row Data Comparison ---")

    # 1. Merge and Identify Record Differences
    merged_df = pd.merge(df_live, df_test,
                         on=key_column_header,
                         how='outer',
                         indicator='Source',
                         suffixes=('_Live', '_Test'))

    missing_in_test = merged_df[merged_df['Source'] == 'left_only'].drop(
        columns=merged_df.filter(regex='_Test$').columns.tolist() + ['Source'])
    extra_in_test = merged_df[merged_df['Source'] == 'right_only'].drop(
        columns=merged_df.filter(regex='_Live$').columns.tolist() + ['Source'])

    results['Record_Missing_in_Test'] = missing_in_test
    results['Record_Extra_in_Test'] = extra_in_test

    # 2. Find ALL Row-Level Differences
    common_records_df = merged_df[merged_df['Source'] == 'both'].drop(columns=['Source'])
    all_diff_rows_df = pd.DataFrame()

    if not common_records_df.empty:
        print(f"Comparing data values for {len(common_records_df)} common records...")

        difference_mask = pd.Series(False, index=common_records_df.index)
        compare_cols = [col for col in df_live.columns if col != key_column_header]

        for col in compare_cols:
            col_live = f"{col}_Live"
            col_test = f"{col}_Test"

            if col_live not in common_records_df.columns or col_test not in common_records_df.columns:
                continue

            if col in NUMERIC_COLUMNS_LIST:
                # --- NUMERIC COMPARISON (Tolerance) ---
                s1 = common_records_df[col_live].fillna(0)
                s2 = common_records_df[col_test].fillna(0)

                mismatched_mask = ~np.isclose(s1, s2, atol=NUMERIC_TOLERANCE, equal_nan=True)

            else:
                # --- STRING/TEXT COMPARISON (STRICTLY ORIGINAL) ---
                s1 = common_records_df[col_live]
                s2 = common_records_df[col_test]

                mismatched_mask = (s1 != s2)

                both_na_mask = s1.isna() & s2.isna()

                mismatched_mask = mismatched_mask & (~both_na_mask)

            difference_mask = difference_mask | mismatched_mask

        # Filter the common records using the final difference mask
        row_differences_df = common_records_df[difference_mask].copy()

        if not row_differences_df.empty:

            # --- NON-PRINTABLE CHARACTER DETECTION AND REPORTING ---
            row_differences_df['Non_Printable_Live_Columns'] = ""
            row_differences_df['Non_Printable_Test_Columns'] = ""
            string_cols = [col for col in df_live.columns if
                           col not in NUMERIC_COLUMNS_LIST and col != key_column_header]

            for col in string_cols:
                col_live = f"{col}_Live"
                col_test = f"{col}_Test"

                # Check Live data for non-printable characters
                live_mask = row_differences_df[col_live].astype(str).str.contains(NON_PRINTABLE_REGEX, regex=True,
                                                                                  na=False)
                row_differences_df.loc[live_mask, 'Non_Printable_Live_Columns'] += col + "; "

                # Check Test data for non-printable characters
                test_mask = row_differences_df[col_test].astype(str).str.contains(NON_PRINTABLE_REGEX, regex=True,
                                                                                  na=False)
                row_differences_df.loc[test_mask, 'Non_Printable_Test_Columns'] += col + "; "

            # Trim trailing semi-colon and space
            row_differences_df['Non_Printable_Live_Columns'] = row_differences_df[
                'Non_Printable_Live_Columns'].str.rstrip('; ')
            row_differences_df['Non_Printable_Test_Columns'] = row_differences_df[
                'Non_Printable_Test_Columns'].str.rstrip('; ')

            # Reorder columns for clear output
            final_cols = [key_column_header] + ['Non_Printable_Live_Columns', 'Non_Printable_Test_Columns']

            live_cols = sorted(
                [col for col in row_differences_df.columns if col.endswith('_Live') and col not in final_cols])
            test_cols = sorted(
                [col for col in row_differences_df.columns if col.endswith('_Test') and col not in final_cols])

            final_cols.extend(live_cols)
            final_cols.extend(test_cols)

            all_diff_rows_df = row_differences_df[final_cols]
            results['Row_Value_Differences'] = all_diff_rows_df
            print(f"Found {len(all_diff_rows_df)} rows with at least one cell discrepancy.")
        else:
            print("\n✅ No data differences found in common records.")
            empty_cols = [key_column_header, 'Non_Printable_Live_Columns', 'Non_Printable_Test_Columns'] + sorted(
                [f"{col}_Live" for col in compare_cols]) + sorted([f"{col}_Test" for col in compare_cols])
            results['Row_Value_Differences'] = pd.DataFrame(columns=empty_cols)

    # 3. Write Results to Excel
    try:
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

            # Summary Sheet
            summary_data = {
                'Description': [
                    'Comparison Status', 'Total Rows in Live Source', 'Total Rows in Test Source',
                    'Records Missing in Test Source (Key Mismatch)', 'Records Extra in Test Source (Key Mismatch)',
                    'Total Mismatching Rows Found'
                ],
                'Count': [
                    'Differences Found' if not (
                                missing_in_test.empty and extra_in_test.empty and all_diff_rows_df.empty) else 'Success',
                    len(df_live), len(df_test), len(missing_in_test), len(extra_in_test),
                    len(all_diff_rows_df)
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            results['Record_Missing_in_Test'].to_excel(writer, sheet_name='Records_Missing_in_Test', index=False)
            results['Record_Extra_in_Test'].to_excel(writer, sheet_name='Records_Extra_in_Test', index=False)
            results['Row_Value_Differences'].to_excel(writer, sheet_name='Row_Value_Differences', index=False)

        print(f"\n✨ **SUCCESS:** Excel source comparison report generated.")
        print(f"File saved to: {output_path}")

    except Exception as e:
        print(f"\n❌ ERROR writing to Excel: {e}")


# --- Execution ---
if __name__ == "__main__":
    df_live = fetch_excel_data(SOURCE_FILE_LIVE, SHEET_NAME)
    df_test = fetch_excel_data(SOURCE_FILE_TEST, SHEET_NAME)

    if df_live is not None and df_test is not None:
        compare_excel_files_to_excel(df_live, df_test, KEY_COLUMN_HEADER, REPORT_FILE_PATH_TEMPLATE)