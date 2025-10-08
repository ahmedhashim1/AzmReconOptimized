import pyodbc
import pandas as pd
import os
from datetime import datetime
import numpy as np

# --- Configuration ---
# ⚠️ UPDATE THESE FILE NAMES OR PATHS ⚠️
DB_FILE_LIVE = rf"D:\Freelance\Azm\DailyTrans.accdb"  # Replace with your actual live database file name
DB_FILE_TEST = rf"D:\Freelance\Azm\DailyTrans - Testing.accdb"  # Replace with your actual testing database file name
TABLE_NAME = "TempForImportMySql"

# The COMBINATION of these columns is used for GROUPING (non-unique allowed).
COMPOSITE_KEY = ["Cust", "Index", "InvoiceNum"]

# Columns to be totaled/aggregated
AGGREGATE_COLS = ["InvAmount", "AmountPaid", "OpFee", "PostPaidShare", "SubBillerShare"]

# Specify the desired output path and file name. Use {timestamp} for a unique name.
REPORT_FILE_PATH_TEMPLATE = r"D:\Freelance\Azm\DB_Comparison_Report_{timestamp}.xlsx"


# --- Helper Function for Connection ---
def get_connection_string(db_path):
    """Generates the ODBC connection string for an Access database."""
    return (
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
        f"DBQ={db_path};"
    )


# --- Data Retrieval ---
def fetch_table_data(db_path, table_name):
    """Connects to the database and fetches all data from the specified table."""
    conn_str = get_connection_string(db_path)
    print(f"Connecting to: {db_path}...")
    try:
        conn = pyodbc.connect(conn_str)
        query = f"SELECT * FROM [{table_name}]"
        df = pd.read_sql(query, conn)
        conn.close()

        # Ensure key columns are strings for consistent grouping/merging
        for col in COMPOSITE_KEY:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().fillna('')

        print(f"Successfully loaded {len(df)} rows from '{table_name}' in {db_path}.")
        return df
    except pyodbc.Error as ex:
        sqlstate = ex.args[0]
        print(f"\n❌ ERROR connecting to {db_path} or reading table:")
        print(f"   SQLState: {sqlstate}")
        print("   Ensure the file path is correct and the Access ODBC driver is installed.")
        return None


# --- Main Comparison Logic ---
def compare_databases_by_totals(df_live, df_test, composite_key, agg_cols, report_path_template):
    """Compares two DataFrames by aggregating specified columns using a composite key."""

    # Generate the final report path with a timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = report_path_template.format(timestamp=timestamp)

    # Ensure all columns exist before proceeding
    if not all(col in df_live.columns for col in composite_key + agg_cols) or \
            not all(col in df_test.columns for col in composite_key + agg_cols):
        missing_live = set(composite_key + agg_cols) - set(df_live.columns)
        missing_test = set(composite_key + agg_cols) - set(df_test.columns)
        print(f"❌ ERROR: Missing columns. Live missing: {missing_live}, Test missing: {missing_test}")
        return

    print("\n--- Starting Grouped Totals Comparison ---")

    # 1. AGGREGATE DATA
    print("Aggregating Live Data...")
    df_agg_live = df_live.groupby(composite_key)[agg_cols].sum().reset_index()

    print("Aggregating Test Data...")
    df_agg_test = df_test.groupby(composite_key)[agg_cols].sum().reset_index()

    # 2. MERGE AGGREGATED TABLES
    merged_agg_df = pd.merge(df_agg_live, df_agg_test,
                             on=composite_key,
                             how='outer',
                             indicator='Source',
                             suffixes=('_Live', '_Test'))

    # 3. IDENTIFY RECORD (GROUP) DIFFERENCES

    # Groups only in Live (Missing in Test)
    missing_groups = merged_agg_df[merged_agg_df['Source'] == 'left_only'].drop(
        columns=merged_agg_df.filter(regex='_Test$').columns.tolist() + ['Source'])

    # Groups only in Test (Extra in Test)
    extra_groups = merged_agg_df[merged_agg_df['Source'] == 'right_only'].drop(
        columns=merged_agg_df.filter(regex='_Live$').columns.tolist() + ['Source'])

    # 4. IDENTIFY VALUE DIFFERENCES in COMMON GROUPS
    common_groups_df = merged_agg_df[merged_agg_df['Source'] == 'both'].drop(columns=['Source'])
    data_diff_list = []

    if common_groups_df.empty:
        print("\nNo common groups (composite keys) to compare total values.")
    else:
        print(f"Comparing totals for {len(common_groups_df)} common groups...")

        for col in agg_cols:
            col_live = f"{col}_Live"
            col_test = f"{col}_Test"

            # Use numpy.isclose for robust float comparison, as sums can introduce minor float errors
            # tolerance is set to a reasonable banking standard (e.g., penny differences are flagged)
            mismatched_mask = ~np.isclose(common_groups_df[col_live], common_groups_df[col_test], atol=0.005,
                                          equal_nan=True)

            if mismatched_mask.any():
                differing_records = common_groups_df[mismatched_mask].copy()

                # Calculate the difference for easy analysis
                differing_records['Difference_Value'] = differing_records[col_live] - differing_records[col_test]

                # Prepare report record
                report_df = differing_records[composite_key + [col_live, col_test, 'Difference_Value']].copy()
                report_df.insert(len(composite_key), 'Difference_Column', col)

                data_diff_list.append(report_df)

    # Prepare final output dataframes
    final_data_diff_df = pd.concat(data_diff_list, ignore_index=True) if data_diff_list else pd.DataFrame(
        columns=composite_key + ['Difference_Column', 'Value_Live', 'Value_Test', 'Difference_Value'])

    # --- 5. Write Results to Excel ---
    try:
        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

            # Summary Sheet
            summary_data = {
                'Description': [
                    'Comparison Status',
                    'Total Records in Live DB',
                    'Total Records in Test DB',
                    'Total Unique Groups (Composite Keys) in Live',
                    'Total Unique Groups (Composite Keys) in Test',
                    'Groups Missing in Test DB (Key Mismatch)',
                    'Groups Extra in Test DB (Key Mismatch)',
                    'Total Value Mismatches Found'
                ],
                'Count': [
                    'Success' if missing_groups.empty and extra_groups.empty and final_data_diff_df.empty else 'Differences Found',
                    len(df_live),
                    len(df_test),
                    len(df_agg_live),
                    len(df_agg_test),
                    len(missing_groups),
                    len(extra_groups),
                    len(final_data_diff_df)
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Write Difference Sheets
            missing_groups.to_excel(writer, sheet_name='Groups_Missing_in_Test', index=False)
            extra_groups.to_excel(writer, sheet_name='Groups_Extra_in_Test', index=False)
            final_data_diff_df.to_excel(writer, sheet_name='Total_Value_Differences', index=False)

        print(f"\n✨ **SUCCESS:** Grouped totals comparison report generated.")
        print(f"File saved to: {output_path}")

    except Exception as e:
        print(f"\n❌ ERROR writing to Excel: {e}")
        print("Possible causes: The directory path is invalid, or the report file is currently open.")


# --- Execution ---
if __name__ == "__main__":
    # 1. Fetch Data
    df_live = fetch_table_data(DB_FILE_LIVE, TABLE_NAME)
    df_test = fetch_table_data(DB_FILE_TEST, TABLE_NAME)

    # 2. Compare Data and Output to Excel
    compare_databases_by_totals(df_live, df_test, COMPOSITE_KEY, AGGREGATE_COLS, REPORT_FILE_PATH_TEMPLATE)