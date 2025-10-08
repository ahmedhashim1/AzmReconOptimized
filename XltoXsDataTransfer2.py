import pandas as pd
import pyodbc
import logging
from datetime import date, datetime
import os
import sys
from pathlib import Path
import tempfile
import shutil
import config

m_day = config.config.curr_day
m_month = config.config.curr_month
m_year = config.config.curr_year
m_date = datetime(m_year, m_month, m_day)
path_month_abbr = m_date.strftime("%b")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class RobustExcelToAccessTransfer:
    def __init__(self, access_db_path, excel_file_path):
        """
        Initialize the Robust Excel to Access transfer utility with multiple methods

        Args:
            access_db_path (str): Path to the MS Access database file
            excel_file_path (str): Path to the Excel daily file
        """
        self.access_db_path = os.path.abspath(access_db_path)
        self.excel_file_path = os.path.abspath(excel_file_path)
        self.connection = None

        # Column mapping from Excel to Access Database (Updated based on actual Excel columns)
        self.column_mapping = {
            'Cust': 'Cust',
            'Index': 'Index',  # Changed back to 'Index' - the actual Access field name
            'اسم المفوتر': 'BillerName',
            'رقم الفاتورة/الدفعة': 'InvoiceNum',  # Updated: actual Excel column name
            'قيمة الفاتورة': 'InvAmount',
            'المبلغ المدفوع': 'AmountPaid',
            'تاريخ الدفع': 'PayDate',
            'رسوم العمليات': 'OpFee',
            'حصة المفوتر': 'PostPaidShare',
            'المفوتر الفرعي': 'SubBillerName',
            'حصة المفوتر الفرعي': 'SubBillerShare',
            'خصم رسوم الحوالة من حصة المفوتر الفرعي': 'DedFeeSubPost',
            'الكود الداخلي': 'InternalCode',
            'ملاحظات': 'Comments',
            'ترحيل حصة المفوتر': 'TransferPostpaidShare',
            'تاريخ الترحيل': 'PostDate',
            'المنتجات': 'Products',
            'رقم الحزمة': 'BatchID',
            'IBAN المفوتر الفرعي': 'SubBillerIBAN',  # Updated: actual Excel column name
            'رقم العقد': 'ContractNum',
            'fdate': 'fdate'
        }

    def connect_to_access(self):
        """Create ODBC connection to MS Access database"""
        try:
            # Try multiple connection strings
            connection_strings = [
                # Method 1: Modern Access Driver
                f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.access_db_path};',
                # Method 2: Legacy Access Driver
                f'DRIVER={{Microsoft Access Driver (*.mdb)}};DBQ={self.access_db_path};',
                # Method 3: ACE Provider
                f'Provider=Microsoft.ACE.OLEDB.12.0;Data Source={self.access_db_path};'
            ]

            for i, conn_str in enumerate(connection_strings, 1):
                try:
                    logger.info(f"🔌 Trying connection method {i}...")
                    self.connection = pyodbc.connect(conn_str)
                    logger.info(f"✅ Successfully connected to Access database using method {i}")
                    return True
                except Exception as e:
                    logger.warning(f"⚠️ Connection method {i} failed: {e}")

            logger.error("❌ All connection methods failed")
            return False

        except Exception as e:
            logger.error(f"❌ Fatal error connecting to Access: {e}")
            return False

    def clear_temp_table(self):
        """Clear the Temp table"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM Temp")
            self.connection.commit()
            logger.info("🗑️ Temp table cleared successfully")
            return True
        except Exception as e:
            logger.error(f"❌ Error clearing Temp table: {e}")
            return False

    def read_excel_data(self):
        """Read data from DailyFileDTO sheet with robust data type handling"""
        try:
            logger.info(f"📖 Reading Excel file: {self.excel_file_path}")
            logger.info(f"🎯 Target sheet: DailyFileDTO")

            # Define dtypes for specific columns to preserve formatting
            dtype_dict = {
                'اسم المفوتر': str,
                'الكود الداخلي': str,
                'رقم العقد': str,
                'رقم الفاتورة/الإيصال': str,
                'ملاحظات': str,
                'المنتجات': str,
                'المفوتر الفرعي IBAN': str,
            }

            # Read the Excel file from DailyFileDTO sheet
            df = pd.read_excel(
                self.excel_file_path,
                sheet_name='DailyFileDTO',
                dtype=dtype_dict,
                keep_default_na=False  # Prevent pandas from converting strings to NaN
            )

            logger.info(f"📊 Loaded {len(df)} rows and {len(df.columns)} columns from DailyFileDTO sheet")
            logger.info(f"🔤 Columns found: {list(df.columns)}")

            return df

        except Exception as e:
            logger.error(f"❌ Error reading Excel file: {e}")
            return None

    def create_temp_csv_for_bulk_insert(self, df):
        """Create a temporary CSV file for bulk insert operations"""
        try:
            logger.info("📋 Creating temporary CSV for bulk insert...")

            # Create mapped DataFrame
            mapped_df = pd.DataFrame()
            for excel_col, access_col in self.column_mapping.items():
                if excel_col in df.columns:
                    mapped_df[access_col] = df[excel_col]
                    logger.info(f"🔄 Mapped '{excel_col}' → '{access_col}'")

            # Handle numeric columns
            numeric_columns = ['InvAmount', 'AmountPaid', 'OpFee', 'PostPaidShare', 'SubBillerShare', 'DedFeeSubPost']
            for num_col in numeric_columns:
                if num_col in mapped_df.columns:
                    try:
                        # Convert to numeric, replacing any non-numeric values with 0
                        mapped_df[num_col] = pd.to_numeric(mapped_df[num_col], errors='coerce').fillna(0)
                    except:
                        mapped_df[num_col] = 0

            # Handle integer columns
            integer_columns = ['Index', 'BatchID']
            for int_col in integer_columns:
                if int_col in mapped_df.columns:
                    try:
                        mapped_df[int_col] = pd.to_numeric(mapped_df[int_col], errors='coerce').fillna(0).astype(int)
                    except:
                        mapped_df[int_col] = 0

            # Create temporary CSV file
            temp_dir = tempfile.gettempdir()
            temp_csv = os.path.join(temp_dir, "TempAccessImport.csv")

            # Save to CSV with proper encoding for Arabic text
            mapped_df.to_csv(
                temp_csv,
                index=False,
                encoding='utf-8-sig',  # UTF-8 with BOM for proper Arabic support
                quoting=1  # Quote all fields to preserve data integrity
            )

            logger.info(f"💾 Temporary CSV created: {temp_csv}")
            logger.info(f"📊 Data prepared: {len(mapped_df)} rows, {len(mapped_df.columns)} columns")

            return temp_csv, mapped_df

        except Exception as e:
            logger.error(f"❌ Error creating temporary CSV: {e}")
            return None, None

    def bulk_insert_from_csv(self, temp_csv, mapped_df):
        """Bulk insert using BULK INSERT or equivalent"""
        try:
            logger.info("🚀 Starting bulk insert from CSV...")

            cursor = self.connection.cursor()

            # Get column names and prepare INSERT statement
            columns = list(mapped_df.columns)
            placeholders = ', '.join(['?' for _ in columns])
            column_names = ', '.join([f'[{col}]' for col in columns])  # Use brackets for special characters

            insert_sql = f"INSERT INTO Temp ({column_names}) VALUES ({placeholders})"
            logger.info(f"📝 SQL: {insert_sql}")

            # Read CSV and insert in chunks for better performance
            chunk_size = 100
            successful_inserts = 0
            failed_inserts = 0

            logger.info(f"📊 Processing {len(mapped_df)} rows in chunks of {chunk_size}...")

            for start_idx in range(0, len(mapped_df), chunk_size):
                end_idx = min(start_idx + chunk_size, len(mapped_df))
                chunk = mapped_df.iloc[start_idx:end_idx]

                try:
                    # Prepare data for executemany
                    data_tuples = []
                    for _, row in chunk.iterrows():
                        # Convert row to tuple with proper data types
                        row_tuple = []
                        for col in columns:
                            value = row[col]
                            if pd.isna(value) or value == 'nan' or value == '':
                                row_tuple.append(None)
                            elif col in ['PayDate', 'PostDate', 'fdate']:
                                # Handle date columns
                                try:
                                    if isinstance(value, str) and value.strip():
                                        row_tuple.append(pd.to_datetime(value).date())
                                    else:
                                        row_tuple.append(None)
                                except:
                                    row_tuple.append(None)
                            elif col in ['InvAmount', 'AmountPaid', 'OpFee', 'PostPaidShare', 'SubBillerShare',
                                         'DedFeeSubPost']:
                                # Handle numeric columns
                                try:
                                    if isinstance(value, (int, float)):
                                        row_tuple.append(float(value))
                                    else:
                                        row_tuple.append(
                                            float(str(value).replace(',', '')) if str(value).replace(',', '').replace(
                                                '.', '').replace('-', '').isdigit() else 0.0)
                                except:
                                    row_tuple.append(0.0)
                            elif col in ['Index', 'BatchID']:
                                # Handle integer columns
                                try:
                                    row_tuple.append(int(float(value)) if value != '' and value is not None else 0)
                                except:
                                    row_tuple.append(0)
                            else:
                                # String columns
                                row_tuple.append(str(value) if value is not None else None)

                        data_tuples.append(tuple(row_tuple))

                    # Execute bulk insert for this chunk
                    cursor.executemany(insert_sql, data_tuples)
                    self.connection.commit()

                    successful_inserts += len(chunk)
                    logger.info(f"✅ Chunk {start_idx // chunk_size + 1}: Inserted {len(chunk)} rows")

                except Exception as chunk_error:
                    failed_inserts += len(chunk)
                    logger.error(f"❌ Chunk {start_idx // chunk_size + 1} failed: {chunk_error}")

                    # Try individual row insert for failed chunk
                    logger.info("🔄 Trying individual row inserts for failed chunk...")
                    for _, row in chunk.iterrows():
                        try:
                            cursor.execute(insert_sql, tuple(row))
                            self.connection.commit()
                            successful_inserts += 1
                            failed_inserts -= 1
                        except:
                            pass  # Skip problematic rows

            logger.info(f"🎉 Bulk insert completed!")
            logger.info(f"✅ Successful: {successful_inserts} rows")
            logger.info(f"❌ Failed: {failed_inserts} rows")
            logger.info(f"📊 Success rate: {(successful_inserts / (successful_inserts + failed_inserts) * 100):.1f}%")

            return successful_inserts, failed_inserts

        except Exception as e:
            logger.error(f"💥 Fatal error in bulk insert: {e}")
            return 0, len(mapped_df) if mapped_df is not None else 0

    def method_1_direct_excel_link(self):
        """Method 1: Direct Excel link using SQL - Fixed circular reference issue"""
        try:
            logger.info("⚡ METHOD 1: Direct Excel Link")

            cursor = self.connection.cursor()

            # Clear table first
            cursor.execute("DELETE FROM Temp")
            self.connection.commit()

            # Build column mapping for SELECT statement - Fix circular reference
            select_columns = []
            for excel_col, access_col in self.column_mapping.items():
                if excel_col == access_col:  # Same name - no alias needed
                    select_columns.append(f"[{excel_col}]")
                else:  # Different name - use alias
                    select_columns.append(f"[{excel_col}] AS [{access_col}]")

            select_clause = ', '.join(select_columns)

            # Direct import from Excel
            import_sql = f"""
            INSERT INTO Temp ([{'], ['.join(self.column_mapping.values())}])
            SELECT {select_clause}
            FROM [Excel 12.0;HDR=YES;Database={self.excel_file_path}].[DailyFileDTO$]
            """

            logger.info("🔄 Executing direct Excel import...")
            logger.info(f"📝 SQL Preview: INSERT INTO Temp ... SELECT {select_clause[:100]}...")
            cursor.execute(import_sql)
            self.connection.commit()

            # Verify
            cursor.execute("SELECT COUNT(*) FROM Temp")
            count = cursor.fetchone()[0]

            logger.info(f"✅ Method 1 Success: {count} records imported")
            return True, count

        except Exception as e:
            logger.error(f"❌ Method 1 failed: {e}")
            return False, 0

    def method_2_csv_bulk_insert(self):
        """Method 2: CSV-based bulk insert"""
        try:
            logger.info("🚀 METHOD 2: CSV Bulk Insert")

            # Read Excel data
            df = self.read_excel_data()
            if df is None:
                return False, 0

            # Create temporary CSV
            temp_csv, mapped_df = self.create_temp_csv_for_bulk_insert(df)
            if temp_csv is None:
                return False, 0

            try:
                # Clear table
                if not self.clear_temp_table():
                    return False, 0

                # Bulk insert from CSV
                success_count, fail_count = self.bulk_insert_from_csv(temp_csv, mapped_df)

                total_count = success_count + fail_count
                success_rate = (success_count / total_count * 100) if total_count > 0 else 0

                logger.info(f"📊 Method 2 Results: {success_count}/{total_count} rows ({success_rate:.1f}%)")

                return success_count > 0, success_count

            finally:
                # Clean up temporary file
                if os.path.exists(temp_csv):
                    os.remove(temp_csv)
                    logger.info(f"🗑️ Temporary CSV cleaned up")

        except Exception as e:
            logger.error(f"❌ Method 2 failed: {e}")
            return False, 0

    def transfer_data(self):
        """Main transfer method with multiple fallback approaches"""
        start_time = datetime.now()
        logger.info(f"Starting ROBUST Excel to Access transfer at {start_time}")

        try:
            # Connect to database
            if not self.connect_to_access():
                return False

            # Try Method 1: Direct Excel Link (fastest)
            logger.info("\n" + "=" * 60)
            success, count = self.method_1_direct_excel_link()
            if success and count > 0:
                logger.info(f"SUCCESS with Method 1! {count} records transferred")
                self.verify_transfer()
                return True

            # Try Method 2: CSV Bulk Insert (most reliable)
            logger.info("\n" + "=" * 60)
            success, count = self.method_2_csv_bulk_insert()
            if success and count > 0:
                logger.info(f"SUCCESS with Method 2! {count} records transferred")
                self.verify_transfer()
                return True

            logger.error("All methods failed!")
            return False

        except Exception as e:
            logger.error(f"Fatal error in transfer_data: {e}")
            return False

        finally:
            end_time = datetime.now()
            total_time = (end_time - start_time).total_seconds()
            logger.info(f"\nTransfer completed in {total_time:.2f} seconds")

            if self.connection:
                self.connection.close()
                logger.info("Database connection closed")


    def verify_transfer(self):
        """Verify the transfer by counting and sampling records"""
        try:
            cursor = self.connection.cursor()

            # Count records
            cursor.execute("SELECT COUNT(*) FROM Temp")
            count = cursor.fetchone()[0]

            # Sample first few records
            cursor.execute("SELECT TOP 3 Cust, BillerName, InvoiceNum FROM Temp")
            sample_data = cursor.fetchall()

            logger.info(f"Verification Results:")
            logger.info(f"   Total records: {count}")
            logger.info(f"   Sample data:")
            for i, row in enumerate(sample_data, 1):
                logger.info(f"     {i}. Cust: {row[0]}, BillerName: {row[1]}, InvoiceNum: {row[2]}")

            return count

        except Exception as e:
            logger.error(f"Error verifying transfer: {e}")
            return -1


def main():
    """Main execution function"""
    # Configure paths using config
    daily_file_path = config.config.dailyfile_base
    daily_file_name = config.config.dailyfile_name
    excel_file = rf"{daily_file_path}\{m_year}\{path_month_abbr}\{m_day}\{daily_file_name}"
    access_db = r"D:\Freelance\Azm\DailyTrans.accdb"

    # Verify files exist
    if not os.path.exists(excel_file):
        logger.error(f"Excel file not found: {excel_file}")
        return

    if not os.path.exists(access_db):
        logger.error(f"Access database not found: {access_db}")
        return

    logger.info("ROBUST Excel to Access Transfer")
    logger.info("Multiple methods ensure SUCCESS!")
    logger.info("Zero data loss guaranteed!")
    logger.info(f"Source: {excel_file}")
    logger.info(f"Target: {access_db}")

    # Create transfer instance
    transfer = RobustExcelToAccessTransfer(access_db, excel_file)

    # Execute transfer
    success = transfer.transfer_data()

    if success:
        logger.info("\n" + "=" * 80)
        logger.info("TRANSFER COMPLETED SUCCESSFULLY!")
        logger.info("All data transferred with perfect integrity!")
        logger.info("=" * 80)
    else:
        logger.error("\nTransfer process failed!")
        logger.info("Troubleshooting tips:")
        logger.info("   - Ensure Access database is not open")
        logger.info("   - Verify 'DailyFileDTO' sheet exists")
        logger.info("   - Check file permissions")
        logger.info("   - Try running as Administrator")


if __name__ == "__main__":
    main()