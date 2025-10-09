import pyodbc
import mysql.connector
from mysql.connector import Error
from datetime import datetime


class AccessToMySQLImporter:
    def __init__(self):
        # MS Access configuration
        self.access_db_path = r'D:\Freelance\Azm\DailyTrans.accdb'
        self.access_source_table = 'TempForImportMySql'

        # MySQL configuration
        self.mysql_config = {
            'host': 'localhost',
            'user': 'root',
            'password': 'root',
            'database': 'azm',
            'port': 3306
        }
        self.mysql_target_table = 'dailyfiledto'

        # Column mapping
        self.columns = [
            'Cust', 'Index', 'BillerName', 'InvoiceNum', 'InvAmount',
            'AmountPaid', 'PayDate', 'OpFee', 'PostPaidShare', 'SubBillerName',
            'SubBillerShare', 'DedFeeSubPost', 'InternalCode', 'Comments',
            'ContractNum', 'fdate'
        ]

    def connect_access(self):
        """Establish connection to MS Access database"""
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.access_db_path};'
            )
            conn = pyodbc.connect(conn_str)
            print("✓ Connected to MS Access database")
            return conn
        except pyodbc.Error as e:
            print(f"✗ Error connecting to Access: {e}")
            raise

    def connect_mysql(self):
        """Establish connection to MySQL database"""
        try:
            conn = mysql.connector.connect(**self.mysql_config)
            print("✓ Connected to MySQL database")
            return conn
        except Error as e:
            print(f"✗ Error connecting to MySQL: {e}")
            raise

    def create_table_if_not_exists(self, mysql_conn):
        """Create the target table if it doesn't exist"""
        create_table_query = f"""
        CREATE TABLE IF NOT EXISTS {self.mysql_target_table} (
            Cust VARCHAR(255),
            `Index` DOUBLE,
            BillerName VARCHAR(255),
            InvoiceNum VARCHAR(255),
            InvAmount DOUBLE,
            AmountPaid DOUBLE,
            PayDate VARCHAR(255),
            OpFee DOUBLE,
            PostPaidShare DOUBLE,
            SubBillerName VARCHAR(255),
            SubBillerShare DOUBLE,
            DedFeeSubPost VARCHAR(255),
            InternalCode VARCHAR(255),
            Comments VARCHAR(255),
            ContractNum VARCHAR(255),
            fdate TIMESTAMP(6)
        ) ENGINE=InnoDB
        """

        try:
            cursor = mysql_conn.cursor()
            cursor.execute(create_table_query)
            mysql_conn.commit()
            print(f"✓ Table '{self.mysql_target_table}' ready")
            cursor.close()
        except Error as e:
            print(f"✗ Error creating table: {e}")
            raise

    def fetch_access_data(self, access_conn):
        """Fetch all data from MS Access table"""
        try:
            cursor = access_conn.cursor()
            columns_str = ', '.join([f'[{col}]' for col in self.columns])
            query = f"SELECT {columns_str} FROM [{self.access_source_table}]"
            cursor.execute(query)
            rows = cursor.fetchall()
            print(f"✓ Fetched {len(rows)} rows from Access")
            cursor.close()
            return rows
        except pyodbc.Error as e:
            print(f"✗ Error fetching data from Access: {e}")
            raise

    def insert_mysql(self, mysql_conn, rows):
        """Insert data in MySQL (INSERT only, no updates)"""
        cursor = mysql_conn.cursor()

        # Prepare simple INSERT query
        columns_str = ', '.join([f'`{col}`' for col in self.columns])
        placeholders = ', '.join(['%s'] * len(self.columns))

        insert_query = f"""
        INSERT INTO {self.mysql_target_table} ({columns_str})
        VALUES ({placeholders})
        """

        success_count = 0
        error_count = 0

        for row in rows:
            try:
                # Convert row to list and handle None values
                row_data = [cell if cell is not None else None for cell in row]
                cursor.execute(insert_query, row_data)
                success_count += 1
            except Error as e:
                error_count += 1
                print(f"✗ Error inserting row: {e}")
                if error_count > 10:  # Stop if too many errors
                    print("Too many errors, aborting...")
                    raise

        mysql_conn.commit()
        print(f"✓ Successfully imported {success_count} rows")
        if error_count > 0:
            print(f"⚠ {error_count} rows failed to import")

        cursor.close()

    def import_data(self):
        """Main import process"""
        access_conn = None
        mysql_conn = None

        try:
            print("Starting data import from Access to MySQL...\n")

            # Connect to databases
            access_conn = self.connect_access()
            mysql_conn = self.connect_mysql()

            # Create table if needed
            self.create_table_if_not_exists(mysql_conn)

            # Fetch data from Access
            rows = self.fetch_access_data(access_conn)

            if not rows:
                print("⚠ No data to import")
                return

            # Insert data in MySQL
            self.insert_mysql(mysql_conn, rows)

            print("\n✓ Import completed successfully!")

        except Exception as e:
            print(f"\n✗ Import failed: {e}")
            raise

        finally:
            # Close connections
            if access_conn:
                access_conn.close()
                print("✓ Access connection closed")
            if mysql_conn:
                mysql_conn.close()
                print("✓ MySQL connection closed")


if __name__ == "__main__":
    importer = AccessToMySQLImporter()
    importer.import_data()