import pyodbc
import threading
from queue import Queue
from typing import List, Optional
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(threadName)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class DatabaseConfig:
    """Database configuration settings"""

    def __init__(self, access_db_path: str, mysql_dsn: str = None,
                 mysql_user: str = None, mysql_password: str = None):
        self.access_db_path = access_db_path
        self.mysql_dsn = mysql_dsn
        self.mysql_user = mysql_user
        self.mysql_password = mysql_password


class AccessDatabaseSearchFilter:
    """Multi-threaded search (MySQL) with single-threaded insert (Access)"""

    def __init__(self, config: DatabaseConfig, num_search_threads: int = 5,
                 batch_size: int = 100, start_date: Optional[str] = None,
                 end_date: Optional[str] = None):
        self.config = config
        self.num_search_threads = num_search_threads
        self.batch_size = batch_size
        self.search_queue = Queue()
        self.results_lock = threading.Lock()
        self.all_results = []
        # New attributes for date range
        self.start_date = start_date
        self.end_date = end_date

        if (self.start_date and not self.end_date) or (self.end_date and not self.start_date):
            logger.warning("Only one date provided. Search will proceed with BillerName only.")
        elif self.start_date and self.end_date:
            logger.info(f"Filtering by date range: {self.start_date} to {self.end_date}")

    def get_access_connection(self):
        """Create connection to Access database"""
        conn_str = (
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={self.config.access_db_path};'
        )
        return pyodbc.connect(conn_str)

    def get_mysql_connection(self):
        """Create direct connection to MySQL database"""
        if not self.config.mysql_dsn:
            raise ValueError("MySQL DSN not configured")

        conn_str = f'DSN={self.config.mysql_dsn};'
        if self.config.mysql_user:
            conn_str += f'UID={self.config.mysql_user};'
        if self.config.mysql_password:
            conn_str += f'PWD={self.config.mysql_password};'

        return pyodbc.connect(conn_str)

    def fetch_biller_names(self) -> List[str]:
        """Fetch all biller names from BillerSrch table"""
        logger.info("Fetching biller names from BillerSrch...")
        conn = self.get_access_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT BillerName FROM BillerSrch")
        biller_names = [row[0] for row in cursor.fetchall()]

        cursor.close()
        conn.close()

        logger.info(f"Found {len(biller_names)} biller names to search")
        return biller_names

    def clear_filtered_table(self):
        """Clear the dailyfiledto_filtered table before inserting new records"""
        logger.info("Clearing dailyfiledto_filtered table...")
        conn = self.get_access_connection()
        cursor = conn.cursor()

        try:
            cursor.execute("DELETE FROM dailyfiledto_filtered")
            conn.commit()
            logger.info("Table cleared successfully")
        except Exception as e:
            logger.error(f"Error clearing table: {e}")
            conn.rollback()
            raise
        finally:
            cursor.close()
            conn.close()

    def search_worker(self, thread_id: int):
        """Worker thread for searching records by BillerName and Date Range - Direct MySQL connection"""
        logger.info(f"Search worker {thread_id} started")

        try:
            conn = self.get_mysql_connection()
            cursor = conn.cursor()
            logger.info(f"Thread {thread_id}: Connected to MySQL directly")
        except Exception as e:
            logger.error(f"Thread {thread_id} failed to connect to MySQL: {e}")
            # Drain the queue even if connection failed
            while True:
                batch = self.search_queue.get()
                if batch is None:
                    break
                self.search_queue.task_done()
            return

        # Prepare the date filter for the query
        date_filter = ""
        date_params = []
        if self.start_date and self.end_date:
            # Assuming 'fdate' is a date/datetime field in MySQL
            date_filter = " AND fdate BETWEEN ? AND ?"
            date_params = [self.start_date, self.end_date]
            logger.info(f"Thread {thread_id}: Applying date filter: {date_filter.strip()}")

        while True:
            batch = self.search_queue.get()
            if batch is None:  # Poison pill
                break

            try:
                # Process in smaller sub-batches
                sub_batch_size = 50

                for i in range(0, len(batch), sub_batch_size):
                    sub_batch = batch[i:i + sub_batch_size]
                    placeholders = ','.join(['?' for _ in sub_batch])

                    # Updated query to include the date filter
                    query = f"""
                        SELECT Cust, `Index`, BillerName, InvoiceNum, InvAmount, 
                               AmountPaid, PayDate, OpFee, PostPaidShare, SubBillerName, 
                               SubBillerShare, DedFeeSubPost, InternalCode, Comments, 
                               ContractNum, fdate
                        FROM DailyFileDTO
                        WHERE Cust IN ({placeholders}) 
                        {date_filter}
                    """

                    # Combine biller names and date parameters for execution
                    exec_params = sub_batch + date_params

                    cursor.execute(query, exec_params)
                    results = cursor.fetchall()

                    if results:
                        with self.results_lock:
                            self.all_results.extend(results)
                        logger.info(f"Thread {thread_id}: Found {len(results)} records in sub-batch")

            except Exception as e:
                logger.error(f"Thread {thread_id} search error: {e}")
                logger.error(f"Batch size was: {len(batch)}")
            finally:
                self.search_queue.task_done()

        cursor.close()
        conn.close()
        logger.info(f"Search worker {thread_id} finished")

    def execute_search(self):
        """Execute multi-threaded search on MySQL"""
        # (This method is simplified as biller_names is fetched in run())
        # The logic inside run() already calls this with biller_names
        # Keep the existing parameter for backward compatibility if needed,
        # but the run method handles passing the fetched list.
        pass  # The execution logic is moved to run or managed by other methods

    def execute_search(self, biller_names: List[str]):
        """Execute multi-threaded search on MySQL"""
        logger.info("Starting multi-threaded search...")

        # Split biller names into batches
        for i in range(0, len(biller_names), self.batch_size):
            batch = biller_names[i:i + self.batch_size]
            self.search_queue.put(batch)

        # Start search worker threads
        search_threads = []
        for i in range(self.num_search_threads):
            t = threading.Thread(target=self.search_worker, args=(i,),
                                 name=f"SearchWorker-{i}")
            t.start()
            search_threads.append(t)

        # Add poison pills to stop workers
        for _ in range(self.num_search_threads):
            self.search_queue.put(None)

        # Wait for all search operations to complete
        for t in search_threads:
            t.join()

        logger.info(f"Search complete. Total records found: {len(self.all_results)}")

    def insert_all_records(self):
        """Insert all records into Access in a single-threaded, reliable way"""
        if not self.all_results:
            logger.warning("No results to insert")
            return

        total_records = len(self.all_results)
        logger.info(f"Starting single-threaded insert of {total_records} records...")

        conn = None
        cursor = None
        inserted_count = 0
        failed_count = 0

        try:
            conn = self.get_access_connection()
            cursor = conn.cursor()

            insert_query = """
                INSERT INTO dailyfiledto_filtered 
                (Cust, [Index], BillerName, InvoiceNum, InvAmount, AmountPaid, 
                 PayDate, OpFee, PostPaidShare, SubBillerName, SubBillerShare, 
                 DedFeeSubPost, InternalCode, Comments, ContractNum, fdate)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """

            # Process in batches for commit efficiency
            for i in range(0, total_records, self.batch_size):
                batch = self.all_results[i:i + self.batch_size]
                batch_num = (i // self.batch_size) + 1
                total_batches = (total_records + self.batch_size - 1) // self.batch_size

                try:
                    # Try batch insert first
                    cursor.executemany(insert_query, batch)
                    conn.commit()
                    inserted_count += len(batch)
                    logger.info(
                        f"Batch {batch_num}/{total_batches}: Inserted {len(batch)} records (Total: {inserted_count}/{total_records})")

                except pyodbc.Error as e:
                    # If batch fails, try one by one
                    logger.warning(f"Batch {batch_num} failed, inserting individually: {e}")
                    conn.rollback()

                    for record in batch:
                        try:
                            cursor.execute(insert_query, record)
                            conn.commit()
                            inserted_count += 1
                        except pyodbc.IntegrityError:
                            # Duplicate key - skip silently
                            failed_count += 1
                        except Exception as e:
                            logger.error(f"Failed to insert record (Index: {record[1]}): {e}")
                            failed_count += 1
                            conn.rollback()

                    logger.info(
                        f"Batch {batch_num}/{total_batches}: Individually inserted, Total: {inserted_count}/{total_records}")

            logger.info("=" * 60)
            logger.info(f"Insert complete: {inserted_count} inserted, {failed_count} failed/skipped")
            logger.info("=" * 60)

        except Exception as e:
            logger.error(f"Fatal error during insert: {e}", exc_info=True)
            if conn:
                conn.rollback()
            raise
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def run(self):
        """Execute the complete search and filter process"""
        start_time = datetime.now()
        logger.info("=" * 60)
        logger.info("Starting search and filter process by Biller Name and Date Range")
        logger.info(f"Date Range: {self.start_date or 'N/A'} to {self.end_date or 'N/A'}")
        logger.info("=" * 60)

        try:
            # Step 1: Fetch biller names to search
            biller_names = self.fetch_biller_names()

            if not biller_names:
                logger.warning("No biller names found to search")
                return

            # Step 2: Clear the filtered table
            self.clear_filtered_table()

            # Step 3: Execute multi-threaded search (directly on MySQL)
            self.execute_search(biller_names)

            # Step 4: Insert all records (single-threaded for reliability)
            self.insert_all_records()

            elapsed = (datetime.now() - start_time).total_seconds()
            logger.info("=" * 60)
            logger.info(f"Process completed successfully in {elapsed:.2f} seconds")
            logger.info(f"Total records found: {len(self.all_results)}")
            logger.info("=" * 60)

        except Exception as e:
            logger.error(f"Fatal error during execution: {e}", exc_info=True)
            raise


def main():
    """Main execution function"""

    # --- INPUT PARAMETERS FOR DATE FILTER ---
    # NOTE: Date format must be compatible with your MySQL database's date/datetime field type.
    # MySQL typically accepts 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM:SS'.
    # For date range search, use the desired format.
    start_date = '2025-10-01'  # Example start date
    end_date = '2025-10-31'  # Example end date
    # To run without a date filter, set both to None:
    # start_date = None
    # end_date = None
    # ----------------------------------------

    # Configure database connections
    config = DatabaseConfig(
        access_db_path=rf'D:\Freelance\Azm\DailyTrans.accdb',
        mysql_dsn='AzmSer5New',
        mysql_user='root',
        mysql_password='root'
    )

    # Create and run the search filter, passing the dates
    search_filter = AccessDatabaseSearchFilter(
        config=config,
        num_search_threads=5,  # Multi-threaded MySQL search
        batch_size=1000,  # Records per batch
        start_date=start_date,  # Pass the start date
        end_date=end_date  # Pass the end date
    )

    search_filter.run()


if __name__ == "__main__":
    main()