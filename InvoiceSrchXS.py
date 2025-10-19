import pyodbc
import threading
from queue import Queue
from typing import List
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
    """Multi-threaded search and filter for Access database"""

    def __init__(self, config: DatabaseConfig, num_search_threads: int = 3,
                 num_insert_threads: int = 2, batch_size: int = 100):
        self.config = config
        self.num_search_threads = num_search_threads
        self.num_insert_threads = num_insert_threads
        self.batch_size = batch_size
        self.search_queue = Queue()
        self.insert_queue = Queue()
        self.results_lock = threading.Lock()
        self.all_results = []

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

    def fetch_invoice_numbers(self) -> List[str]:
        """Fetch all invoice numbers from InvoiceNumSrch table"""
        logger.info("Fetching invoice numbers from InvoiceNumSrch...")
        conn = self.get_access_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT InvoiceNum FROM InvoiceNumSrch")
        invoice_nums = [row[0] for row in cursor.fetchall()]

        cursor.close()
        conn.close()

        logger.info(f"Found {len(invoice_nums)} invoice numbers to search")
        return invoice_nums

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
        """Worker thread for searching records - Direct MySQL connection"""
        logger.info(f"Search worker {thread_id} started")

        # Connect directly to MySQL instead of through Access linked table
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

                    # Query the actual MySQL table name (not the Access linked name)
                    # You may need to adjust the table name if it's different in MySQL
                    query = f"""
                        SELECT Cust, `Index`, BillerName, InvoiceNum, InvAmount, 
                               AmountPaid, PayDate, OpFee, PostPaidShare, SubBillerName, 
                               SubBillerShare, DedFeeSubPost, InternalCode, Comments, 
                               ContractNum, fdate
                        FROM DailyFileDTO
                        WHERE InvoiceNum IN ({placeholders}) 
                           OR ContractNum IN ({placeholders})
                    """

                    params = sub_batch + sub_batch
                    cursor.execute(query, params)
                    results = cursor.fetchall()

                    if results:
                        with self.results_lock:
                            self.all_results.extend(results)
                        logger.info(f"Thread {thread_id}: Found {len(results)} records")

            except Exception as e:
                logger.error(f"Thread {thread_id} search error: {e}")
                logger.error(f"Batch size was: {len(batch)}")
            finally:
                self.search_queue.task_done()

        cursor.close()
        conn.close()
        logger.info(f"Search worker {thread_id} finished")

    def insert_worker(self, thread_id: int):
        """Worker thread for inserting records into Access"""
        logger.info(f"Insert worker {thread_id} started")
        conn = self.get_access_connection()
        cursor = conn.cursor()

        insert_query = """
            INSERT INTO dailyfiledto_filtered 
            (Cust, [Index], BillerName, InvoiceNum, InvAmount, AmountPaid, 
             PayDate, OpFee, PostPaidShare, SubBillerName, SubBillerShare, 
             DedFeeSubPost, InternalCode, Comments, ContractNum, fdate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        while True:
            batch = self.insert_queue.get()
            if batch is None:  # Poison pill
                break

            try:
                cursor.executemany(insert_query, batch)
                conn.commit()
                logger.info(f"Thread {thread_id}: Inserted {len(batch)} records")
            except Exception as e:
                logger.error(f"Thread {thread_id} insert error: {e}")
                conn.rollback()
            finally:
                self.insert_queue.task_done()

        cursor.close()
        conn.close()
        logger.info(f"Insert worker {thread_id} finished")

    def execute_search(self, invoice_numbers: List[str]):
        """Execute multi-threaded search"""
        logger.info("Starting multi-threaded search...")

        # Split invoice numbers into batches
        for i in range(0, len(invoice_numbers), self.batch_size):
            batch = invoice_numbers[i:i + self.batch_size]
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

    def execute_insert(self):
        """Execute multi-threaded insert"""
        if not self.all_results:
            logger.warning("No results to insert")
            return

        logger.info("Starting multi-threaded insert...")

        # Split results into batches for insertion
        for i in range(0, len(self.all_results), self.batch_size):
            batch = self.all_results[i:i + self.batch_size]
            self.insert_queue.put(batch)

        # Start insert worker threads
        insert_threads = []
        for i in range(self.num_insert_threads):
            t = threading.Thread(target=self.insert_worker, args=(i,),
                                 name=f"InsertWorker-{i}")
            t.start()
            insert_threads.append(t)

        # Add poison pills to stop workers
        for _ in range(self.num_insert_threads):
            self.insert_queue.put(None)

        # Wait for all insert operations to complete
        for t in insert_threads:
            t.join()

        logger.info("Insert complete")

    def run(self):
        """Execute the complete search and filter process"""
        start_time = datetime.now()
        logger.info("=" * 60)
        logger.info("Starting search and filter process")
        logger.info("=" * 60)

        try:
            # Step 1: Fetch invoice numbers to search
            invoice_numbers = self.fetch_invoice_numbers()

            if not invoice_numbers:
                logger.warning("No invoice numbers found to search")
                return

            # Step 2: Clear the filtered table
            self.clear_filtered_table()

            # Step 3: Execute multi-threaded search (directly on MySQL)
            self.execute_search(invoice_numbers)

            # Step 4: Execute multi-threaded insert (into Access)
            self.execute_insert()

            elapsed = (datetime.now() - start_time).total_seconds()
            logger.info("=" * 60)
            logger.info(f"Process completed successfully in {elapsed:.2f} seconds")
            logger.info(f"Total records processed: {len(self.all_results)}")
            logger.info("=" * 60)

        except Exception as e:
            logger.error(f"Fatal error during execution: {e}", exc_info=True)
            raise


def main():
    """Main execution function"""

    # Configure database connections
    # The DSN 'AzmSer5New' should be configured in ODBC Data Sources
    # Control Panel > Administrative Tools > ODBC Data Sources (64-bit)

    config = DatabaseConfig(
        access_db_path=rf'D:\Freelance\Azm\DailyTrans.accdb',
        mysql_dsn='AzmSer5New',  # Your ODBC DSN name
        mysql_user='root',  # Your MySQL username
        mysql_password='root'  # Your MySQL password
    )

    # Create and run the search filter
    # Search directly from MySQL, insert into Access

    search_filter = AccessDatabaseSearchFilter(
        config=config,
        num_search_threads=3,  # Threads for MySQL queries
        num_insert_threads=2,  # Threads for Access inserts
        batch_size=100  # Records per batch
    )

    search_filter.run()


if __name__ == "__main__":
    main()