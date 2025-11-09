import asyncio
import aiomysql
import sqlite3
import pyodbc
import logging
import os
import psutil
from datetime import datetime
from typing import List, Optional

# ---------------- Logging ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


# ---------------- Config ----------------
class DatabaseConfig:
    """Database configuration settings"""
    def __init__(self, access_db_path: str, mysql_host: str, mysql_port: int,
                 mysql_db: str, mysql_user: str, mysql_password: str):
        self.access_db_path = access_db_path
        self.mysql_host = mysql_host
        self.mysql_port = mysql_port
        self.mysql_db = mysql_db
        self.mysql_user = mysql_user
        self.mysql_password = mysql_password


# ---------------- Main Class ----------------
class AsyncMySQLToSQLite:
    """Async MySQL search -> SQLite insert (cleared each run) -> optional Access export"""

    def __init__(self, config: DatabaseConfig, batch_size: int = 500,
                 start_date: Optional[str] = None, end_date: Optional[str] = None,
                 sqlite_path: str = "dailyfiledto_filtered.sqlite"):
        self.config = config
        self.batch_size = batch_size
        self.start_date = start_date
        self.end_date = end_date
        self.sqlite_path = sqlite_path
        self.all_results = []

        cpu_cores = os.cpu_count() or 4
        self.max_concurrency = max(4, min(2 * cpu_cores, 32))
        logger.info(f"Adaptive concurrency = {self.max_concurrency}")

    # ---------------- Connections ----------------
    def get_access_connection(self):
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.config.access_db_path};"
        )
        return pyodbc.connect(conn_str)

    async def get_mysql_pool(self, concurrency: int):
        pool = await aiomysql.create_pool(
            host=self.config.mysql_host,
            port=self.config.mysql_port,
            user=self.config.mysql_user,
            password=self.config.mysql_password,
            db=self.config.mysql_db,
            minsize=max(2, concurrency // 2),
            maxsize=concurrency,
            autocommit=True,
        )
        return pool

    # ---------------- Access (for export only) ----------------
    def clear_access_table(self):
        conn = self.get_access_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM dailyfiledto_filtered")
            conn.commit()
            logger.info("Cleared Access table dailyfiledto_filtered")
        finally:
            cursor.close()
            conn.close()

    # ---------------- Async MySQL Search ----------------
    async def search_batch(self, pool, batch, batch_id):
        """Async search for a batch of biller names"""
        results = []
        placeholders = ",".join(["%s"] * len(batch))
        date_filter = ""
        params = list(batch)

        if self.start_date and self.end_date:
            date_filter = " AND fdate BETWEEN %s AND %s"
            params.extend([self.start_date, self.end_date])

        query = f"""
            SELECT Cust, `Index`, BillerName, InvoiceNum, InvAmount, AmountPaid,
                   PayDate, OpFee, PostPaidShare, SubBillerName, SubBillerShare,
                   DedFeeSubPost, InternalCode, Comments, ContractNum, fdate
            FROM DailyFileDTO
            WHERE Cust IN ({placeholders}) {date_filter}
        """

        try:
            async with pool.acquire() as conn:
                async with conn.cursor() as cur:
                    await cur.execute(query, params)
                    rows = await cur.fetchall()
                    results.extend(rows)
                    logger.info(f"Batch {batch_id}: Found {len(rows)} records.")
        except Exception as e:
            logger.error(f"MySQL batch {batch_id} error: {e}")
        return results

    async def execute_search_async(self, biller_names: List[str]):
        """Run async MySQL searches concurrently with adaptive control"""
        logger.info("Starting async MySQL search...")

        concurrency = self.max_concurrency
        pool = await self.get_mysql_pool(concurrency)

        # Simple load check
        load = psutil.cpu_percent(interval=1)
        if load > 80:
            concurrency = max(4, concurrency // 2)
            logger.warning(f"High CPU load {load}%, reducing concurrency to {concurrency}")
        elif load < 40:
            concurrency = min(concurrency + 4, self.max_concurrency)
            logger.info(f"Low CPU load {load}%, increasing concurrency to {concurrency}")

        sem = asyncio.Semaphore(concurrency)
        tasks = []

        async def limited_batch_search(batch, batch_id):
            async with sem:
                return await self.search_batch(pool, batch, batch_id)

        for i in range(0, len(biller_names), self.batch_size):
            batch = biller_names[i:i + self.batch_size]
            batch_id = (i // self.batch_size) + 1
            tasks.append(asyncio.create_task(limited_batch_search(batch, batch_id)))

        all_batches = await asyncio.gather(*tasks)
        pool.close()
        await pool.wait_closed()

        for batch_results in all_batches:
            self.all_results.extend(batch_results)

        logger.info(f"Async search complete. Total records: {len(self.all_results)}")

    # ---------------- SQLite Writer ----------------
    def write_to_sqlite(self):
        """Insert all results into SQLite (safe, fast, cleared before each run)"""
        if not self.all_results:
            logger.warning("No results to insert into SQLite.")
            return

        total_records = len(self.all_results)
        os.makedirs(os.path.dirname(os.path.abspath(self.sqlite_path)), exist_ok=True)

        logger.info(f"Writing {total_records} records to SQLite database {self.sqlite_path}")

        conn = sqlite3.connect(self.sqlite_path, timeout=30)
        cur = conn.cursor()

        # Create table if not exists
        create_sql = """
        CREATE TABLE IF NOT EXISTS dailyfiledto_filtered (
            Cust TEXT,
            [Index] INTEGER,
            BillerName TEXT,
            InvoiceNum TEXT,
            InvAmount REAL,
            AmountPaid REAL,
            PayDate TEXT,
            OpFee REAL,
            PostPaidShare REAL,
            SubBillerName TEXT,
            SubBillerShare REAL,
            DedFeeSubPost REAL,
            InternalCode TEXT,
            Comments TEXT,
            ContractNum TEXT,
            fdate TEXT
        )
        """
        cur.execute(create_sql)
        conn.commit()

        # Clear the table before inserting
        try:
            cur.execute("DELETE FROM dailyfiledto_filtered")
            conn.commit()
            logger.info("Cleared SQLite table dailyfiledto_filtered before inserting new records.")
        except Exception as e:
            logger.warning(f"Could not clear SQLite table (ignored): {e}")

        insert_sql = """
        INSERT INTO dailyfiledto_filtered
        (Cust, [Index], BillerName, InvoiceNum, InvAmount, AmountPaid,
         PayDate, OpFee, PostPaidShare, SubBillerName, SubBillerShare,
         DedFeeSubPost, InternalCode, Comments, ContractNum, fdate)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        inserted = 0
        batch_size = max(200, self.batch_size)
        for i in range(0, total_records, batch_size):
            batch = self.all_results[i:i + batch_size]
            cur.executemany(insert_sql, batch)
            conn.commit()
            inserted += len(batch)
            logger.info(f"SQLite: Inserted {inserted}/{total_records}")

        conn.close()
        logger.info("✅ SQLite insert complete.")

    # ---------------- Export SQLite -> Access ----------------
    def export_to_access(self):
        """
        Export only records from SQLite → Access where ANY field matches
        an invoice number found in Access table [InvoiceNumSrch].
        """
        logger.info("Preparing filtered export (match ANY field) from SQLite → Access")

        if not os.path.exists(self.sqlite_path):
            logger.error(f"SQLite file not found: {self.sqlite_path}")
            return

        # ---------------- Fetch invoice numbers from Access ----------------
        conn_access_src = self.get_access_connection()
        cur_access_src = conn_access_src.cursor()
        cur_access_src.execute("SELECT InvoiceNum FROM InvoiceNumSrch")
        invoice_nums = [str(row[0]).strip() for row in cur_access_src.fetchall() if row[0]]
        cur_access_src.close()
        conn_access_src.close()

        if not invoice_nums:
            logger.warning("No invoice numbers found in Access.InvoiceNumSrch — skipping export.")
            return

        logger.info(f"Fetched {len(invoice_nums)} invoice numbers from Access.InvoiceNumSrch")

        # ---------------- Discover SQLite table structure ----------------
        conn_sqlite = sqlite3.connect(self.sqlite_path)
        cur_sqlite = conn_sqlite.cursor()
        cur_sqlite.execute("PRAGMA table_info(dailyfiledto_filtered)")
        columns = [row[1] for row in cur_sqlite.fetchall()]
        # Quote each column safely for SQLite (handles reserved words like Index)
        quoted_columns = [f'"{col}"' for col in columns]
        logger.info(f"SQLite columns detected: {columns}")

        # ---------------- Search Matches Across All Columns ----------------
        matched_rows = []
        chunk_size = 300  # for SQLite parameter limit safety
        total_invoices = len(invoice_nums)

        logger.info("Searching SQLite for any matches across all columns...")
        for i in range(0, total_invoices, chunk_size):
            chunk = invoice_nums[i:i + chunk_size]

            # Build WHERE clause for this chunk
            placeholders = ",".join("?" * len(chunk))
            conditions = [f"CAST({col} AS TEXT) IN ({placeholders})" for col in quoted_columns]
            query = f"SELECT * FROM dailyfiledto_filtered WHERE {' OR '.join(conditions)}"

            # Parameters repeated for each column in OR set
            params = chunk * len(columns)
            cur_sqlite.execute(query, params)
            rows = cur_sqlite.fetchall()
            matched_rows.extend(rows)
            logger.info(f"Chunk {i // chunk_size + 1}: Found {len(rows)} matching rows")

        if not matched_rows:
            logger.warning("No matching records found in SQLite — nothing to export.")
            conn_sqlite.close()
            return

        logger.info(f"Total matched records to export: {len(matched_rows)}")

        # ---------------- Export to Access ----------------
        self.clear_access_table()
        conn_access_dest = self.get_access_connection()
        cur_access_dest = conn_access_dest.cursor()

        insert_query = """
            INSERT INTO dailyfiledto_filtered 
            (Cust, [Index], BillerName, InvoiceNum, InvAmount, AmountPaid, 
             PayDate, OpFee, PostPaidShare, SubBillerName, SubBillerShare, 
             DedFeeSubPost, InternalCode, Comments, ContractNum, fdate)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        total = len(matched_rows)
        batch_size = 50
        inserted = 0
        logger.info("Starting export to Access (any-field matches)...")

        for i in range(0, total, batch_size):
            batch = matched_rows[i:i + batch_size]
            cur_access_dest.executemany(insert_query, batch)
            conn_access_dest.commit()
            inserted += len(batch)
            logger.info(f"Access export: {inserted}/{total} inserted")

        cur_access_dest.close()
        conn_access_dest.close()
        conn_sqlite.close()
        logger.info("✅ Filtered export to Access (any-field match) complete.")

    # ---------------- Main Run ----------------
    async def run(self):
        start = datetime.now()
        logger.info("=" * 60)
        logger.info("Starting full async MySQL → SQLite (cleared each run) → optional Access export process")
        logger.info("=" * 60)

        # Fetch biller names from Access
        names = self.fetch_biller_names()
        if not names:
            logger.warning("No biller names found in Access table BillerSrch.")
            return

        # Run async MySQL search
        await self.execute_search_async(names)

        # Write to SQLite (cleared first)
        self.write_to_sqlite()

        elapsed = (datetime.now() - start).total_seconds()
        logger.info(f"Process completed successfully in {elapsed:.2f}s")

    # ---------------- Access helper ----------------
    def fetch_biller_names(self) -> List[str]:
        conn = self.get_access_connection()
        cur = conn.cursor()
        cur.execute("SELECT BillerName FROM BillerSrch")
        names = [row[0] for row in cur.fetchall()]
        cur.close()
        conn.close()
        return names


# ---------------- Entry Point ----------------
def main():
    config = DatabaseConfig(
        access_db_path=r"D:\Freelance\Azm\DailyTrans.accdb",
        mysql_host="localhost",
        mysql_port=3306,
        mysql_db="azm",
        mysql_user="root",
        mysql_password="root"
    )

    start_date = "2025-11-01"
    end_date = "2025-11-06"

    pipeline = AsyncMySQLToSQLite(
        config=config,
        batch_size=5000,
        start_date=start_date,
        end_date=end_date,
        sqlite_path=r"D:\Freelance\Azm\dailyfiledto_filtered.sqlite"
    )

    asyncio.run(pipeline.run())

    # Optional export step
    export_to_access = True  # set False to skip
    if export_to_access:
        pipeline.export_to_access()


if __name__ == "__main__":
    main()
