import asyncio
import aiomysql
import sqlite3
import pyodbc
import logging
import os
from datetime import datetime


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class DatabaseConfig:
    def __init__(self, access_db_path: str, mysql_host: str, mysql_port: int,
                 mysql_db: str, mysql_user: str, mysql_password: str):
        self.access_db_path = access_db_path
        self.mysql_host = mysql_host
        self.mysql_port = mysql_port
        self.mysql_db = mysql_db
        self.mysql_user = mysql_user
        self.mysql_password = mysql_password


class MonthlySummaryPipeline:
    """Fetch grouped transactions for all billers from MySQL → SQLite → Access"""

    def __init__(self, config: DatabaseConfig,
                 start_date: str, end_date: str,
                 sqlite_path: str = "monthly_summary.sqlite"):
        self.config = config
        self.start_date = start_date
        self.end_date = end_date
        self.sqlite_path = sqlite_path
        self.summary_data = []

    # ---------------- Connections ----------------
    def get_access_connection(self):
        conn_str = (
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={self.config.access_db_path};"
        )
        return pyodbc.connect(conn_str)

    async def get_mysql_pool(self):
        pool = await aiomysql.create_pool(
            host=self.config.mysql_host,
            port=self.config.mysql_port,
            user=self.config.mysql_user,
            password=self.config.mysql_password,
            db=self.config.mysql_db,
            autocommit=True,
            minsize=1,
            maxsize=10
        )
        return pool

    # ---------------- Fetch from Access ----------------
    def fetch_biller_names(self):
        """Read biller names from Access table BillerSrch"""
        conn = self.get_access_connection()
        cur = conn.cursor()
        cur.execute("SELECT BillerName FROM BillerSrch")
        billers = [str(row[0]).strip() for row in cur.fetchall() if row[0]]
        cur.close()
        conn.close()

        if not billers:
            logger.warning("No biller names found in Access table BillerSrch.")
        else:
            logger.info(f"Fetched {len(billers)} biller names from BillerSrch.")
        return billers

    def fetch_calendar_dates(self):
        """Fetch all fdate values from CalendarDates within range"""
        conn = self.get_access_connection()
        cur = conn.cursor()
        query = f"""
            SELECT fdate FROM CalendarDates 
            WHERE fdate BETWEEN #{self.start_date}# AND #{self.end_date}#
            ORDER BY fdate
        """
        cur.execute(query)
        dates = [row[0] for row in cur.fetchall()]
        cur.close()
        conn.close()
        logger.info(f"Fetched {len(dates)} calendar dates from Access.CalendarDates")
        return dates

    # ---------------- Async MySQL Aggregation ----------------
    async def fetch_mysql_summary(self, pool, biller_name: str):
        """Run GROUP BY query in MySQL asynchronously for one biller"""
        query = """
            SELECT fdate, Cust, 
                   SUM(OpFee) AS OpFee,
                   COUNT(InvoiceNum) AS InvoiceCount
            FROM DailyFileDTO
            WHERE Cust = %s AND fdate BETWEEN %s AND %s
            GROUP BY fdate, Cust
            ORDER BY fdate
        """
        try:
            async with pool.acquire() as conn:
                async with conn.cursor() as cur:
                    await cur.execute(query, (biller_name, self.start_date, self.end_date))
                    rows = await cur.fetchall()
                    logger.info(f"[{biller_name}] → {len(rows)} MySQL rows fetched")
                    return {r[0]: (r[1], r[2], r[3]) for r in rows}  # map by fdate
        except Exception as e:
            logger.error(f"MySQL query failed for biller '{biller_name}': {e}")
            return {}

    # ---------------- SQLite Stage ----------------
    def init_sqlite(self):
        """Create SQLite table (and clear data if exists)"""
        conn = sqlite3.connect(self.sqlite_path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS monthly_summary (
                fdate TEXT,
                Cust TEXT,
                OpFee REAL,
                InvoiceCount INTEGER
            )
        """)
        conn.commit()
        cur.execute("DELETE FROM monthly_summary")
        conn.commit()
        conn.close()
        logger.info("SQLite.monthly_summary table initialized and cleared.")

    def write_to_sqlite(self, all_rows):
        """Write merged data (with missing dates) into SQLite"""
        if not all_rows:
            logger.warning("No rows to write into SQLite.")
            return
        conn = sqlite3.connect(self.sqlite_path)
        cur = conn.cursor()
        insert_sql = """
            INSERT INTO monthly_summary (fdate, Cust, OpFee, InvoiceCount)
            VALUES (?, ?, ?, ?)
        """
        cur.executemany(insert_sql, all_rows)
        conn.commit()
        conn.close()
        logger.info(f"Inserted {len(all_rows)} total rows into SQLite.monthly_summary.")

    # ---------------- Export to Access ----------------
    def export_to_access(self):
        """Export summarized SQLite data into Access table WalletUsage"""
        logger.info("Exporting monthly summary to Access (WalletUsage)...")

        if not os.path.exists(self.sqlite_path):
            logger.error(f"SQLite file not found: {self.sqlite_path}")
            return

        conn_sqlite = sqlite3.connect(self.sqlite_path)
        cur_sqlite = conn_sqlite.cursor()
        cur_sqlite.execute("SELECT * FROM monthly_summary")
        rows = cur_sqlite.fetchall()
        conn_sqlite.close()

        if not rows:
            logger.warning("No data in SQLite to export.")
            return

        conn_access = self.get_access_connection()
        cur_access = conn_access.cursor()

        # Ensure target table exists
        try:
            cur_access.execute("""
                CREATE TABLE WalletUsage (
                    fdate DATETIME,
                    Cust TEXT,
                    OpFee DOUBLE,
                    InvoiceCount LONG
                )
            """)
            conn_access.commit()
            logger.info("Created Access table WalletUsage.")
        except Exception:
            logger.info("Access table WalletUsage already exists, continuing...")

        cur_access.execute("DELETE FROM WalletUsage")
        conn_access.commit()
        logger.info("Cleared old records from Access.WalletUsage")

        insert_query = """
            INSERT INTO WalletUsage (fdate, Cust, OpFee, InvoiceCount)
            VALUES (?, ?, ?, ?)
        """

        batch_size = 100
        for i in range(0, len(rows), batch_size):
            batch = rows[i:i + batch_size]
            cur_access.executemany(insert_query, batch)
            conn_access.commit()
        cur_access.close()
        conn_access.close()
        logger.info(f"✅ Export complete — inserted {len(rows)} rows into WalletUsage")

    # ---------------- Full Run ----------------
    async def run(self):
        start = datetime.now()
        logger.info("=" * 70)
        logger.info("Generating Monthly Summary for All Billers (with Calendar Join)")
        logger.info("=" * 70)

        self.init_sqlite()
        billers = self.fetch_biller_names()
        calendar_dates = self.fetch_calendar_dates()
        if not billers or not calendar_dates:
            logger.error("Missing billers or calendar dates — aborting.")
            return

        pool = await self.get_mysql_pool()
        try:
            # Run all biller queries in parallel
            tasks = [self.fetch_mysql_summary(pool, biller) for biller in billers]
            results = await asyncio.gather(*tasks)
        finally:
            pool.close()
            await pool.wait_closed()

        # Combine all results and fill missing dates
        merged_rows = []
        for biller, data in zip(billers, results):
            for fdate in calendar_dates:
                if fdate in data:
                    _, opfee, invoicecount = data[fdate]
                    merged_rows.append((fdate, biller, opfee or 0, invoicecount or 0))
                else:
                    merged_rows.append((fdate, biller, 0, 0))

        self.write_to_sqlite(merged_rows)
        self.export_to_access()

        elapsed = (datetime.now() - start).total_seconds()
        logger.info(f"✅ Process completed successfully in {elapsed:.2f} seconds")


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

    start_date = "2025-10-01"
    end_date = "2025-10-31"

    pipeline = MonthlySummaryPipeline(
        config=config,
        start_date=start_date,
        end_date=end_date,
        sqlite_path=r"D:\Freelance\Azm\monthly_summary.sqlite"
    )

    asyncio.run(pipeline.run())


if __name__ == "__main__":
    main()
