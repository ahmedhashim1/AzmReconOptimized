import asyncio
import aiomysql
import sqlite3
import pyodbc
import logging
import os
import psutil
import pandas as pd
import tempfile
import win32com.client
import time
from datetime import datetime
from typing import List, Optional

# ---------------- Logging ----------------
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

class AsyncMySQLToSQLite:
    def __init__(self, config: DatabaseConfig, batch_size: int = 500,
                 start_date: Optional[str] = None, end_date: Optional[str] = None,
                 sqlite_path: str = "hyperpay_filtered.sqlite"):
        self.config = config
        self.batch_size = batch_size
        self.start_date = start_date
        self.end_date = end_date
        self.sqlite_path = sqlite_path
        self.all_results = []

        cpu_cores = os.cpu_count() or 4
        self.max_concurrency = max(4, min(2 * cpu_cores, 32))
        logger.info(f"Adaptive concurrency = {self.max_concurrency}")

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

    def clear_access_table(self):
        conn = self.get_access_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM Hyperpay_filtered")
            conn.commit()
            logger.info("Cleared Access table Hyperpay_filtered")
        finally:
            cursor.close()
            conn.close()

    async def search_batch(self, pool, batch, batch_id):
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
        logger.info("Starting async MySQL search...")

        concurrency = self.max_concurrency
        pool = await self.get_mysql_pool(concurrency)

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

    def write_to_sqlite(self):
        if not self.all_results:
            logger.warning("No results to insert into SQLite.")
            return

        total_records = len(self.all_results)
        os.makedirs(os.path.dirname(os.path.abspath(self.sqlite_path)), exist_ok=True)

        logger.info(f"Writing {total_records} records to SQLite database {self.sqlite_path}")

        conn = sqlite3.connect(self.sqlite_path, timeout=30)
        cur = conn.cursor()

        create_sql = """
        CREATE TABLE IF NOT EXISTS Hyperpay_filtered (
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

        try:
            cur.execute("DELETE FROM Hyperpay_filtered")
            conn.commit()
            logger.info("Cleared SQLite table Hyperpay_filtered before inserting new records.")
        except Exception as e:
            logger.warning(f"Could not clear SQLite table (ignored): {e}")

        insert_sql = """
        INSERT INTO Hyperpay_filtered
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

    def export_to_access(self):
        logger.info("Starting Unicode-safe ultra-fast export: SQLite → Excel → Access")

        if not os.path.exists(self.sqlite_path):
            logger.error(f"SQLite file not found: {self.sqlite_path}")
            return

        conn_sqlite = sqlite3.connect(self.sqlite_path)
        df = pd.read_sql_query("SELECT * FROM Hyperpay_filtered", conn_sqlite)
        conn_sqlite.close()
        total_records = len(df)
        logger.info(f"Total records to export: {total_records}")

        if total_records == 0:
            logger.warning("No records found to export.")
            return

        temp_xlsx = os.path.join(tempfile.gettempdir(), "hyperpay_export.xlsx")
        logger.info(f"Exporting SQLite data to Excel: {temp_xlsx}")
        df.to_excel(temp_xlsx, index=False)

        self.clear_access_table()

        time.sleep(2)  # allow ODBC to release lock

        logger.info("Importing Excel into Access via TransferSpreadsheet...")
        try:
            access_app = win32com.client.Dispatch("Access.Application")

            for attempt in range(3):
                try:
                    access_app.OpenCurrentDatabase(self.config.access_db_path, False)
                    access_app.Visible = False  # 👈 set visibility AFTER opening
                    break
                except Exception as e:
                    logger.warning(f"Access still locked (attempt {attempt+1}/3). Retrying in 2s...")
                    time.sleep(2)
            else:
                raise Exception("Failed to open Access database after 3 attempts.")

            access_app.DoCmd.TransferSpreadsheet(
                TransferType=0,
                SpreadsheetType=10,
                TableName="Hyperpay_filtered",
                FileName=temp_xlsx,
                HasFieldNames=True
            )

            access_app.CloseCurrentDatabase()
            access_app.Quit()
            logger.info(f"✅ Ultra-fast Excel import complete ({total_records} records).")

        except Exception as e:
            logger.error(f"TransferSpreadsheet import failed: {e}")
        finally:
            try:
                os.remove(temp_xlsx)
            except:
                pass

    async def run(self):
        start = datetime.now()
        logger.info("=" * 60)
        logger.info("Starting full async MySQL → SQLite (cleared each run) → Access export process")
        logger.info("=" * 60)

        names = self.fetch_biller_names()
        if not names:
            logger.warning("No biller names found in Access table BillerSrch.")
            return

        await self.execute_search_async(names)
        self.write_to_sqlite()

        elapsed = (datetime.now() - start).total_seconds()
        logger.info(f"Process completed successfully in {elapsed:.2f}s")

    def fetch_biller_names(self) -> List[str]:
        conn = self.get_access_connection()
        cur = conn.cursor()
        cur.execute("SELECT BillerName FROM BillerSrch")
        names = [row[0] for row in cur.fetchall()]
        cur.close()
        conn.close()
        return names

def main():
    config = DatabaseConfig(
        access_db_path=r"D:\Freelance\Azm\DailyTrans.accdb",
        mysql_host="localhost",
        mysql_port=3306,
        mysql_db="azm",
        mysql_user="root",
        mysql_password="root"
    )

    start_date = "2024-01-01"
    end_date = "2025-10-31"

    pipeline = AsyncMySQLToSQLite(
        config=config,
        batch_size=5000,
        start_date=start_date,
        end_date=end_date,
        sqlite_path=r"D:\Freelance\Azm\hyperpay_filtered.sqlite"
    )

    asyncio.run(pipeline.run())
    pipeline.export_to_access()

if __name__ == "__main__":
    main()
