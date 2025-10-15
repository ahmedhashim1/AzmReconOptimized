import os
from datetime import datetime
from pathlib import Path
import pyodbc
import pymysql
from typing import Optional, List, Dict
import logging
from config import config
import win32com.client
import pythoncom

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_sender.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class EmailSender:
    def __init__(self,
                 sender_email: str = "billersreport@edaat.sa",
                 access_db_path: str = None,
                 mysql_host: str = None,
                 mysql_user: str = None,
                 mysql_password: str = None,
                 mysql_database: str = None):
        """
        Initialize the Email Sender using MS Outlook.

        Args:
            sender_email: Sender email account configured in Outlook
            access_db_path: Path to MS Access database file
            mysql_host: MySQL server host
            mysql_user: MySQL username
            mysql_password: MySQL password
            mysql_database: MySQL database name
        """
        self.sender_email = sender_email
        self.access_db_path = access_db_path

        # MySQL connection parameters
        self.mysql_host = mysql_host or os.getenv('MYSQL_HOST')
        self.mysql_user = mysql_user or os.getenv('MYSQL_USER')
        self.mysql_password = mysql_password or os.getenv('MYSQL_PASSWORD')
        self.mysql_database = mysql_database or os.getenv('MYSQL_DATABASE')

        if not self.access_db_path:
            raise ValueError("MS Access database path must be provided")

        if not os.path.exists(self.access_db_path):
            raise FileNotFoundError(f"Access database not found: {self.access_db_path}")

        if not all([self.mysql_host, self.mysql_user, self.mysql_password, self.mysql_database]):
            raise ValueError("MySQL connection parameters must be provided or set in environment variables")

        # Initialize Outlook
        self.outlook_app = None
        self.outlook_account = None
        self._initialize_outlook()

        # Load email signature
        self.signature = self.load_signature()
        self.auto_response = self.load_auto_response()

    def _initialize_outlook(self):
        """Initialize Outlook application and get the specific account."""
        try:
            pythoncom.CoInitialize()
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")

            # Get the specific account
            logger.info(f"Searching for Outlook account: {self.sender_email}")
            for account in self.outlook_app.Session.Accounts:
                logger.info(f"Found account: {account.DisplayName} ({account.SmtpAddress})")
                if account.SmtpAddress.lower() == self.sender_email.lower():
                    self.outlook_account = account
                    logger.info(f"Successfully configured Outlook account: {self.sender_email}")
                    return

            if not self.outlook_account:
                error_msg = f"Account {self.sender_email} not found in Outlook. Please configure this account in Outlook first."
                logger.error(error_msg)
                raise ValueError(error_msg)

        except Exception as e:
            logger.error(f"Error initializing Outlook: {e}")
            raise

    def load_signature(self, signature_file: str = "Edaat.htm") -> str:
        """Load email signature from HTML file in Outlook Signatures folder."""
        try:
            # Get the Outlook Signatures folder path (same as VBA)
            appdata_dir = os.getenv('APPDATA')
            signatures_dir = os.path.join(appdata_dir, 'Microsoft', 'Signatures')
            signature_path = os.path.join(signatures_dir, signature_file)

            logger.info(f"Looking for signature at: {signature_path}")

            if os.path.exists(signature_path):
                with open(signature_path, 'r', encoding='utf-8') as f:
                    sig = f.read()

                # Fix relative references to images (same as VBA)
                # Replace relative paths with absolute paths
                file_name = signature_file.replace('.htm', '') + '_files/'
                absolute_path = os.path.join(signatures_dir, file_name)
                sig = sig.replace(file_name, absolute_path)

                logger.info(f"Successfully loaded signature from {signature_file}")
                return sig
            else:
                logger.warning(f"Signature file not found at {signature_path}. Using default signature.")
                return "<p>Best Regards,<br>Edaat Team</p>"
        except Exception as e:
            logger.error(f"Error loading signature: {e}")
            return "<p>Best Regards,<br>Edaat Team</p>"

    def load_auto_response(self) -> str:
        """Load auto-response footer text."""
        # <p style='background-color: yellow; color: #000; font-size: 11px; padding: 10px;'>
        # <h3><center><b><i> This is an automated message from an unmonitored email account. For any inquiries or discrepancies, please contact us at support@edaat.sa within three (3) days of receiving this email. After this period, the records will be considered final. </h3></center></b></i> "
        # </p>
        # """
        return """
        <div style='background-color: yellow; color: #000; font-size: 17px; padding: 10px; text-align: center;'>
        <b><i>This is an automated message from an unmonitored email account. For any inquiries or discrepancies, please contact us at support@edaat.sa within three (3) days of receiving this email. After this period, the records will be considered final.</i></b>
        </div>
        """
    def get_mysql_connection(self):
        """
        Create and return a connection to MySQL database.

        Returns:
            pymysql connection object
        """
        try:
            conn = pymysql.connect(
                host=self.mysql_host,
                user=self.mysql_user,
                password=self.mysql_password,
                database=self.mysql_database,
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor
            )
            return conn
        except pymysql.Error as e:
            logger.error(f"Error connecting to MySQL database: {e}")
            raise

    def get_access_connection(self):
        """
        Create and return a connection to MS Access database.

        Returns:
            pyodbc connection object
        """
        try:
            # Connection string for MS Access (no MySQL credentials needed)
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.access_db_path};'
            )
            conn = pyodbc.connect(conn_str)
            return conn
        except pyodbc.Error as e:
            logger.error(f"Error connecting to Access database: {e}")
            raise

    def get_customers_from_mysql(self, date: datetime) -> List[str]:
        """
        Get distinct customers from MySQL DailyFileDTO table for the specified date.

        Args:
            date: Transaction date

        Returns:
            List of customer names
        """
        mysql_conn = None
        try:
            mysql_conn = self.get_mysql_connection()
            cursor = mysql_conn.cursor()

            # Format date for MySQL
            date_str = date.strftime("%Y-%m-%d")

            # Query to get distinct customers for the date
            query = "SELECT DISTINCT Cust FROM DailyFileDTO WHERE fdate = %s"
            cursor.execute(query, (date_str,))

            #TEMPORARY CODE
            # cursor.execute(f"EXPLAIN {query}", (date_str,))
            # explain_result = cursor.fetchall()
            # logger.info(f"Query execution plan: {explain_result}")

            customers = [row['Cust'] for row in cursor.fetchall()]
            logger.info(f"Found {len(customers)} customers in MySQL for date {date_str}")

            return customers

        except pymysql.Error as e:
            logger.error(f"MySQL error: {e}")
            raise
        finally:
            if mysql_conn:
                mysql_conn.close()

    def get_email_info_from_access(self, customers: List[str], trans_type: str) -> List[Dict]:
        """
        Get email information from Access EmailDB table for specified customers.

        Args:
            customers: List of customer names
            trans_type: Type of transaction (Manual, B2B, VIP, WithRef, All)

        Returns:
            List of dictionaries containing email information
        """
        access_conn = None
        cursor = None
        recipients = []

        try:
            if not customers:
                logger.warning("No customers provided to query Access database")
                return recipients

            access_conn = self.get_access_connection()
            cursor = access_conn.cursor()

            # Build query with placeholders for customers
            placeholders = ','.join(['?' for _ in customers])
            query = f"""
                SELECT Cust, Email, EmailCC, IsB2b, isVip, isRef 
                FROM EmailDB 
                WHERE Cust IN ({placeholders})
            """

            # Add transaction type filter
            if trans_type == "Manual":
                query += " AND IsB2b = False"
            elif trans_type == "B2B":
                query += " AND IsB2b = True"
            elif trans_type == "VIP":
                query += " AND isVip = True"
            elif trans_type == "WithRef":
                query += " AND isRef = True"
            # "All" - no additional filtering

            logger.info(f"Access Query: {query}")
            cursor.execute(query, customers)

            # Fetch column names
            columns = [column[0] for column in cursor.description]

            # Fetch all rows and convert to list of dictionaries
            for row in cursor.fetchall():
                recipient = dict(zip(columns, row))
                recipients.append(recipient)

            logger.info(f"Retrieved {len(recipients)} recipients from Access database")

        except pyodbc.Error as e:
            logger.error(f"Access database error: {e}")
            raise
        finally:
            if cursor:
                cursor.close()
            if access_conn:
                access_conn.close()

        return recipients

    def get_email_recipients(self, date: datetime, trans_type: str) -> List[Dict]:
        """
        Retrieve email recipients by querying MySQL and Access separately.

        Args:
            date: Transaction date
            trans_type: Type of transaction (Manual, B2B, VIP, WithRef, All)

        Returns:
            List of dictionaries containing recipient information
        """
        try:
            # Step 1: Get customers from MySQL who have data for this date
            customers = self.get_customers_from_mysql(date)

            if not customers:
                logger.warning(f"No customers found in MySQL for date {date}")
                return []

            logger.info(f"MySQL returned {len(customers)} customers with transactions on {date}")

            # Step 2: Get email information from Access for these customers
            recipients = self.get_email_info_from_access(customers, trans_type)

            if len(recipients) < len(customers):
                logger.info(
                    f"Filtered from {len(customers)} customers to {len(recipients)} recipients based on transaction type '{trans_type}'")

            return recipients

        except Exception as e:
            logger.error(f"Error getting email recipients: {e}")
            raise

    def get_attachment_path(self, customer_name: str, date: datetime) -> Optional[str]:
        """
        Construct the file path for the biller report attachment.

        Args:
            customer_name: Name of the customer/biller
            date: Report date

        Returns:
            File path if exists, None otherwise
        """
        curr_year = date.year
        curr_month_abbr = date.strftime("%b")  # Jan, Feb, etc.
        curr_month_full = date.strftime("%B")  # January, February, etc.
        curr_day = date.strftime("%d")  # 01, 02, etc.

        # Construct primary file path
        file_name = f"{customer_name} Report {curr_day}-{curr_month_full}.xlsx"
        file_path = os.path.join(
            config.biller_base,
            customer_name,
            str(curr_year),
            curr_month_abbr,
            file_name
        )

        # Check if primary file exists
        if os.path.exists(file_path):
            return file_path

        # Try alternative path (WITHOUT Reference ID)
        alt_file_name = f"(WITHOUT Reference ID) {customer_name} Report {curr_day}-{curr_month_full}.xlsx"
        alt_file_path = os.path.join(
            config.biller_base,
            customer_name,
            str(curr_year),
            curr_month_abbr,
            alt_file_name
        )

        if os.path.exists(alt_file_path):
            logger.info(f"Using alternative file path for {customer_name}")
            return alt_file_path

        logger.warning(f"Attachment not found for {customer_name}: {file_path}")
        return None

    def create_email_subject(self, customer_name: str, date: datetime, trans_type: str) -> str:
        """Create email subject based on transaction type."""
        # date_str = date.strftime("%m-%d-%Y")
        day = date.day
        suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        date_str = f"{day}{suffix} {date.strftime('%b %Y')}"

        subject_suffix = {
            "Manual": " (M)",
            "B2B": " (B)",
            "VIP": " (M)",
            "All": " (B/M)",
            "WithRef": " (With Ref Numbers)"
        }

        suffix = subject_suffix.get(trans_type, "")
        return f"{customer_name} Reconciliation File {date_str}{suffix}"

    def create_email_body(self, customer_name: str, date: datetime) -> str:
        """Create HTML email body."""
        # date_str = date.strftime("%m-%d-%Y")
        day = date.day
        suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        date_str = f"{day}{suffix} {date.strftime('%b %Y')}"

        body = f"""
        {self.signature}
        <b>Dear {customer_name} company,</b><br><br>
        Please be informed that the due amounts have been successfully transferred to your account. 
        Kindly find attached the payment report reflecting the transactions received from your 
        customers dated {date_str}<br><br><br><br><br><br><br>
        {self.auto_response}
        """

        return body

    def send_email(self, to_email: str, cc_email: str, subject: str,
                   body: str, attachment_path: Optional[str] = None,
                   save_to_drafts: bool = False) -> bool:
        """
        Create an email in the correct account's Outbox or Drafts in Outlook.

        Args:
            to_email: Recipient email address
            cc_email: CC email address(es)
            subject: Email subject
            body: HTML email body
            attachment_path: Path to attachment file
            save_to_drafts: If True, saves to Drafts instead of sending

        Returns:
            True if successful, False otherwise
        """
        mail = None
        try:
            # Get the namespace first
            namespace = self.outlook_app.GetNamespace("MAPI")

            # Find the store for billersreport@edaat.sa account
            target_store = None
            for store in namespace.Stores:
                try:
                    # Check if this store belongs to our account
                    if self.sender_email.lower() in str(store.DisplayName).lower():
                        target_store = store
                        logger.info(f"Found target store: {store.DisplayName}")
                        break
                except:
                    continue

            if not target_store:
                logger.error(f"Could not find store for {self.sender_email}")
                return False

            # Get the Drafts folder for this specific account
            if save_to_drafts:
                target_folder = target_store.GetDefaultFolder(16)  # 16 = olFolderDrafts
                logger.info(f"Using Drafts folder from {target_store.DisplayName}")
            else:
                target_folder = target_store.GetDefaultFolder(4)  # 4 = olFolderOutbox
                logger.info(f"Using Outbox folder from {target_store.DisplayName}")

            # Create new mail item in the specific folder
            mail = target_folder.Items.Add(0)  # 0 = olMailItem

            # Set the sending account
            if not self.outlook_account:
                logger.error(f"Outlook account {self.sender_email} not configured!")
                return False

            mail.SendUsingAccount = self.outlook_account

            # Set mail properties
            mail.To = to_email
            if cc_email and cc_email.strip():
                mail.CC = cc_email
            mail.Subject = subject
            mail.HTMLBody = body

            # Attach file if exists
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                logger.info(f"Attached file: {os.path.basename(attachment_path)}")
            else:
                if attachment_path:
                    logger.warning(f"Attachment not found: {attachment_path}")

            if save_to_drafts:
                # Save to Drafts
                mail.Save()
                logger.info(f"Email saved to {self.sender_email} Drafts for {to_email}")
            else:
                # Send the email
                mail.Send()
                logger.info(f"Email sent from {self.sender_email} for {to_email}")

            return True

        except Exception as e:
            logger.error(f"Failed to create email for {to_email}: {e}")
            if mail:
                try:
                    mail.Close(0)  # Close without saving
                except:
                    pass
            return False

    def send_batch_emails(self, date: datetime, trans_type: str, save_to_drafts: bool = False) -> dict:
        """
        Create emails in Outlook Outbox or Drafts for multiple recipients based on transaction type.

        Args:
            date: Transaction date
            trans_type: Type of transaction
            save_to_drafts: If True, saves to Drafts instead of sending to Outbox

        Returns:
            Dictionary with success/failure statistics
        """
        folder_name = "Drafts" if save_to_drafts else "Outbox"
        logger.info(f"Starting batch email creation for {trans_type} on {date} (saving to {folder_name})")

        results = {
            'total': 0,
            'successful': 0,
            'failed': 0,
            'failed_emails': [],
            'mysql_customers': 0,
            'filtered_recipients': 0
        }

        try:
            # Get recipients (queries MySQL and Access)
            recipients = self.get_email_recipients(date, trans_type)
            results['filtered_recipients'] = len(recipients)
            results['total'] = len(recipients)

            if not recipients:
                logger.warning(f"No recipients found for {trans_type} on {date}")
                logger.info(f"This means either:")
                logger.info(f"  1. No customers in EmailDB match the transaction type filter ({trans_type})")
                logger.info(f"  2. Customers in MySQL don't exist in Access EmailDB table")
                return results

            logger.info(f"Creating {len(recipients)} emails in {self.sender_email} {folder_name}...")

            # Create emails for each recipient
            for recipient in recipients:
                customer_name = recipient['Cust'].strip()
                to_email = recipient['Email'].strip()
                cc_email = recipient.get('EmailCC', '').strip() if recipient.get('EmailCC') else ''

                # Get attachment
                attachment_path = self.get_attachment_path(customer_name, date)

                # Create subject and body
                subject = self.create_email_subject(customer_name, date, trans_type)
                body = self.create_email_body(customer_name, date)

                # Create email in Outlook
                success = self.send_email(to_email, cc_email, subject, body, attachment_path, save_to_drafts)

                if success:
                    results['successful'] += 1
                else:
                    results['failed'] += 1
                    results['failed_emails'].append(to_email)

            logger.info(
                f"Batch email creation completed. Success: {results['successful']}, Failed: {results['failed']}")
            if save_to_drafts:
                logger.info(f"Emails are now in {self.sender_email} Drafts folder for review")
            else:
                logger.info(f"Emails are now in {self.sender_email} Outbox and will be sent automatically")

        except Exception as e:
            logger.error(f"Error in batch email creation: {e}")
            raise

        return results

    def __del__(self):
        """Cleanup COM objects."""
        try:
            pythoncom.CoUninitialize()
        except:
            pass


# Example usage
if __name__ == "__main__":
    # Path to your MS Access database
    ACCESS_DB_PATH = r"D:\Freelance\Azm\DailyTrans.accdb"  # Update this path

    # MySQL connection details (or set as environment variables)
    MYSQL_HOST = "localhost"
    MYSQL_USER = "root"
    MYSQL_PASSWORD = "root"
    MYSQL_DATABASE = "azm"

    try:
        # Initialize email sender with Outlook
        sender = EmailSender(
            sender_email="billersreport@edaat.sa",
            access_db_path=ACCESS_DB_PATH,
            mysql_host=MYSQL_HOST,
            mysql_user=MYSQL_USER,
            mysql_password=MYSQL_PASSWORD,
            mysql_database=MYSQL_DATABASE
        )

        # Verify the account was found
        if sender.outlook_account:
            print(f"✓ Using Outlook account: {sender.outlook_account.DisplayName}")
            print(f"  Email: {sender.outlook_account.SmtpAddress}\n")

        # Example: Create emails in Outbox for a specific date and type
        report_date = datetime(config.curr_year, config.curr_month, config.curr_day)
        transaction_type = "All"  # Can be: Manual, B2B, VIP, WithRef, All

        # Create batch emails in Outlook Outbox
        # Set save_to_drafts=True to save in Drafts folder instead
        results = sender.send_batch_emails(
            date=report_date,
            trans_type=transaction_type,
            save_to_drafts=False  # Change to True to save in Drafts folder
        )

        print(f"\n{'=' * 60}")
        print(f"Email Creation Summary for {transaction_type}")
        print(f"{'=' * 60}")
        print(f"Date: {report_date.strftime('%Y-%m-%d')}")
        print(f"Customers in MySQL: {results.get('mysql_customers', 'N/A')}")
        print(f"Recipients after filtering: {results['filtered_recipients']}")
        print(f"Successfully created in Outbox: {results['successful']}")
        print(f"Failed: {results['failed']}")
        if results['failed_emails']:
            print(f"\nFailed emails:")
            for email in results['failed_emails']:
                print(f"  - {email}")
        print(f"\n✓ Emails are now in {sender.sender_email} {('Drafts' if False else 'Outbox')}")
        if False:  # save_to_drafts
            print(f"  Review and send them manually from Outlook.")
        else:
            print(f"  They will be sent automatically by Outlook.")
        print(f"{'=' * 60}")

    except Exception as e:
        print(f"\n✗ Error: {e}")
        logger.error(f"Fatal error: {e}", exc_info=True)