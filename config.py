class AppConfig:
    def __init__(self):

        #For reconciliation
        self.curr_day = 3
        self.curr_month = 11
        self.curr_year = 2025

        #For Email sender
        self.curr_day_Email = 30
        self.curr_month_Email = 10
        self.curr_year_Email = 2025

        def pad_number(number, width=2, fillchar='0'):
            """
            Converts a number to a padded string.

            Args:
              number: The number to pad.
              width: The desired width of the padded string.
              fillchar: The character to use for padding.

            Returns:
              The padded string representation of the number.
            """
            return str(number).zfill(width)

        #Actual Production links
        self.invoice_base = rf"D:\Freelance\Azm\OneDrive - AZM Saudi\Customers\Reconcilation Reports"
        self.dailyfile_base = rf"D:\Freelance\Azm"
        self.biller_base = rf"D:\Freelance\Azm\OneDrive - AZM Saudi\Customers\Biller Reports"

        #Testing code environment links
        # self.invoice_base = rf"E:\ReconTest\Reconcilation Reports"
        # self.dailyfile_base = rf"E:\ReconTest\DailyFiles"
        # self.biller_base = rf"E:\ReconTest\Biller Reports"


        self.dailyfile_name = rf"AllCustomersDailyFile_{pad_number(self.curr_day)}.xlsx"
        self.billerRepEmailDay = rf"{pad_number(self.curr_day_Email)}"


        # BILLER_REPORT_BASE = rf"E:\ReconTest\Biller Reports"
        # DAILY_FILE_BASE = rf"E:\ReconTest\DailyFiles"
        # INVOICE_BASE = rf"E:\ReconTest\Reconcilation Reports"

        self.debug_mode = False

config = AppConfig()  # Create a single instance