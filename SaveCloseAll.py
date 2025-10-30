import win32com.client
import datetime
import os
from config import config


def get_active_workbook_name():
    """
    Build the name of the active workbook based on config settings.

    Returns:
        str: The name of the active workbook (filename only, not full path)
    """
    m_day = config.curr_day
    m_month = config.curr_month
    m_year = config.curr_year

    date = datetime.datetime(m_year, m_month, m_day)

    # Format date components
    path_day = date.strftime("%d")
    path_month_abbr = date.strftime("%b")
    path_month_full = date.strftime("%B")
    path_year = date.strftime("%Y")

    # Build the biller summary name (the active workbook)
    biller_summary_name = f"All Billers Reconciliation Summary - {path_month_full}.xlsm"

    return biller_summary_name


def get_active_workbook_path():
    """
    Build the full path of the active workbook based on config settings.

    Returns:
        str: The full path to the active workbook
    """
    m_day = config.curr_day
    m_month = config.curr_month
    m_year = config.curr_year

    date = datetime.datetime(m_year, m_month, m_day)

    # Format date components
    path_day = date.strftime("%d")
    path_month_abbr = date.strftime("%b")
    path_month_full = date.strftime("%B")
    path_year = date.strftime("%Y")

    # Build paths
    invoice_base_folder = config.invoice_base
    biller_summary_name = f"All Billers Reconciliation Summary - {path_month_full}.xlsm"
    biller_summary_path_name = rf"Billers Summary\{path_year}\{path_month_abbr}\{biller_summary_name}"
    biller_summary_path = os.path.join(invoice_base_folder, biller_summary_path_name)

    return biller_summary_path


def save_close_all():
    """
    Close all workbooks except the active biller summary workbook and PERSONAL.XLSB,
    saving changes if the active workbook is unsaved.
    """
    # Get Excel application instance
    excel = win32com.client.Dispatch("Excel.Application")

    # Get the expected active workbook name from config
    active_workbook_name = get_active_workbook_name()

    # Enable optimized mode
    optimized_mode(excel, True)

    try:
        # Check if active workbook has unsaved changes
        if excel.ActiveWorkbook is not None:
            active_wb = excel.ActiveWorkbook
            active_wb_saved = active_wb.Saved

            if config.debug_mode:
                print(f"Active workbook: {active_wb.Name}")
                print(f"Active workbook saved: {active_wb_saved}")
                print(f"Expected active workbook: {active_workbook_name}")

            # Only proceed if active workbook has unsaved changes
            if not active_wb_saved:
                # Iterate through all open workbooks
                for wb in excel.Workbooks:
                    wb_name = wb.Name

                    if config.debug_mode:
                        print(f"Processing workbook: {wb_name}")

                    # Skip if it's the active workbook, PERSONAL.XLSB, or starts with "AllCustomersDailyFile"
                    if (wb_name != active_workbook_name and
                            wb_name != "PERSONAL.XLSB" and
                            not wb_name.startswith("AllCustomersDailyFile")):
                        if config.debug_mode:
                            print(f"  -> Closing workbook: {wb_name}")
                        wb.Close(SaveChanges=True)
                    else:
                        if config.debug_mode:
                            print(f"  -> Skipping workbook: {wb_name}")

    except Exception as e:
        print(f"Error in save_close_all: {str(e)}")
        raise

    finally:
        # Disable optimized mode (restore normal settings)
        optimized_mode(excel, False)


def optimized_mode(excel_app, enable):
    """
    Toggle Excel application optimization settings.

    Args:
        excel_app: Excel Application COM object
        enable (bool): True to enable optimization, False to restore normal mode
    """
    try:
        excel_app.EnableEvents = not enable
        excel_app.ScreenUpdating = not enable
        excel_app.DisplayStatusBar = not enable
        excel_app.PrintCommunication = not enable

        # Set calculation mode
        if enable:
            excel_app.Calculation = -4135  # xlCalculationManual
        else:
            excel_app.Calculation = -4105  # xlCalculationAutomatic

        # EnableAnimations (may not be available in all Excel versions)
        try:
            excel_app.EnableAnimations = not enable
        except AttributeError:
            pass  # Skip if not available in this Excel version

    except Exception as e:
        print(f"Error in optimized_mode: {str(e)}")
        # Continue even if some settings fail


# Alternative version using xlwings (cleaner syntax)
def save_close_all_xlwings():
    """
    Alternative implementation using xlwings library.
    Requires: pip install xlwings
    """
    import xlwings as xw

    app = xw.apps.active
    active_workbook_name = get_active_workbook_name()

    # Enable optimized mode
    optimized_mode_xlwings(app, True)

    try:
        active_wb = app.books.active

        if config.debug_mode:
            print(f"Active workbook: {active_wb.name}")
            print(f"Expected active workbook: {active_workbook_name}")

        if not active_wb.api.Saved:
            for wb in app.books:
                wb_name = wb.name

                if config.debug_mode:
                    print(f"Processing workbook: {wb_name}")

                if wb_name != active_workbook_name and wb_name != "PERSONAL.XLSB":
                    if config.debug_mode:
                        print(f"  -> Closing workbook: {wb_name}")
                    wb.save()
                    wb.close()
                else:
                    if config.debug_mode:
                        print(f"  -> Skipping workbook: {wb_name}")

    except Exception as e:
        print(f"Error in save_close_all_xlwings: {str(e)}")
        raise

    finally:
        optimized_mode_xlwings(app, False)


def optimized_mode_xlwings(app, enable):
    """
    Toggle Excel optimization settings using xlwings.

    Args:
        app: xlwings App object
        enable (bool): True to enable optimization, False to restore normal mode
    """
    try:
        app.api.EnableEvents = not enable
        app.screen_updating = not enable
        app.api.DisplayStatusBar = not enable
        app.api.PrintCommunication = not enable

        if enable:
            app.calculation = 'manual'
        else:
            app.calculation = 'automatic'

        try:
            app.api.EnableAnimations = not enable
        except AttributeError:
            pass

    except Exception as e:
        print(f"Error in optimized_mode_xlwings: {str(e)}")


# Usage example
if __name__ == "__main__":
    # Using win32com
    try:
        print(f"Expected active workbook: {get_active_workbook_name()}")
        print(f"Expected active workbook path: {get_active_workbook_path()}")
        print("\nClosing all workbooks except active one...")
        save_close_all()
        print("Done!")
    except Exception as e:
        print(f"Failed: {str(e)}")

    # Or using xlwings (uncomment if xlwings is installed)
    # save_close_all_xlwings()