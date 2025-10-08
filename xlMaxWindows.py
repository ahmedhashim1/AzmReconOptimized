import win32com.client
import time

def maximize_all_excel_workbook_windows():
    """Loops through all opened Excel workbooks and maximizes their individual windows."""
    try:
        excel = win32com.client.GetObject(None, "Excel.Application")
        if excel:
            for workbook in excel.Workbooks:
                try:
                    for window in workbook.Windows:
                        window.WindowState = win32com.client.constants.xlMaximized
                        print(workbook.Name)
                        # Optional: Add a small delay if needed
                        time.sleep(0.1)
                except Exception as e:
                    print(f"Error maximizing windows for '{workbook.Name}': {e}")
            print("Successfully attempted to maximize all opened Excel workbook windows.")
        else:
            print("Excel application is not running.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    maximize_all_excel_workbook_windows()