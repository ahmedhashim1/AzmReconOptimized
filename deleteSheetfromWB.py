import xlwings as xw

def delete_sheet_if_exists(sheet_name):
    # Suppress Excel alerts
    app = xw.apps.active
    # app.display_alerts = False

    try:
        # Loop through all open workbooks
        for workbook in app.books:
            if workbook.name != "PERSONAL.XLSB":  # Key change: Exclude this file.
                print(f"Checking workbook: {workbook.name}")


                # Check if the sheet exists in the workbook
                if sheet_name in [sheet.name for sheet in workbook.sheets]:
                    # print(f"Sheet '{sheet_name}' found in workbook '{workbook.name}'. Deleting...")
                    # Delete the sheet
                    workbook.sheets[sheet_name].delete()
                    print(f"Sheet '{sheet_name}' found and deleted from workbook '{workbook.name}'.")
                    workbook.save()
                    workbook.close()
                else:
                    print(f"Sheet '{sheet_name}' not found in workbook '{workbook.name}'.")

    except Exception as e:
        print(f"An error occurred: {e}")

    # finally:
        # Restore Excel alerts
        # app.display_alerts = True


sheet_name_to_delete = "10-Oct"  # Replace with the name of the sheet you want to delete
delete_sheet_if_exists(sheet_name_to_delete)