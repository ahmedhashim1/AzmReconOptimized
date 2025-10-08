import xlwings as xw



for wb in xw.apps.active.books:
    if not wb.name.lower().endswith("personal.xlsb"):
        wb.app.api.Windows(wb.name).Activate()