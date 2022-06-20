import openpyxl as opxl

wb = opxl.load_workbook(r"money_data.xlsx")

fees_ws = wb["内訳"]

member_ws = wb["名簿"]



ws_new = wb.create_sheet(title="")

wb.save(r"money_data.xlsx")