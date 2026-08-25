import openpyxl

wb = openpyxl.load_workbook("Quant Strategy(2026-27).xlsx", read_only=True)
print("Sheet names:", wb.sheetnames)
