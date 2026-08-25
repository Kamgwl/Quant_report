import openpyxl

wb = openpyxl.load_workbook("Quant Strategy(2026-27).xlsx", data_only=False)
ws = wb["FY-(26-27)Q1"]

for r in [67, 73]:
    row_formulas = [ws.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r}: {row_formulas}")
