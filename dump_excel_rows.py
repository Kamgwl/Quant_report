import openpyxl
wb = openpyxl.load_workbook("Quant Strategy(2026-27).xlsx", data_only=True)
ws = wb["FY-(26-27)Q1"]
for r in range(65, 75):
    row_vals = [ws.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r}: {row_vals}")
