import openpyxl

wb = openpyxl.load_workbook("Quant Strategy(2026-27).xlsx", data_only=True)
ws = wb["FY-(26-27)Q1"]

header_row = [ws.cell(3, c).value for c in range(1, 15)]
print("Header row values:")
for idx, val in enumerate(header_row, 1):
    print(f"Col {idx}: {val}")
