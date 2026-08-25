import re
import json

# Read data.js
with open("quant-dashboard/src/data.js", "r", encoding="utf-8") as f:
    js_content = f.read()

# Extract the RAW array content
m = re.search(r'export const RAW = \[(.*?)\];', js_content, re.DOTALL)
if not m:
    print("RAW array not found!")
    exit(1)

body = m.group(1)
records = []
for line in body.strip().split('\n'):
    line = line.strip()
    if not line or line.startswith('//'):
        continue
    line_json = re.sub(r'([a-zA-Z0-9_]+)\s*:', r'"\1":', line)
    if line_json.endswith(','):
        line_json = line_json[:-1]
    
    try:
        record = json.loads(line_json)
        records.append(record)
    except Exception as e:
        print(f"Failed to parse line: {line}\nError: {e}")

print(f"Parsed {len(records)} records from data.js.\n")

# 1. Verification of the KPI Card sums (Normal Section vs Deepesh Section)
normal_section = [r for r in records if r.get("group") == 1]
deepesh_section = [r for r in records if r.get("group") == 2 and r.get("code") != "P3361"]

fo_total_fund = sum(r.get("fund", 0.0) for r in normal_section)
fo_apr = sum(r.get("apr", 0) for r in normal_section)
fo_may = sum(r.get("may", 0) for r in normal_section)
fo_jun = sum(r.get("jun", 0) for r in normal_section)

cash_total_fund = sum(r.get("fund", 0.0) for r in deepesh_section)
cash_apr = sum(r.get("apr", 0) for r in deepesh_section)
cash_may = sum(r.get("may", 0) for r in deepesh_section)
cash_jun = sum(r.get("jun", 0) for r in deepesh_section)

print("--- DASHBOARD KPI CARD SUMS VERIFICATION ---")
print(f"F&O (Normal Section) KPI Totals:")
print(f"  Total AUM: {fo_total_fund} Cr (Expected: 156.4 Cr with override)")
print(f"  April P&L: {fo_apr:,} (Expected: 18,164,639)")
print(f"  May P&L  : {fo_may:,} (Expected: -608,314)")
print(f"  June P&L : {fo_jun:,} (Expected: -855,168)")

print(f"\nCash / ATS (Deepesh Section) KPI Totals:")
print(f"  Total AUM: {cash_total_fund} Cr (Expected: 32.0 Cr)")
print(f"  April P&L: {cash_apr:,} (Expected: 6,980,972)")
print(f"  May P&L  : {cash_may:,} (Expected: -191,500)")
print(f"  June P&L : {cash_jun:,} (Expected: 795,000)")

# 2. Verification of segment segregation (Jinesh Jain / Ramakar Jha)
CASH_STRATEGIES = ["CASH", "ETF", "NA"]
def is_cash(r):
    seg = (r.get("segment") or "").strip().upper()
    strat = (r.get("strategy") or "").strip().upper()
    name = r.get("name", "").lower()
    return seg == "CASH" or seg == "ETF" or strat in CASH_STRATEGIES or "cash" in name or "etf" in name

jinesh = next(r for r in records if r.get("name") == "Jinesh Jain")
ramakar = next(r for r in records if r.get("name") == "Ramakar Jha")

print("\n--- SEGREGATION STATUS ---")
print(f"Jinesh Jain: Segment={jinesh.get('segment')}, Strategy={jinesh.get('strategy')}, isCash={is_cash(jinesh)}")
print(f"Ramakar Jha: Segment={ramakar.get('segment')}, Strategy={ramakar.get('strategy')}, isCash={is_cash(ramakar)}")
