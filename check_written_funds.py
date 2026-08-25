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

CASH_STRATEGIES = ["CASH", "ETF", "NA"]
def is_cash(r):
    strat = (r.get("strategy") or "").strip().upper()
    name = r.get("name", "").lower()
    return strat in CASH_STRATEGIES or "cash" in name or "etf" in name

fo_fund = 0
cash_fund = 0

for r in records:
    c = is_cash(r)
    fund = r.get("fund", 0.0)
    if c:
        cash_fund += fund
    else:
        fo_fund += fund

print(f"F&O (non-cash) Total Fund: {fo_fund} Cr")
print(f"Cash (cash) Total Fund: {cash_fund} Cr")
print(f"Combined Total Fund: {fo_fund + cash_fund} Cr")
