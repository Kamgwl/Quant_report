"""
update_dashboard.py
───────────────────
Reads "Quant Strategy(2026-27).xlsx"  (sheet: FY-(26-27)Q1)
and updates the RAW data array in:
  • index.html                       (standalone CDN page → GitHub Pages)
  • indx.html                        (JSX source)
  • quant-dashboard/src/data.js      (Vite project module)

Run from  d:\\file_dash  then deploy.bat does the git push.
"""

import openpyxl
import re
import sys
import os
from datetime import datetime

# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_FILE  = "Quant Strategy(2026-27).xlsx"
SHEET_NAME  = "FY-(26-27)Q1"

# Column positions in the FY-(26-27)Q1 sheet  (1-based)
COL_CODE     = 2
COL_NAME     = 3
COL_STRATEGY = 5
COL_FUND     = 7
COL_APR      = 8
COL_APR_ROI  = 9
COL_MAY      = 10
COL_MAY_ROI  = 11
COL_JUN      = 12
COL_JUN_ROI  = 13

DATA_START_ROW = 4   # row 3 is the header row; data from row 4

# NAME_OVERRIDES:  "CODE": (display_name, strategy, fund_override_or_None)
# fund_override is in Crores; set None to use value from Excel.
NAME_OVERRIDES = {
    # code          display_name              strategy         fund_override
    "P3224":  ("Ansh Cash",           "CASH",          None),
    "P2827":  ("Bharat ETF",          "ETF",           None),
    "P2827_2":("Bharat Cash",         "CASH",          None),
    "P3090":  ("Bharat Straddle",     "STRADDLE",      None),
    "P2954":  ("Bharat Munjal_5",     "STRADDLE",      None),
    "P3196":  ("Bharat Munjal BFO",   "STRADDLE",      None),
    "P2954_2":("Bharat Munjal S1515", "STRADDLE",      None),
    "P3079":  ("Yogesh Cash",         "CASH",          None),
    "P2777":  ("Yogesh Kumar",        "STRADDLE",      None),
    "P2826":  ("Yogesh ATS",          "",              None),
    "P3039":  ("Abhishek",            "",              None),
    "P2967":  ("Aman",                "Money Circle",  None),
    "P2951":  ("Aman 20-60",          "20_60",         None),
    "P3166":  ("Aman Dohre",          "NSE FO",        None),
    "P3390":  ("Aman Dohre XTS34",    "NSE FO",        None),
    "P3323":  ("Amit&Kartik",         "NSE FO",        None),
    "P3323_1":("Amit&Kartik XTS35",   "",              None),
    "PSW023": ("Ansh SWV",            "Falcon",        None),
    "PHPO06": ("Ansh HPO",            "Falcon",        None),
    "PSW019": ("Kartaram SWV097",     "Spider 2.0",    None),
    "SWV0096":("Kartaram SWV096",     "Spider 2.0",    None),
    "PHPO07": ("Kartaram HPO",        "Spider 2.0",    None),
    "XMR0548":("Kartaram XMR",        "Spider 2.0",    None),
    "PA528":  ("Archana Ma'am",       "SPIDER 2.0",    None),
    "P3181":  ("Harshit XTS32",       "SELF",          None),
    "P3240":  ("Harshit XTS38",       "SELF",          None),
    "P3311":  ("Himanshu Pal",        "SELF",          None),
    "P3109":  ("Jyoti Prakesh XTS30", "Straddle",      None),
    "P3313":  ("Jyoti Prakesh ATS",   "TRANDING",      None),
    "P3146":  ("Jinesh Jain",         "STOCK FO",      None),
    "P2105":  ("Kamlesh",             "Wanda",         None),
    "P3024":  ("Kartaram_1",          "Spider 2.0",    None),
    "P3202":  ("Gagandeep",           "SELF",          None),
    "P3020":  ("Mahavir Jindal",      "SELF",          None),
    "P3218":  ("Muhunthan",           "CASH",          None),
    "P3341":  ("Maneesh Yadav",       "NSE FO",        None),
    "P2971":  ("Neeraj Garg_1",       "Straddle",      None),
    "P2971_2":("Neeraj Garg_2",       "ACE1",          None),
    "P2940":  ("Prince Cash",         "cash",          None),
    "P3347":  ("Prince FO",           "FO",            None),
    "P3070":  ("Prateek_2",           "multipul",      None),
    "P2999":  ("Prateek ETF",         "ETF",           None),
    "P3186":  ("Prateek NFO",         "NFO",           None),
    # P3135 fund is 3 Cr in reality; Excel shows 1 Cr (data entry error)
    "P3135":  ("Piyush Singhal",      "Strangle",      3),
    "PSW024": ("Piyush SWV",          "Strangle",      None),
    "P3168":  ("Prabjot Singh",       "NSE FO",        None),
    "P3342":  ("Raghav Tuli",         "",              None),
    "P3082":  ("Ramakar Jha",         "STOCK FO",      None),
    "P3048":  ("Shahid",              "SELF",          None),
    "P3208":  ("Sajal Sharma ATS",    "",              None),
    "P3335":  ("Sajal Sharma_2",      "",              None),
    "P3119":  ("Sajal Sharma_3",      "",              None),
    "P3297":  ("Sudeep",              "ATS",           None),
    "P3360":  ("Vidhya Sagar",        "SELF",          None),
    "P3385":  ("Vikas Gupta XTS13",   "SELF",          None),
    "P3386":  ("Vikas Gupta XTS15",   "SELF",          None),
    "P3334":  ("Varun Tondan XTS10",  "self",          None),
    "P3110":  ("Varun Tondan",        "self",          None),
    "P3113":  ("Vaishali",            "self",          None),
    "P2817":  ("DK Sir_M",            "Maximum",       None),
    "P2792":  ("DK Sir_T",            "ThanosNF",      None),
    "P3013":  ("Deepesh_1",           "NA",            None),
    "P3112":  ("Deepesh_2",           "NA",            None),
    "P3361":  ("Nishaanth",           "NA",            0),    # no fund denominator
}

DEEPESH_CODES = {"P3013", "P3112", "P3361"}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def safe_float(v, default=0.0):
    try:
        return float(v) if v is not None else default
    except (ValueError, TypeError):
        return default

def safe_int(v, default=0):
    try:
        return int(round(float(v))) if v is not None else default
    except (ValueError, TypeError):
        return default

def fmt_num(v):
    if v == 0:
        return "0"
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(round(v, 4))

def is_data_row(ws, r):
    code = ws.cell(r, COL_CODE).value
    if code is None:
        return False
    code = str(code).strip()
    if code.upper() in ("CODE", "", "NONE"):
        return False
    if not re.match(r'^[A-Z0-9]{3,10}$', code.upper()):
        return False
    return True

# ── EXTRACTION ────────────────────────────────────────────────────────────────
def extract_accounts(excel_path):
    print(f"[1/4] Reading {os.path.basename(excel_path)} ...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[SHEET_NAME]

    accounts = []
    for r in range(DATA_START_ROW, ws.max_row + 1):
        if not is_data_row(ws, r):
            continue

        code     = str(ws.cell(r, COL_CODE).value).strip()
        raw_name = str(ws.cell(r, COL_NAME).value or "").strip()
        raw_strat= str(ws.cell(r, COL_STRATEGY).value or "").strip()
        fund     = safe_float(ws.cell(r, COL_FUND).value)
        apr      = safe_int(ws.cell(r, COL_APR).value)
        apr_roi  = safe_float(ws.cell(r, COL_APR_ROI).value)
        may      = safe_int(ws.cell(r, COL_MAY).value)
        may_roi  = safe_float(ws.cell(r, COL_MAY_ROI).value)
        jun      = safe_int(ws.cell(r, COL_JUN).value)
        jun_roi  = safe_float(ws.cell(r, COL_JUN_ROI).value)

        if code in NAME_OVERRIDES:
            ov = NAME_OVERRIDES[code]
            disp_name = ov[0]
            strat     = ov[1]
            if ov[2] is not None:
                fund = ov[2]
        else:
            disp_name = raw_name
            strat     = raw_strat

        accounts.append({
            "code":       code,
            "name":       disp_name,
            "strategy":   strat,
            "fund":       fund,
            "apr":        apr,
            "apr_roi":    round(apr_roi, 4),
            "may":        may,
            "may_roi":    round(may_roi, 4),
            "jun":        jun,
            "jun_roi":    round(jun_roi, 4),
            "is_deepesh": code in DEEPESH_CODES,
        })

    print(f"    Extracted {len(accounts)} accounts.")
    return accounts

# ── JS RAW ARRAY BUILDER ──────────────────────────────────────────────────────
def build_raw_js(accounts):
    lines = []
    in_deepesh = False
    for i, acc in enumerate(accounts):
        comma = "," if i < len(accounts) - 1 else ""

        if acc["is_deepesh"] and not in_deepesh:
            lines.append("// Deepesh group")
            in_deepesh = True

        # Escape single quotes in name (JS string uses double quotes so safe)
        name_js  = acc["name"].replace('"', '\\"')
        strat_js = acc["strategy"].replace('"', '\\"')

        line = (
            f'{{ code:"{acc["code"]}", name:"{name_js}", strategy:"{strat_js}", '
            f'fund:{fmt_num(acc["fund"])}, '
            f'apr:{acc["apr"]}, apr_roi:{fmt_num(acc["apr_roi"])}, '
            f'may:{acc["may"]}, may_roi:{fmt_num(acc["may_roi"])}, '
            f'jun:{acc["jun"]}, jun_roi:{fmt_num(acc["jun_roi"])} }}{comma}'
        )
        lines.append(line)

    return lines

# ── FILE UPDATER ──────────────────────────────────────────────────────────────
# Matches both:   const RAW = [...]
#            and: export const RAW = [...]
RAW_PATTERN = re.compile(
    r'((?:export\s+)?const\s+RAW\s*=\s*\[)(.*?)(\];)',
    re.DOTALL
)

def update_file(filepath, new_raw_lines, label, is_data_js=False):
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    raw_body = "\n".join(new_raw_lines)

    if is_data_js and not RAW_PATTERN.search(content):
        # Rebuild data.js from scratch
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_content = (
            f'// Raw account data from "{EXCEL_FILE}" (Q1: Apr-Jun 2026).\n'
            f'// Auto-updated: {ts}\n'
            f'export const RAW = [\n{raw_body}\n];\n'
        )
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(new_content)
        print(f"    Rebuilt {label} ({len(new_raw_lines)} data lines).")
        return True

    new_content, count = RAW_PATTERN.subn(
        lambda m: f"{m.group(1)}\n{raw_body}\n{m.group(3)}",
        content,
        count=1
    )
    if count == 0:
        print(f"    [WARN] RAW array not found in {label} — skipping.")
        return False

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"    Updated {label} ({len(new_raw_lines)} data lines).")
    return True

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    base = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base, EXCEL_FILE)

    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found:\n  {excel_path}")
        sys.exit(1)

    # 1. Extract
    accounts = extract_accounts(excel_path)

    # 2. Build JS
    print("[2/4] Building JS RAW array ...")
    raw_lines = build_raw_js(accounts)

    # 3. Update files
    print("[3/4] Updating dashboard files ...")
    targets = [
        ("index.html",                                    False),
        ("indx.html",                                     False),
        (os.path.join("quant-dashboard", "src", "data.js"), True),
    ]
    updated = 0
    for relpath, is_data_js in targets:
        full = os.path.join(base, relpath)
        label = os.path.basename(relpath)
        if not os.path.exists(full):
            print(f"    [SKIP] {label} not found.")
            continue
        if update_file(full, raw_lines, label, is_data_js):
            updated += 1

    # 4. Summary
    total_apr = sum(a["apr"] for a in accounts)
    total_may = sum(a["may"] for a in accounts)
    total_jun = sum(a["jun"] for a in accounts)
    total_q1  = total_apr + total_may + total_jun

    print(f"\n[4/4] Done -- {updated} file(s) updated.  {len(accounts)} accounts.")
    print(f"      Apr P&L: Rs. {total_apr:>14,.0f}")
    print(f"      May P&L: Rs. {total_may:>14,.0f}")
    print(f"      Jun P&L: Rs. {total_jun:>14,.0f}")
    print(f"      Q1  P&L: Rs. {total_q1:>14,.0f}")

if __name__ == "__main__":
    main()
