#!/usr/bin/env python3
"""Klein Finance - Monthly Updater v5.0
Uses xlwings to write to Excel while it's open.
Excel never closes. Button/macro always survives."""
import os, sys, re, shutil, time
from pathlib import Path
from datetime import datetime

# Check dependencies
missing = []
try: import pandas as pd
except: missing.append("pandas")
try: import xlwings as xw
except: missing.append("xlwings")
if missing:
    print("Installing: " + " ".join(missing))
    os.system("python -m pip install " + " ".join(missing) + " --break-system-packages -q")
    import pandas as pd
    import xlwings as xw

BASE    = Path(__file__).parent
MONTHLY = BASE / "monthly"
BACKUPS = BASE / "backups"
EXCEL   = BASE / "\u05de\u05d0\u05d6\u05df_\u05e7\u05dc\u05d9\u05d9\u05df.xlsm"

ANTHROPIC_KEY = ""
key_file = BASE / "api_key.txt"
if key_file.exists():
    ANTHROPIC_KEY = key_file.read_text(encoding="utf-8").strip()

G="\033[32m"; Y="\033[33m"; C="\033[36;1m"; X="\033[0m"
def ok(t):   print(f"  {G}\u2713{X} {t}")
def warn(t): print(f"  {Y}\u26a0{X} {t}")
def hdr(t):  print(f"\n{C}\u2500\u2500 {t} \u2500\u2500{X}")

def num(val):
    try: return float(str(val).replace(",","").replace("\u20aa","").replace("$","").replace("%","").replace(" ","").strip())
    except: return 0.0

def detect_type(path):
    name = path.name

    # --- Filename-based detection (fast, reliable) ---
    if "\u05e2\u05d5\u05e9" in name or "\u05dc\u05d0\u05d5\u05de\u05d9" in name: return "bank"
    if "\u05d4\u05ea\u05de\u05d5\u05e0\u05d4 \u05d4\u05de\u05dc\u05d0\u05d4" in name:
        return "pension_liat" if "(11)" in name else "pension_dror"
    if "\u05d0\u05d7\u05d6\u05e7\u05d5\u05ea" in name: return "invest"
    if "\u05e8\u05d9\u05db\u05d5\u05d6" in name and "\u05d9\u05ea\u05e8\u05d5\u05ea" in name: return "balance"
    if "5647" in name or "\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8" in name.lower(): return "isracard"
    if "transaction-details" in name.lower(): return "credit"

    # --- Content-based detection (fallback when filename changed) ---
    try:
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, nrows=5, engine=engine)
        text = " ".join(str(v) for row in df.values for v in row if str(v) != "nan")

        if "\u05d9\u05ea\u05e8\u05d4 \u05de\u05e6\u05d8\u05d1\u05e8\u05ea" in text: return "bank"
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05db\u05d9\u05e9\u05d4" in text and "\u05e9\u05dd \u05d1\u05d9\u05ea \u05e2\u05e1\u05e7" in text: return "isracard"
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e2\u05e1\u05e7\u05d4" in text and "\u05e1\u05d5\u05d2 \u05e2\u05e1\u05e7\u05d4" in text: return "credit"
        if "\u05e9\u05dd \u05d4\u05e0\u05d9\u05d9\u05e8" in text or "\u05de\u05d1\u05d8 \u05d0\u05d9\u05e9\u05d9" in text: return "invest"

        # Pension: check if file has the ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¤ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¦ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ©ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ sheet
        xl = pd.ExcelFile(path, engine=engine)
        if "\u05e4\u05e8\u05d8\u05d9 \u05d4\u05de\u05d5\u05e6\u05e8\u05d9\u05dd \u05e9\u05dc\u05d9" in xl.sheet_names:
            # Determine Dror vs Liat by row count in the pension sheet
            df2 = pd.read_excel(path, sheet_name="\u05e4\u05e8\u05d8\u05d9 \u05d4\u05de\u05d5\u05e6\u05e8\u05d9\u05dd \u05e9\u05dc\u05d9", header=None, engine=engine)
            return "pension_liat" if len(df2) <= 8 else "pension_dror"
    except: pass
    return None


def find_files():
    hdr("Scanning monthly folder")
    MONTHLY.mkdir(exist_ok=True)
    found = {}
    for f in list(MONTHLY.glob("*.xls")) + list(MONTHLY.glob("*.xlsx")) + list(MONTHLY.glob("*.xlsm")):
        ft = detect_type(f)
        if ft and ft not in found:
            found[ft] = f; ok(f"Found {ft}: {f.name}")
        elif not ft:
            warn(f"Could not identify: {f.name}")
    for f in list(MONTHLY.glob("*.png")) + list(MONTHLY.glob("*.jpg")) + list(MONTHLY.glob("*.jpeg")):
        found["rsu_image"] = f; ok(f"Found RSU image: {f.name}")
    return found

def parse_pension(path):
    SHEET = "\u05e4\u05e8\u05d8\u05d9 \u05d4\u05de\u05d5\u05e6\u05e8\u05d9\u05dd \u05e9\u05dc\u05d9"
    df = None
    try:
        df = pd.read_excel(path, sheet_name=SHEET, header=None, engine='xlrd')
        ok(f"  Pension: read sheet '{SHEET}', shape={df.shape}")
    except Exception as e:
        warn(f"  Pension sheet error: {e}, trying default sheet")
        try: df = pd.read_excel(path, header=None, engine='xlrd')
        except: pass
    if df is None: return {}
    
    pension = provident = 0
    products_list = []
    for i, row in df.iterrows():
        if i == 0: continue
        row = list(row)
        if not row[0] or str(row[0]) == "nan": continue
        name = str(row[0])
        t = float(row[4]) if len(row) > 4 and str(row[4]) != "nan" else 0
        if t < 100: continue
        products_list.append({"product": name, "total": t})
        if "\u05e4\u05e0\u05e1\u05d9\u05d4" in name: pension += t
        elif "\u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea" in name or "\u05e7\u05e8\u05df" in name: provident += t
    ok(f"  pension={pension:,.0f}, provident={provident:,.0f}")
    return {"pension": pension, "provident": provident, "products": products_list}


def parse_bank(path):
    engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
    df = pd.read_excel(path, header=None, engine=engine)
    try: return {"balance": float(str(df.iloc[2,0]).replace("\u20aa","").replace(",","").replace(" ",""))}
    except: return {}

def parse_invest(path):
    for enc in ["windows-1255", "utf-8"]:
        try:
            for df in pd.read_html(str(path), encoding=enc):
                for _, row in df.iterrows():
                    for val in row:
                        v = num(str(val))
                        if 500000 < v < 20000000: return {"total": v}
        except: pass
    try:
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine)
        return {"total": num(str(df.iloc[2,3]).replace(",",""))}
    except: return {}

def parse_rsu(path):
    if not ANTHROPIC_KEY:
        warn("No API key. Enter RSU manually:")
        try:
            u = float(input("  Unvested ($): ").replace(",","").replace("$",""))
            a = float(input("  Available ($): ").replace(",","").replace("$",""))
            return {"unvested": u, "available": a}
        except: return {}
    try:
        import base64, json as _j, urllib.request
        with open(path,"rb") as f: b64 = base64.b64encode(f.read()).decode()
        mime = "image/png" if str(path).lower().endswith(".png") else "image/jpeg"
        payload = {"model":"claude-sonnet-4-20250514","max_tokens":100,
            "messages":[{"role":"user","content":[
                {"type":"image","source":{"type":"base64","media_type":mime,"data":b64}},
                {"type":"text","text":"RSU equity overview screenshot. Find Unvested and Shares/Available dollar amounts. JSON only: {\"unvested\": 187148, \"available\": 170600}"}
            ]}]}
        req = urllib.request.Request("https://api.anthropic.com/v1/messages",
            data=_j.dumps(payload).encode(),
            headers={"Content-Type":"application/json","anthropic-version":"2023-06-01","x-api-key":ANTHROPIC_KEY})
        with urllib.request.urlopen(req, timeout=30) as r:
            text = _j.loads(r.read())["content"][0]["text"].strip()
            m = re.search(r"\{[^}]+\}", text)
            if m: return _j.loads(m.group())
    except Exception as e: warn(f"RSU API: {e}")
    warn("Enter RSU manually:")
    try:
        u = float(input("  Unvested ($): ").replace(",","").replace("$",""))
        a = float(input("  Available ($): ").replace(",","").replace("$",""))
        return {"unvested": u, "available": a}
    except: return {}


def update_excel_xlwings(values, found):
    hdr("Updating Excel with xlwings")
    try:
        app = xw.apps.active
        wb = None
        for book in app.books:
            if "\u05de\u05d0\u05d6\u05df" in book.name or "klein" in book.name.lower():
                wb = book; break
        if wb is None:
            warn("Excel workbook not found - opening it")
            wb = app.books.open(str(EXCEL))
        
        dash = wb.sheets["\u05d3\u05e9\u05d1\u05d5\u05e8\u05d3"]
        
        def write(cell, val, label):
            if val and val != 0:
                dash[cell].value = val
                ok(f"  {label} = {val:,.0f}")
        
        write("D10", values.get("dror_pension",0),   "D10 \u05e4\u05e0\u05e1\u05d9\u05d4 \u05d3\u05e8\u05d5\u05e8")
        write("D11", values.get("liat_pension",0),   "D11 \u05e4\u05e0\u05e1\u05d9\u05d4 \u05dc\u05d9\u05d0\u05ea")
        write("D12", values.get("dror_provident",0), "D12 \u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05d3\u05e8\u05d5\u05e8")
        write("D13", values.get("liat_provident",0), "D13 \u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05dc\u05d9\u05d0\u05ea")
        write("D14", values.get("invest",0),         "D14 \u05ea\u05d9\u05e7 \u05d4\u05e9\u05e7\u05e2\u05d5\u05ea")
        write("D18", values.get("bank",0),           "D18 \u05e2\u05d5\"\u05e9")


        rsu_avail = values.get("rsu_available", 0)
        rsu_unves = values.get("rsu_unvested", 0)
        if rsu_avail or rsu_unves:
            rsu_sheet = wb.sheets["ALIGN RSU"]
            rsu_sheet["H13"].value = rsu_avail
            rsu_sheet["H14"].value = rsu_unves
            ok(f"  RSU H13={rsu_avail}, H14={rsu_unves}")

        dash["A2"].value = f"\u05e2\u05d3\u05db\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df: {datetime.now().strftime('%d/%m/%Y')}"
        wb.save()
        # Update transaction sheets
        hdr("Updating transaction sheets")
        if found.get("credit"): update_max_sheets(wb, found["credit"])
        if found.get("bank"): update_bank_sheet(wb, found["bank"])
        if found.get("isracard"): update_isracard_sheet(wb, found["isracard"])
        # Force recalculation for graphs
        wb.app.calculate()
        ok("Excel saved - button and macro preserved!")
        save_history_snapshot(wb)
        return True
    except Exception as e:
        warn(f"xlwings error: {e}")
        return False

def update_max_sheets(wb, credit_path):
    """Replace MAX credit card sheets with downloaded data."""
    for sname in ["\u05e2\u05e1\u05e7\u05d0\u05d5\u05ea \u05d1\u05de\u05d5\u05e2\u05d3 \u05d4\u05d7\u05d9\u05d5\u05d1", '\u05e2\u05e1\u05e7\u05d0\u05d5\u05ea \u05d7\u05d5"\u05dc \u05d5\u05de\u05d8"\u05d7']:
        try:
            df = pd.read_excel(credit_path, sheet_name=sname, header=None, engine='openpyxl')
            ws = wb.sheets[sname]
            ws.clear_contents()
            for r, row in enumerate(df.values.tolist(), start=1):
                for c, val in enumerate(row, start=1):
                    if val is not None and str(val) != 'nan':
                        ws.cells(r, c).value = val
            ok(f"  Updated {sname}: {len(df)} rows")
        except Exception as e:
            warn(f"  MAX error ({sname}): {e}")


def update_bank_sheet(wb, bank_path):
    """Replace raw data in cols A-E, preserve and extend formulas in cols F-I."""
    import re as re2
    try:
        df = pd.read_excel(bank_path, sheet_name='\u05e2\u05d5\u05e9', header=None, engine='openpyxl')
        ws = wb.sheets['\u05e2\u05d5\u05e9']
        # Get formula templates from row 3
        templates = {}
        for col in [6, 7, 8, 9]:
            f = ws.cells(3, col).formula
            if f: templates[col] = f
        # Data starts at row index 2 (skip 2 header rows)
        data = df.iloc[2:].values.tolist()
        # Clear rows 3+ all cols
        used = ws.api.UsedRange.Rows.Count
        if used >= 3:
            ws.range(ws.cells(3, 1), ws.cells(used + 10, 9)).clear_contents()
        # Write data to cols A-E
        for r_idx, row in enumerate(data, start=3):
            for c_idx, val in enumerate(row[:5], start=1):
                if val is not None and str(val) != 'nan':
                    ws.cells(r_idx, c_idx).value = val
        # Extend formulas F-I
        for r_idx in range(3, 3 + len(data)):
            for col, tmpl in templates.items():
                new_f = re2.sub(r'([A-Z]+)3(?=[^0-9]|$)', lambda m: m.group(1) + str(r_idx), tmpl)
                ws.cells(r_idx, col).formula = new_f
        ok(f"  Updated \u05e2\u05d5\u05e9: {len(data)} rows")
    except Exception as e:
        warn(f"  \u05e2\u05d5\u05e9 error: {e}")


def update_isracard_sheet(wb, isr_path):
    """Replace Isracard data cols A-H, preserve and extend col I category formula."""
    import re as re2
    try:
        df = pd.read_excel(isr_path, sheet_name='\u05e4\u05d9\u05e8\u05d5\u05d8 \u05e2\u05e1\u05e7\u05d0\u05d5\u05ea', header=None, engine='openpyxl')
        ws = wb.sheets['\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8']
        # Get category formula template from row 2
        cat_formula = ws.cells(2, 9).formula or ws.cells(2, 9).formula_array
        # Find data start row in downloaded file
        data_start = 0
        for i, row in df.iterrows():
            if '\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05db\u05d9\u05e9\u05d4' in str(list(row)):
                data_start = i + 1; break
        # Collect valid data rows (date pattern DD.MM.YY)
        data_rows = []
        for i in range(data_start, len(df)):
            row = list(df.iloc[i])
            first = str(row[0])
            if any(x in first for x in ['\u05e1\u05d4"\u05db', '\u05ea\u05e0\u05d0\u05d9\u05dd']) or first.strip() in ('', 'nan'): continue
            if re2.match(r'\d{2}\.\d{2}\.\d{2}', first): data_rows.append(row)
        # Clear rows 2+ all cols
        used = ws.api.UsedRange.Rows.Count
        if used >= 2:
            ws.range(ws.cells(2, 1), ws.cells(used + 5, 9)).clear_contents()
        # Write data cols A-H
        for r_idx, row in enumerate(data_rows, start=2):
            for c_idx, val in enumerate(row[:8], start=1):
                if val is not None and str(val) != 'nan':
                    ws.cells(r_idx, c_idx).value = val
        # Extend category formula col I
        if cat_formula:
            for r_idx in range(2, 2 + len(data_rows)):
                new_f = re2.sub(r'([A-Z]+)2(?=[^0-9]|$)', lambda m: m.group(1) + str(r_idx), cat_formula)
                ws.cells(r_idx, 9).formula = new_f
        ok(f"  Updated \u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8: {len(data_rows)} rows")
    except Exception as e:
        warn(f"  \u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8 error: {e}")


def save_history_snapshot(wb):
    hdr("Saving history snapshot")
    try:
        hist = wb.sheets["\u05d4\u05d9\u05e1\u05d8\u05d5\u05e8\u05d9\u05d4"]
        target_pos = hist["P2"].value
        if not target_pos or not isinstance(target_pos, (int, float)):
            warn("Could not determine target month column"); return
        target_col = int(target_pos) + 1
        month_name = hist.cells(4, target_col).value or str(target_col)
        existing = hist.cells(6, target_col).value
        if existing and existing != 0 and existing != "":
            import ctypes
            res = ctypes.windll.user32.MessageBoxW(0,
                f"\u05d4\u05d7\u05d5\u05d3\u05e9 {month_name} \u05db\u05d1\u05e8 \u05de\u05db\u05d9\u05dc \u05e0\u05ea\u05d5\u05e0\u05d9\u05dd ({int(existing):,}).\n\u05dc\u05e2\u05d3\u05db\u05df \u05d1\u05db\u05dc \u05d6\u05d0\u05ea?",
                "Klein Finance", 4)
            if res != 6:
                warn(f"History skipped for {month_name}"); return
        rows = [6,7,8,10,11,12,14,15,16,18,19,20,22,23,25,27,28]
        for row in rows:
            val = hist.cells(row, 14).value
            if val is not None:
                hist.cells(row, target_col).value = val
        hist.cells(26, target_col).value = datetime.now()
        ok(f"History saved for {month_name}")
    except Exception as e:
        warn(f"History error: {e}")


def main():
    print(f"\n{C}  Klein Family Finance v5.0{X}")
    print(f"  {datetime.now().strftime('%d %B %Y, %H:%M')}\n")
    
    if not EXCEL.exists():
        print(f"ERROR: Excel not found: {EXCEL}"); sys.exit(1)

    found = find_files()
    
    hdr("Reading files")
    dror   = parse_pension(found["pension_dror"]) if "pension_dror" in found else {}
    liat   = parse_pension(found["pension_liat"]) if "pension_liat" in found else {}
    bank   = parse_bank(found["bank"])             if "bank"         in found else {}
    invest = parse_invest(found["invest"])          if "invest"       in found else {}
    rsu    = parse_rsu(found["rsu_image"])          if "rsu_image"    in found else {}

    ok(f"Dror: pension={dror.get('pension',0):,.0f}, provident={dror.get('provident',0):,.0f}")
    ok(f"Liat: pension={liat.get('pension',0):,.0f}, provident={liat.get('provident',0):,.0f}")
    ok(f"Bank: {bank.get('balance',0):,.0f}")
    ok(f"Invest: {invest.get('total',0):,.0f}")
    ok(f"RSU: available={rsu.get('available',0)}, unvested={rsu.get('unvested',0)}")

    hdr("Backup")
    BACKUPS.mkdir(exist_ok=True)
    shutil.copy2(EXCEL, BACKUPS / f"\u05de\u05d0\u05d6\u05df_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsm")
    ok("Backup created")

    values = {
        "dror_pension":   dror.get("pension", 0),
        "liat_pension":   liat.get("pension", 0),
        "dror_provident": dror.get("provident", 0),
        "liat_provident": liat.get("provident", 0),
        "dror_products":  dror.get("products", []),
        "liat_products":  liat.get("products", []),
        "invest":         invest.get("total", 0),
        "bank":           bank.get("balance", 0),
        "rsu_available":  rsu.get("available", 0),
        "rsu_unvested":   rsu.get("unvested", 0),
    }

    success = update_excel_xlwings(values, found)
    if success:
        print(f"\n{G}  Done! Excel updated. Button and macro preserved.{X}\n")
    else:
        print(f"\n{Y}  Update failed. Check the errors above.{X}\n")

if __name__ == "__main__":
    main()
