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
    if "\u05e2\u05d5\u05e9" in name or "\u05dc\u05d0\u05d5\u05de\u05d9" in name: return "bank"
    if "\u05d4\u05ea\u05de\u05d5\u05e0\u05d4 \u05d4\u05de\u05dc\u05d0\u05d4" in name:
        return "pension_liat" if "(11)" in name else "pension_dror"
    if "\u05d0\u05d7\u05d6\u05e7\u05d5\u05ea" in name: return "invest"
    if "\u05e8\u05d9\u05db\u05d5\u05d6" in name and "\u05d9\u05ea\u05e8\u05d5\u05ea" in name: return "balance"
    if "5647" in name or "\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8" in name.lower(): return "isracard"
    if "transaction-details" in name.lower(): return "credit"
    try:
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, nrows=5, engine=engine)
        text = " ".join(str(v) for row in df.values for v in row if str(v)!="nan")
        if "\u05d9\u05ea\u05e8\u05d4 \u05de\u05e6\u05d8\u05d1\u05e8\u05ea" in text: return "bank"
        if "\u05e9\u05dd \u05de\u05d5\u05e6\u05e8" in text: return "pension_dror"
        if "\u05e9\u05dd \u05d4\u05e0\u05d9\u05d9\u05e8" in text or "\u05de\u05d1\u05d8 \u05d0\u05d9\u05e9\u05d9" in text: return "invest"
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05db\u05d9\u05e9\u05d4" in text: return "isracard"
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
    df = None
    # Try HTML first (Mislaka files often HTML disguised as .xls)
    for enc in ["windows-1255", "utf-8", "iso-8859-8"]:
        try:
            tables = pd.read_html(str(path), encoding=enc)
            if tables: df = tables[0]; break
        except: pass
    # Fall back to Excel
    if df is None:
        for engine in ["xlrd", "openpyxl"]:
            try: df = pd.read_excel(path, header=None, engine=engine); break
            except: pass
    if df is None: return {}

    # Find column indices by scanning for Hebrew headers
    product_col = total_col = -1
    header_row = 0
    for i, row in df.iterrows():
        row_vals = [str(v) for v in row]
        row_text = " ".join(row_vals)
        # Look for the savings column header
        if any(k in row_text for k in ["\u05e1\u05da \u05d4\u05db\u05dc", "\u05d7\u05d9\u05e1\u05db\u05d5\u05df"]):
            header_row = i
            for j, val in enumerate(row_vals):
                if "\u05e9\u05dd \u05de\u05d5\u05e6\u05e8" in val or "\u05de\u05d5\u05e6\u05e8" in val:
                    product_col = j
                if "\u05e1\u05da \u05d4\u05db\u05dc" in val or ("\u05d7\u05d9\u05e1\u05db\u05d5\u05df" in val and "\u05e6\u05e4\u05d5\u05d9" not in val):
                    total_col = j
            break

    if product_col == -1: product_col = 0
    if total_col == -1: total_col = 4

    ok(f"  Pension columns: product={product_col}, total={total_col}, header_row={header_row}")

    pension = provident = 0
    products_list = []
    for i, row in df.iterrows():
        if i <= header_row: continue
        row = list(row)
        if len(row) <= max(product_col, total_col): continue
        name = str(row[product_col])
        if not name or name == "nan": continue
        t = num(str(row[total_col]))
        if t == 0 or t < 100: continue
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

def update_pension_table(wb, table_name, sheet_name, products):
    """Update Excel table with pension products so formulas in D10-D13 recalculate."""
    try:
        ws = wb.sheets[sheet_name]
        # Find the table and update rows starting from row 2
        for i, p in enumerate(products, 2):
            ws.cells(i, 1).value = p.get("product", "")
            ws.cells(i, 5).value = p.get("total", 0)
        ok(f"  Updated table: {table_name}")
    except Exception as e:
        warn(f"  Table update error ({table_name}): {e}")

def update_excel_xlwings(values):
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

        # Update pension tables so D10-D13 formulas recalculate correctly
        dror_products = values.get("dror_products", [])
        liat_products = values.get("liat_products", [])
        if dror_products:
            update_pension_table(wb, "Tbl_ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¡ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ§ÃÂÃÂÃÂÃÂ_ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨", "ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¨ - ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¡ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ§ÃÂÃÂÃÂÃÂ", dror_products)
        if liat_products:
            update_pension_table(wb, "Tbl_ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¡ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ§ÃÂÃÂÃÂÃÂ_ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂª", "ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂª - ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ¡ÃÂÃÂÃÂÃÂÃÂÃÂÃÂÃÂ§ÃÂÃÂÃÂÃÂ", liat_products)

        rsu_avail = values.get("rsu_available", 0)
        rsu_unves = values.get("rsu_unvested", 0)
        if rsu_avail or rsu_unves:
            rsu_sheet = wb.sheets["ALIGN RSU"]
            rsu_sheet["H13"].value = rsu_avail
            rsu_sheet["H14"].value = rsu_unves
            ok(f"  RSU H13={rsu_avail}, H14={rsu_unves}")

        dash["A2"].value = f"\u05e2\u05d3\u05db\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df: {datetime.now().strftime('%d/%m/%Y')}"
        wb.save()
        ok("Excel saved - button and macro preserved!")
        return True
    except Exception as e:
        warn(f"xlwings error: {e}")
        return False

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

    success = update_excel_xlwings(values)
    if success:
        print(f"\n{G}  Done! Excel updated. Button and macro preserved.{X}\n")
    else:
        print(f"\n{Y}  Update failed. Check the errors above.{X}\n")

if __name__ == "__main__":
    main()
