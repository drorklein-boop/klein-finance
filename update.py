#!/usr/bin/env python3
"""Klein Finance - Monthly Updater v4.0
Updates ONLY the 8 input cells in the dashboard. Never touches sheets."""
import os, sys, re, shutil, time
from pathlib import Path
from datetime import datetime

missing = []
try: import pandas as pd
except: missing.append("pandas")
try: import openpyxl
except: missing.append("openpyxl")
if missing:
    print("Run: python -m pip install " + " ".join(missing))
    input("Press Enter..."); sys.exit(1)

from openpyxl import load_workbook

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

# ââ File detection ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def detect_type(path):
    name = path.name
    if "\u05e2\u05d5\u05e9" in name or "\u05dc\u05d0\u05d5\u05de\u05d9" in name: return "bank"
    if "\u05d4\u05ea\u05de\u05d5\u05e0\u05d4 \u05d4\u05de\u05dc\u05d0\u05d4" in name:
        return "pension_liat" if "(11)" in name else "pension_dror"
    if "\u05d0\u05d7\u05d6\u05e7\u05d5\u05ea" in name: return "invest"
    if "5647" in name or "\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8" in name.lower(): return "isracard"
    if "transaction-details" in name.lower(): return "credit"
    try:
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, nrows=6, engine=engine)
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
    found = {k: None for k in ["bank","credit","pension_dror","pension_liat","invest","rsu_image"]}
    for f in list(MONTHLY.glob("*.xls"))+list(MONTHLY.glob("*.xlsx"))+list(MONTHLY.glob("*.xlsm")):
        ft = detect_type(f)
        if ft and found.get(ft) is None: found[ft]=f; ok(f"Found {ft}: {f.name}")
        elif ft is None: warn(f"Could not identify: {f.name}")
    for f in list(MONTHLY.glob("*.png"))+list(MONTHLY.glob("*.jpg"))+list(MONTHLY.glob("*.jpeg")):
        found["rsu_image"]=f; ok(f"Found RSU image: {f.name}")
    return found

# ââ Parsers âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def parse_bank(path):
    engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
    df = pd.read_excel(path, header=None, engine=engine)
    balance = 0
    try: balance = float(str(df.iloc[2,0]).replace("\u20aa","").replace(",","").replace(" ",""))
    except: pass
    return {"balance": balance}

def parse_pension(path):
    df = None
    # Try HTML first (Mislaka files are often HTML disguised as .xls)
    for enc in ["windows-1255", "utf-8", "iso-8859-8"]:
        try:
            tables = pd.read_html(str(path), encoding=enc)
            if tables: df = tables[0]; break
        except: pass
    # Fall back to Excel engines
    if df is None:
        for engine in ["xlrd", "openpyxl"]:
            try: df = pd.read_excel(path, header=None, engine=engine); break
            except: pass
    if df is None: return {}
    pension = provident = 0
    for i, row in df.iterrows():
        row = list(row)
        if not row[0] or str(row[0]) == "nan": continue
        # Find the total savings column - look for values > 1000
        t = 0
        for col_idx in [4, 3, 5, 2]:
            if len(row) > col_idx:
                candidate = num(str(row[col_idx]))
                if 1000 < candidate < 10000000:
                    t = candidate; break
        if t == 0: continue
        name = str(row[0])
        if "\u05e4\u05e0\u05e1\u05d9\u05d4" in name: pension += t
        elif "\u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea" in name or "\u05e7\u05e8\u05df" in name: provident += t
    return {"pension": pension, "provident": provident}

def parse_invest(path):
    try:
        for enc in ["windows-1255","utf-8"]:
            try:
                tables = pd.read_html(str(path), encoding=enc)
                if tables:
                    for _, row in tables[0].iterrows():
                        for val in row:
                            v = num(str(val))
                            if 500000 < v < 20000000: return {"total": v}
            except: pass
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, engine=engine)
        total = num(str(df.iloc[2,3]).replace(",",""))
        return {"total": total}
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
                {"type":"text","text":"RSU screenshot. Find Unvested and Available/Shares dollar amounts. JSON only: {\"unvested\": 187148, \"available\": 170600}"}
            ]}]}
        req = urllib.request.Request("https://api.anthropic.com/v1/messages",
            data=_j.dumps(payload).encode(),
            headers={"Content-Type":"application/json","anthropic-version":"2023-06-01","x-api-key":ANTHROPIC_KEY})
        with urllib.request.urlopen(req, timeout=30) as r:
            res = _j.loads(r.read())
            text = res["content"][0]["text"].strip()
            m = re.search(r"\{[^}]+\}", text)
            if m: return _j.loads(m.group())
    except Exception as e: warn(f"RSU API error: {e}")
    warn("Enter RSU manually:")
    try:
        u = float(input("  Unvested ($): ").replace(",","").replace("$",""))
        a = float(input("  Available ($): ").replace(",","").replace("$",""))
        return {"unvested": u, "available": a}
    except: return {}

# ââ Main update â ONLY writes to 8 cells âââââââââââââââââââââââââââââââââââââ
def update_excel(found):
    hdr("Closing Excel")
    os.system("taskkill /f /im excel.exe 2>nul")
    time.sleep(3)

    hdr("Reading files")
    bank    = parse_bank(found["bank"]) if found.get("bank") else {}
    dror    = parse_pension(found["pension_dror"]) if found.get("pension_dror") else {}
    liat    = parse_pension(found["pension_liat"]) if found.get("pension_liat") else {}
    invest  = parse_invest(found["invest"]) if found.get("invest") else {}
    rsu     = parse_rsu(found["rsu_image"]) if found.get("rsu_image") else {}

    ok(f"Bank balance: {bank.get('balance',0):,.0f}")
    ok(f"Dror pension: {dror.get('pension',0):,.0f} | provident: {dror.get('provident',0):,.0f}")
    ok(f"Liat pension: {liat.get('pension',0):,.0f} | provident: {liat.get('provident',0):,.0f}")
    ok(f"Invest: {invest.get('total',0):,.0f}")
    ok(f"RSU: available={rsu.get('available',0)}, unvested={rsu.get('unvested',0)}")

    hdr("Backup")
    BACKUPS.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    shutil.copy2(EXCEL, BACKUPS / f"\u05de\u05d0\u05d6\u05df_{stamp}.xlsm")
    ok("Backup created")

    hdr("Updating Excel â 8 cells only")
    wb = load_workbook(EXCEL, keep_vba=True)
    dash = wb["\u05d3\u05e9\u05d1\u05d5\u05e8\u05d3"]

    def write(row, col, val, label):
        if val and val != 0:
            dash.cell(row=row, column=col).value = val
            ok(f"  {label} = {val:,.0f}")

    write(10, 4, dror.get("pension",0),   "D10 \u05e4\u05e0\u05e1\u05d9\u05d4 \u05d3\u05e8\u05d5\u05e8")
    write(11, 4, liat.get("pension",0),   "D11 \u05e4\u05e0\u05e1\u05d9\u05d4 \u05dc\u05d9\u05d0\u05ea")
    write(12, 4, dror.get("provident",0), "D12 \u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05d3\u05e8\u05d5\u05e8")
    write(13, 4, liat.get("provident",0), "D13 \u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05dc\u05d9\u05d0\u05ea")
    write(14, 4, invest.get("total",0),   "D14 \u05ea\u05d9\u05e7 \u05d4\u05e9\u05e7\u05e2\u05d5\u05ea")
    write(18, 4, bank.get("balance",0),   'D18 \u05e2\u05d5"\u05e9')

    if rsu.get("available") or rsu.get("unvested"):
        rsu_ws = wb["ALIGN RSU"]
        rsu_ws.cell(row=13, column=8).value = rsu.get("available", 0)
        rsu_ws.cell(row=14, column=8).value = rsu.get("unvested", 0)
        ok(f"  ALIGN RSU H13={rsu.get('available',0)}, H14={rsu.get('unvested',0)}")

    dash.cell(row=2, column=1).value = f"\u05e2\u05d3\u05db\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df: {datetime.now().strftime('%d/%m/%Y')}"
    wb.save(EXCEL)
    ok("Excel saved")

def main():
    print(f"\n{C}  Klein Family Finance v4.0{X}")
    print(f"  {datetime.now().strftime('%d %B %Y, %H:%M')}\n")
    if not EXCEL.exists():
        print(f"ERROR: Excel not found at {EXCEL}"); sys.exit(1)
    found = find_files()
    update_excel(found)
    hdr("Opening Excel")
    os.startfile(str(EXCEL))
    ok("Done!")
    print()

if __name__ == "__main__":
    main()
