#!/usr/bin/env python3
"""Klein Finance - Monthly Updater v3.0"""
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
    input("Press Enter...")
    sys.exit(1)

from openpyxl import load_workbook

BASE    = Path(__file__).parent
MONTHLY = BASE / "monthly"
BACKUPS = BASE / "backups"
EXCEL   = BASE / "\u05de\u05d0\u05d6\u05df_\u05e7\u05dc\u05d9\u05d9\u05df.xlsm"

ANTHROPIC_KEY = ""
key_file = BASE / "api_key.txt"
if key_file.exists():
    ANTHROPIC_KEY = key_file.read_text(encoding="utf-8").strip()

G="\033[32m"; Y="\033[33m"; R="\033[31m"; C="\033[36;1m"; X="\033[0m"
def ok(t):   print(f"  {G}\u2713{X} {t}")
def warn(t): print(f"  {Y}\u26a0{X} {t}")
def err(t):  print(f"  {R}\u2717{X} {t}")
def hdr(t):  print(f"\n{C}\u2500\u2500 {t} \u2500\u2500{X}")

def num(val):
    try: return float(str(val).replace(",","").replace("\u20aa","").replace("$","").replace("%","").replace(" ","").strip())
    except: return 0.0

def read_html_xls(path):
    for enc in ["windows-1255","utf-8","iso-8859-8"]:
        try:
            tables = pd.read_html(str(path), encoding=enc)
            if tables: return tables[0]
        except: pass
    return None

def detect_type(path):
    name = path.name
    if "\u05e2\u05d5\u05e9" in name or "\u05dc\u05d0\u05d5\u05de\u05d9 \u05e9\u05e2" in name: return "bank"
    if "\u05d4\u05ea\u05de\u05d5\u05e0\u05d4 \u05d4\u05de\u05dc\u05d0\u05d4" in name:
        return "pension_liat" if "(11)" in name else "pension_dror"
    if "\u05d0\u05d7\u05d6\u05e7\u05d5\u05ea" in name: return "invest"
    if "\u05e8\u05d9\u05db\u05d5\u05d6" in name and "\u05d9\u05ea\u05e8\u05d5\u05ea" in name: return "balance"
    if "5647" in name or "\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8" in name.lower(): return "isracard"
    if "transaction-details" in name.lower():
        try:
            df = pd.read_excel(path, header=None, nrows=10, engine="openpyxl")
            text = " ".join(str(v) for row in df.values for v in row if str(v)!="nan")
            return "foreign" if '\u05d7\u05d5"\u05dc' in text or '\u05de\u05d8"\u05d7' in text else "credit"
        except: return "credit"
    try:
        engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
        df = pd.read_excel(path, header=None, nrows=6, engine=engine)
        text = " ".join(str(v) for row in df.values for v in row if str(v)!="nan")
        if "\u05d9\u05ea\u05e8\u05d4 \u05de\u05e6\u05d8\u05d1\u05e8\u05ea" in text: return "bank"
        if "\u05e9\u05dd \u05de\u05d5\u05e6\u05e8" in text and "\u05e4\u05d5\u05dc\u05d9\u05e1\u05d4" in text:
            df2 = pd.read_excel(path, header=None, engine=engine)
            rows = sum(1 for _,r in df2.iterrows() if any(str(v)!="nan" for v in r))
            return "pension_liat" if rows<=8 else "pension_dror"
        if "\u05e9\u05dd \u05d4\u05e0\u05d9\u05d9\u05e8" in text or "\u05de\u05d1\u05d8 \u05d0\u05d9\u05e9\u05d9" in text: return "invest"
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05db\u05d9\u05e9\u05d4" in text and "\u05e9\u05dd \u05d1\u05d9\u05ea \u05e2\u05e1\u05e7" in text: return "isracard"
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e2\u05e1\u05e7\u05d4" in text and "\u05e1\u05d5\u05d2 \u05e2\u05e1\u05e7\u05d4" in text:
            return "foreign" if "\u05de\u05d8\u05d1\u05e2 \u05e2\u05e1\u05e7\u05d4 \u05de\u05e7\u05d5\u05e8\u05d9" in text else "credit"
    except: pass
    return None

def find_files():
    hdr("Scanning monthly folder")
    MONTHLY.mkdir(exist_ok=True)
    found = {k: None for k in ["bank","credit","foreign","isracard","pension_dror","pension_liat","invest","balance","rsu_image"]}
    for f in list(MONTHLY.glob("*.xls"))+list(MONTHLY.glob("*.xlsx"))+list(MONTHLY.glob("*.xlsm")):
        ft = detect_type(f)
        if ft and found[ft] is None: found[ft]=f; ok(f"Found {ft}: {f.name}")
        elif ft is None: warn(f"Could not identify: {f.name}")
    for f in list(MONTHLY.glob("*.png"))+list(MONTHLY.glob("*.jpg"))+list(MONTHLY.glob("*.jpeg")):
        found["rsu_image"]=f; ok(f"Found RSU image: {f.name}")
    return found

def parse_bank(path):
    engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
    df = pd.read_excel(path, header=None, engine=engine)
    balance=0
    try: balance=float(str(df.iloc[2,0]).replace("\u20aa","").replace(",","").replace(" ",""))
    except: pass
    txs=[]
    for i,row in df.iterrows():
        if i<2: continue
        row=list(row); d=num(row[6]) if len(row)>6 else 0; c=num(row[7]) if len(row)>7 else 0
        if d==0 and c==0: continue
        txs.append({"debit":d,"credit":c})
    return {"balance":balance,"income":sum(t["credit"] for t in txs),"expense":sum(t["debit"] for t in txs),"raw_df":df}

def parse_credit(path):
    engine = "xlrd" if str(path).endswith(".xls") else "openpyxl"
    df=pd.read_excel(path,header=None,engine=engine); txs=[]; hf=False
    for i,row in df.iterrows():
        row=list(row)
        if "\u05ea\u05d0\u05e8\u05d9\u05da \u05e2\u05e1\u05e7\u05d4" in str(row[0]) or "\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05db\u05d9\u05e9\u05d4" in str(row[0]):
            hf=True; continue
        if not hf or not row[0] or str(row[0])=="nan": continue
        a=num(row[5]) if len(row)>5 else num(row[2])
        if a==0: continue
        txs.append({"category":str(row[2] or "\u05e9\u05d5\u05e0\u05d5\u05ea"),"amount":a})
    bc={}
    for t in txs: bc[t["category"]]=bc.get(t["category"],0)+t["amount"]
    return {"transactions":txs,"by_category":bc,"total":sum(bc.values()),"raw_df":df}

def parse_isracard(path):
    engine="xlrd" if str(path).endswith(".xls") else "openpyxl"
    df=pd.read_excel(path,header=None,engine=engine); txs=[]
    for i,row in df.iterrows():
        if i==0: continue
        row=list(row)
        if not row[0] or str(row[0])=="nan": continue
        a=num(row[2])
        if a==0: continue
        cat=str(row[8]) if len(row)>8 and str(row[8])!="nan" else "\u05e9\u05d5\u05e0\u05d5\u05ea"
        txs.append({"category":cat,"amount":a})
    bc={}
    for t in txs: bc[t["category"]]=bc.get(t["category"],0)+t["amount"]
    return {"transactions":txs,"by_category":bc,"total":sum(bc.values()),"raw_df":df}

def parse_pension(path):
    df=None
    for engine in ["xlrd","openpyxl"]:
        try: df=pd.read_excel(path,header=None,engine=engine); break
        except: pass
    if df is None: df=read_html_xls(path)
    if df is None: return {}
    products=[]
    for i,row in df.iterrows():
        if i==0: continue
        row=list(row)
        if not row[0] or str(row[0])=="nan": continue
        t=num(row[4]) if len(row)>4 else 0
        if t==0:
            for v in row[2:8]:
                c=num(str(v))
                if 1000<c<10000000: t=c; break
        if t==0: continue
        products.append({"product":str(row[0]),"total":t,
            "fee_deposit":num(row[11]) if len(row)>11 else 0,
            "fee_accum":num(row[12]) if len(row)>12 else 0,
            "return_ytd":num(row[13]) if len(row)>13 else 0})
    p_kw=["\u05e4\u05e0\u05e1\u05d9\u05d4"]
    h_kw=["\u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea","\u05e7\u05e8\u05df \u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea"]
    pt=sum(p["total"] for p in products if any(k in p["product"] for k in p_kw))
    ht=sum(p["total"] for p in products if any(k in p["product"] for k in h_kw))
    if pt==0 and ht==0: pt=sum(p["total"] for p in products)
    return {"pension":pt,"provident":ht,"products":products,"raw_df":df}

def parse_invest(path):
    df=read_html_xls(path)
    if df is not None:
        total=0
        for _,row in df.iterrows():
            for val in row:
                v=num(str(val))
                if 500000<v<20000000: total=v; break
            if total: break
        return {"total":total,"raw_df":df}
    try:
        engine="xlrd" if str(path).endswith(".xls") else "openpyxl"
        df=pd.read_excel(path,header=None,engine=engine)
        total=0
        try: total=num(str(df.iloc[2,3]).replace(",",""))
        except: pass
        return {"total":total,"raw_df":df}
    except: return {}

def parse_rsu_image(path):
    if not ANTHROPIC_KEY:
        warn("No API key. Enter RSU manually:")
        try:
            u=float(input("  Unvested ($): ").replace(",","").replace("$",""))
            a=float(input("  Available ($): ").replace(",","").replace("$",""))
            return {"unvested":u,"available":a}
        except: return {"unvested":0,"available":0}
    try:
        import base64,json as _j,urllib.request
        with open(path,"rb") as f: b64=base64.b64encode(f.read()).decode()
        mime="image/png" if str(path).lower().endswith(".png") else "image/jpeg"
        payload={"model":"claude-sonnet-4-20250514","max_tokens":100,
            "messages":[{"role":"user","content":[
                {"type":"image","source":{"type":"base64","media_type":mime,"data":b64}},
                {"type":"text","text":"RSU equity overview. Find Unvested and Shares/Available dollar amounts. JSON only: {\"unvested\": 187148, \"available\": 170600}"}
            ]}]}
        req=urllib.request.Request("https://api.anthropic.com/v1/messages",
            data=_j.dumps(payload).encode(),
            headers={"Content-Type":"application/json","anthropic-version":"2023-06-01","x-api-key":ANTHROPIC_KEY})
        with urllib.request.urlopen(req,timeout=30) as r:
            res=_j.loads(r.read())
            text=res["content"][0]["text"].strip()
            m=re.search(r"\{[^}]+\}",text)
            if m: return _j.loads(m.group())
    except Exception as e: warn(f"RSU API error: {e}")
    warn("RSU: enter manually:")
    try:
        u=float(input("  Unvested ($): ").replace(",","").replace("$",""))
        a=float(input("  Available ($): ").replace(",","").replace("$",""))
        return {"unvested":u,"available":a}
    except: return {"unvested":0,"available":0}

def parse_all(found):
    hdr("Parsing files")
    data={}
    parsers={"bank":parse_bank,"credit":parse_credit,"foreign":parse_credit,
             "isracard":parse_isracard,"pension_dror":parse_pension,"pension_liat":parse_pension,
             "invest":parse_invest,"balance":parse_invest,"rsu_image":parse_rsu_image}
    for key,path in found.items():
        if path is None: data[key]={}; warn(f"Missing: {key}"); continue
        try: data[key]=parsers[key](path); ok(f"Parsed: {key}")
        except Exception as e: warn(f"Error ({key}): {e}"); data[key]={}
    return data

def update_mislaka(wb,sheet,pension_data):
    if sheet not in wb.sheetnames or not pension_data: return
    products=pension_data.get("products",[])
    if not products: return
    ws=wb[sheet]
    for i,p in enumerate(products,2):
        ws.cell(row=i,column=1).value=p.get("product","")
        ws.cell(row=i,column=5).value=p.get("total",0)
        ws.cell(row=i,column=12).value=p.get("fee_deposit",0)
        ws.cell(row=i,column=13).value=p.get("fee_accum",0)
        ws.cell(row=i,column=14).value=p.get("return_ytd",0)

def update_excel(data):
    hdr("Updating Excel")
    # Close Excel first to avoid permission errors
    os.system("taskkill /f /im excel.exe 2>nul")
    time.sleep(2)
    wb=load_workbook(EXCEL,keep_vba=True)

    # Update dashboard input cells
    ws=wb["\u05d3\u05e9\u05d1\u05d5\u05e8\u05d3"]
    def sc(row,col,val,label):
        if val and val!=0:
            ws.cell(row=row,column=col).value=val
            ok(f"  {label}: {val:,.0f}")
    sc(10,4,data.get("pension_dror",{}).get("pension",0),"\u05e4\u05e0\u05e1\u05d9\u05d4 \u05d3\u05e8\u05d5\u05e8")
    sc(11,4,data.get("pension_liat",{}).get("pension",0),"\u05e4\u05e0\u05e1\u05d9\u05d4 \u05dc\u05d9\u05d0\u05ea")
    sc(12,4,data.get("pension_dror",{}).get("provident",0),"\u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05d3\u05e8\u05d5\u05e8")
    sc(13,4,data.get("pension_liat",{}).get("provident",0),"\u05d4\u05e9\u05ea\u05dc\u05de\u05d5\u05ea \u05dc\u05d9\u05d0\u05ea")
    sc(14,4,data.get("invest",{}).get("total",0),"\u05ea\u05d9\u05e7 \u05d4\u05e9\u05e7\u05e2\u05d5\u05ea")
    sc(18,4,data.get("bank",{}).get("balance",0),'\u05e2\u05d5"\u05e9')
    ws.cell(row=2,column=1).value=f"\u05e2\u05d3\u05db\u05d5\u05df \u05d0\u05d7\u05e8\u05d5\u05df: {datetime.now().strftime('%d/%m/%Y')}"

    # Update RSU — write to ALIGN RSU sheet H13/H14
    rsu=data.get("rsu_image",{})
    avail=rsu.get("available",0); unves=rsu.get("unvested",0)
    if avail or unves:
        if "ALIGN RSU" in wb.sheetnames:
            rws=wb["ALIGN RSU"]
            rws.cell(row=13,column=8).value=avail
            rws.cell(row=14,column=8).value=unves
            ok(f"  RSU: available={avail}, unvested={unves}")

    # Update mislaka sheets (pension data)
    update_mislaka(wb,"\u05d3\u05e8\u05d5\u05e8 - \u05de\u05e1\u05dc\u05e7\u05d4",data.get("pension_dror",{}))
    update_mislaka(wb,"\u05dc\u05d9\u05d0\u05ea - \u05de\u05e1\u05dc\u05e7\u05d4",data.get("pension_liat",{}))

    # Update transaction sheets
    for sheet,key in {
        "\u05e2\u05d5\u05e9":"bank",
        "\u05e2\u05e1\u05e7\u05d0\u05d5\u05ea \u05d1\u05de\u05d5\u05e2\u05d3 \u05d4\u05d7\u05d9\u05d5\u05d1":"credit",
        '\u05e2\u05e1\u05e7\u05d0\u05d5\u05ea \u05d7\u05d5"\u05dc \u05d5\u05de\u05d8"\u05d7':"foreign",
        "\u05d0\u05d9\u05e9\u05e8\u05d0\u05db\u05e8\u05d8":"isracard"
    }.items():
        df=data.get(key,{}).get("raw_df")
        if df is None or sheet not in wb.sheetnames: continue
        wsh=wb[sheet]
        for row in wsh.iter_rows(min_row=1,max_row=wsh.max_row):
            for cell in row: cell.value=None
        for r,row in enumerate(df.values,1):
            for c,val in enumerate(row,1):
                if val is not None and str(val)!="nan":
                    try: wsh.cell(row=r,column=c).value=val
                    except: pass
        ok(f"  Updated sheet: {sheet}")

    wb.save(EXCEL)
    ok("Excel saved")

def main():
    print(f"\n{C}  Klein Family Finance v3.0{X}")
    print(f"  {datetime.now().strftime('%d %B %Y, %H:%M')}\n")
    if not EXCEL.exists():
        err(f"Excel not found: {EXCEL}"); sys.exit(1)
    found=find_files()
    data=parse_all(found)
    hdr("Backup")
    BACKUPS.mkdir(exist_ok=True)
    shutil.copy2(EXCEL,BACKUPS/f"\u05de\u05d0\u05d6\u05df_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsm")
    ok("Backup created")
    update_excel(data)
    hdr("Opening Excel")
    os.startfile(str(EXCEL))
    ok("Done!")
    print()

if __name__=="__main__":
    main()
