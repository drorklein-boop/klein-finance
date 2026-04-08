# Klein Finance - Monthly Sheet Updater v7.6
import sys, shutil, datetime, json, base64
from pathlib import Path
from collections import defaultdict

BASE    = Path(__file__).parent
MONTHLY = BASE / "MONTHLY"
TRACKER = BASE / "processed_files.json"
API_KEY_FILE = BASE / "api_key.txt"

DROR_POLICY = '35995836'
LIAT_POLICY = '6650891010'

# ââ Tracker âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def load_tracker():
    if TRACKER.exists():
        try: return json.loads(TRACKER.read_text(encoding='utf-8'))
        except: pass
    return {}

def save_tracker(tracker):
    TRACKER.write_text(json.dumps(tracker, indent=2, ensure_ascii=False), encoding='utf-8')

def file_sig(fpath):
    s = fpath.stat()
    return {"mtime": round(s.st_mtime, 2), "size": s.st_size}

def is_new(fpath, tracker):
    key = fpath.name
    sig = file_sig(fpath)
    if key not in tracker: return True
    t = tracker[key]
    return t.get('mtime') != sig['mtime'] or t.get('size') != sig['size']

# ââ Detection âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def detect_by_name(fname):
    if fname.startswith('~'): return None
    name = fname.lower()
    if name.endswith(('.jpeg','.jpg','.png')): return 'rsu_image'
    if 'transaction-details' in name: return 'credit'
    if '×¢××©' in name or '×××××' in name: return 'bank'
    if '××××§××ª' in name: return 'invest'
    if '××ª××× × ×××××' in name: return 'pension_check'
    if '5647' in name or '×××©×¨×××¨×' in name: return 'isracard'
    if '×¨××××' in name and '××ª×¨××ª' in name: return 'balance'
    return None

def detect_by_content(fpath):
    fname = fpath.name.lower()
    is_xlsx = fname.endswith('.xlsx') or fname.endswith('.xls.xlsx')
    is_xls  = fname.endswith('.xls') and not fname.endswith('.xls.xlsx')
    try:
        if is_xlsx:
            import openpyxl
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            sheets = set(wb.sheetnames)
            first_rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))[:3]
            wb.close()
            if '×¢×¡×§×××ª ××××¢× ×××××' in sheets or '×¢×¡×§×××ª ××"× ×××"×' in sheets: return 'credit'
            if '×¤××¨×× ×¢×¡×§×××ª' in sheets: return 'isracard'
            if '×¢××©' in sheets: return 'bank'
            for row in first_rows:
                if any('××× ×××©×' in str(v or '') for v in row): return 'invest'
        elif is_xls:
            import xlrd
            wb = xlrd.open_workbook(fpath)
            if '×¤×¨×× ××××¦×¨×× ×©××' in set(wb.sheet_names()):
                ws = wb.sheet_by_name('×¤×¨×× ××××¦×¨×× ×©××')
                vals = ' '.join(str(ws.cell(r,c).value) for r in range(ws.nrows) for c in range(ws.ncols))
                if DROR_POLICY in vals: return 'pension_dror'
                if LIAT_POLICY in vals: return 'pension_liat'
    except: pass
    return None

def detect(fpath):
    by_name = detect_by_name(fpath.name)
    if by_name == 'pension_check' or by_name is None:
        return detect_by_content(fpath)
    return by_name

# ââ RSU via Anthropic API âââââââââââââââââââââââââââââââââââââââââââââââââââââ

def read_rsu_from_image(img_path):
    try:
        import anthropic
    except ImportError:
        import subprocess
        subprocess.run([sys.executable,'-m','pip','install','anthropic','--quiet'], check=True)
        import anthropic

    api_key = API_KEY_FILE.read_text().strip()
    img_b64 = base64.standard_b64encode(img_path.read_bytes()).decode()
    ext = img_path.suffix.lower().lstrip('.')
    media_type = 'image/jpeg' if ext in ('jpg','jpeg') else f'image/{ext}'

    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=256,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
                {"type": "text", "text": (
                    "This is an RSU portfolio screenshot. Extract exactly two numbers:\n"
                    "1. ×××× ××××××© (available/vested) dollar amount\n"
                    "2. ××¨× ×××©×× (unvested) dollar amount\n"
                    "Reply with JSON only, no explanation: {\"available\": 170600.00, \"unvested\": 187148.20}"
                )}
            ]
        }]
    )
    result = json.loads(msg.content[0].text)
    return float(result['available']), float(result['unvested'])

# ââ Reading âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def clean_val(val):
    if isinstance(val, str):
        for ch in ('\u200e','\u200f','\u202b','\u202c'): val = val.replace(ch,'')
        val = val.strip()
        try: return float(val.replace(',',''))
        except: return val if val else None
    return val

def read_full_xlsx(path, sheet_name=None):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name or wb.sheetnames[0]]
    data = [[clean_val(v) for v in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return data

def read_full_xls(path, sheet_name):
    import xlrd
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    return [[clean_val(ws.cell(r,c).value) for c in range(ws.ncols)] for r in range(ws.nrows)]

def read_from_header(path, header_val, sheet_name=None):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name or wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    start = next((i for i,r in enumerate(rows) if r[0] and str(r[0]).strip()==header_val), None)
    if start is None: raise ValueError(f"Header '{header_val}' not found")
    return [[clean_val(v) for v in row] for row in rows[start:]]

def read_file(ftype, fpath):
    fname = fpath.name.lower()
    is_xlsx = fname.endswith('.xlsx') or fname.endswith('.xls.xlsx')
    is_xls  = fname.endswith('.xls') and not fname.endswith('.xls.xlsx')

    if ftype == 'credit' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        out = {s: [[clean_val(v) for v in row] for row in wb[s].iter_rows(values_only=True)] for s in wb.sheetnames}
        wb.close()
        return out
    elif ftype == 'bank' and is_xlsx:
        return {'×¢××©': read_full_xlsx(fpath, '×¢××©')}
    elif ftype == 'invest' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        for s in wb.sheetnames:
            first = next(wb[s].iter_rows(values_only=True), [])
            if first and '××× ×××©×' in str(first[0] or ''):
                data = [[clean_val(v) for v in row] for row in wb[s].iter_rows(values_only=True)]
                wb.close()
                return {'×ª××§ ××©×§×¢××ª ×¢××× ×': data}
        wb.close()
        return {}
    elif ftype == 'pension_dror' and is_xls:
        return {'××¨××¨ - ××¡××§×': read_full_xls(fpath, '×¤×¨×× ××××¦×¨×× ×©××')}
    elif ftype == 'pension_liat' and is_xls:
        return {'××××ª - ××¡××§×': read_full_xls(fpath, '×¤×¨×× ××××¦×¨×× ×©××')}
    elif ftype == 'isracard' and is_xlsx:
        return {'×××©×¨×××¨×': read_from_header(fpath, '×ª××¨×× ×¨×××©×')}
    elif ftype == 'balance' and is_xlsx:
        return {'×¨×××× ××ª×¨××ª ×××××': read_from_header(fpath, '×¡×× ×¤×¢××××ª')}
    return {}

# ââ Write âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def write_sheet(xw_wb, name, data):
    ws = xw_wb.sheets[name]
    ws.clear_contents()
    if data:
        xw_wb.app.screen_updating = False
        xw_wb.app.calculation = 'manual'
        ws.range('A1').value = data
        xw_wb.app.calculation = 'automatic'
        xw_wb.app.screen_updating = True

# ââ Main ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

def main():
    print("\n  Klein Finance - Monthly Update v7.7")
    print("  =====================================")

    try:
        import xlwings as xw
    except ImportError:
        import subprocess
        subprocess.run([sys.executable,'-m','pip','install','xlwings','--quiet'], check=True)
        import xlwings as xw
    try:
        import openpyxl, xlrd
    except ImportError:
        import subprocess
        subprocess.run([sys.executable,'-m','pip','install','openpyxl','xlrd','--quiet'], check=True)

    app = xw.apps.active
    if not app:
        print("  ERROR: Excel is not open.")
        input("\n  Press Enter to close..."); return

    wb = next((b for b in app.books if '××××_×§××××' in b.name), None)
    if not wb:
        print("  ERROR: ××××_×§××××.xlsm not open.")
        input("\n  Press Enter to close..."); return

    print(f"  Workbook: {wb.name}")

    backup_dir = BASE / "backups"
    backup_dir.mkdir(exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    src = BASE / "××××_×§××××.xlsm"
    if src.exists():
        shutil.copy2(src, backup_dir / f"××××_{ts}.xlsm")
        print("  Backup saved")

    tracker = load_tracker()
    all_files = [f for f in MONTHLY.iterdir()
                 if f.is_file() and not f.name.startswith('~')
                 and f.suffix.lower() not in ('.xml',)]
    new_files = [f for f in all_files if is_new(f, tracker)]

    if not new_files:
        print("\n  No new or changed files since last run.")
        input("\n  Press Enter to close..."); return

    print(f"\n  New/changed files: {len(new_files)}")

    typed = defaultdict(list)
    for f in new_files:
        t = detect(f)
        if t: typed[t].append(f)
    for t in typed:
        typed[t].sort(key=lambda f: f.stat().st_mtime, reverse=True)

    if not typed:
        print("  None of the new files could be identified.")
        input("\n  Press Enter to close..."); return

    print("  Detected:")
    for ftype, files in typed.items():
        print(f"    {ftype}: {files[0].name}")

    results = []
    target_sheets = [s.name for s in wb.sheets]

    for ftype, files in typed.items():
        fpath = files[0]

        # RSU image â special handling
        if ftype == 'rsu_image':
            try:
                available, unvested = read_rsu_from_image(fpath)
                ws_rsu = wb.sheets['ALIGN RSU']
                ws_rsu['H13'].value = available
                ws_rsu['H14'].value = unvested
                results.append(f"  OK    ALIGN RSU H13={available:,.2f}  H14={unvested:,.2f}")
            except Exception as e:
                results.append(f"  ERROR rsu_image ({fpath.name}): {e}")
            continue

        try:
            sheets_data = read_file(ftype, fpath)
            for tname, data in sheets_data.items():
                if tname not in target_sheets:
                    results.append(f"  SKIP  '{tname}' not in workbook"); continue
                write_sheet(wb, tname, data)
                results.append(f"  OK    {tname} ({len(data)} rows)")
        except Exception as e:
            results.append(f"  ERROR {ftype} ({fpath.name}): {e}")

    wb.save()

    updated_tracker = {f.name: file_sig(f) for f in all_files}
    save_tracker(updated_tracker)

    print("\n  Results:")
    for r in results: print(r)
    print("\n  Done.")
    input("\n  Press Enter to close...")

if __name__ == '__main__':
    main()
