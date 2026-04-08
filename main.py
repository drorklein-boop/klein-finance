# Klein Finance - Monthly Sheet Updater v7.4
import sys, shutil, datetime
from pathlib import Path
from collections import defaultdict

BASE    = Path(__file__).parent
MONTHLY = BASE / "MONTHLY"

DROR_POLICY = '35995836'
LIAT_POLICY = '6650891010'

def detect_by_name(fname):
    if fname.startswith('~'): return None
    name = fname.lower()
    if 'transaction-details' in name: return 'credit'
    if 'עוש' in name or 'לאומי' in name: return 'bank'
    if 'אחזקות' in name: return 'invest'
    if 'התמונה המלאה' in name: return 'pension_check'  # content must confirm
    if '5647' in name or 'אישראכרט' in name: return 'isracard'
    if 'ריכוז' in name and 'יתרות' in name: return 'balance'
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
            if 'עסקאות במועד החיוב' in sheets or 'עסקאות חו"ל ומט"ח' in sheets:
                return 'credit'
            if 'פירוט עסקאות' in sheets:
                return 'isracard'
            if 'עוש' in sheets:
                return 'bank'
            for row in first_rows:
                if any('מבט אישי' in str(v or '') for v in row):
                    return 'invest'
        elif is_xls:
            import xlrd
            wb = xlrd.open_workbook(fpath)
            sheets = set(wb.sheet_names())
            if 'פרטי המוצרים שלי' in sheets:
                ws = wb.sheet_by_name('פרטי המוצרים שלי')
                all_vals = ' '.join(str(ws.cell(r,c).value)
                                    for r in range(ws.nrows) for c in range(ws.ncols))
                if DROR_POLICY in all_vals: return 'pension_dror'
                if LIAT_POLICY in all_vals: return 'pension_liat'
    except:
        pass
    return None

def detect(fpath):
    by_name = detect_by_name(fpath.name)
    if by_name == 'pension_check' or by_name is None:
        return detect_by_content(fpath)
    return by_name

def clean_val(val):
    if isinstance(val, str):
        for ch in ('\u200e','\u200f','\u202b','\u202c'):
            val = val.replace(ch, '')
        val = val.strip()
        try: return float(val.replace(',',''))
        except: return val if val else None
    return val

def read_full_xlsx(path, sheet_name=None):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sname = sheet_name or wb.sheetnames[0]
    ws = wb[sname]
    data = [[clean_val(v) for v in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return data

def read_full_xls(path, sheet_name):
    import xlrd
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    return [[clean_val(ws.cell(r,c).value) for c in range(ws.ncols)] for r in range(ws.nrows)]

def read_from_header(path, header_col0_value, sheet_name=None):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sname = sheet_name or wb.sheetnames[0]
    ws = wb[sname]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    start = next((i for i, r in enumerate(all_rows)
                  if r[0] and str(r[0]).strip() == header_col0_value), None)
    if start is None:
        raise ValueError(f"Header '{header_col0_value}' not found in {path.name}")
    return [[clean_val(v) for v in row] for row in all_rows[start:]]

def collect_files():
    typed = defaultdict(list)
    for f in MONTHLY.iterdir():
        if not f.is_file(): continue
        if f.suffix.lower() in ('.jpeg', '.jpg', '.png', '.xml'): continue
        t = detect(f)
        if t:
            typed[t].append(f)
    for t in typed:
        typed[t].sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return typed

def read_file(ftype, fpath):
    fname_lower = fpath.name.lower()
    is_xlsx = fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls.xlsx')
    is_xls  = fname_lower.endswith('.xls') and not fname_lower.endswith('.xls.xlsx')

    if ftype == 'credit' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        sheets = {sname: [[clean_val(v) for v in row]
                           for row in wb[sname].iter_rows(values_only=True)]
                  for sname in wb.sheetnames}
        wb.close()
        return sheets

    elif ftype == 'bank' and is_xlsx:
        return {'עוש': read_full_xlsx(fpath, 'עוש')}

    elif ftype == 'invest' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            first = next(ws.iter_rows(values_only=True), [])
            if first and 'מבט אישי' in str(first[0] or ''):
                data = [[clean_val(v) for v in row] for row in ws.iter_rows(values_only=True)]
                wb.close()
                return {'תיק השקעות עדכני': data}
        wb.close()
        return {}

    elif ftype == 'pension_dror' and is_xls:
        return {'דרור - מסלקה': read_full_xls(fpath, 'פרטי המוצרים שלי')}

    elif ftype == 'pension_liat' and is_xls:
        return {'ליאת - מסלקה': read_full_xls(fpath, 'פרטי המוצרים שלי')}

    elif ftype == 'isracard' and is_xlsx:
        return {'אישראכרט': read_from_header(fpath, 'תאריך רכישה')}

    elif ftype == 'balance' and is_xlsx:
        return {'ריכוז יתרות לאומי': read_from_header(fpath, 'סוג פעילות')}

    return {}

def write_sheet(xw_wb, target_name, data):
    ws = xw_wb.sheets[target_name]
    ws.clear_contents()
    if data:
        xw_wb.app.screen_updating = False
        xw_wb.app.calculation = 'manual'
        ws.range('A1').value = data
        xw_wb.app.calculation = 'automatic'
        xw_wb.app.screen_updating = True

def main():
    print("\n  Klein Finance - Monthly Update v7.4")
    print("  =====================================")

    try:
        import xlwings as xw
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'xlwings', '--quiet'], check=True)
        import xlwings as xw

    try:
        import openpyxl, xlrd
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'openpyxl', 'xlrd', '--quiet'], check=True)

    app = xw.apps.active
    if not app:
        print("  ERROR: Excel is not open.")
        input("\n  Press Enter to close...")
        return

    wb = None
    for book in app.books:
        if 'מאזן_קליין' in book.name:
            wb = book
            break

    if wb is None:
        print("  ERROR: מאזן_קליין.xlsm not open.")
        input("\n  Press Enter to close...")
        return

    print(f"  Workbook: {wb.name}")

    backup_dir = BASE / "backups"
    backup_dir.mkdir(exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    src = BASE / "מאזן_קליין.xlsm"
    if src.exists():
        shutil.copy2(src, backup_dir / f"מאזן_{ts}.xlsm")
        print(f"  Backup saved")

    typed = collect_files()
    if not typed:
        print("\n  No recognized files found in MONTHLY folder.")
        input("\n  Press Enter to close...")
        return

    print(f"\n  Detected files:")
    for ftype, files in typed.items():
        print(f"    {ftype}: {files[0].name}")

    results = []
    target_sheet_names = [s.name for s in wb.sheets]

    for ftype, files in typed.items():
        fpath = files[0]
        try:
            sheets_data = read_file(ftype, fpath)
            for target_name, data in sheets_data.items():
                if target_name not in target_sheet_names:
                    results.append(f"  SKIP  '{target_name}' not in workbook")
                    continue
                write_sheet(wb, target_name, data)
                results.append(f"  OK    {target_name} ({len(data)} rows)")
        except Exception as e:
            results.append(f"  ERROR {ftype} ({fpath.name}): {e}")

    wb.save()

    print("\n  Results:")
    for r in results: print(r)
    print("\n  Done.")
    input("\n  Press Enter to close...")

if __name__ == '__main__':
    main()
