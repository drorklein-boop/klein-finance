# Klein Finance - Monthly Sheet Updater v7.2
import sys, shutil, datetime
from pathlib import Path
from collections import defaultdict

BASE    = Path(__file__).parent
MONTHLY = BASE / "MONTHLY"

def detect_type(fname):
    if fname.startswith('~'): return None
    name = fname.lower()
    if 'transaction-details' in name: return 'credit'
    if 'עוש' in name or 'לאומי' in name: return 'bank'
    if 'אחזקות' in name: return 'invest'
    if 'התמונה המלאה' in name:
        return 'pension_liat' if '(11)' in fname else 'pension_dror'
    if '5647' in name or 'אישראכרט' in name: return 'isracard'
    if 'ריכוז' in name and 'יתרות' in name: return 'balance'
    return None

def clean_val(val):
    """Strip invisible RTL markers; convert number-strings to floats."""
    if isinstance(val, str):
        cleaned = val.replace('\u200e', '').replace('\u200f', '').strip()
        try:
            return float(cleaned.replace(',', ''))
        except ValueError:
            return cleaned if cleaned else None
    return val

def read_all_rows(path, sheet_name, is_xls=False):
    """Read ALL rows exactly as-is, preserving positions. Strip RTL artifacts only."""
    if is_xls:
        import xlrd
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_name(sheet_name)
        return [[clean_val(ws.cell(r, c).value) for c in range(ws.ncols)]
                for r in range(ws.nrows)]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet_name]
        data = [[clean_val(v) for v in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
        return data

def collect_files():
    typed = defaultdict(list)
    for f in MONTHLY.iterdir():
        if not f.is_file(): continue
        t = detect_type(f.name)
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
        return sheets  # sheet name matches target exactly

    elif ftype == 'bank' and is_xlsx:
        return {'עוש': read_all_rows(fpath, 'עוש')}

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
        return {'דרור - מסלקה': read_all_rows(fpath, 'פרטי המוצרים שלי', is_xls=True)}

    elif ftype == 'pension_liat' and is_xls:
        return {'ליאת - מסלקה': read_all_rows(fpath, 'פרטי המוצרים שלי', is_xls=True)}

    elif ftype == 'isracard' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        sname = wb.sheetnames[0]
        data = [[clean_val(v) for v in row] for row in wb[sname].iter_rows(values_only=True)]
        wb.close()
        return {'אישראכרט': data}

    elif ftype == 'balance' and is_xlsx:
        import openpyxl
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        sname = wb.sheetnames[0]
        data = [[clean_val(v) for v in row] for row in wb[sname].iter_rows(values_only=True)]
        wb.close()
        return {'ריכוז יתרות לאומי': data}

    return {}

def write_sheet(xw_wb, target_name, data):
    """Clear sheet, write all rows from A1 — preserves source cell positions."""
    ws = xw_wb.sheets[target_name]
    ws.clear_contents()
    if data:
        xw_wb.app.screen_updating = False
        xw_wb.app.calculation = 'manual'
        ws.range('A1').value = data
        xw_wb.app.calculation = 'automatic'
        xw_wb.app.screen_updating = True

def main():
    print("\n  Klein Finance - Monthly Update v7.2")
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

    print(f"\n  Processing {sum(len(v) for v in typed.values())} file(s)...")
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
    for r in results:
        print(r)
    print("\n  Done.")
    input("\n  Press Enter to close...")

if __name__ == '__main__':
    main()
