# Klein Finance - Dashboard Generator v1.0
import sys, json, warnings, webbrowser, datetime
warnings.filterwarnings('ignore')
from pathlib import Path

BASE = Path(__file__).parent

def read_workbook():
    import openpyxl
    xlsm = next((f for f in BASE.glob('*.xlsm') if not f.name.startswith('~')), None)
    if not xlsm:
        print("ERROR: No .xlsm file found in", BASE)
        return None
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    data = {}

    # Dashboard cells
    ws = wb['דשבורד']
    data['pension_dror']   = ws['D10'].value or 0
    data['pension_liat']   = ws['D11'].value or 0
    data['hishtalmut_dror']= ws['D12'].value or 0
    data['hishtalmut_liat']= ws['D13'].value or 0
    data['portfolio']      = ws['D14'].value or 0
    data['bank']           = ws['D18'].value or 0

    # RSU
    rsu = wb['ALIGN RSU']
    data['rsu_vested']   = rsu['H13'].value or 0
    data['rsu_unvested'] = rsu['H14'].value or 0

    # דוח חודשי
    ws_r = wb['דוח חודשי']
    rows = list(ws_r.iter_rows(values_only=True))
    def find_val(label, col=1):
        for row in rows:
            if row[0] and label in str(row[0]):
                return row[col]
        return None
    data['income']        = find_val('הכנסות') or 0
    data['expenses']      = find_val('הוצאות') or 0
    data['surplus']       = find_val('יתרה חודשית') or 0
    data['savings_rate']  = find_val('יחס חיסכון') or 0
    data['portfolio_val'] = find_val('שווי תיק מסחר', 1) or 0
    data['kaspiyot']      = find_val('קרן ביטחון', 3) or 0
    data['net_worth']     = find_val('סה"כ נכסים נטו', 1) or 0
    if not data['net_worth']:
        data['net_worth'] = find_val('סה"כ נכסים נטו', 1) or 7110995

    # Report date
    for row in rows:
        if row[0] and 'תאריך הפקה' in str(row[0]):
            data['report_date'] = str(row[1]) if row[1] else ''
            data['period']      = str(row[3]) if row[3] else ''
            break
    else:
        data['report_date'] = datetime.date.today().strftime('%d/%m/%Y')
        data['period'] = datetime.date.today().strftime('%B %Y')

    # Top 5 expenses
    expenses = []
    in_top5 = False
    for row in rows:
        if row[0] and 'Top 5' in str(row[0]): in_top5 = True; continue
        if in_top5 and row[0] and str(row[0]).isdigit() and row[1] and row[2]:
            expenses.append({'name': str(row[1]), 'amount': float(row[2] or 0), 'source': str(row[3] or '')})
        if in_top5 and len(expenses) >= 5: break
    data['top_expenses'] = expenses

    # Top 3 categories
    cats = []
    in_cats = False
    for row in rows:
        if row[0] and 'Top 3 קטגוריות' in str(row[0]): in_cats = True; continue
        if in_cats and row[0] and row[1] and isinstance(row[1], (int,float)):
            cats.append({'name': str(row[0]), 'amount': float(row[1]), 'pct': float(row[2] or 0)})
        if in_cats and len(cats) >= 3: break
    data['top_categories'] = cats

    # Investment holdings
    holdings = []
    ws_inv = wb['תיק השקעות עדכני']
    header_found = False
    for row in ws_inv.iter_rows(values_only=True):
        if not header_found:
            if row[0] and str(row[0]).isdigit() and row[1] and row[4] and row[5]:
                header_found = True
        if header_found and row[0] and str(row[0]).isdigit() and row[1]:
            pct_chg = float(row[7]) if row[7] is not None else 0
            holdings.append({
                'name': str(row[1]),
                'value': float(row[5] or 0),
                'pct_portfolio': float(row[9] or 0),
                'pct_change': pct_chg,
            })
    data['holdings'] = holdings

    # ריכוז יתרות
    balances = []
    ws_bal = wb['ריכוז יתרות לאומי']
    for row in ws_bal.iter_rows(values_only=True):
        if row[0] and row[2] and str(row[0]) not in ('סוג פעילות','סה"כ') and isinstance(row[2], (int,float)):
            balances.append({'type': str(row[0]), 'value': float(row[2])})
    data['balances'] = balances

    # Mortgage
    for row in ws_bal.iter_rows(values_only=True):
        if row[0] and 'משכנתא' in str(row[0]) and row[2] and isinstance(row[2], (int,float)):
            data['mortgage'] = float(row[2])
            break
    else:
        data['mortgage'] = 0

    wb.close()
    return data


def generate_html(data):
    def fmt_ils(n, decimals=0):
        try:
            n = float(n)
            if decimals:
                return f'₪{n:,.{decimals}f}'
            return f'₪{n:,.0f}'
        except: return '₪0'

    def fmt_pct(n):
        try: return f'{float(n)*100:.1f}%'
        except: return '0%'

    def fmt_usd(n):
        try: return f'${float(n):,.0f}'
        except: return '$0'

    def sign_color(n):
        try:
            v = float(n)
            if v > 0: return '#3bbd6e'
            if v < 0: return '#e05555'
            return '#6b7897'
        except: return '#6b7897'

    holdings_html = ''
    for h in data.get('holdings', []):
        pct = h['pct_change']
        arrow = '▲' if pct >= 0 else '▼'
        clr = '#3bbd6e' if pct >= 0 else '#e05555'
        bar_w = min(int(h['pct_portfolio'] * 100 * 2), 100)
        holdings_html += f'''
        <div class="holding-row">
          <div class="holding-name">{h["name"]}</div>
          <div class="holding-bar-wrap"><div class="holding-bar" style="width:{bar_w}%"></div></div>
          <div class="holding-pct-port">{fmt_pct(h["pct_portfolio"])}</div>
          <div class="holding-val">{fmt_ils(h["value"])}</div>
          <div class="holding-chg" style="color:{clr}">{arrow} {abs(pct)*100:.1f}%</div>
        </div>'''

    expenses_html = ''
    for i, e in enumerate(data.get('top_expenses', []), 1):
        expenses_html += f'''
        <div class="expense-row">
          <span class="exp-rank">{i}</span>
          <span class="exp-name">{e["name"]}</span>
          <span class="exp-source">{e["source"]}</span>
          <span class="exp-amount">{fmt_ils(e["amount"])}</span>
        </div>'''

    cats_html = ''
    for c in data.get('top_categories', []):
        bar_w = min(int(c['pct'] * 300), 100)
        cats_html += f'''
        <div class="cat-row">
          <span class="cat-name">{c["name"]}</span>
          <div class="cat-bar-wrap"><div class="cat-bar" style="width:{bar_w}%"></div></div>
          <span class="cat-pct">{fmt_pct(c["pct"])}</span>
          <span class="cat-amount">{fmt_ils(c["amount"])}</span>
        </div>'''

    pension_total = (data['pension_dror'] + data['pension_liat'] +
                     data['hishtalmut_dror'] + data['hishtalmut_liat'])
    rsu_ils = (data['rsu_vested'] + data['rsu_unvested']) * 3.65  # approx

    return f'''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>קליין — דשבורד {data.get("period","")}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;800;900&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:#0b0e14; --bg2:#111520; --bg3:#161b28; --bg4:#1c2235;
  --border:#1e2438; --text:#dce4f0; --muted:#6b7897;
  --gold:#c8a45e; --gold2:#e8c47e;
  --green:#3bbd6e; --yellow:#e8b84b; --red:#e05555; --blue:#5b8dee;
  --mono:'JetBrains Mono',monospace;
}}
*,*::before,*::after{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:var(--bg);color:var(--text);font-family:'Heebo',sans-serif;direction:rtl;min-height:100vh;}}

.header{{
  background:linear-gradient(135deg,#0e1220,#141928);
  border-bottom:1px solid var(--border);
  padding:24px 40px;
  display:flex;justify-content:space-between;align-items:center;
}}
.brand span{{font-size:11px;color:var(--muted);letter-spacing:3px;text-transform:uppercase;}}
.brand strong{{display:block;font-size:26px;color:var(--gold);font-weight:900;letter-spacing:-0.5px;}}
.nw-block{{text-align:center;}}
.nw-label{{font-size:10px;color:var(--muted);letter-spacing:2px;text-transform:uppercase;}}
.nw-value{{font-family:var(--mono);font-size:38px;font-weight:600;color:var(--gold2);letter-spacing:-1px;margin-top:6px;}}
.meta{{text-align:left;font-size:12px;color:var(--muted);line-height:2;}}
.meta strong{{color:var(--text);}}

.strip{{display:grid;grid-template-columns:repeat(5,1fr);border-bottom:1px solid var(--border);background:var(--bg2);}}
.strip-cell{{padding:18px 24px;border-left:1px solid var(--border);}}
.strip-cell:last-child{{border-left:none;}}
.strip-label{{font-size:10px;color:var(--muted);letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;}}
.strip-value{{font-family:var(--mono);font-size:20px;font-weight:600;color:var(--text);}}
.strip-sub{{font-size:11px;color:var(--muted);margin-top:4px;}}

.main{{padding:32px 40px;display:grid;grid-template-columns:1fr 1fr;gap:24px;}}
.card{{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:24px;}}
.card-full{{grid-column:1/-1;}}
.card-title{{font-size:11px;color:var(--muted);letter-spacing:3px;text-transform:uppercase;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;}}
.card-title-accent{{color:var(--gold);font-size:18px;font-weight:800;letter-spacing:0;}}

.holding-row{{display:grid;grid-template-columns:2fr 1fr 60px 100px 80px;gap:12px;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);font-size:13px;}}
.holding-row:last-child{{border-bottom:none;}}
.holding-name{{color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.holding-bar-wrap{{background:var(--bg3);border-radius:4px;height:6px;overflow:hidden;}}
.holding-bar{{height:100%;background:linear-gradient(90deg,var(--blue),var(--gold));border-radius:4px;transition:width 0.5s;}}
.holding-pct-port{{color:var(--muted);font-size:12px;text-align:center;}}
.holding-val{{font-family:var(--mono);font-size:13px;text-align:left;}}
.holding-chg{{font-size:12px;text-align:left;font-family:var(--mono);}}

.pension-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px;}}
.pension-item{{background:var(--bg3);border-radius:8px;padding:14px;}}
.pension-item-label{{font-size:11px;color:var(--muted);margin-bottom:6px;}}
.pension-item-value{{font-family:var(--mono);font-size:16px;font-weight:600;color:var(--gold2);}}

.rsu-row{{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid var(--border);}}
.rsu-row:last-child{{border-bottom:none;}}
.rsu-label{{font-size:13px;color:var(--muted);}}
.rsu-val{{font-family:var(--mono);font-size:15px;font-weight:600;}}

.expense-row{{display:grid;grid-template-columns:28px 1fr 80px 100px;gap:12px;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);font-size:13px;}}
.expense-row:last-child{{border-bottom:none;}}
.exp-rank{{color:var(--muted);font-family:var(--mono);}}
.exp-name{{color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.exp-source{{color:var(--muted);font-size:11px;}}
.exp-amount{{font-family:var(--mono);color:var(--red);text-align:left;}}

.cat-row{{display:grid;grid-template-columns:1fr 100px 60px 80px;gap:12px;align-items:center;padding:10px 0;border-bottom:1px solid var(--border);font-size:13px;}}
.cat-row:last-child{{border-bottom:none;}}
.cat-name{{color:var(--text);}}
.cat-bar-wrap{{background:var(--bg3);border-radius:4px;height:6px;overflow:hidden;}}
.cat-bar{{height:100%;background:var(--yellow);border-radius:4px;}}
.cat-pct{{color:var(--muted);font-size:12px;text-align:center;}}
.cat-amount{{font-family:var(--mono);text-align:left;}}

.balance-row{{display:flex;justify-content:space-between;padding:10px 0;border-bottom:1px solid var(--border);font-size:13px;}}
.balance-row:last-child{{border-bottom:none;}}
.balance-type{{color:var(--muted);}}
.balance-val{{font-family:var(--mono);font-weight:600;}}

.footer{{text-align:center;padding:24px;color:var(--muted);font-size:11px;letter-spacing:2px;border-top:1px solid var(--border);}}

@media(max-width:900px){{
  .main{{grid-template-columns:1fr;padding:16px;}}
  .strip{{grid-template-columns:repeat(2,1fr);}}
  .header{{flex-direction:column;gap:16px;text-align:center;}}
  .holding-row{{grid-template-columns:1fr 80px 80px;}}
  .holding-bar-wrap,.holding-pct-port{{display:none;}}
}}
</style>
</head>
<body>

<div class="header">
  <div class="brand">
    <span>Klein Finance</span>
    <strong>מאזן משפחתי</strong>
  </div>
  <div class="nw-block">
    <div class="nw-label">סה"כ נכסים נטו</div>
    <div class="nw-value">{fmt_ils(data["net_worth"])}</div>
  </div>
  <div class="meta">
    <strong>תקופה:</strong> {data.get("period","")}<br>
    <strong>הופק:</strong> {data.get("report_date","")}<br>
    <strong>גרסה:</strong> Klein Finance v1.0
  </div>
</div>

<div class="strip">
  <div class="strip-cell">
    <div class="strip-label">הכנסות</div>
    <div class="strip-value">{fmt_ils(data["income"])}</div>
    <div class="strip-sub">החודש</div>
  </div>
  <div class="strip-cell">
    <div class="strip-label">הוצאות</div>
    <div class="strip-value" style="color:var(--red)">{fmt_ils(data["expenses"])}</div>
    <div class="strip-sub">החודש</div>
  </div>
  <div class="strip-cell">
    <div class="strip-label">עודף</div>
    <div class="strip-value" style="color:var(--green)">{fmt_ils(data["surplus"])}</div>
    <div class="strip-sub">יתרה חודשית</div>
  </div>
  <div class="strip-cell">
    <div class="strip-label">חיסכון</div>
    <div class="strip-value" style="color:var(--gold)">{fmt_pct(data["savings_rate"])}</div>
    <div class="strip-sub">מהכנסה</div>
  </div>
  <div class="strip-cell">
    <div class="strip-label">תיק מסחר</div>
    <div class="strip-value">{fmt_ils(data["portfolio_val"])}</div>
    <div class="strip-sub">כולל כספית</div>
  </div>
</div>

<div class="main">

  <!-- Holdings -->
  <div class="card card-full">
    <div class="card-title">
      <span>תיק השקעות</span>
      <span class="card-title-accent">{fmt_ils(data["portfolio_val"])}</span>
    </div>
    {holdings_html}
  </div>

  <!-- Pension -->
  <div class="card">
    <div class="card-title">
      <span>פנסיה והשתלמות</span>
      <span class="card-title-accent">{fmt_ils(pension_total)}</span>
    </div>
    <div class="pension-grid">
      <div class="pension-item">
        <div class="pension-item-label">פנסיה — דרור</div>
        <div class="pension-item-value">{fmt_ils(data["pension_dror"])}</div>
      </div>
      <div class="pension-item">
        <div class="pension-item-label">פנסיה — ליאת</div>
        <div class="pension-item-value">{fmt_ils(data["pension_liat"])}</div>
      </div>
      <div class="pension-item">
        <div class="pension-item-label">השתלמות — דרור</div>
        <div class="pension-item-value">{fmt_ils(data["hishtalmut_dror"])}</div>
      </div>
      <div class="pension-item">
        <div class="pension-item-label">השתלמות — ליאת</div>
        <div class="pension-item-value">{fmt_ils(data["hishtalmut_liat"])}</div>
      </div>
    </div>
  </div>

  <!-- RSU -->
  <div class="card">
    <div class="card-title">
      <span>RSU — Align Technology</span>
      <span class="card-title-accent">{fmt_usd(data["rsu_vested"] + data["rsu_unvested"])}</span>
    </div>
    <div class="rsu-row">
      <span class="rsu-label">🟢 זמין למימוש</span>
      <span class="rsu-val" style="color:var(--green)">{fmt_usd(data["rsu_vested"])}</span>
    </div>
    <div class="rsu-row">
      <span class="rsu-label">🔵 טרם הבשיל</span>
      <span class="rsu-val" style="color:var(--blue)">{fmt_usd(data["rsu_unvested"])}</span>
    </div>
    <div class="rsu-row">
      <span class="rsu-label" style="color:var(--text)">סה"כ (בדולרים)</span>
      <span class="rsu-val" style="color:var(--gold)">{fmt_usd(data["rsu_vested"] + data["rsu_unvested"])}</span>
    </div>
  </div>

  <!-- Top expenses -->
  <div class="card">
    <div class="card-title"><span>Top 5 הוצאות</span></div>
    {expenses_html}
  </div>

  <!-- Categories -->
  <div class="card">
    <div class="card-title"><span>Top 3 קטגוריות הוצאה</span></div>
    {cats_html}
  </div>

  <!-- Leumi balances -->
  <div class="card">
    <div class="card-title"><span>ריכוז יתרות לאומי</span></div>
    {''.join(f"""<div class="balance-row"><span class="balance-type">{b["type"]}</span><span class="balance-val" style="color:{'var(--red)' if b['value']<0 else 'var(--text)'}">{fmt_ils(b["value"])}</span></div>""" for b in data.get("balances",[]))}
    <div class="balance-row" style="border-top:2px solid var(--border);margin-top:8px;padding-top:14px;">
      <span class="balance-type" style="color:var(--text);font-weight:600;">משכנתא</span>
      <span class="balance-val" style="color:var(--red)">{fmt_ils(data.get("mortgage",0))}</span>
    </div>
  </div>

</div>

<div class="footer">KLEIN FAMILY FINANCE &nbsp;·&nbsp; {data.get("period","")} &nbsp;·&nbsp; CONFIDENTIAL</div>

</body></html>'''


def upload_to_fileio(html_path):
    import urllib.request
    with open(html_path, 'rb') as f:
        content = f.read()
    boundary = b'----KleinBoundary'
    body = (b'--' + boundary + b'\r\n' +
            b'Content-Disposition: form-data; name="file"; filename="dashboard.html"\r\n' +
            b'Content-Type: text/html\r\n\r\n' +
            content + b'\r\n' +
            b'--' + boundary + b'--\r\n')
    req = urllib.request.Request(
        'https://file.io/?expires=14d',
        data=body,
        headers={'Content-Type': f'multipart/form-data; boundary={boundary.decode()}'},
        method='POST')
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            result = json.loads(r.read())
            if result.get('success'):
                return result.get('link', '')
    except Exception as e:
        print(f"  Upload failed: {e}")
    return None


def main():
    print("\n  Klein Finance - Dashboard Generator v1.0")
    print("  =========================================")

    data = read_workbook()
    if not data:
        input("\n  Press Enter to close..."); return

    html = generate_html(data)
    out = BASE / 'dashboard.html'
    out.write_text(html, encoding='utf-8')
    print(f"  Dashboard saved: {out}")

    # Open in browser
    webbrowser.open(out.as_uri())
    print("  Opened in browser.")

    # Upload for sharing
    print("  Uploading for sharing...")
    link = upload_to_fileio(out)
    if link:
        print(f"\n  Shareable link (14 days, single download):")
        print(f"  {link}")
    else:
        print("  Upload failed — dashboard available locally only.")

    input("\n  Press Enter to close...")

if __name__ == '__main__':
    main()
