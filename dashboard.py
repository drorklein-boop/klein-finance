# Klein Finance - Dashboard Generator v2.0
import sys, json, warnings, webbrowser, datetime, re
warnings.filterwarnings('ignore')
from pathlib import Path

BASE = Path(__file__).parent

def read_workbook():
    import openpyxl
    xlsm = next((f for f in BASE.glob('*.xlsm') if not f.name.startswith('~')), None)
    if not xlsm:
        print("ERROR: No .xlsm file found in", BASE); return None
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)

    d = {}
    ws = wb['דשבורד']
    d['pension_dror']    = float(ws['D10'].value or 0)
    d['pension_liat']    = float(ws['D11'].value or 0)
    d['hishtalmut_dror'] = float(ws['D12'].value or 0)
    d['hishtalmut_liat'] = float(ws['D13'].value or 0)
    d['portfolio']       = float(ws['D14'].value or 0)
    d['bank']            = float(ws['D18'].value or 0)

    rsu = wb['ALIGN RSU']
    d['rsu_vested']   = float(rsu['H13'].value or 0)
    d['rsu_unvested'] = float(rsu['H14'].value or 0)

    ws_r = wb['דוח חודשי']
    rows = list(ws_r.iter_rows(values_only=True))
    def v(label, col=1, exact=False):
        for row in rows:
            if row[0] and (str(row[0]) == label if exact else label in str(row[0])):
                val = row[col] if col < len(row) else None
                return float(val) if isinstance(val, (int, float)) else val
        return None

    d['period']         = v('תקופה:', 3) or ''
    d['report_date']    = v('תאריך הפקה:', 1) or datetime.date.today().strftime('%d/%m/%Y')
    d['income']         = v('הכנסות') or 0
    d['expenses']       = v('הוצאות') or 0
    d['surplus']        = v('יתרה חודשית') or 0
    d['savings_rate']   = v('יחס חיסכון') or 0
    d['portfolio_val']  = v('שווי תיק מסחר', 1) or 0
    d['kaspiyot']       = v('שווי תיק מסחר', 3) or 0
    d['portfolio_chg']  = v('שינוי מחודש קודם') or 0
    d['cumulative_gain']= v('רווח מצטבר') or 0
    d['return_pct']     = v('תשואה מקניה', 3) or 0
    d['monthly_return'] = v('תשואה חודשית:', 3) or 0
    d['target_dist']    = v('מרחק ליעד') or 0
    d['net_worth']      = v('סה"כ נכסים נטו', 10) or 0
    d['invest_surplus'] = v('עודף להשקעה', 10) or 0
    d['bank_free']      = v('יתרה פנויה', 10) or 0
    d['mortgage']       = 0

    # Pension detail
    d['pension_dror_change'] = v('פנסיה - דרור (מגדל)', 11) or 0
    d['pension_total']       = v('פנסיה - סה"כ', 10) or (d['pension_dror']+d['pension_liat'])
    d['hishtalmut_total']    = v('השתלמות - סה"כ', 10) or (d['hishtalmut_dror']+d['hishtalmut_liat'])
    d['pension_monthly_dep'] = v('סה"כ הפקדות חודשיות:', 10) or 0

    # Balance sheet for mortgage
    try:
        ws_bal = wb['ריכוז יתרות לאומי']
        for row in ws_bal.iter_rows(values_only=True):
            if row[0] and 'משכנתא' in str(row[0]) and row[2] and isinstance(row[2], (int,float)):
                d['mortgage'] = abs(float(row[2]))
                break
    except: pass

    # Top 5 expenses
    d['top5'] = []
    for row in rows:
        if row[0] and isinstance(row[0], int) and 1 <= row[0] <= 5 and row[1] and row[2]:
            d['top5'].append({'name': str(row[1]), 'amount': float(row[2]), 'source': str(row[3] or '')})

    # Top 3 categories
    d['top3_cats'] = []
    in_cats = False
    for row in rows:
        if row[0] and 'Top 3 קטגוריות' in str(row[0]): in_cats = True; continue
        if row[0] == 'קטגוריה': continue
        if in_cats and row[0] and isinstance(row[1], (int,float)):
            d['top3_cats'].append({'name': str(row[0]), 'amount': float(row[1]), 'pct': float(row[2] or 0)})
        if len(d['top3_cats']) >= 3: break

    # Investment holdings
    d['holdings'] = []
    try:
        ws_inv = wb['תיק השקעות עדכני']
        for row in ws_inv.iter_rows(values_only=True):
            if row[0] and str(row[0]).isdigit() and row[1] and row[5]:
                chg = float(row[7]) if row[7] is not None else 0
                d['holdings'].append({
                    'name': str(row[1]),
                    'value': float(row[5]),
                    'pct': float(row[9] or 0),
                    'chg': chg
                })
    except: pass

    wb.close()
    return d


def ils(n, show_sign=False):
    try:
        n = float(n)
        sign = '+' if (show_sign and n > 0) else ''
        return f'{sign}₪{abs(n):,.0f}' if not show_sign else f'{sign}₪{n:,.0f}'.replace('₪-','−₪')
    except: return '₪0'

def pct(n, decimals=1):
    try: return f'{float(n)*100:.{decimals}f}%'
    except: return '0%'

def usd(n):
    try: return f'${float(n):,.0f}'
    except: return '$0'


def generate_html(d):
    # Derived values
    pension_total    = d['pension_dror'] + d['pension_liat']
    hishtalmut_total = d['hishtalmut_dror'] + d['hishtalmut_liat']
    pension_all      = pension_total + hishtalmut_total
    rsu_total_usd    = d['rsu_vested'] + d['rsu_unvested']
    # Approx ILS from workbook RSU net values
    rsu_net_ils      = d.get('rsu_net_ils', rsu_total_usd * 3.65)
    liquid           = d['portfolio_val'] + d['bank'] + rsu_total_usd * 3.65
    net_worth        = d['net_worth'] or (liquid + pension_all - d['mortgage'])
    gross            = net_worth + d['mortgage']
    stocks_pct       = 71  # from workbook
    mm_pct           = 29
    kaspiyot_target  = 350000
    kaspiyot_excess  = max(0, d['kaspiyot'] - kaspiyot_target)
    invest_target_pct= int((d['portfolio_val'] / 2000000) * 100)

    # Holdings table rows
    holdings_rows = ''
    for h in d['holdings']:
        arrow = '▲' if h['chg'] >= 0 else '▼'
        clr = 'var(--green)' if h['chg'] >= 0 else 'var(--red)'
        holdings_rows += f'''
              <tr>
                <td>{h["name"]}</td>
                <td style="font-family:var(--mono)">{ils(h["value"])}</td>
                <td style="font-family:var(--mono)">{pct(h["pct"])}</td>
                <td style="color:{clr};font-family:var(--mono)">{arrow} {abs(h["chg"]*100):.1f}%</td>
              </tr>'''

    # Top 5 rows
    top5_rows = ''
    for e in d['top5']:
        top5_rows += f'''
      <div class="exp-row">
        <span class="exp-row-name">{e["name"]}</span>
        <span class="exp-row-badge">{e["source"]}</span>
        <span class="exp-row-amount">{ils(e["amount"])}</span>
      </div>'''

    # Top 3 cats
    cat_rows = ''
    for c in d['top3_cats']:
        bar_w = min(int(c['pct'] * 450), 100)
        cat_rows += f'''
        <div class="cat-row">
          <span class="cat-name">{c["name"]}</span>
          <div class="cat-bar-wrap"><div class="cat-bar-fill" style="width:{bar_w}%"></div></div>
          <span class="cat-pct c-muted" style="font-size:12px;min-width:36px;text-align:left">{pct(c["pct"])}</span>
          <span class="cat-amount" style="font-family:var(--mono);font-size:13px;min-width:72px;text-align:left">{ils(c["amount"])}</span>
        </div>'''

    # Pension detail rows
    pension_rows = f'''
              <tr><td>פנסיה - דרור (מגדל מקפת)</td><td class="c-blue" style="font-family:var(--mono)">{ils(d["pension_dror"])}</td><td class="c-muted" style="font-family:var(--mono)">{pct(d["pension_dror_change"])}</td></tr>
              <tr><td>פנסיה - ליאת (הפניקס)</td><td class="c-blue" style="font-family:var(--mono)">{ils(d["pension_liat"])}</td><td></td></tr>
              <tr><td style="font-weight:600;color:var(--gold)">פנסיה - סה"כ</td><td class="c-gold" style="font-family:var(--mono)">{ils(pension_total)}</td><td></td></tr>
              <tr><td>השתלמות - דרור (אלטשולר)</td><td style="font-family:var(--mono)">{ils(d["hishtalmut_dror"])}</td><td></td></tr>
              <tr><td>השתלמות - ליאת (אלטשולר)</td><td style="font-family:var(--mono)">{ils(d["hishtalmut_liat"])}</td><td></td></tr>
              <tr><td style="font-weight:600;color:var(--gold)">השתלמות - סה"כ</td><td class="c-gold" style="font-family:var(--mono)">{ils(hishtalmut_total)}</td><td></td></tr>
              <tr style="border-top:1px solid var(--border)"><td>סה"כ הפקדות חודשיות</td><td style="font-family:var(--mono)">{ils(d["pension_monthly_dep"])}</td><td></td></tr>'''

    return f'''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>מאזן קליין — {d["period"]}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;800;900&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:#0b0e14; --bg2:#111520; --bg3:#161b28; --bg4:#1c2235;
  --border:#1e2438; --text:#dce4f0; --muted:#6b7897;
  --gold:#c8a45e; --gold2:#e8c47e; --green:#3bbd6e; --yellow:#e8b84b; --red:#e05555; --blue:#7baad4;
  --mono:'JetBrains Mono',monospace;
}}
*,*::before,*::after{{margin:0;padding:0;box-sizing:border-box}}
html,body{{background:var(--bg);color:var(--text);font-family:'Heebo',sans-serif;direction:rtl;min-height:100vh;width:100%;overflow-x:hidden;font-size:15px}}
.header{{background:linear-gradient(135deg,#0d1019,#141928);border-bottom:1px solid var(--border);padding:18px 16px 14px}}
.header-top{{display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px}}
.header-brand{{font-size:11px;color:var(--muted);letter-spacing:2px;text-transform:uppercase}}
.header-brand strong{{display:block;font-size:20px;color:var(--gold);font-weight:800;margin-top:2px}}
.header-meta{{font-size:13px;color:var(--muted);line-height:2;text-align:left}}
.header-meta strong{{color:var(--text)}}
.nw-split{{display:flex;flex-direction:column;margin-top:14px;border:1px solid var(--border);border-radius:10px;overflow:hidden}}
@media(min-width:520px){{.nw-split{{flex-direction:row}}.nw-block{{border-left:1px solid var(--border);border-bottom:none!important}}.nw-block:last-child{{border-left:none}}}}
.nw-block{{flex:1;padding:13px 16px;border-bottom:1px solid var(--border)}}
.nw-block:last-child{{border-bottom:none}}
.nw-label{{font-size:11px;color:var(--muted);letter-spacing:1px;text-transform:uppercase;margin-bottom:4px}}
.nw-val{{font-family:var(--mono);font-size:20px;font-weight:600}}
.nw-sub{{font-size:11px;color:var(--muted);margin-top:3px}}
.strip-wrap{{background:var(--bg2);border-bottom:1px solid var(--border)}}
.strip{{display:grid;grid-template-columns:repeat(3,1fr)}}
@media(min-width:700px){{.strip{{grid-template-columns:repeat(6,1fr)}}}}
.strip-cell{{padding:12px 13px;border-left:1px solid var(--border);border-bottom:1px solid var(--border)}}
.strip-cell:nth-child(3n+1){{border-left:none}}
@media(min-width:700px){{.strip-cell:nth-child(3n+1){{border-left:1px solid var(--border)}}.strip-cell:first-child{{border-left:none}}.strip-cell{{border-bottom:none}}}}
.strip-cell.clickable{{cursor:pointer;transition:background 0.15s}}
.strip-cell.clickable:hover{{background:var(--bg3)}}
.strip-cell.active{{background:var(--bg3);border-bottom:2px solid var(--gold)!important}}
.sc-label{{font-size:11px;color:var(--muted);margin-bottom:3px}}
.sc-value{{font-family:var(--mono);font-size:15px;font-weight:600}}
.sc-sub{{font-size:11px;color:var(--muted);margin-top:2px}}
.strip-panel{{display:none;background:var(--bg3);border-bottom:1px solid var(--border);padding:18px 16px}}
.strip-panel.open{{display:block}}
@media(min-width:700px){{.strip-panel{{padding:20px 28px}}}}
.main{{padding:14px;width:100%}}
@media(min-width:700px){{.main{{padding:20px 28px}}}}
.section-label{{font-size:11px;color:var(--muted);letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;margin-top:22px}}
.portfolio-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:10px}}
@media(min-width:700px){{.portfolio-grid{{grid-template-columns:repeat(4,1fr)}}}}
.pcard{{background:var(--bg2);border:1px solid var(--border);border-radius:9px;overflow:hidden;transition:border-color 0.2s}}
.pcard.expandable{{cursor:pointer}}
.pcard.expandable:hover{{border-color:#2a3150}}
.pcard.open{{border-color:#2a3458;box-shadow:0 4px 18px rgba(0,0,0,0.3);grid-column:1/-1}}
.pcard-head{{padding:13px}}
.pcard-label{{font-size:11px;color:var(--muted);margin-bottom:5px;display:flex;justify-content:space-between;align-items:center}}
.pcard-chevron{{font-size:10px;color:var(--muted);transition:transform 0.25s}}
.pcard.open .pcard-chevron{{transform:rotate(180deg)}}
.pcard-value{{font-family:var(--mono);font-size:18px;font-weight:600}}
.pcard-sub{{font-size:11px;margin-top:4px}}
.pcard-body{{display:none;padding:0 13px 16px;border-top:1px solid var(--border)}}
.pcard.open .pcard-body{{display:block}}
.exp-grid{{display:grid;grid-template-columns:1fr;gap:10px}}
@media(min-width:600px){{.exp-grid{{grid-template-columns:1fr 1fr}}}}
.exp-card{{background:var(--bg2);border:1px solid var(--border);border-radius:9px;padding:14px}}
.exp-card-title{{font-size:11px;color:var(--muted);letter-spacing:1px;text-transform:uppercase;margin-bottom:10px}}
.exp-row{{display:flex;align-items:center;padding:8px 0;border-bottom:1px solid var(--border);gap:7px}}
.exp-row:last-child{{border-bottom:none}}
.exp-row-name{{color:var(--text);flex:1;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.exp-row-badge{{font-size:10px;color:var(--muted);background:var(--bg4);padding:2px 6px;border-radius:8px;flex-shrink:0}}
.exp-row-amount{{font-family:var(--mono);font-size:15px;font-weight:600;color:var(--gold2);flex-shrink:0}}
.cat-row{{display:flex;align-items:center;gap:8px;padding:8px 0;border-bottom:1px solid var(--border)}}
.cat-row:last-child{{border-bottom:none}}
.cat-bar-wrap{{flex:1;height:6px;background:var(--bg4);border-radius:3px;overflow:hidden}}
.cat-bar-fill{{height:100%;background:var(--gold);border-radius:3px}}
.cat-name{{font-size:13px;min-width:110px}}
.d-row{{background:var(--bg2);border:1px solid var(--border);border-radius:9px;margin-bottom:10px;overflow:hidden}}
.d-head{{display:flex;justify-content:space-between;align-items:center;padding:13px;cursor:pointer}}
.d-head:hover{{background:var(--bg3)}}
.d-title{{font-size:13px;font-weight:600;display:flex;align-items:center;gap:8px}}
.d-status{{font-size:11px;color:var(--muted)}}
.d-body{{display:none;padding:14px;border-top:1px solid var(--border)}}
.d-body.open{{display:block}}
.dcards{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}}
.dcard{{background:var(--bg3);border-radius:7px;padding:10px 14px;flex:1;min-width:130px}}
.dcard-label{{font-size:11px;color:var(--muted);margin-bottom:4px}}
.dcard-value{{font-family:var(--mono);font-size:16px;font-weight:600}}
.note{{font-size:12px;color:var(--muted);background:var(--bg4);padding:10px 12px;border-radius:6px;line-height:1.7}}
.note p{{margin-bottom:4px}}.note p:last-child{{margin-bottom:0}}
.prog-bar{{background:var(--bg4);border-radius:4px;height:8px;overflow:hidden;margin:8px 0}}
.prog-fill{{height:100%;background:linear-gradient(90deg,var(--blue),var(--green));border-radius:4px}}
.prog-label{{font-size:11px;color:var(--muted);margin-bottom:4px}}
.prog-info{{display:flex;justify-content:space-between;font-size:11px;color:var(--muted);margin-top:3px}}
.tbl{{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px}}
.tbl th{{text-align:right;font-size:11px;color:var(--muted);border-bottom:1px solid var(--border);padding:6px 0;font-weight:500}}
.tbl td{{padding:8px 0;border-bottom:1px solid var(--border)}}
.tbl tr:last-child td{{border-bottom:none}}
.signal{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-left:6px}}
.sig-green{{background:var(--green);box-shadow:0 0 6px var(--green)}}
.sig-yellow{{background:var(--yellow);box-shadow:0 0 6px var(--yellow)}}
.c-gold{{color:var(--gold2)}} .c-blue{{color:var(--blue)}} .c-green{{color:var(--green)}}
.c-red{{color:var(--red)}} .c-yellow{{color:var(--yellow)}} .c-muted{{color:var(--muted)}}
.income-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:14px}}
@media(min-width:600px){{.income-grid{{grid-template-columns:repeat(4,1fr)}}}}
.income-card{{background:var(--bg4);border-radius:7px;padding:10px 12px}}
.income-label{{font-size:11px;color:var(--muted);margin-bottom:4px}}
.income-value{{font-family:var(--mono);font-size:16px;font-weight:600}}
.footer{{text-align:center;padding:20px;color:var(--muted);font-size:11px;letter-spacing:2px;border-top:1px solid var(--border);margin-top:20px}}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-top">
    <div class="header-brand">Klein Finance<strong>מאזן משפחתי</strong></div>
    <div class="header-meta">
      <div>תקופה: <strong>{d["period"]}</strong></div>
      <div>הופק: <strong>{d["report_date"]}</strong></div>
    </div>
  </div>
  <div class="nw-split">
    <div class="nw-block">
      <div class="nw-label">נכסים נזילים</div>
      <div class="nw-val c-gold">{ils(liquid)}</div>
      <div class="nw-sub">תיק + עו"ש + RSU</div>
    </div>
    <div class="nw-block">
      <div class="nw-label">פנסיה והשתלמות</div>
      <div class="nw-val c-blue">{ils(pension_all)}</div>
      <div class="nw-sub">כולל 4 קרנות</div>
    </div>
    <div class="nw-block">
      <div class="nw-label">סה"כ נכסים לפני מס</div>
      <div class="nw-val c-gold">{ils(net_worth)}</div>
      <div class="nw-sub">משכנתא: <span class="c-red">−{ils(d["mortgage"])}</span> | ברוטו: {ils(gross)}</div>
    </div>
  </div>
</div>

<!-- STRIP -->
<div class="strip-wrap">
  <div class="strip">
    <div class="strip-cell clickable" onclick="togglePanel('income')">
      <div class="sc-label">הכנסות</div>
      <div class="sc-value c-gold">{ils(d["income"])}</div>
    </div>
    <div class="strip-cell clickable" onclick="togglePanel('expenses')">
      <div class="sc-label">הוצאות</div>
      <div class="sc-value">{ils(d["expenses"])}</div>
    </div>
    <div class="strip-cell">
      <div class="sc-label">יתרה חודשית</div>
      <div class="sc-value c-green">{ils(d["surplus"])}</div>
    </div>
    <div class="strip-cell">
      <div class="sc-label">חיסכון</div>
      <div class="sc-value c-green">{pct(d["savings_rate"])}</div>
    </div>
    <div class="strip-cell">
      <div class="sc-label">עו"ש פנוי</div>
      <div class="sc-value c-green">{ils(d["bank"])}</div>
      <div class="sc-sub">פנוי: {ils(d["bank_free"])}</div>
    </div>
    <div class="strip-cell">
      <div class="sc-label">משכנתא</div>
      <div class="sc-value c-red">{ils(d["mortgage"])}</div>
    </div>
  </div>

  <div id="panel-income" class="strip-panel">
    <div class="income-grid">
      <div class="income-card"><div class="income-label">משכורת + בונוס</div><div class="income-value c-gold">{ils(d["income"])}</div><div class="sc-sub">רגיל: ₪41-50K</div></div>
      <div class="income-card"><div class="income-label">שכר דירה</div><div class="income-value c-blue">—</div></div>
      <div class="income-card"><div class="income-label">החזר הוצאות</div><div class="income-value">—</div></div>
      <div class="income-card"><div class="income-label">אחר</div><div class="income-value">—</div></div>
    </div>
  </div>

  <div id="panel-expenses" class="strip-panel">
    <div class="note">סה"כ הוצאות החודש: {ils(d["expenses"])} &nbsp;|&nbsp; עודף להשקעה: {ils(d["invest_surplus"])}</div>
  </div>
</div>

<!-- PORTFOLIO CARDS -->
<div class="main">
  <div class="section-label">תיק נכסים</div>
  <div class="portfolio-grid">

    <!-- עו"ש -->
    <div class="pcard">
      <div class="pcard-head">
        <div class="pcard-label">ריכוז יתרות לאומי</div>
        <div class="pcard-value">{ils(d["bank"])}</div>
        <div class="pcard-sub c-muted">עו"ש + ניירות ערך</div>
      </div>
    </div>

    <!-- תיק מסחר -->
    <div class="pcard expandable" onclick="toggleCard(this)">
      <div class="pcard-head">
        <div class="pcard-label">חשבון מסחר לאומי <span class="pcard-chevron">▼</span></div>
        <div class="pcard-value c-gold">{ils(d["portfolio_val"])}</div>
        <div class="pcard-sub c-green">{int(d["portfolio_val"]/2000000*100)}% מיעד ₪2M &nbsp;|&nbsp; לחץ לפירוט אחזקות</div>
      </div>
      <div class="pcard-body">
        <div style="font-size:12px;color:var(--muted);margin:10px 0 6px">כספית: {ils(d["kaspiyot"])} &nbsp;|&nbsp; שינוי: {ils(d["portfolio_chg"],True)}</div>
        <table class="tbl">
          <tr><th>נייר</th><th>שווי</th><th>% תיק</th><th>שינוי</th></tr>
          {holdings_rows}
        </table>
        <div class="prog-label" style="margin-top:14px">התקדמות ליעד ₪2,000,000</div>
        <div class="prog-bar"><div class="prog-fill" style="width:{invest_target_pct}%"></div></div>
        <div class="prog-info"><span>₪0</span><span class="c-green" style="font-weight:700">{invest_target_pct}%</span><span>₪2M</span></div>
        <div class="note" style="margin-top:10px">רווח מצטבר: {ils(d["cumulative_gain"])} | תשואה מקנייה: {pct(d["return_pct"])} | תשואה חודשית: {pct(d["monthly_return"])}</div>
      </div>
    </div>

    <!-- RSU -->
    <div class="pcard expandable" onclick="toggleCard(this)">
      <div class="pcard-head">
        <div class="pcard-label">RSU — Align Technology <span class="pcard-chevron">▼</span></div>
        <div class="pcard-value">{usd(rsu_total_usd)}</div>
        <div class="pcard-sub c-muted">ALGN</div>
      </div>
      <div class="pcard-body">
        <table class="tbl">
          <tr><th>סוג</th><th>$ USD</th></tr>
          <tr><td>🟢 זמין למימוש</td><td class="c-green" style="font-family:var(--mono)">{usd(d["rsu_vested"])}</td></tr>
          <tr><td>🔵 טרם הבשיל</td><td class="c-blue" style="font-family:var(--mono)">{usd(d["rsu_unvested"])}</td></tr>
          <tr><td style="font-weight:600">סה"כ</td><td class="c-gold" style="font-family:var(--mono)">{usd(rsu_total_usd)}</td></tr>
        </table>
      </div>
    </div>

    <!-- פנסיה -->
    <div class="pcard expandable" onclick="toggleCard(this)">
      <div class="pcard-head">
        <div class="pcard-label">פנסיה והשתלמות <span class="pcard-chevron">▼</span></div>
        <div class="pcard-value c-blue">{ils(pension_all)}</div>
        <div class="pcard-sub c-muted">פנסיה {ils(pension_total)} | השתלמות {ils(hishtalmut_total)}</div>
      </div>
      <div class="pcard-body">
        <table class="tbl">
          <tr><th>קרן</th><th>שווי</th><th>שינוי</th></tr>
          {pension_rows}
        </table>
      </div>
    </div>

  </div>

  <!-- EXPENSES -->
  <div class="section-label">הוצאות החודש</div>
  <div class="exp-grid">
    <div class="exp-card">
      <div class="exp-card-title">Top 5 הוצאות</div>
      {top5_rows}
    </div>
    <div class="exp-card">
      <div class="exp-card-title">Top 3 קטגוריות</div>
      {cat_rows}
    </div>
  </div>

  <!-- DECISIONS -->
  <div class="section-label">החלטות חודשיות</div>

  <div class="d-row">
    <div class="d-head" onclick="toggleDecision(this)">
      <div class="d-title"><span class="signal sig-green"></span> השקעה חודשית — ACWI</div>
      <div class="d-status">עודף: {ils(d["invest_surplus"])} | מרחק ליעד: {ils(d["target_dist"])} | {stocks_pct}% מניות / {mm_pct}% כספית</div>
    </div>
    <div class="d-body">
      <div class="dcards">
        <div class="dcard"><div class="dcard-label">להשקעה ב-15 לחודש</div><div class="dcard-value c-green">{ils(d["invest_surplus"])}</div></div>
        <div class="dcard"><div class="dcard-label">מרחק ליעד ₪2M</div><div class="dcard-value c-yellow">{ils(d["target_dist"])}</div></div>
      </div>
      <div class="prog-label">התקדמות ליעד ₪2,000,000</div>
      <div class="prog-bar"><div class="prog-fill" style="width:{invest_target_pct}%"></div></div>
      <div class="prog-info"><span>₪0</span><span class="c-green" style="font-weight:700">{invest_target_pct}%</span><span>₪2M</span></div>
      <div class="note" style="margin-top:10px">
        <p>• {ils(d["invest_surplus"])} העודף החודשי — כנס ב-15 לחודש ללא תנאי.</p>
      </div>
    </div>
  </div>

  <div class="d-row">
    <div class="d-head" onclick="toggleDecision(this)">
      <div class="d-title"><span class="signal sig-{'yellow' if kaspiyot_excess > 0 else 'green'}"></span> קרן ביטחון — כספית</div>
      <div class="d-status">קיים: {ils(d["kaspiyot"])} | יעד: ₪350K | עודף לפריסה: {ils(kaspiyot_excess)}</div>
    </div>
    <div class="d-body">
      <div class="dcards">
        <div class="dcard"><div class="dcard-label">כספית נוכחית</div><div class="dcard-value c-yellow">{ils(d["kaspiyot"])}</div></div>
        <div class="dcard"><div class="dcard-label">יעד (6 חודשי הוצאות)</div><div class="dcard-value">{ils(kaspiyot_target)}</div></div>
        <div class="dcard"><div class="dcard-label">פנוי לפריסה</div><div class="dcard-value c-green">{ils(kaspiyot_excess)}</div></div>
      </div>
      <div class="note">
        <p>{"עודף מעל היעד — שקול לפרוס לתוך ACWI בהדרגה." if kaspiyot_excess > 0 else "הכספית בטווח היעד. אין צורך בפעולה."}</p>
      </div>
    </div>
  </div>

</div>

<div class="footer">KLEIN FAMILY FINANCE &nbsp;·&nbsp; {d["period"]} &nbsp;·&nbsp; CONFIDENTIAL</div>

<script>
function togglePanel(id) {{
  const panel = document.getElementById('panel-' + id);
  const cell = panel.previousElementSibling.querySelector('[onclick*="' + id + '"]');
  const isOpen = panel.classList.contains('open');
  document.querySelectorAll('.strip-panel').forEach(p => p.classList.remove('open'));
  document.querySelectorAll('.strip-cell').forEach(c => c.classList.remove('active'));
  if (!isOpen) {{
    panel.classList.add('open');
    if (cell) cell.classList.add('active');
  }}
}}
function toggleCard(card) {{
  const isOpen = card.classList.contains('open');
  document.querySelectorAll('.pcard').forEach(c => c.classList.remove('open'));
  if (!isOpen) card.classList.add('open');
}}
function toggleDecision(head) {{
  const body = head.nextElementSibling;
  body.classList.toggle('open');
}}
</script>
</body></html>'''


def upload_to_fileio(html_path):
    import urllib.request
    with open(html_path, 'rb') as f:
        content = f.read()
    boundary = b'----KleinFinanceBoundary7x'
    body = (b'--' + boundary + b'\r\n' +
            b'Content-Disposition: form-data; name="file"; filename="dashboard.html"\r\n' +
            b'Content-Type: text/html\r\n\r\n' +
            content + b'\r\n--' + boundary + b'--\r\n')
    req = urllib.request.Request('https://file.io/?expires=14d', data=body,
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
    print("\n  Klein Finance - Dashboard Generator v2.0")
    print("  =========================================")
    data = read_workbook()
    if not data:
        input("\n  Press Enter to close..."); return

    html = generate_html(data)
    out = BASE / 'dashboard.html'
    out.write_text(html, encoding='utf-8')
    print(f"  Dashboard saved: {out}")
    webbrowser.open(out.as_uri())
    print("  Opened in browser.")
    print("  Uploading for sharing...")
    link = upload_to_fileio(out)
    if link:
        print(f"\n  Shareable link (14 days):\n  {link}")
    else:
        print("  Upload skipped — dashboard available locally.")
    input("\n  Press Enter to close...")

if __name__ == '__main__':
    main()
