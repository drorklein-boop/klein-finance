# Klein Finance - Dashboard Generator v3.0.1
import sys, json, warnings, webbrowser, datetime, urllib.request
warnings.filterwarnings('ignore')
from pathlib import Path

BASE = Path(__file__).parent
TEMPLATE_URL = 'https://raw.githubusercontent.com/drorklein-boop/klein-finance/main/dashboard_template.html'
TEMPLATE_LOCAL = BASE / 'dashboard_template.html'

def get_template():
    try:
        url = TEMPLATE_URL + '?t=' + str(int(__import__('time').time()))
        req = urllib.request.Request(url, headers={'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req, timeout=15) as r:
            content = r.read().decode('utf-8')
            TEMPLATE_LOCAL.write_text(content, encoding='utf-8')
            return content
    except Exception as e:
        print(f'  Warning: could not download template ({e}), using local copy')
        if TEMPLATE_LOCAL.exists():
            return TEMPLATE_LOCAL.read_text(encoding='utf-8')
        raise RuntimeError('No template available')

def ils(n):
    try: return f'₪{abs(float(n)):,.0f}'
    except: return '₪0'

def ils_signed(n):
    try:
        n = float(n)
        return ('+₪' if n >= 0 else '−₪') + f'{abs(n):,.0f}'
    except: return '₪0'

def pct(n, decimals=1):
    try: return f'{float(n)*100:.{decimals}f}%'
    except: return '0%'

def usd(n):
    try: return f'${float(n):,.0f}'
    except: return '$0'

def k(n):
    try: return f'₪{float(n)/1000:,.0f}K'
    except: return '₪0K'

def read_workbook():
    import xlwings as xw
    import openpyxl
    xlsm = next((f for f in BASE.glob('*.xlsm') if not f.name.startswith('~')), None)
    if not xlsm:
        print('ERROR: No .xlsm file found'); return None
    # Use xlwings for formula cells (live values), openpyxl for data-only sheets
    xw_app = xw.apps.active
    if not xw_app:
        print('  ERROR: Excel must be open to generate dashboard')
        return None
    xw_wb = next((b for b in xw_app.books if b.name.lower().endswith('.xlsm')), None)
    if not xw_wb:
        print('  ERROR: No .xlsm workbook open in Excel')
        return None

    ws_xw = xw_wb.sheets['דוח חודשי']

    def r(row_1based, col_1based):
        val = ws_xw.range((row_1based, col_1based)).value
        return val

    # Also open with openpyxl for sheets that don't have formulas
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    rsu  = wb['ALIGN RSU']

    d = {}
    d['period']          = r(3, 4) or datetime.date.today().strftime('%B %Y')
    d['report_date']     = str(r(3, 2) or datetime.date.today().strftime('%d.%m.%Y'))
    d['income']          = float(r(6, 2) or 0)   # B6
    d['expenses']        = float(r(7, 2) or 0)   # B7
    d['surplus']         = float(r(8, 2) or 0)   # B8
    d['savings_rate']    = float(r(9, 2) or 0)   # B9
    d['portfolio_val']   = float(r(25, 2) or 0)  # B25
    d['kaspiyot']        = float(r(25, 4) or 0)  # D25
    d['portfolio_chg']   = float(r(26, 2) or 0)  # B26
    d['cumulative_gain'] = float(r(27, 2) or 0)  # B27
    d['return_pct']      = float(r(27, 4) or 0)  # D27
    d['monthly_return']  = float(r(29, 4) or 0)  # D29
    d['target_dist']     = float(r(31, 2) or 0)  # B31
    d['invest_surplus']  = float(r(29, 11) or 0) # K29
    d['bank_free']       = float(r(27, 11) or 0) # K27
    d['bank']            = float(r(24, 11) or 0) # K24
    d['exp_credit']      = abs(float(r(25, 11) or 0)) + abs(float(r(26, 11) or 0))  # K25+K26
    d['net_worth']       = float(r(32, 11) or 0) # K32
    d['pension_total']   = float(r(13, 11) or 0) # K13
    d['hishtalmut_total']= float(r(14, 11) or 0) # K14
    d['pension_dror']    = float(r(17, 11) or 0) # K17
    d['pension_liat']    = float(r(18, 11) or 0) # K18
    d['hishtalmut_dror'] = float(r(19, 11) or 0) # K19
    d['hishtalmut_liat'] = float(r(20, 11) or 0) # K20
    d['pension_monthly'] = float(r(21, 11) or 0) # K21
    d['pension_dror_chg']= float(r(17, 12) or 0) # L17
    d['hishtalmut_dror_f1'] = float(r(19, 11) or 0) * 0.47
    d['hishtalmut_dror_f2'] = float(r(19, 11) or 0) * 0.53

    d['rsu_vested']      = float(rsu['H13'].value or 0)
    d['rsu_unvested']    = float(rsu['H14'].value or 0)

    d['pension_dror_chg']  = float(r(17, 12) or 0)  # L17
    d['pension_monthly']   = float(r(21, 11) or 0)  # K21
    d['hishtalmut_dror_f1']= d['hishtalmut_dror'] * 0.47
    d['hishtalmut_dror_f2']= d['hishtalmut_dror'] * 0.53
    # Read actual Migdal Makefet fund split from מסלקה sheet
    try:
        ws_dror = wb['דרור - מסלקה']
        dror_rows = list(ws_dror.iter_rows(values_only=True))
        makefet_vals = [float(r[4]) for r in dror_rows
                        if r[1] and 'מגדל מקפת' in str(r[1]) and r[4] and isinstance(r[4], (int,float)) and float(r[4]) > 1000]
        if len(makefet_vals) >= 2:
            d['pension_dror_f1'] = makefet_vals[0]
            d['pension_dror_f2'] = makefet_vals[1]
        elif len(makefet_vals) == 1:
            d['pension_dror_f1'] = makefet_vals[0]
            d['pension_dror_f2'] = d['pension_dror'] - makefet_vals[0]
        else:
            d['pension_dror_f1'] = d['pension_dror'] * 0.5
            d['pension_dror_f2'] = d['pension_dror'] * 0.5
    except:
        d['pension_dror_f1'] = d['pension_dror'] * 0.5
        d['pension_dror_f2'] = d['pension_dror'] * 0.5

    # Mortgage from balance sheet
    d['mortgage'] = 244691.90
    try:
        ws_bal = wb['ריכוז יתרות לאומי']
        for row in ws_bal.iter_rows(values_only=True):
            if row[0] and 'משכנתא' in str(row[0]) and row[2] and isinstance(row[2], (int, float)):
                d['mortgage'] = abs(float(row[2]))
                break
    except: pass

    # Top 5
    d['top5'] = []
    for row_i in range(13, 18):  # rows 13-17
        rank = ws_xw.range((row_i, 1)).value
        name = ws_xw.range((row_i, 2)).value
        amt  = ws_xw.range((row_i, 3)).value
        src  = ws_xw.range((row_i, 4)).value
        if rank and isinstance(rank, (int,float)) and name and amt:
            d['top5'].append({'name': str(name), 'amount': float(amt), 'source': str(src or '')})

    # Top 3 cats
    d['top3'] = []
    for row_i in range(20, 23):  # rows 20-22
        name = ws_xw.range((row_i, 1)).value
        amt  = ws_xw.range((row_i, 2)).value
        pct_ = ws_xw.range((row_i, 3)).value
        prev = ws_xw.range((row_i, 4)).value
        if name and amt and isinstance(amt, (int,float)):
            d['top3'].append({'name': str(name), 'amount': float(amt),
                              'pct': float(pct_ or 0), 'prev_pct': float(prev or 0)})

    # Chart historical data — read via xlwings for live formula values
    try:
        ws_tz = xw_wb.sheets['ניתוח תזרים']
        def tz(row_1, col_1):
            v = ws_tz.range((row_1, col_1)).value
            return int(float(v)) if isinstance(v, (int,float)) and v else 0
        def tz_str(row_1, col_1):
            v = ws_tz.range((row_1, col_1)).value
            return str(v) if v else ''
        data_cols = [c for c in range(2, 20)
                     if tz(47, c) > 0 and 'סה' not in tz_str(46, c) and tz_str(46, c)]
        d['chart_months']   = [tz_str(46, c) for c in data_cols]
        d['chart_income']   = [tz(47, c) for c in data_cols]
        d['chart_expenses'] = [tz(48, c) for c in data_cols]
        d['chart_invest']   = [tz(49, c) for c in data_cols]
        d['chart_surplus']  = [tz(52, c) for c in data_cols]
        print(f'  Chart: {len(data_cols)} months: {d["chart_months"]}')
    except Exception as e:
        print(f'  Chart data read failed: {e}')
        d['chart_months'] = []
        d['chart_income'] = d['chart_expenses'] = d['chart_invest'] = d['chart_surplus'] = []
    # Holdings
    d['holdings'] = []
    try:
        ws_inv = wb['תיק השקעות עדכני']
        for row in ws_inv.iter_rows(values_only=True):
            if row[0] and str(row[0]).isdigit() and row[1] and row[5]:
                d['holdings'].append({
                    'name': str(row[1]), 'value': float(row[5]),
                    'pct': float(row[9] or 0), 'chg': float(row[7] or 0)
                })
    except: pass

    wb.close()

    # Derived
    d['pension_all']       = d['pension_total'] + d['hishtalmut_total']
    d['rsu_total']         = d['rsu_vested'] + d['rsu_unvested']
    d['rate']              = 3.65
    d['rsu_ils']           = d['rsu_total'] * d['rate']
    d['rsu_vested_ils']    = d['rsu_vested'] * d['rate']
    d['liquid']            = d['portfolio_val'] + d['bank'] + d['rsu_ils']
    d['gross']             = d['liquid'] + d['pension_all']
    # net_worth read directly from K32; gross computed for display
    d['kaspiyot_excess']   = max(0, d['kaspiyot'] - 350000)
    d['invest_pct']        = int(d['portfolio_val'] / 2000000 * 100)
    d['stocks_pct']        = 71
    d['mm_pct']            = 29
    # Approx ALGN price from RSU vested / vested shares
    d['algn_price']        = 163.38  # use last known; update annually

    return d


def build_top5_html(top5):
    rows = ''
    for e in top5:
        rows += (f'      <div class="exp-row">'
                 f'<span class="exp-row-name">{e["name"]}</span>'
                 f'<span class="exp-row-badge">{e["source"]}</span>'
                 f'<span class="exp-row-amount">{ils(e["amount"])}</span></div>\n')
    return rows

def build_top3_html(top3):
    if not top3: return ''
    max_amt = max(c['amount'] for c in top3)
    rows = ''
    for c in top3:
        bar_w = int(c['amount'] / max_amt * 100) if max_amt else 0
        pct_chg = c['pct'] - c['prev_pct']
        arrow = '▲' if pct_chg >= 0 else '▼'
        clr = 'var(--red)' if pct_chg >= 0 else 'var(--green)'
        rows += (f'      <div class="cat-row">\n'
                 f'        <div class="cat-name">{c["name"]}</div>\n'
                 f'        <div class="cat-bar-wrap"><div class="cat-bar-fill" style="width:{bar_w}%"></div></div>\n'
                 f'        <div class="cat-amount">{ils(c["amount"])}</div>\n'
                 f'        <div class="cat-pct">{pct(c["pct"])}</div>\n'
                 f'      </div>\n')
    return rows

def build_holdings_html(holdings):
    rows = ''
    for h in holdings:
        arrow = '▲' if h['chg'] >= 0 else '▼'
        clr = 'c-green' if h['chg'] >= 0 else 'c-red'
        rows += (f'            <tr>\n'
                 f'              <td>{h["name"]}</td>\n'
                 f'              <td style="font-family:var(--mono)">{ils(h["value"])}</td>\n'
                 f'              <td style="font-family:var(--mono);color:var(--muted)">{pct(h["pct"])}%</td>\n'
                 f'              <td style="font-family:var(--mono);color:var(--{clr.replace("c-","")})">{arrow} {abs(h["chg"]*100):.1f}%</td>\n'
                 f'            </tr>\n')
    return rows

_MONTH_MAP = {
    'January':'ינו','February':'פבר','March':'מרץ','April':'אפר',
    'May':'מאי','June':'יוני','July':'יולי','August':'אוג',
    'September':'ספט','October':'אוק','November':'נוב','December':'דצמ'
}
def _month_he(period):
    for eng, heb in _MONTH_MAP.items():
        if eng in period: return heb
    return period[:3]

def fill_template(template, d):
    # Build dynamic HTML blocks
    top5_html     = build_top5_html(d['top5'])
    top3_html     = build_top3_html(d['top3'])
    holdings_html = build_holdings_html(d['holdings'])

    # Kaspiyot deploy note
    monthly_chunk = d['kaspiyot_excess'] / 3
    deploy_note = f'פרוס: {ils(monthly_chunk)} × 3 חודשים לתוך ACWI.' if d['kaspiyot_excess'] > 0 else 'הכספית בטווח היעד — אין צורך בפעולה.'

    # RSU pct from target
    target = 300
    rsu_pct_from_target = f'+{int((target - d["algn_price"]) / d["algn_price"] * 100)}% ליעד'

    substitutions = {
        '__PERIOD__':             str(d['period']),
        '__CURRENT_MONTH__':       _month_he(str(d['period'])),
        '__REPORT_DATE__':        str(d['report_date']),
        '__LIQUID__':             ils(d['liquid']),
        '__PENSION_ALL__':        ils(d['pension_all']),
        '__NET_WORTH__':          ils(d['net_worth']),
        '__MORTGAGE_NEG__':       '−' + ils(d['mortgage']),
        '__GROSS__':              ils(d['gross']),
        '__RATE__':               f'₪{d["rate"]:.4f}',
        '__RATE_DATE__':          str(d['report_date']),
        '__INCOME__':             ils(d['income']),
        '__EXPENSES__':           ils(d['expenses']),
        '__SURPLUS__':            ils(d['surplus']),
        '__SAVINGS_RATE__':       pct(d['savings_rate']),
        '__BANK_FREE_STRIP__':    ils(d['invest_surplus']),
        '__BANK_FREE__':          ils(d['bank_free']),
        '__MORTGAGE__':           ils(d['mortgage']),
        '__INCOME_ALIGN__':       ils(d['income']),
        '__INCOME_OTHER__':       '—',
        '__EXP_CREDIT__':         ils(d['exp_credit']),
        '__AVG_MONTHLY__':        ils(d['expenses']),
        '__BANK__':               ils(d['bank']),
        '__SEC_VAL__':            ils(d['portfolio_val'] - d['kaspiyot']),
        '__CREDIT_NEG__':         '—',
        '__PORTFOLIO__':          ils(d['portfolio_val']),
        '__INVEST_PCT__':         str(d['invest_pct']),
        '__CUMUL_GAIN__':         ils(d['cumulative_gain']),
        '__RETURN_PCT__':         pct(d['return_pct']),
        '__MONTHLY_RET__':        pct(d['monthly_return']),
        '__RSU_USD__':            usd(d['rsu_total']),
        '__RSU_ILS__':            ils(d['rsu_ils']),
        '__ALGN_PRICE__':         f'${d["algn_price"]:,.2f}',
        '__RSU_VESTED_USD__':     usd(d['rsu_vested']),
        '__RSU_VESTED_ILS__':     ils(d['rsu_vested_ils']),
        '__PENSION_TOTAL_K__':    k(d['pension_total']),
        '__HISHTALMUT_TOTAL_K__': k(d['hishtalmut_total']),
        '__PENSION_CHG__':        ils_signed(d['pension_dror_chg']),
        '__PENSION_DROR_F1__':    ils(d.get('pension_dror_f1', d['pension_dror'] * 0.5)),
        '__PENSION_DROR_F2__':    ils(d.get('pension_dror_f2', d['pension_dror'] * 0.5)),
        '__PENSION_DROR__':       ils(d['pension_dror']),
        '__PENSION_LIAT__':       ils(d['pension_liat']),
        '__HISHTALMUT_DROR_F1__': ils(d['hishtalmut_dror_f1']),
        '__HISHTALMUT_DROR_F2__': ils(d['hishtalmut_dror_f2']),
        '__HISHTALMUT_LIAT__':    ils(d['hishtalmut_liat']),
        '__HISHTALMUT_TOTAL__':   ils(d['hishtalmut_total']),
        '__PENSION_MONTHLY__':    ils(d['pension_monthly']),
        '__INVEST_SURPLUS__':     ils(d['invest_surplus']),
        '__TARGET_DIST__':        ils(d['target_dist']),
        '__STOCKS_PCT__':         str(d['stocks_pct']),
        '__MM_PCT__':             str(d['mm_pct']),
        '__KASPIYOT__':           ils(d['kaspiyot']),
        '__KASPIYOT_EXCESS__':    ils(d['kaspiyot_excess']),
        '__KASPIYOT_DEPLOY_NOTE__': deploy_note,
        '__RSU_PCT_FROM_TARGET__': rsu_pct_from_target,
        '__TOP5_ROWS__':          top5_html,
        '__CHART_MONTHS__':        ','.join(f"'{m}'" for m in d.get('chart_months', [])),
        '__CHART_INCOME__':        ','.join(str(v) for v in d.get('chart_income', [])),
        '__CHART_EXPENSES__':      ','.join(str(v) for v in d.get('chart_expenses', [])),
        '__CHART_INVEST__':        ','.join(str(v) for v in d.get('chart_invest', [])),
        '__CHART_SURPLUS__':       ','.join(str(v) for v in d.get('chart_surplus', [])),
        '__TOP3_ROWS__':          top3_html,
        '__HOLDINGS_ROWS__':      holdings_html,
    }

    for placeholder, value in substitutions.items():
        template = template.replace(placeholder, str(value))

    return template


def upload_to_fileio(html_path):
    import ssl
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with open(html_path, 'rb') as f:
        content = f.read()
    boundary = b'----KleinFinanceBoundary'
    body = (b'--' + boundary + b'\r\n' +
            b'Content-Disposition: form-data; name="file"; filename="dashboard.html"\r\n' +
            b'Content-Type: text/html\r\n\r\n' +
            content + b'\r\n--' + boundary + b'--\r\n')
    req = urllib.request.Request('https://file.io/?expires=14d', data=body,
        headers={'Content-Type': f'multipart/form-data; boundary={boundary.decode()}'},
        method='POST')
    try:
        with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
            result = json.loads(r.read())
            if result.get('success'):
                return result.get('link', '')
    except Exception as e:
        print(f'  Upload failed: {e}')
    return None


def main():
    print('\n  Klein Finance - Dashboard Generator v3.0')
    print('  =========================================')

    data = read_workbook()
    if not data:
        input('\n  Press Enter to close...'); return

    print('  Fetching template...')
    template = get_template()

    print('  Filling data...')
    html = fill_template(template, data)

    out = BASE / 'dashboard.html'
    out.write_text(html, encoding='utf-8')
    print(f'  Dashboard saved: {out}')

    webbrowser.open(out.as_uri())
    print('  Opened in browser.')

    print('  Uploading for sharing...')
    link = upload_to_fileio(out)
    if link:
        print(f'\n  Shareable link (14 days):\n  {link}')
    else:
        print('  Upload failed — dashboard available locally.')

    input('\n  Press Enter to close...')

if __name__ == '__main__':
    main()
