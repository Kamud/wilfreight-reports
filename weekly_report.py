#!/usr/bin/env python3
"""Build the weekly truck report (Monday to Sunday, compared with the week before).

The weekly workbook holds its own data: the daily report's log tabs (Trips, Expense Log, Receipts Log,
Transfers Log) from the first of the month (or from the Monday of the comparison week, if earlier) to
Sunday, with Cash & Banks, Customers and Assumptions. Every summary figure is a formula over those
logs, so a manual edit to a log line flows through the whole report. The only typed numbers outside
the logs are inputs: opening balances at the first log date, and the Assumptions tab.

It applies no rules of its own: the logs come from daily-report workbooks made by fill_report.py.

  1. Freeze each monthly workbook at its as-at date BEFORE recalculating it:
       python3 weekly_report.py freeze --workbook FILE --as-at 2026-11-22
  2. Build from one or two recalculated monthly workbooks (two when the comparison week starts in
     the previous month) and the fill_report.py JSON summary for the month holding Sunday:
       python3 weekly_report.py build --week-end 2026-11-22 --monthly "out/weekly/month 2026-11.xlsx" \
           [--monthly "out/weekly/month 2026-10.xlsx"] --summary out/weekly/summary_2026-11.json \
           [--out "out/Wilfreight Logistics Weekly Report 16-22 November 2026.xlsx"]
  3. Recalculate the weekly workbook (xlsx skill recalc.py), then:
       python3 weekly_report.py verify --weekly "out/Wilfreight Logistics Weekly Report 16-22 November 2026.xlsx" \
           --monthly "out/weekly/month 2026-11.xlsx"
     It checks the "should be 0" block on Week Summary and compares month to date, cash and
     customer balances with the daily-report workbook.
"""
import argparse, json, re
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = 'Arial'
MONEY = '#,##0.00;(#,##0.00);-'
PCT = '0.0%;(0.0%);-'
CHG_PCT = '+0.0%;-0.0%;-'
PTS = '+0.0" pts";-0.0" pts";-'
NUM = '#,##0;(#,##0);-'
DATE = 'ddd d mmm yy'
HEAD_FILL = PatternFill('solid', start_color='1F3864')
SUB_FILL = PatternFill('solid', start_color='D9E1F2')
TOTAL_FILL = PatternFill('solid', start_color='F2F2F2')
THIN = Side(style='thin', color='BFBFBF')
BLUE = '0000FF'
FIRST, LAST = 6, 2000   # log rows

# Exception groups shown on Week Summary, most important first: (match in fill_report type, group label, action)
ATTENTION = [
    ('no fuel treatment set', 'client-paid fuel lines with no fuel rule (not counted)', 'set the fuel treatment on the Client rules sheet'),
    ('client rule is', 'client fuel lines that conflict with the client rule', 'check the fuel log or the rule'),
    ('not yet clean', 'Vuma expenses not yet clean', 'clear them in Vuma'),
    ('Category changed because of the description', 'expenses re-filed from their description', 'fix the category in Vuma'),
    ('New customer added', 'new customers added', 'enter their opening balance on Customers'),
    ('New account added', 'new accounts added', 'enter their opening balance on Cash & Banks'),
]
OTHER = 'Other note'


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and re.match(r'\d{4}-\d{2}-\d{2}', v):
        return datetime.strptime(v[:10], '%Y-%m-%d').date()
    return None


def period_text(s, e):
    if (s.year, s.month) == (e.year, e.month):
        return f'{s.day}-{e.day} {e:%B %Y}'
    if s.year == e.year:
        return f'{s.day} {s:%B} - {e.day} {e:%B %Y}'
    return f'{s.day} {s:%B %Y} - {e.day} {e:%B %Y}'


def dt(d):
    return datetime(d.year, d.month, d.day)


# ---------------------------------------------------------------- freeze
def cmd_freeze(a):
    wb = openpyxl.load_workbook(a.workbook)
    as_at = to_date(a.as_at)
    wb['Assumptions']['B4'].value = dt(as_at)
    wb['Assumptions']['C4'].value = f'Frozen at {as_at:%d %b %Y} for the weekly report.'
    wb.save(a.workbook)
    print(json.dumps({'frozen': a.workbook, 'as_at': str(as_at)}))


# ---------------------------------------------------------------- read a monthly (daily-report) workbook
def read_monthly(path):
    wv = openpyxl.load_workbook(path, data_only=True)
    wf = openpyxl.load_workbook(path)
    asm = wv['Assumptions']
    if asm['B5'].value is None:
        raise SystemExit(f'{path} has no calculated values: recalculate it first')
    m = {'path': path, 'month_start': to_date(asm['B3'].value), 'as_at': to_date(asm['B4'].value),
         'days_in_month': asm['B5'].value}
    m['month_end'] = m['month_start'].replace(day=m['days_in_month'])

    def log(name, ncols, need_col):
        out = []
        ws, fs = wv[name], wf[name]
        for r in range(6, ws.max_row + 1):
            d = to_date(ws.cell(r, 1).value)
            if d is None or ws.cell(r, need_col).value in (None, ''):
                continue
            row = [fs.cell(r, c).value for c in range(1, ncols + 1)]
            row[0] = d
            out.append(row)
        return out
    # inputs only; formula cells are rebuilt in the weekly workbook
    m['trips'] = log('Trips', 11, 2)
    m['costs'] = log('Expense Log', 12, 4)
    m['receipts'] = log('Receipts Log', 6, 3)
    m['transfers'] = log('Transfers Log', 6, 4)
    a = wf['Assumptions']
    m['asm_trucks'] = [[a.cell(r, c).value for c in range(1, 9)] for r in range(12, 24)]
    m['asm_truck_cats'] = [a.cell(r, 11).value for r in range(12, 20)]
    m['asm_overheads'] = [[a.cell(r, 13).value, a.cell(r, 14).value, a.cell(r, 15).value] for r in range(12, 28)]
    m['asm_notes'] = [a.cell(r, 1).value for r in range(26, 40) if a.cell(r, 1).value]
    cb = wv['Cash & Banks']
    m['accounts'] = [(cb.cell(r, 1).value, float(cb.cell(r, 2).value or 0)) for r in range(7, 23) if cb.cell(r, 1).value]
    cu = wv['Customers']
    m['customers'] = [(cu.cell(r, 1).value, float(cu.cell(r, 2).value or 0)) for r in range(7, 47) if cu.cell(r, 1).value]
    ts = wv['Truck Summary']
    m['mtd'] = {ts.cell(4, c).value: ts.cell(5, c).value for c in range(2, 10)}
    return m


# ---------------------------------------------------------------- small writing helpers
def hdr(ws, row, c1, labels, fill=HEAD_FILL, color='FFFFFF'):
    for i, h in enumerate(labels):
        cell = ws.cell(row, c1 + i)
        cell.value = h
        cell.font = Font(name=FONT, bold=True, color=color)
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left' if i == 0 else 'center')


def band(ws, row, c1, c2, text, fill=HEAD_FILL, color='FFFFFF'):
    ws.cell(row, c1).value = text
    for c in range(c1, c2 + 1):
        ws.cell(row, c).fill = fill
        ws.cell(row, c).font = Font(name=FONT, bold=True, color=color)


def title(ws, text, sub=None):
    ws['A1'].value = text
    ws['A1'].font = Font(name=FONT, bold=True, size=14, color='1F3864')
    if sub:
        ws['A2'].value = sub
        ws['A2'].font = Font(name=FONT, italic=True, color='595959')


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def note(ws, r, text, c1=1, c2=None, size=9, italic=True, height=None):
    ws.cell(r, c1).value = text
    ws.cell(r, c1).font = Font(name=FONT, size=size, italic=italic, color='595959')
    ws.cell(r, c1).alignment = Alignment(wrap_text=True, vertical='top')
    if c2 and c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    if height:
        ws.row_dimensions[r].height = height


def total_row(ws, r, c1, c2):
    for c in range(c1, c2 + 1):
        ws.cell(r, c).font = Font(name=FONT, bold=True)
        ws.cell(r, c).fill = TOTAL_FILL
        ws.cell(r, c).border = Border(top=THIN)


def arial(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FONT:
                f = c.font
                c.font = Font(name=FONT, bold=f.bold, italic=f.italic, size=f.size, color=f.color)


def colour_change(ws, ref, good_up=True):
    green, red = Font(name=FONT, color='008000'), Font(name=FONT, color='C00000')
    ws.conditional_formatting.add(ref, CellIsRule(operator='greaterThan', formula=['0'], font=green if good_up else red))
    ws.conditional_formatting.add(ref, CellIsRule(operator='lessThan', formula=['0'], font=red if good_up else green))


# references used by many formulas
WS_ = "'Week Summary'!"
TW = (WS_ + '$B$3', WS_ + '$C$3')
LW = (WS_ + '$B$4', WS_ + '$C$4')
MS = WS_ + '$B$5'          # month start (for month to date)


def L(sheet, col):
    return f"'{sheet}'!${col}${FIRST}:${col}${LAST}"


def win(sheet, s, e):
    return f'{L(sheet, "A")},">="&{s},{L(sheet, "A")},"<="&{e}'


TR, EX, RC, TF = 'Trips', 'Expense Log', 'Receipts Log', 'Transfers Log'
TRUCK = f',{L(EX, "K")},"Truck"'
OVH = f',{L(EX, "K")},"Overhead"'
REP = f',{L(EX, "C")},"Repairs"'
FUEL = f',{L(EX, "C")},"Fuel"'
PAID = f',{L(EX, "E")},"<>"'
CREDIT = f',{L(EX, "L")},"<>"'


def f_rev(s, e, extra=''):
    return f'SUMIFS({L(TR, "G")},{win(TR, s, e)}{extra})'


def f_cost(s, e, extra='', col='D'):
    return f'SUMIFS({L(EX, col)},{win(EX, s, e)}{extra})'


# ---------------------------------------------------------------- build
def cmd_build(a):
    week_end = to_date(a.week_end)
    if week_end.weekday() != 6:
        raise SystemExit('--week-end must be a Sunday')
    w_start = week_end - timedelta(days=6)
    p_start = w_start - timedelta(days=7)
    data_start = min(p_start, week_end.replace(day=1))

    months = sorted((read_monthly(p) for p in a.monthly), key=lambda m: m['month_start'])
    cur = next((m for m in months if m['month_start'] <= week_end <= m['month_end']), None)
    if cur is None:
        raise SystemExit('No monthly workbook covers the week-end date')
    if cur['as_at'] != week_end:
        raise SystemExit(f'{cur["path"]} is as at {cur["as_at"]}, not {week_end}: freeze it before recalculating')
    for d in (data_start, w_start, week_end):
        if not any(m['month_start'] <= d <= m['month_end'] for m in months):
            raise SystemExit(f'No monthly workbook covers {d}: pass the other month with a second --monthly')
    for m in months:
        if m is not cur and m['as_at'] < m['month_end']:
            raise SystemExit(f'{m["path"]} stops at {m["as_at"]}; the earlier month must run to its last day')

    inside = lambda d: data_start <= d <= week_end
    pick = lambda key: sorted([r for m in months for r in m[key] if inside(r[0])], key=lambda x: x[0])
    trips, costs, receipts, transfers = pick('trips'), pick('costs'), pick('receipts'), pick('transfers')
    summary = json.load(open(a.summary, encoding='utf-8')) if a.summary else {}

    # opening balances at the first log date, worked back from the daily report's Sunday balances
    acc_flow, cust_flow = {}, {}
    for r in receipts:
        acc_flow[r[4]] = acc_flow.get(r[4], 0) + float(r[2])
        cust_flow[r[1]] = cust_flow.get(r[1], 0) - float(r[2])
    for t in transfers:
        acc_flow[t[2]] = acc_flow.get(t[2], 0) + float(t[3])
        acc_flow[t[1]] = acc_flow.get(t[1], 0) - float(t[3])
    for c in costs:
        if c[4]:
            acc_flow[c[4]] = acc_flow.get(c[4], 0) - float(c[3])
        if c[11]:
            cust_flow[c[11]] = cust_flow.get(c[11], 0) - float(c[3])
    cfg = json.load(open(a.config, encoding='utf-8'))
    rate = cfg['bakers_inn_rate_per_km']
    company = cfg.get('company_name', 'Wilfreight Logistics')
    if not a.out:
        a.out = f'out/{company} Weekly Report {period_text(w_start, week_end)}.xlsx'
    for t in trips:
        g = t[6]
        val = round(float(t[8] or 0) * rate, 2) if isinstance(g, str) and g.startswith('=') else float(g or 0)
        cust_flow[t[3]] = cust_flow.get(t[3], 0) + val
    accounts = [(n, round(bal - acc_flow.get(n, 0), 2)) for n, bal in cur['accounts']]
    customers = [(n, round(bal - cust_flow.get(n, 0), 2)) for n, bal in cur['customers']]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Week Summary'
    names = ['By Truck', 'By Customer', 'By Day', 'Costs', 'Exceptions', 'Trips', 'Expense Log', 'Receipts Log',
             'Transfers Log', 'Cash & Banks', 'Customers', 'Assumptions']
    sh = {n: wb.create_sheet(n) for n in names}

    # ================= Assumptions (inputs, same cells as the daily template)
    asm = sh['Assumptions']
    title(asm, 'Assumptions', 'Copied from the daily report template. Blue cells are inputs; change them here and every tab follows.')
    asm['A3'], asm['B3'] = 'Logs start', dt(data_start)
    asm['A4'], asm['B4'] = 'Logs end (Sunday)', f'={TW[1]}'
    asm['C3'] = 'First date held in the log tabs. Opening balances on Cash & Banks and Customers are at this date.'
    for c in ('B3', 'B4'):
        asm[c].number_format = 'd mmm yyyy'
    hdr(asm, 11, 1, ['Truck', 'Model', 'Class', 'Status', 'Licence & insurance per 4 months ($)', 'Cartrack per month ($)',
                     'Remuneration accrual per month ($)', 'Depreciation per month ($)', 'Monthly running total ($)'])
    hdr(asm, 11, 11, ['Cost categories'])
    hdr(asm, 11, 13, ['Overhead category', 'Fixed monthly amount ($)', 'Note'])
    for i, row in enumerate(cur['asm_trucks']):
        r = 12 + i
        for c, v in enumerate(row, 1):
            asm.cell(r, c).value = v
            if c >= 5 and v is not None:
                asm.cell(r, c).font = Font(name=FONT, color=BLUE)
                asm.cell(r, c).number_format = MONEY
        if row[0]:
            asm.cell(r, 9).value = f'=E{r}/4+F{r}+G{r}+H{r}'
            asm.cell(r, 9).number_format = MONEY
    for i, v in enumerate(cur['asm_truck_cats']):
        asm.cell(12 + i, 11).value = v
    for i, (cat, amt, nt) in enumerate(cur['asm_overheads']):
        asm.cell(12 + i, 13).value, asm.cell(12 + i, 14).value, asm.cell(12 + i, 15).value = cat, amt, nt
        if amt is not None:
            asm.cell(12 + i, 14).font = Font(name=FONT, color=BLUE)
            asm.cell(12 + i, 14).number_format = MONEY
    asm.cell(24, 1).value, asm.cell(24, 2).value = 'Overhead', 'Company overhead (not a truck)'
    for i, t in enumerate(cur['asm_notes']):
        note(asm, 26 + i, t, 1, 9)
    note(asm, 26 + len(cur['asm_notes']) + 1,
         'When the comparison week falls in the previous month, these current figures are used for it too.', 1, 9)
    widths(asm, [22, 20, 10, 11, 14, 12, 14, 13, 13, 2, 18, 2, 30, 13, 60])

    # ================= Log tabs (data; formula columns rebuilt as in the daily template)
    def log_tab(name, sub, heads, rows_, writer, wdth):
        s = sh[name]
        title(s, name, sub)
        hdr(s, 5, 1, heads)
        for i, row in enumerate(rows_):
            writer(s, FIRST + i, row)
        s.freeze_panes = 'A6'
        widths(s, wdth)
        return s

    def w_trip(s, r, v):
        s.cell(r, 1).value = dt(v[0])
        for c in range(2, 12):
            s.cell(r, c).value = v[c - 1]
        g = v[6]
        if isinstance(g, str) and g.startswith('='):
            s.cell(r, 7).value = f'=ROUND(I{r}*{rate},2)'
        s.cell(r, 8).value = f'=IF(G{r}="","",IF(F{r}<>"","Invoiced",IF(G{r}>0,"Estimate","No value")))'
        s.cell(r, 12).value = f'=IF(AND(B{r}<>"",N(G{r})=0),ROW(),"")'
        for c, f in ((1, DATE), (7, MONEY), (9, NUM), (10, 'd mmm yy')):
            s.cell(r, c).number_format = f

    log_tab(TR, 'One line per completed trip. Revenue counts on the completed date. Column L marks trips with no value (used by Week Summary).',
            ['Completed date', 'Trip #', 'Truck', 'Customer', 'Route', 'Invoice #', 'Revenue ($)', 'Basis', 'Km (logged)',
             'Invoice date', 'In Zoho? (Y/N)', 'No-value marker'],
            trips, w_trip, [15, 11, 11, 24, 32, 11, 12, 10, 10, 11, 9, 9])

    def w_cost(s, r, v):
        s.cell(r, 1).value = dt(v[0])
        for c in range(2, 9):
            s.cell(r, c).value = v[c - 1]
        s.cell(r, 12).value = v[11]
        s.cell(r, 9).value = (f'=IF(D{r}="","",IF(OR(B{r}="",C{r}=""),"Incomplete",IF(AND(E{r}="",L{r}=""),"No account",'
                              f'IF(IF(B{r}="Overhead",COUNTIF(Assumptions!$M$12:$M$27,C{r}),COUNTIF(Assumptions!$K$12:$K$19,C{r}))=0,"Check","OK"))))')
        s.cell(r, 10).value = (f'=IF(D{r}="","",IF(C{r}="Truck insurance (prepaid)",0,IF(B{r}="Overhead",'
                               f'IFERROR(IF(INDEX(Assumptions!$N$12:$N$27,MATCH(C{r},Assumptions!$M$12:$M$27,0))>0,0,D{r}),D{r}),D{r})))')
        s.cell(r, 11).value = f'=IF(B{r}="","",IF(B{r}="Overhead","Overhead","Truck"))'
        s.cell(r, 13).value = (f'=IF(AND(N(D{r})>0,A{r}>={TW[0]},A{r}<={TW[1]},ISERROR(SEARCH("fuel",C{r}))),D{r}+ROW()/1000000,"")')
        for c, f in ((1, DATE), (4, MONEY), (10, MONEY)):
            s.cell(r, c).number_format = f

    log_tab(EX, 'Cash expenses from Vuma and client fuel deducted by credit note. Columns I to K and M are formulas.',
            ['Date', 'Applies to (truck or Overhead)', 'Category', 'Amount ($)', 'Paid from (account)', 'Description',
             'Invoice / trip #', 'In Zoho? (Y/N)', 'Check', 'Counted in profit ($)', 'Type', 'Deducted from customer (credit note)',
             'Biggest-cost key (this week, not fuel)'],
            costs, w_cost, [15, 14, 24, 11, 24, 46, 12, 8, 8, 11, 10, 18, 12])

    def w_simple(s, r, v):
        s.cell(r, 1).value = dt(v[0])
        for c in range(2, 7):
            s.cell(r, c).value = v[c - 1]
        s.cell(r, 1).number_format = DATE

    log_tab(RC, 'Money received from customers (Vuma collections).',
            ['Date', 'Customer', 'Amount ($)', 'Reference / remittance #', 'Received into (account)', 'In Zoho? (Y/N)'],
            receipts, w_simple, [15, 26, 12, 40, 26, 9])
    for r in range(FIRST, FIRST + len(receipts)):
        sh[RC].cell(r, 3).number_format = MONEY
    log_tab(TF, 'Money moved between accounts and payouts (Vuma). They net to zero across all accounts.',
            ['Date', 'From account', 'To account', 'Amount ($)', 'Note', 'In Zoho? (Y/N)'],
            transfers, w_simple, [15, 26, 26, 12, 60, 9])
    for r in range(FIRST, FIRST + len(transfers)):
        sh[TF].cell(r, 4).number_format = MONEY

    # ================= Cash & Banks and Customers (opening inputs + formulas)
    def bal_formula(r, upto):
        return (f'=IF($A{r}="","",N(B{r})+SUMIFS({L(RC, "C")},{L(RC, "E")},$A{r},{L(RC, "A")},"<="&{upto})'
                f'+SUMIFS({L(TF, "D")},{L(TF, "C")},$A{r},{L(TF, "A")},"<="&{upto})'
                f'-SUMIFS({L(TF, "D")},{L(TF, "B")},$A{r},{L(TF, "A")},"<="&{upto})'
                f'-SUMIFS({L(EX, "D")},{L(EX, "E")},$A{r},{L(EX, "A")},"<="&{upto}))')

    cb = sh['Cash & Banks']
    title(cb, 'Cash & Banks', f'Balance at {data_start:%d %b %Y} (blue, input) + money in - money out. '
                              'Worked out at the end of each week from the log tabs.')
    hdr(cb, 5, 1, ['Account', f'Balance at {data_start:%d %b} ($)', 'Receipts in ($)', 'Transfers in ($)', 'Transfers out ($)',
                   'Expenses paid ($)', 'Balance end of this week ($)', 'Balance end of last week ($)'])
    n_acc = len(accounts) + 4
    for i in range(n_acc):
        r = 7 + i
        if i < len(accounts):
            cb.cell(r, 1).value, cb.cell(r, 2).value = accounts[i]
            cb.cell(r, 2).font = Font(name=FONT, color=BLUE)
        upto = TW[1]
        cb.cell(r, 3).value = f'=IF($A{r}="","",SUMIFS({L(RC, "C")},{L(RC, "E")},$A{r},{L(RC, "A")},"<="&{upto}))'
        cb.cell(r, 4).value = f'=IF($A{r}="","",SUMIFS({L(TF, "D")},{L(TF, "C")},$A{r},{L(TF, "A")},"<="&{upto}))'
        cb.cell(r, 5).value = f'=IF($A{r}="","",SUMIFS({L(TF, "D")},{L(TF, "B")},$A{r},{L(TF, "A")},"<="&{upto}))'
        cb.cell(r, 6).value = f'=IF($A{r}="","",SUMIFS({L(EX, "D")},{L(EX, "E")},$A{r},{L(EX, "A")},"<="&{upto}))'
        cb.cell(r, 7).value = f'=IF($A{r}="","",N(B{r})+C{r}+D{r}-E{r}-F{r})'
        cb.cell(r, 8).value = bal_formula(r, LW[1])
        for c in range(2, 9):
            cb.cell(r, c).number_format = MONEY
    cb_last = 7 + n_acc - 1
    cb['A6'] = 'TOTAL'
    for c in range(2, 9):
        L_ = get_column_letter(c)
        cb.cell(6, c).value = f'=SUM({L_}7:{L_}{cb_last})'
        cb.cell(6, c).number_format = MONEY
    total_row(cb, 6, 1, 8)
    cb.cell(cb_last + 2, 1).value = 'Transfers check (should be 0)'
    cb.cell(cb_last + 2, 2).value = '=ROUND(D6-E6,2)'
    note(cb, cb_last + 3, f'Opening balances were worked back from the daily report\'s Sunday balances. '
                          'Add a new account in an empty row and type its opening balance.', 1, 8)
    widths(cb, [30, 14, 13, 13, 13, 13, 15, 15])
    cb.freeze_panes = 'B7'

    cu = sh['Customers']
    title(cu, 'Customers', f'Owed at {data_start:%d %b %Y} (blue, input) + revenue - collected - credit notes.')
    hdr(cu, 5, 1, ['Customer', f'Owed at {data_start:%d %b} ($)', 'Revenue ($)', 'Collected ($)', 'Credit notes / client fuel ($)',
                   'Owed end of this week ($)', 'Owed end of last week ($)'])
    n_cu = len(customers) + 6
    for i in range(n_cu):
        r = 7 + i
        if i < len(customers):
            cu.cell(r, 1).value, cu.cell(r, 2).value = customers[i]
            cu.cell(r, 2).font = Font(name=FONT, color=BLUE)
        for c, (sheet, col, key) in ((3, (TR, 'G', 'D')), (4, (RC, 'C', 'B')), (5, (EX, 'D', 'L'))):
            cu.cell(r, c).value = f'=IF($A{r}="","",SUMIFS({L(sheet, col)},{L(sheet, key)},$A{r},{L(sheet, "A")},"<="&{TW[1]}))'
        cu.cell(r, 6).value = f'=IF($A{r}="","",N(B{r})+C{r}-D{r}-E{r})'
        cu.cell(r, 7).value = (f'=IF($A{r}="","",N(B{r})+SUMIFS({L(TR, "G")},{L(TR, "D")},$A{r},{L(TR, "A")},"<="&{LW[1]})'
                               f'-SUMIFS({L(RC, "C")},{L(RC, "B")},$A{r},{L(RC, "A")},"<="&{LW[1]})'
                               f'-SUMIFS({L(EX, "D")},{L(EX, "L")},$A{r},{L(EX, "A")},"<="&{LW[1]}))')
        for c in range(2, 8):
            cu.cell(r, c).number_format = MONEY
    cu_last = 7 + n_cu - 1
    cu['A6'] = 'TOTAL'
    for c in range(2, 8):
        L_ = get_column_letter(c)
        cu.cell(6, c).value = f'=SUM({L_}7:{L_}{cu_last})'
        cu.cell(6, c).number_format = MONEY
    total_row(cu, 6, 1, 7)
    note(cu, cu_last + 2, 'Opening amounts were worked back from the daily report\'s Sunday balances. '
                          'A negative amount means the customer has paid ahead.', 1, 7)
    widths(cu, [30, 14, 13, 13, 15, 15, 15])
    cu.freeze_panes = 'B7'

    # ================= By Day (also the daily accrual of running costs and fixed overheads)
    bd = sh['By Day']
    title(bd, 'Day by day: this week beside the same day last week',
          'Running costs and fixed overheads accrue per day: monthly amounts on Assumptions / days in that month.')
    bd.merge_cells('B4:I4')
    bd.merge_cells('J4:Q4')
    band(bd, 4, 1, 19, '')
    bd['B4'], bd['J4'] = 'THIS WEEK', 'LAST WEEK'
    hdr(bd, 5, 1, ['Day', 'Date', 'Revenue ($)', 'Trip costs ($)', 'Repairs ($)', 'Running costs ($)', 'Overheads ($)', 'Net profit ($)',
                   'Collected ($)', 'Date', 'Revenue ($)', 'Trip costs ($)', 'Repairs ($)', 'Running costs ($)', 'Overheads ($)',
                   'Net profit ($)', 'Collected ($)', 'Share of month (this)', 'Share of month (last)'])
    for i in range(7):
        r = 6 + i
        bd.cell(r, 1).value = f'=TEXT(B{r},"dddd")'
        for base, start, shc in ((2, TW[0], 18), (10, LW[0], 19)):
            D_ = get_column_letter(base)
            bd.cell(r, base).value = f'={start}+{i}'
            bd.cell(r, base).number_format = 'd mmm'
            day = f'{D_}{r}'
            day_win = f'{L(TR, "A")},{day}'
            rev = f'=SUMIFS({L(TR, "G")},{day_win})'
            rep = f'=SUMIFS({L(EX, "D")},{L(EX, "A")},{day}{TRUCK}{REP})'
            trip = f'=SUMIFS({L(EX, "D")},{L(EX, "A")},{day}{TRUCK})-{get_column_letter(base + 3)}{r}'
            run = f'=SUM(Assumptions!$I$12:$I$23)*{get_column_letter(shc)}{r}'
            ovh = f'=SUM(Assumptions!$N$12:$N$27)*{get_column_letter(shc)}{r}+SUMIFS({L(EX, "J")},{L(EX, "A")},{day}{OVH})'
            c = [get_column_letter(base + k) for k in range(1, 8)]
            net = f'={c[0]}{r}-{c[1]}{r}-{c[2]}{r}-{c[3]}{r}-{c[4]}{r}'
            col = f'=SUMIFS({L(RC, "C")},{L(RC, "A")},{day})'
            for k, f in enumerate([rev, trip, rep, run, ovh, net, col]):
                bd.cell(r, base + 1 + k).value = f
                bd.cell(r, base + 1 + k).number_format = MONEY
            bd.cell(r, shc).value = f'=1/DAY(EOMONTH({day},0))'
            bd.cell(r, shc).number_format = '0.0000'
    bd['A13'] = 'WEEK'
    for c in list(range(3, 10)) + list(range(11, 18)) + [18, 19]:
        L_ = get_column_letter(c)
        bd.cell(13, c).value = f'=SUM({L_}6:{L_}12)'
        bd.cell(13, c).number_format = '0.0000' if c >= 18 else MONEY
    total_row(bd, 13, 1, 19)
    widths(bd, [11, 8, 11, 11, 10, 11, 11, 11, 11, 8, 11, 11, 10, 11, 11, 11, 11, 11, 11])
    bd.freeze_panes = 'B6'

    # ================= By Truck
    bt = sh['By Truck']
    title(bt, 'By truck: this week against last week',
          'Running costs: monthly total on Assumptions x share of the month in the week. Company overheads are not split by truck.')
    groups = [('Trips', 'trips'), ('Revenue ($)', 'rev'), ('Trip costs ($)', 'trip'), ('Repairs ($)', 'rep'),
              ('Running costs ($)', 'run'), ('Truck net profit ($)', 'net')]
    hdr(bt, 4, 1, ['', ''] + [g for name, _ in groups for g in (name, '')] + ['', '', '', ''])
    hdr(bt, 5, 1, ['Truck', 'Status'] + ['This week', 'Last week'] * len(groups) +
        ['Change in net profit ($)', 'Contribution this week ($)', 'Contribution last week ($)', 'Repairs change (for the summary)'])
    gc = {}
    for i, (_, key) in enumerate(groups):
        c = 3 + 2 * i
        bt.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
        gc[key] = (get_column_letter(c), get_column_letter(c + 1))
    for i in range(12):
        r = 6 + i
        ar = 12 + i
        bt.cell(r, 1).value = f'=IF(Assumptions!A{ar}="","",Assumptions!A{ar})'
        bt.cell(r, 2).value = f'=IF(Assumptions!A{ar}="","",Assumptions!D{ar})'
        for key, (ca, cb_) in gc.items():
            for col, (s, e), shcol in ((ca, TW, 'R'), (cb_, LW, 'S')):
                x = 0 if col == ca else 1
                if key == 'trips':
                    f = f'=COUNTIFS({L(TR, "C")},$A{r},{win(TR, s, e)})'
                elif key == 'rev':
                    f = f'=SUMIFS({L(TR, "G")},{L(TR, "C")},$A{r},{win(TR, s, e)})'
                elif key == 'rep':
                    f = f'=SUMIFS({L(EX, "D")},{L(EX, "B")},$A{r},{win(EX, s, e)}{TRUCK}{REP})'
                elif key == 'trip':
                    f = f'=SUMIFS({L(EX, "D")},{L(EX, "B")},$A{r},{win(EX, s, e)}{TRUCK})-{gc["rep"][x]}{r}'
                elif key == 'run':
                    f = f"=N(Assumptions!I{ar})*'By Day'!{shcol}13"
                else:
                    rc, tc, pc, uc = (gc[k][x] for k in ('rev', 'trip', 'rep', 'run'))
                    f = f'={rc}{r}-{tc}{r}-{pc}{r}-{uc}{r}'
                bt[f'{col}{r}'] = f
                bt[f'{col}{r}'].number_format = NUM if key == 'trips' else MONEY
        bt[f'O{r}'] = f'={gc["net"][0]}{r}-{gc["net"][1]}{r}'
        bt[f'P{r}'] = f'={gc["rev"][0]}{r}-{gc["trip"][0]}{r}-{gc["rep"][0]}{r}'
        bt[f'Q{r}'] = f'={gc["rev"][1]}{r}-{gc["trip"][1]}{r}-{gc["rep"][1]}{r}'
        bt[f'R{r}'] = f'=IF($A{r}="","",ABS({gc["rep"][0]}{r}-{gc["rep"][1]}{r})+ROW()/1000000)'
        for c in 'OPQ':
            bt[f'{c}{r}'].number_format = MONEY
        bt[f'R{r}'].number_format = MONEY
    bt['A18'] = 'FLEET'
    for c in range(3, 18):
        L_ = get_column_letter(c)
        bt.cell(18, c).value = f'=SUM({L_}6:{L_}17)'
        bt.cell(18, c).number_format = NUM if c in (3, 4) else MONEY
    total_row(bt, 18, 1, 17)
    colour_change(bt, 'O6:O17')
    widths(bt, [22, 11] + [10] * 12 + [12, 12, 12, 10])
    bt.freeze_panes = 'C6'

    # ================= By Customer
    bc = sh['By Customer']
    title(bc, 'By customer: this week against last week', 'Customers come from the Customers tab.')
    cgroups = [('Trips', 'trips'), ('Revenue ($)', 'rev'), ('Collected ($)', 'col'), ('Fuel deducted by credit note ($)', 'cn')]
    hdr(bc, 4, 1, [''] + [g for name, _ in cgroups for g in (name, '')] + ['', '', '', ''])
    hdr(bc, 5, 1, ['Customer'] + ['This week', 'Last week'] * len(cgroups) +
        ['Change in revenue ($)', 'Owed end of this week ($)', 'Rank key: revenue this week', 'Rank key: revenue change'])
    cc = {}
    for i, (_, key) in enumerate(cgroups):
        c = 2 + 2 * i
        bc.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 1)
        cc[key] = (get_column_letter(c), get_column_letter(c + 1))
    for i in range(n_cu):
        r = 6 + i
        cr = 7 + i
        bc.cell(r, 1).value = f'=IF(Customers!A{cr}="","",Customers!A{cr})'
        for key, (ca, cb_) in cc.items():
            for col, (s, e) in ((ca, TW), (cb_, LW)):
                if key == 'trips':
                    f = f'=IF($A{r}="","",COUNTIFS({L(TR, "D")},$A{r},{win(TR, s, e)}))'
                elif key == 'rev':
                    f = f'=IF($A{r}="","",SUMIFS({L(TR, "G")},{L(TR, "D")},$A{r},{win(TR, s, e)}))'
                elif key == 'col':
                    f = f'=IF($A{r}="","",SUMIFS({L(RC, "C")},{L(RC, "B")},$A{r},{win(RC, s, e)}))'
                else:
                    f = f'=IF($A{r}="","",SUMIFS({L(EX, "D")},{L(EX, "L")},$A{r},{win(EX, s, e)}))'
                bc[f'{col}{r}'] = f
                bc[f'{col}{r}'].number_format = NUM if key == 'trips' else MONEY
        bc[f'J{r}'] = f'=IF($A{r}="","",D{r}-E{r})'
        bc[f'K{r}'] = f'=IF($A{r}="","",Customers!F{cr})'
        bc[f'L{r}'] = f'=IF($A{r}="","",N(D{r})+(N(E{r})+N(F{r})+N(G{r}))/10000+1/(ROW()*1000000))'
        bc[f'M{r}'] = f'=IF($A{r}="","",ABS(N(J{r}))+1/(ROW()*1000))'
        for c in 'JKLM':
            bc[f'{c}{r}'].number_format = MONEY
    bc_tot = 6 + n_cu
    bc.cell(bc_tot, 1).value = 'TOTAL'
    for c in range(2, 12):
        L_ = get_column_letter(c)
        bc.cell(bc_tot, c).value = f'=SUM({L_}6:{L_}{bc_tot - 1})'
        bc.cell(bc_tot, c).number_format = NUM if c in (2, 3) else MONEY
    total_row(bc, bc_tot, 1, 11)
    colour_change(bc, f'J6:J{bc_tot - 1}')
    widths(bc, [28] + [11] * 8 + [13, 13, 11, 11])
    bc.freeze_panes = 'B6'

    # ================= Costs
    co = sh['Costs']
    title(co, 'Where the money went', 'Truck costs by category, company overheads, and the largest costs of the week.')
    hdr(co, 4, 1, ['Category', 'This week ($)', 'Last week ($)', 'Change ($)', 'Change %', 'Rank key (for the summary)'])
    r = 5
    band(co, r, 1, 6, 'TRUCK COSTS', SUB_FILL, '1F3864')
    r += 1
    t0 = r
    for i in range(8):
        ar = 12 + i
        co.cell(r, 1).value = f'=IF(Assumptions!K{ar}="","",Assumptions!K{ar})'
        for col, (s, e) in (('B', TW), ('C', LW)):
            co[f'{col}{r}'] = f'=IF($A{r}="","",SUMIFS({L(EX, "D")},{L(EX, "C")},$A{r}{TRUCK},{win(EX, s, e)}))'
        r += 1
    t1 = r - 1
    co.cell(r, 1).value = 'Total truck costs'
    for col in 'BC':
        co[f'{col}{r}'] = f'=SUM({col}{t0}:{col}{t1})'
    total_row(co, r, 1, 5)
    t_tot = r
    r += 2
    band(co, r, 1, 6, 'COMPANY OVERHEADS', SUB_FILL, '1F3864')
    r += 1
    o0 = r
    co.cell(r, 1).value = 'Fixed overheads accrued (Assumptions)'
    co[f'B{r}'] = "=SUM(Assumptions!$N$12:$N$27)*'By Day'!R13"
    co[f'C{r}'] = "=SUM(Assumptions!$N$12:$N$27)*'By Day'!S13"
    r += 1
    for i in range(16):
        ar = 12 + i
        co.cell(r, 1).value = f'=IF(Assumptions!M{ar}="","",Assumptions!M{ar})'
        for col, (s, e) in (('B', TW), ('C', LW)):
            co[f'{col}{r}'] = f'=IF($A{r}="","",SUMIFS({L(EX, "J")},{L(EX, "C")},$A{r}{OVH},{win(EX, s, e)}))'
        r += 1
    o1 = r - 1
    co.cell(r, 1).value = 'Total company overheads'
    for col in 'BC':
        co[f'{col}{r}'] = f'=SUM({col}{o0}:{col}{o1})'
    total_row(co, r, 1, 5)
    o_tot = r
    for rr in list(range(t0, t_tot + 1)) + list(range(o0, o_tot + 1)):
        co[f'D{rr}'] = f'=IF(A{rr}="","",B{rr}-C{rr})'
        co[f'E{rr}'] = f'=IF(OR(A{rr}="",N(C{rr})=0),"",(B{rr}-C{rr})/ABS(C{rr}))'
        for col in 'BCD':
            co[f'{col}{rr}'].number_format = MONEY
        co[f'E{rr}'].number_format = CHG_PCT
        if rr not in (t_tot, o_tot, o0):
            co[f'F{rr}'] = f'=IF(OR(A{rr}="",A{rr}="Repairs"),"",ABS(N(D{rr}))+1/(ROW()*1000))'
            co[f'F{rr}'].number_format = MONEY
    colour_change(co, f'D{t0}:D{t_tot}', good_up=False)
    colour_change(co, f'D{o0}:D{o_tot}', good_up=False)
    r += 2
    band(co, r, 1, 6, 'LARGEST COSTS THIS WEEK (fuel left out: it follows the trips)', SUB_FILL, '1F3864')
    r += 1
    hdr(co, r, 1, ['Applies to', 'Date', 'Amount ($)', 'Category', 'Description'])
    co.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    r += 1
    top0 = r
    EXM = L(EX, 'M')
    for k in range(1, 11):
        key = f'LARGE({EXM},{k})'
        row_ = f'MATCH({key},{EXM},0)'
        get = lambda col: f"=IFERROR(INDEX({L(EX, col)},{row_})" + ('' if col in 'AD' else '&""') + ",\"\")"
        co[f'A{r}'] = get('B')
        co[f'B{r}'] = get('A')
        co[f'C{r}'] = get('D')
        co[f'D{r}'] = get('C')
        co[f'E{r}'] = get('F')
        co.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        co[f'B{r}'].number_format = 'ddd d mmm'
        co[f'B{r}'].alignment = Alignment(horizontal='left')
        co[f'C{r}'].number_format = MONEY
        r += 1
    widths(co, [36, 13, 13, 12, 10, 40])

    # ================= Exceptions (data from the daily fill)
    ex = sh['Exceptions']
    title(ex, f'Exceptions open at {week_end:%d %b %Y}',
          'From the daily report fill for the month. "When" is a formula: items dated in the week are marked This week.')
    hdr(ex, 4, 1, ['Group', 'Type (as the daily report names it)', 'Date', 'When', 'Item', 'Amount ($)'])
    r = 5
    for kind, items in summary.get('exceptions', {}).items():
        hit = next((x for x in ATTENTION if x[0].lower() in kind.lower()), None)
        group = hit[1] if hit else (None if 'Completed trip with no revenue' in kind else OTHER)
        if group is None:
            group = 'trips completed with no revenue'
        for item in items:
            m_ = re.search(r'\d{4}-\d{2}-\d{2}', item)
            amt = re.search(r'\$(-?[\d,]+\.\d{2})', item)
            ex.cell(r, 1).value, ex.cell(r, 2).value, ex.cell(r, 5).value = group, kind, item
            if m_:
                ex.cell(r, 3).value = dt(to_date(m_.group(0)))
                ex.cell(r, 3).number_format = 'd mmm yy'
            ex.cell(r, 4).value = f'=IF(C{r}="","",IF(AND(C{r}>={TW[0]},C{r}<={TW[1]}),"This week","Earlier"))'
            if amt:
                ex.cell(r, 6).value = float(amt.group(1).replace(',', ''))
                ex.cell(r, 6).number_format = MONEY
            for c in (2, 5):
                ex.cell(r, c).alignment = Alignment(wrap_text=True, vertical='top')
            r += 1
    if r == 5:
        ex.cell(5, 1).value = 'No exceptions.'
    widths(ex, [34, 44, 11, 11, 70, 12])
    ex.freeze_panes = 'A5'

    # ================= Week Summary
    W = 8
    ws['A1'] = (f'="{company} Weekly Report "&IF(MONTH(B3)=MONTH(C3),TEXT(B3,"d")&"-"&TEXT(C3,"d mmmm yyyy"),'
                f'IF(YEAR(B3)=YEAR(C3),TEXT(B3,"d mmmm")&" - "&TEXT(C3,"d mmmm yyyy"),TEXT(B3,"d mmmm yyyy")&" - "&TEXT(C3,"d mmmm yyyy")))')
    ws['A1'].font = Font(name=FONT, bold=True, size=14, color='1F3864')
    ws['A2'] = (f'=TEXT({TW[0]},"ddd d mmm")&" to "&TEXT({TW[1]},"ddd d mmm yyyy")&", compared with "'
                f'&TEXT({LW[0]},"d mmm")&" to "&TEXT({LW[1]},"d mmm")&". USD."')
    ws['A2'].font = Font(name=FONT, italic=True, color='595959')
    ws['A3'], ws['B3'], ws['C3'] = 'This week (type the Monday)', dt(w_start), '=B3+6'
    ws['A4'], ws['B4'], ws['C4'] = 'Last week', '=B3-7', '=B3-1'
    ws['A5'], ws['B5'], ws['C5'] = 'Month-to-date period', '=DATE(YEAR(C3),MONTH(C3),1)', '=C3'
    for rr in (3, 4, 5):
        ws.cell(rr, 1).font = Font(name=FONT, size=9, color='595959')
        for c in (2, 3):
            ws.cell(rr, c).number_format = 'd mmm yyyy'
            ws.cell(rr, c).font = Font(name=FONT, size=9, color=BLUE if (rr, c) == (3, 2) else '000000')
    note(ws, 3, f'Logs in this file run from {data_start:%d %b} to {week_end:%d %b %Y}; pick a Monday inside that range.', 4, W)

    # headline (row 7), formula built from the workings in column K onwards
    ws.row_dimensions[7].height = 34
    ws.merge_cells('A7:H7')
    ws['A7'].font = Font(name=FONT, bold=True, size=12)
    ws['A7'].alignment = Alignment(wrap_text=True, vertical='center')

    row = {}
    r = 9

    def kpi_block(label, items):
        nonlocal r
        band(ws, r, 1, W, label)
        r += 1
        hdr(ws, r, 1, ['', 'This week', 'Last week', 'Change', 'Change %', 'Biggest change', '', ''], SUB_FILL, '1F3864')
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=W)
        ws.cell(r, 6).alignment = Alignment(horizontal='left')
        r += 1
        for key, lab, fthis, flast, fmt, good in items:
            row[key] = r
            ws.cell(r, 1).value = lab
            ws[f'B{r}'], ws[f'C{r}'] = fthis, flast
            for col in 'BC':
                ws[f'{col}{r}'].number_format = fmt
            if fmt == PCT:
                ws[f'D{r}'] = f'=IF(OR(B{r}="",C{r}=""),"",(B{r}-C{r})*100)'
                ws[f'D{r}'].number_format = PTS
            else:
                ws[f'D{r}'] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
                ws[f'D{r}'].number_format = fmt
                ws[f'E{r}'] = f'=IF(OR(B{r}="",C{r}="",N(C{r})=0),"",(B{r}-C{r})/ABS(C{r}))'
                ws[f'E{r}'].number_format = CHG_PCT
            if good is not None:
                colour_change(ws, f'D{r}:E{r}', good)
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=W)
            ws.cell(r, 6).font = Font(name=FONT, size=9, color='404040')
            ws.cell(r, 6).alignment = Alignment(wrap_text=True, vertical='center')
            r += 1
        r += 1

    two = lambda f: (f(*TW), f(*LW))
    kpi_block('ACTIVITY', [
        ('trips', 'Trips completed', *two(lambda s, e: f'=COUNTIFS({win(TR, s, e)})'), NUM, True),
        ('noval', '...of which with no value yet', *two(lambda s, e: f'=COUNTIFS({win(TR, s, e)},{L(TR, "G")},0)'), NUM, False),
        ('km', 'Km logged', *two(lambda s, e: f'=SUMIFS({L(TR, "I")},{win(TR, s, e)})'), NUM, True),
        ('trucks', 'Trucks that completed a trip', "=COUNTIF('By Truck'!C6:C17,\">0\")", "=COUNTIF('By Truck'!D6:D17,\">0\")", NUM, True),
    ])
    kpi_block('PROFIT AND LOSS', [
        ('rev', 'Revenue', *two(lambda s, e: '=' + f_rev(s, e)), MONEY, True),
        ('trip', 'Trip costs (fuel, tolls, allowances, fines)', *two(lambda s, e: f'={f_cost(s, e, TRUCK)}-{f_cost(s, e, TRUCK + REP)}'), MONEY, False),
        ('rep', 'Repairs', *two(lambda s, e: '=' + f_cost(s, e, TRUCK + REP)), MONEY, False),
        ('contrib', 'Contribution after repairs', None, None, MONEY, True),
        ('run', 'Truck running costs (accrued)', "='By Day'!F13", "='By Day'!N13", MONEY, False),
        ('ovh', 'Company overheads', "='By Day'!G13", "='By Day'!O13", MONEY, False),
        ('net', 'Net profit', None, None, MONEY, True),
        ('margin', 'Net margin', None, None, PCT, True),
    ])
    kpi_block('KEY RATIOS', [
        ('fuel', 'Fuel (cash fuel and client fuel)', *two(lambda s, e: '=' + f_cost(s, e, TRUCK + FUEL)), MONEY, False),
        ('fuelpct', 'Fuel as % of revenue', None, None, PCT, False),
        ('rpk', 'Revenue per km', None, None, MONEY, True),
        ('rpt', 'Average revenue per trip', None, None, MONEY, True),
    ])
    kpi_block('CASH', [
        ('col', 'Collected from customers', *two(lambda s, e: f'=SUMIFS({L(RC, "C")},{win(RC, s, e)})'), MONEY, True),
        ('spent', 'Cash spent (paid from an account)', *two(lambda s, e: '=' + f_cost(s, e, PAID)), MONEY, False),
        ('cn', 'Client fuel deducted by credit note', *two(lambda s, e: '=' + f_cost(s, e, CREDIT)), MONEY, None),
        ('cash', 'Cash and bank balance at end of week', "='Cash & Banks'!G6", "='Cash & Banks'!H6", MONEY, True),
        ('owed', 'Owed by customers at end of week', '=Customers!F6', '=Customers!G6', MONEY, False),
    ])
    for col in 'BC':
        g = lambda k: f'{col}{row[k]}'
        ws[g('contrib')] = f'={g("rev")}-{g("trip")}-{g("rep")}'
        ws[g('net')] = f'={g("contrib")}-{g("run")}-{g("ovh")}'
        ws[g('margin')] = f'=IFERROR({g("net")}/{g("rev")},"")'
        ws[g('fuelpct')] = f'=IFERROR({g("fuel")}/{g("rev")},"")'
        ws[g('rpk')] = f'=IFERROR({g("rev")}/{g("km")},"")'
        ws[g('rpt')] = f'=IFERROR({g("rev")}/{g("trips")},"")'
    for k in ('rev', 'net'):
        for c in range(1, 6):
            ws.cell(row[k], c).font = Font(name=FONT, bold=True)
            ws.cell(row[k], c).fill = TOTAL_FILL

    # main reasons (formulas)
    money = lambda x: f'TEXT({x},"+$#,##0;-$#,##0")'
    top = lambda rng_, names_: f'INDEX({names_},MATCH(MAX({rng_}),{rng_},0))'
    bc_rows = f'6:{bc_tot - 1}'
    cust_name = top(f"'By Customer'!$M${bc_rows.split(':')[0]}:$M${bc_rows.split(':')[1]}", f"'By Customer'!$A$6:$A${bc_tot - 1}")
    cust_chg = top(f"'By Customer'!$M$6:$M${bc_tot - 1}", f"'By Customer'!$J$6:$J${bc_tot - 1}")
    ws[f'F{row["rev"]}'] = (f'=IFERROR({cust_name}&" "&{money(cust_chg)},"")'
                           f'&IF(B{row["noval"]}>0,"; "&B{row["noval"]}&" trips not valued","")')
    cat_rng = f"Costs!$F${t0}:$F${t1}"
    ws[f'F{row["trip"]}'] = (f'=IFERROR({top(cat_rng, f"Costs!$A${t0}:$A${t1}")}&" "'
                            f'&{money(top(cat_rng, f"Costs!$D${t0}:$D${t1}"))},"")')
    BTR, BTA = "'By Truck'!$R$6:$R$17", "'By Truck'!$A$6:$A$17"
    rep_chg = top(BTR, "'By Truck'!$I$6:$I$17") + '-' + top(BTR, "'By Truck'!$J$6:$J$17")
    ws[f'F{row["rep"]}'] = f'=IFERROR({top(BTR, BTA)}&" "&{money(rep_chg)},"")'
    ws[f'F{row["run"]}'] = 'Fixed daily accrual (Assumptions)'
    oc_rng = f"Costs!$F${o0 + 1}:$F${o1}"
    ws[f'F{row["ovh"]}'] = (f'=IFERROR({top(oc_rng, f"Costs!$A${o0 + 1}:$A${o1}")}&" "'
                           f'&{money(top(oc_rng, f"Costs!$D${o0 + 1}:$D${o1}"))},"")')
    ws[f'F{row["cash"]}'] = 'Cash & Banks tab, at each Sunday'
    ws[f'F{row["owed"]}'] = 'Customers tab, at each Sunday'

    # headline workings (outside the print area)
    K = 11
    ws.cell(8, K).value = 'Workings for the headline (not printed)'
    ws.cell(8, K).font = Font(name=FONT, size=9, italic=True, color='808080')
    items = [('revenue', 'rev', 1), ('trip costs', 'trip', -1), ('repairs', 'rep', -1), ('running costs', 'run', -1), ('overheads', 'ovh', -1)]
    for i, (lab, key, sign) in enumerate(items):
        rr = 9 + i
        ws.cell(rr, K).value = lab
        ws.cell(rr, K + 1).value = f'=D{row[key]}'
        ws.cell(rr, K + 2).value = f'=ABS(N(D{row[key]}))+1/(ROW()*1000)'
        ws.cell(rr, K + 3).value = f'={get_column_letter(K)}{rr}&" "&IF(N(D{row[key]})>0,"rose","fell")&" "&TEXT(ABS(N(D{row[key]})),"$#,##0")'
    kk = get_column_letter(K + 2)
    kt = get_column_letter(K + 3)
    pick_ = lambda n: f'INDEX(${kt}$9:${kt}$13,MATCH(LARGE(${kk}$9:${kk}$13,{n}),${kk}$9:${kk}$13,0))'
    ws['A7'] = (f'="Net "&IF(B{row["net"]}<0,"loss","profit")&" of "&TEXT(ABS(B{row["net"]}),"$#,##0")&" this week, against a "'
                f'&IF(C{row["net"]}<0,"loss","profit")&" of "&TEXT(ABS(C{row["net"]}),"$#,##0")&" last week: "&{pick_(1)}&" and "&{pick_(2)}&"."')

    # trucks section (links to By Truck)
    band(ws, r, 1, W, 'TRUCKS (contribution = revenue less trip costs and repairs; detail on By Truck)')
    r += 1
    hdr(ws, r, 1, ['Truck', 'Trips', 'Revenue', 'Trip costs', 'Repairs', 'Contribution', 'Last week', 'Change'], SUB_FILL, '1F3864')
    r += 1
    tk0 = r
    for i in range(12):
        br = 6 + i
        for c, src in enumerate(['A', 'C', 'E', 'G', 'I', 'P', 'Q'], 1):
            ws.cell(r, c).value = f"=IF('By Truck'!$A{br}=\"\",\"\",'By Truck'!{src}{br})"
            ws.cell(r, c).number_format = NUM if c == 2 else MONEY
        ws.cell(r, 8).value = f'=IF(A{r}="","",F{r}-G{r})'
        ws.cell(r, 8).number_format = MONEY
        r += 1
    ws.cell(r, 1).value = 'All trucks'
    for c in range(2, 9):
        L_ = get_column_letter(c)
        ws.cell(r, c).value = f'=SUM({L_}{tk0}:{L_}{r - 1})'
        ws.cell(r, c).number_format = NUM if c == 2 else MONEY
    total_row(ws, r, 1, W)
    tk_tot = r
    colour_change(ws, f'H{tk0}:H{r - 1}')
    ws.conditional_formatting.add(f'F{tk0}:F{r - 1}', CellIsRule(operator='lessThan', formula=['0'], font=Font(name=FONT, color='C00000')))
    r += 2

    # customers section: top 5 by revenue this week + the rest
    band(ws, r, 1, W, 'CUSTOMERS (top five by revenue this week; all customers on By Customer)')
    r += 1
    hdr(ws, r, 1, ['Customer', 'Trips', 'Revenue', 'Last week', 'Collected', 'Last week', 'Credit notes', 'Owed at Sunday'], SUB_FILL, '1F3864')
    r += 1
    cs0 = r
    BCL = f"'By Customer'!$L$6:$L${bc_tot - 1}"
    for k in range(1, 6):
        m_ = f'MATCH(LARGE({BCL},{k}),{BCL},0)'
        for c, src in enumerate(['A', 'B', 'D', 'E', 'F', 'G', 'H', 'K'], 1):
            ws.cell(r, c).value = f"=IFERROR(INDEX('By Customer'!${src}$6:${src}${bc_tot - 1},{m_}),\"\")"
            ws.cell(r, c).number_format = NUM if c == 2 else MONEY
        r += 1
    ws.cell(r, 1).value = 'All other customers'
    for c, src in enumerate(['B', 'D', 'E', 'F', 'G', 'H', 'K'], 2):
        L_ = get_column_letter(c)
        ws.cell(r, c).value = f"='By Customer'!{src}{bc_tot}-SUM({L_}{cs0}:{L_}{r - 1})"
        ws.cell(r, c).number_format = NUM if c == 2 else MONEY
    ws.cell(r, 1).font = Font(name=FONT, italic=True)
    r += 1
    ws.cell(r, 1).value = 'All customers'
    for c in range(2, 9):
        L_ = get_column_letter(c)
        ws.cell(r, c).value = f'=SUM({L_}{cs0}:{L_}{r - 1})'
        ws.cell(r, c).number_format = NUM if c == 2 else MONEY
    total_row(ws, r, 1, W)
    cs_tot = r
    r += 2

    # biggest costs (links to Costs top list)
    band(ws, r, 1, W, 'BIGGEST COSTS THIS WEEK (fuel left out; ten on the Costs tab)')
    r += 1
    hdr(ws, r, 1, ['Applies to', 'Date', 'Amount', 'Category', '', 'Description', '', ''], SUB_FILL, '1F3864')
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=W)
    for c in (2, 4, 6):
        ws.cell(r, c).alignment = Alignment(horizontal='left')
    r += 1
    for k in range(5):
        tr_ = top0 + k
        ws[f'A{r}'] = f'=Costs!A{tr_}'
        ws[f'B{r}'] = f'=Costs!B{tr_}'
        ws[f'C{r}'] = f'=Costs!C{tr_}'
        ws[f'D{r}'] = f'=Costs!D{tr_}'
        ws[f'F{r}'] = f'=Costs!E{tr_}'
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=W)
        ws[f'B{r}'].number_format = 'ddd d mmm'
        ws[f'B{r}'].alignment = Alignment(horizontal='left')
        ws[f'C{r}'].number_format = MONEY
        ws[f'D{r}'].alignment = Alignment(indent=1)
        r += 1
    r += 1

    # needs attention
    band(ws, r, 1, W, 'NEEDS ATTENTION')
    r += 1
    TRL = L(TR, 'L')
    ws[f'A{r}'] = (f'=IF(COUNT({TRL})=0,"No trips without a value.","•  "&COUNT({TRL})&" trips completed this month with no value in Wilogix"'
                   f'&IF(B{row["noval"]}>0," ("&B{row["noval"]}&" this week)","")&": enter the invoice or trip revenue. They are counted at $0.")')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=W)
    r += 1
    for k in range(1, 6):
        idx = f'SMALL({TRL},{k})-{FIRST - 1}'
        g = lambda col: f'INDEX({L(TR, col)},{idx})'
        ws[f'A{r}'] = (f'=IFERROR("      "&{g("B")}&"  "&{g("D")}&", "&{g("C")}&", "&TEXT({g("I")},"#,##0")&" km, completed "'
                       f'&TEXT({g("A")},"ddd d mmm"),"")')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=W)
        ws[f'A{r}'].font = Font(name=FONT, size=9, color='404040')
        r += 1
    EXA, EXD, EXF = "Exceptions!$A$5:$A$500", "Exceptions!$D$5:$D$500", "Exceptions!$F$5:$F$500"
    for _, label, action in ATTENTION:
        n = f'COUNTIF({EXA},"{label}")'
        new = f'COUNTIFS({EXA},"{label}",{EXD},"This week")'
        amt = f'SUMIFS({EXF},{EXA},"{label}")'
        ws[f'A{r}'] = (f'=IF({n}=0,"","•  "&{n}&" {label}"&IF({new}>0," ("&{new}&" this week)","")'
                       f'&IF({amt}<>0,", "&TEXT({amt},"$#,##0"),"")&": {action}.")')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=W)
        r += 1
    ws[f'A{r}'] = f'=IF(COUNTIF({EXA},"{OTHER}")=0,"",COUNTIF({EXA},"{OTHER}")&" other notes are on the Exceptions tab.")'
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=W)
    ws[f'A{r}'].font = Font(name=FONT, size=9, italic=True, color='595959')
    r += 2

    # month to date (formulas over the logs)
    band(ws, r, 1, W, '="MONTH TO DATE: "&UPPER(TEXT(B5,"d mmm"))&" TO "&UPPER(TEXT(C5,"d mmm yyyy"))')
    r += 1
    hdr(ws, r, 1, ['', 'Revenue', 'Trip costs and repairs', 'Running costs', 'Overheads', 'Net profit', 'Net margin', 'Collected'],
        SUB_FILL, '1F3864')
    r += 1
    mrow = r
    share = 'DAY($C$5)/DAY(EOMONTH($C$5,0))'
    ws[f'A{r}'] = 'Month to date'
    ws[f'B{r}'] = '=' + f_rev('$B$5', '$C$5')
    ws[f'C{r}'] = '=' + f_cost('$B$5', '$C$5', TRUCK)
    ws[f'D{r}'] = f'=SUM(Assumptions!$I$12:$I$23)*{share}'
    ws[f'E{r}'] = f'=SUM(Assumptions!$N$12:$N$27)*{share}+{f_cost("$B$5", "$C$5", OVH, "J")}'
    ws[f'F{r}'] = f'=B{r}-C{r}-D{r}-E{r}'
    ws[f'G{r}'] = f'=IFERROR(F{r}/B{r},"")'
    ws[f'H{r}'] = f'=SUMIFS({L(RC, "C")},{win(RC, "$B$5", "$C$5")})'
    for c in 'BCDEFH':
        ws[f'{c}{r}'].number_format = MONEY
        ws[f'{c}{r}'].font = Font(name=FONT, bold=True)
    ws[f'G{r}'].number_format = PCT
    ws[f'G{r}'].font = Font(name=FONT, bold=True)
    r += 1
    note(ws, r, 'Revenue counts on the trip completion date. Running costs and fixed overheads accrue daily from the Assumptions tab. '
                'Not included: anything not recorded in Wilogix or Vuma, depreciation until it is entered on Assumptions, and Zoho Books.',
         1, W, height=26)
    r += 2

    # checks
    band(ws, r, 1, W, 'CHECKS (each should be 0)', SUB_FILL, '1F3864')
    r += 1
    checks = [
        ('Trucks contribution equals P&L contribution', f'=ROUND(F{tk_tot}-B{row["contrib"]},2)'),
        ('Customers revenue equals P&L revenue', f'=ROUND(C{cs_tot}-B{row["rev"]},2)'),
        ('By Day net profit equals P&L net profit (this week, last week)', f"=ROUND('By Day'!H13-B{row['net']},2)+ROUND('By Day'!P13-C{row['net']},2)"),
        ('Costs tab totals equal P&L costs', f'=ROUND(Costs!B{t_tot}+Costs!B{o_tot}-B{row["trip"]}-B{row["rep"]}-B{row["ovh"]},2)'),
        ('Transfers net to zero across accounts', "='Cash & Banks'!B" + str(cb_last + 2)),
        ('Expense lines not OK in the Check column', f'=COUNTIF({L(EX, "I")},"<>OK")-COUNTBLANK({L(EX, "I")})'),
    ]
    chk0 = r
    for lab, f in checks:
        ws.cell(r, 1).value = lab
        ws.cell(r, 1).font = Font(name=FONT, size=9)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(r, 6).value = f
        ws.cell(r, 6).number_format = '0.00;-0.00;0'
        r += 1
    ws.conditional_formatting.add(f'F{chk0}:F{r - 1}', CellIsRule(operator='notEqual', formula=['0'],
                                                                   font=Font(name=FONT, bold=True, color='C00000')))

    widths(ws, [40, 13, 13, 13, 12, 13, 12, 13, 2, 2, 14, 10, 10, 40])
    ws.freeze_panes = 'A9'
    ws.sheet_view.showGridLines = False
    ws.print_area = f'A1:H{r - 1}'

    for s in wb.worksheets:
        arial(s)
        s.page_setup.orientation = 'portrait' if s.title == 'Week Summary' else 'landscape'
        s.page_setup.paperSize = s.PAPERSIZE_A4
        s.page_setup.fitToWidth, s.page_setup.fitToHeight = 1, 0
        s.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 0
    wb.save(a.out)

    print(json.dumps({
        'week': f'{w_start} to {week_end}', 'previous_week': f'{p_start} to {week_end - timedelta(days=7)}',
        'logs_from': str(data_start), 'output': a.out, 'monthly_workbooks': [m['path'] for m in months],
        'rows': {'trips': len(trips), 'expense_lines': len(costs), 'receipts': len(receipts), 'transfers': len(transfers)},
        'checks_cells': f"Week Summary F{chk0}:F{r - 1}", 'mtd_row': mrow, 'summary_rows': row,
    }, indent=1, default=str))


# ---------------------------------------------------------------- verify (after recalculation)
def cmd_verify(a):
    w = openpyxl.load_workbook(a.weekly, data_only=True)
    ws = w['Week Summary']
    if ws['B3'].value is None or ws['A7'].value is None:
        raise SystemExit('Recalculate the weekly workbook first')
    out = {'checks': {}, 'against_daily_report': {}}
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == 'CHECKS (each should be 0)':
            rr = r + 1
            while ws.cell(rr, 1).value:
                out['checks'][ws.cell(rr, 1).value] = ws.cell(rr, 6).value
                rr += 1
        if ws.cell(r, 1).value == 'Month to date':
            mtd = [ws.cell(r, c).value for c in range(2, 9)]
    m = read_monthly(a.monthly)
    daily = m['mtd']
    pairs = [('Revenue MTD', mtd[0]), ('Trip costs and repairs', mtd[1]), ('Truck running costs', mtd[2]),
             ('Company overheads', mtd[3]), ('Net profit MTD', mtd[4]), ('Collected MTD', mtd[6]),
             ('Cash & bank balance', w['Cash & Banks']['G6'].value)]
    for k, v in pairs:
        out['against_daily_report'][k] = {'weekly': round(v or 0, 2), 'daily': round(daily.get(k) or 0, 2),
                                          'difference': round((v or 0) - (daily.get(k) or 0), 2)}
    owed_daily = sum(b for _, b in m['customers'])
    out['against_daily_report']['Owed by customers'] = {'weekly': round(w['Customers']['F6'].value, 2), 'daily': round(owed_daily, 2),
                                                        'difference': round(w['Customers']['F6'].value - owed_daily, 2)}
    bad = [k for k, v in out['checks'].items() if v not in (0, 0.0)] + \
          [k for k, v in out['against_daily_report'].items() if abs(v['difference']) > 0.01]
    out['status'] = 'ok' if not bad else 'needs review: ' + '; '.join(bad)
    out['headline'] = ws['A7'].value
    print(json.dumps(out, indent=1, default=str))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    f = sub.add_parser('freeze')
    f.add_argument('--workbook', required=True)
    f.add_argument('--as-at', required=True)
    b = sub.add_parser('build')
    b.add_argument('--week-end', required=True)
    b.add_argument('--monthly', action='append', required=True)
    b.add_argument('--summary')
    b.add_argument('--config', default='config.json')
    b.add_argument('--out', help='default: out/<company> Weekly Report <period>.xlsx, e.g. "16-22 November 2026"')
    v = sub.add_parser('verify')
    v.add_argument('--weekly', required=True)
    v.add_argument('--monthly', required=True)
    a = ap.parse_args()
    {'freeze': cmd_freeze, 'build': cmd_build, 'verify': cmd_verify}[a.cmd](a)


if __name__ == '__main__':
    main()
