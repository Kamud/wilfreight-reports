#!/usr/bin/env python3
"""Build the weekly truck report (Monday to Sunday, compared with the week before).

The weekly report does not apply any rules of its own. It reads filled and recalculated
daily-report workbooks (made by fill_report.py) and summarises them, so the two reports
always agree.

Two steps:

  1. Freeze each monthly workbook at its as-at date BEFORE recalculating it, so a Monday run
     does not accrue Monday's running costs into Sunday's figures:
       python3 weekly_report.py freeze --workbook FILE --as-at 2026-11-22

  2. Build the weekly workbook from one or two recalculated monthly workbooks (two when the
     fortnight crosses a month end) and the JSON summary fill_report.py printed for the
     month that holds the week's Sunday:
       python3 weekly_report.py build --week-end 2026-11-22 \
           --monthly "out/month 2026-11.xlsx" [--monthly "out/month 2026-10.xlsx"] \
           --summary out/summary_2026-11.json \
           --out "out/Saraquel Weekly Truck Report 2026-11-22.xlsx"

Recalculate the weekly workbook afterwards; it holds formulas over its (hidden) data tabs, and
its hidden Checks tab must show 0 on every line.
"""
import argparse, json, re, sys
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = 'Arial'
MONEY = '#,##0.00;(#,##0.00);-'
MONEY0 = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
NUM = '#,##0;(#,##0);-'
DATE = 'ddd d mmm yy'
HEAD_FILL = PatternFill('solid', start_color='1F3864')
SUB_FILL = PatternFill('solid', start_color='D9E1F2')
TOTAL_FILL = PatternFill('solid', start_color='F2F2F2')
THIN = Side(style='thin', color='BFBFBF')

EXCEPTION_ORDER = [
    'Completed trip with no revenue',
    'no fuel treatment set',
    'client rule is',
    'Category changed because of the description',
    'not yet clean',
    'New customer added',
    'New account added',
]


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and re.match(r'\d{4}-\d{2}-\d{2}', v):
        return datetime.strptime(v[:10], '%Y-%m-%d').date()
    return None


# ---------------------------------------------------------------- freeze
def cmd_freeze(a):
    wb = openpyxl.load_workbook(a.workbook)
    as_at = to_date(a.as_at)
    wb['Assumptions']['B4'].value = datetime(as_at.year, as_at.month, as_at.day)
    wb['Assumptions']['C4'].value = f'Frozen at {as_at:%d %b %Y} for the weekly report.'
    wb.save(a.workbook)
    print(json.dumps({'frozen': a.workbook, 'as_at': str(as_at)}))


# ---------------------------------------------------------------- read monthly workbooks
def rows(ws, first, ncols, key_col=1):
    out = []
    for r in range(first, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ncols + 1)]
        if to_date(vals[key_col - 1]) is None:
            continue
        out.append(vals)
    return out


def read_monthly(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    asm = wb['Assumptions']
    if asm['B5'].value is None:
        raise SystemExit(f'{path} has no calculated values: recalculate it first')
    m = {
        'path': path,
        'month_start': to_date(asm['B3'].value),
        'as_at': to_date(asm['B4'].value),
        'days_in_month': asm['B5'].value,
        'trucks': [],
        'trips': [], 'costs': [], 'receipts': [], 'daily': [],
    }
    for r in range(12, 24):
        name = asm.cell(r, 1).value
        if name:
            m['trucks'].append((name, asm.cell(r, 4).value or '', float(asm.cell(r, 9).value or 0)))
    for v in rows(wb['Trips'], 6, 11):
        if v[1] is None and v[2] is None:
            continue
        m['trips'].append([to_date(v[0]), v[1], v[2], v[3], v[4], v[5], float(v[6] or 0), v[7], float(v[8] or 0)])
    for v in rows(wb['Expense Log'], 6, 12):
        if v[3] in (None, ''):
            continue
        m['costs'].append([to_date(v[0]), v[1], v[2], float(v[3]), v[4] or '', v[5] or '', v[6] or '',
                           float(v[9] or 0), v[10], v[11] or ''])
    for v in rows(wb['Receipts Log'], 6, 5):
        if v[2] in (None, ''):
            continue
        m['receipts'].append([to_date(v[0]), v[1] or '', float(v[2]), v[3] or '', v[4] or ''])
    for v in rows(wb['Daily Trend'], 6, 10):
        m['daily'].append([to_date(v[0])] + [float(x or 0) for x in v[1:10]])
    # balances at the as-at date
    cash = wb['Cash & Banks']
    m['cash_total'] = None
    for r in range(1, 30):
        if cash.cell(r, 1).value == 'TOTAL':
            m['cash_total'] = float(cash.cell(r, 2).value or 0)
            break
    m['owed_total'] = float(wb['Customers']['B6'].value or 0)
    cu = wb['Customers']
    m['owed'] = {cu.cell(r, 1).value: float(cu.cell(r, 2).value or 0) for r in range(7, 47) if cu.cell(r, 1).value}
    ts = wb['Truck Summary']
    m['mtd'] = {ts.cell(4, c).value: ts.cell(5, c).value for c in range(2, 10)}
    return m


# ---------------------------------------------------------------- helpers for writing
def style_header(ws, row, c1, c2, fill=HEAD_FILL, color='FFFFFF'):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name=FONT, bold=True, color=color)
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center' if c > c1 else 'left')


def title(ws, text, sub=None):
    ws['A1'].value = text
    ws['A1'].font = Font(name=FONT, bold=True, size=14, color='1F3864')
    if sub:
        ws['A2'].value = sub
        ws['A2'].font = Font(name=FONT, italic=True, color='595959')


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def font_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and (c.font is None or c.font.name != FONT):
                f = c.font
                c.font = Font(name=FONT, bold=f.bold, italic=f.italic, size=f.size, color=f.color)


def write_table(ws, header, data, formats):
    for c, h in enumerate(header, 1):
        ws.cell(1, c).value = h
    style_header(ws, 1, 1, len(header))
    for r, row in enumerate(data, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c)
            cell.value = datetime(v.year, v.month, v.day) if isinstance(v, date) else v
            if formats[c - 1]:
                cell.number_format = formats[c - 1]
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(header))}{max(2, len(data) + 1)}'


# ---------------------------------------------------------------- build
def fm(x):
    """Money for sentences: whole dollars from $100 up."""
    s = f'${abs(x):,.0f}' if abs(x) >= 100 else f'${abs(x):,.2f}'
    return ('-' if x < 0 else '') + s


def signed(x):
    return ('+' if x >= 0 else '-') + fm(abs(x))


ATTENTION = [
    # (match in exception type, short label, action)
    ('Completed trip with no revenue', 'trips completed with no revenue in Wilogix', 'enter the invoice or trip revenue in Wilogix'),
    ('no fuel treatment set', 'client-paid fuel lines with no fuel rule (not counted)', 'set the fuel treatment on the Client rules sheet'),
    ('client rule is', 'client fuel lines that conflict with the client rule', 'check the fuel log or the rule'),
    ('not yet clean', 'Vuma expenses not yet clean', 'clear them in Vuma'),
    ('Category changed because of the description', 'expenses whose category was changed from the description', 'fix the category in Vuma'),
    ('New customer added', 'new customers added', 'enter their opening balances'),
    ('New account added', 'new accounts added', 'enter their opening balances'),
]


def cmd_build(a):
    week_end = to_date(a.week_end)
    if week_end.weekday() != 6:
        raise SystemExit('--week-end must be a Sunday')
    w_start = week_end - timedelta(days=6)
    p_start, p_end = w_start - timedelta(days=7), w_start - timedelta(days=1)

    months = [read_monthly(p) for p in a.monthly]
    months.sort(key=lambda m: m['month_start'])
    month_end = lambda m: m['month_start'].replace(day=m['days_in_month'])
    cur = next((m for m in months if m['month_start'] <= week_end <= month_end(m)), None)
    if cur is None:
        raise SystemExit('No monthly workbook covers the week-end date')
    if cur['as_at'] != week_end:
        raise SystemExit(f'{cur["path"]} is as at {cur["as_at"]}, not {week_end}: run the freeze step before recalculating')
    for d in (p_start, w_start, week_end):
        if not any(m['month_start'] <= d <= month_end(m) for m in months):
            raise SystemExit(f'No monthly workbook covers {d}: pass the other month with a second --monthly')
    for m in months:
        if m is not cur and m['as_at'] < month_end(m):
            raise SystemExit(f'{m["path"]} stops at {m["as_at"]}; the earlier month must run to its last day')

    inside = lambda d: d is not None and p_start <= d <= week_end
    wk = lambda d: 'W' if d >= w_start else 'P'
    trips = sorted([t for m in months for t in m['trips'] if inside(t[0])], key=lambda x: (x[0], x[1] or ''))
    costs = sorted([c for m in months for c in m['costs'] if inside(c[0])], key=lambda x: x[0])
    receipts = sorted([r for m in months for r in m['receipts'] if inside(r[0])], key=lambda x: x[0])
    daily = sorted([d for m in months for d in m['daily'] if inside(d[0])], key=lambda x: x[0])
    summary = json.load(open(a.summary, encoding='utf-8')) if a.summary else {}
    all_trips = {t[1]: t for m in months for t in m['trips']}

    # ---------- figures worked out here only to write the sentences (the page itself uses formulas)
    agg = {k: {'W': 0.0, 'P': 0.0} for k in ('rev', 'trip', 'rep', 'run', 'oh')}
    by_cust, by_cat, by_truck = {}, {}, {}
    for t in trips:
        agg['rev'][wk(t[0])] += t[6]
        by_cust.setdefault(t[3] or '(none)', {'W': 0.0, 'P': 0.0})[wk(t[0])] += t[6]
        by_truck.setdefault(t[2], {'W': 0.0, 'P': 0.0})
    for c in costs:
        if c[8] == 'Truck':
            key = 'rep' if c[2] == 'Repairs' else 'trip'
            agg[key][wk(c[0])] += c[3]
            if key == 'trip':
                by_cat.setdefault(c[2], {'W': 0.0, 'P': 0.0})[wk(c[0])] += c[3]
            by_truck.setdefault(c[1], {'W': 0.0, 'P': 0.0})
    for d in daily:
        agg['run'][wk(d[0])] += d[5]
        agg['oh'][wk(d[0])] += d[6]
    net = {k: agg['rev'][k] - agg['trip'][k] - agg['rep'][k] - agg['run'][k] - agg['oh'][k] for k in 'WP'}
    delta = {k: agg[k]['W'] - agg[k]['P'] for k in agg}
    zero_w = [t for t in trips if wk(t[0]) == 'W' and not t[6]]

    def top_changes(d, n=2):
        items = sorted(((k, v['W'] - v['P']) for k, v in d.items()), key=lambda x: -abs(x[1]))
        return [f'{k} {signed(v)}' for k, v in items[:n] if abs(v) >= 0.5]

    def top_lines(pred, n=2, field=3):
        ls = sorted([c for c in costs if wk(c[0]) == 'W' and pred(c) and c[field]], key=lambda x: -x[field])[:n]
        out = []
        for c in ls:
            what = c[1] if c[8] == 'Truck' else c[2]
            desc = (c[5] or '').split('(')[0].strip()
            out.append(f'{what} {fm(c[field])}' + (f' ({desc[:22].lower()})' if desc and c[8] == 'Truck' else ''))
        return out

    reasons = {
        'rev': ', '.join(top_changes(by_cust)) + (f'; {len(zero_w)} trip(s) not yet valued' if zero_w else ''),
        'trip': '; '.join(top_changes(by_cat)),
        'rep': ', '.join(top_lines(lambda c: c[8] == 'Truck' and c[2] == 'Repairs')) if agg['rep']['W'] else '',
        'run': 'Fixed daily accrual' if abs(delta['run']) < 0.01 else 'Daily accrual; months of different length',
        'oh': ', '.join(top_lines(lambda c: c[8] == 'Overhead', field=7)),
    }
    names = {'rev': 'revenue', 'trip': 'trip costs', 'rep': 'repairs', 'run': 'running costs', 'oh': 'overheads'}
    effects = sorted(((k, delta[k] if k == 'rev' else -delta[k]) for k in names), key=lambda x: -abs(x[1]))
    drivers = [f"{names[k]} {'rose' if delta[k] > 0 else 'fell'} {fm(abs(delta[k]))}" for k, e in effects[:2] if abs(e) >= 1]
    headline = (f'Net {"profit" if net["W"] >= 0 else "loss"} of {fm(abs(net["W"]))} this week, against a '
                f'{"profit" if net["P"] >= 0 else "loss"} of {fm(abs(net["P"]))} last week'
                + (f': {drivers[0]}' + (f' and {drivers[1]}' if len(drivers) > 1 else '') + '.' if drivers else '.'))

    # ---------- workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Weekly Report'
    data_names = ['Trips data', 'Costs data', 'Receipts data', 'Daily data', 'Checks']
    sheets = {n: wb.create_sheet(n) for n in data_names}

    write_table(sheets['Trips data'], ['Completed date', 'Trip #', 'Truck', 'Customer', 'Route', 'Invoice #', 'Revenue ($)', 'Basis', 'Km'],
                trips, [DATE, None, None, None, None, None, MONEY, None, NUM])
    write_table(sheets['Costs data'], ['Date', 'Applies to', 'Category', 'Amount ($)', 'Paid from', 'Description', 'Ref',
                                       'Counted in profit ($)', 'Type', 'Deducted from customer'],
                costs, [DATE, None, None, MONEY, None, None, None, MONEY, None, None])
    write_table(sheets['Receipts data'], ['Date', 'Customer', 'Amount ($)', 'Reference', 'Into account'],
                receipts, [DATE, None, MONEY, None, None])
    write_table(sheets['Daily data'], ['Date', 'Revenue ($)', 'Trip costs ($)', 'Repairs ($)', 'Contribution ($)', 'Running costs ($)',
                                       'Overheads ($)', 'Net profit ($)', 'Cumulative MTD ($)', 'Collected ($)'],
                daily, [DATE] + [MONEY] * 9)

    T, C, R, D = "'Trips data'", "'Costs data'", "'Receipts data'", "'Daily data'"
    rng = lambda sh, col: f'{sh}!${col}$2:${col}$5000'
    win = lambda sh, s, e: f'{rng(sh, "A")},">="&{s},{rng(sh, "A")},"<="&{e}'
    TRUCK = ',' + rng(C, 'I') + ',"Truck"'
    REP = ',' + rng(C, 'C') + ',"Repairs"'
    PAID = ',' + rng(C, 'E') + ',"<>"'
    CREDIT = ',' + rng(C, 'J') + ',"<>"'
    WK = {'B': ('$B$3', '$C$3'), 'C': ('$B$4', '$C$4')}

    def costs_sum(s, e, extra=''):
        return f'SUMIFS({rng(C, "D")},{win(C, s, e)}{extra})'

    W = 8  # last column used (H)

    def section(r, text):
        ws.cell(r, 1).value = text
        for c in range(1, W + 1):
            ws.cell(r, c).fill = HEAD_FILL
            ws.cell(r, c).font = Font(name=FONT, bold=True, color='FFFFFF')

    def heads(r, labels, start=1):
        for i, h in enumerate(labels):
            cell = ws.cell(r, start + i)
            cell.value = h
            cell.font = Font(name=FONT, bold=True, color='1F3864')
            cell.fill = SUB_FILL
            cell.alignment = Alignment(horizontal='left' if i == 0 else 'right', wrap_text=True)
        for c in range(start + len(labels), W + 1):
            ws.cell(r, c).fill = SUB_FILL

    def note(r, text, c1=1, c2=W, size=9, italic=True, height=None):
        ws.cell(r, c1).value = text
        ws.cell(r, c1).font = Font(name=FONT, size=size, italic=italic, color='595959')
        ws.cell(r, c1).alignment = Alignment(wrap_text=True, vertical='top')
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        if height:
            ws.row_dimensions[r].height = height

    # header
    ws['A1'] = 'Saraquel Logistics: weekly truck report'
    ws['A1'].font = Font(name=FONT, bold=True, size=15, color='1F3864')
    ws['A2'] = f'Week {w_start:%a %d %b} to {week_end:%a %d %b %Y}, compared with {p_start:%d %b} to {p_end:%d %b}. USD.'
    ws['A2'].font = Font(name=FONT, italic=True, color='595959')
    for r, lab, (s, e) in ((3, 'This week', (w_start, week_end)), (4, 'Last week', (p_start, p_end))):
        ws.cell(r, 1).value = lab
        ws.cell(r, 1).font = Font(name=FONT, size=9, color='595959')
        for c, d in ((2, s), (3, e)):
            ws.cell(r, c).value = datetime(d.year, d.month, d.day)
            ws.cell(r, c).number_format = 'd mmm yyyy'
            ws.cell(r, c).font = Font(name=FONT, size=9, color='0000FF')
    ws['A6'] = headline
    ws['A6'].font = Font(name=FONT, bold=True, size=12)
    ws['A6'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.merge_cells('A6:H6')
    ws.row_dimensions[6].height = 34

    # ---------- profit and loss
    r = 8
    section(r, 'PROFIT AND LOSS')
    heads(r + 1, ['', 'This week', 'Last week', 'Change', 'Main reason'])
    ws.merge_cells(start_row=r + 1, start_column=5, end_row=r + 1, end_column=W)
    ws.cell(r + 1, 5).alignment = Alignment(horizontal='left')
    pl = [('Revenue', 'rev', lambda s, e: f'=SUMIFS({rng(T, "G")},{win(T, s, e)})'),
          ('Trip costs (fuel, tolls, allowances)', 'trip', lambda s, e: f'={costs_sum(s, e, TRUCK)}-{costs_sum(s, e, TRUCK + REP)}'),
          ('Repairs', 'rep', lambda s, e: f'={costs_sum(s, e, TRUCK + REP)}'),
          ('Truck running costs', 'run', lambda s, e: f'=SUMIFS({rng(D, "F")},{win(D, s, e)})'),
          ('Company overheads', 'oh', lambda s, e: f'=SUMIFS({rng(D, "G")},{win(D, s, e)})')]
    row = {}
    r += 2
    for label, key, f in pl:
        row[key] = r
        ws.cell(r, 1).value = label
        for col, (s, e) in WK.items():
            ws[f'{col}{r}'] = f(s, e)
        r += 1
    row['net'] = r
    ws.cell(r, 1).value = 'Net profit'
    for col in 'BC':
        ws[f'{col}{r}'] = f'={col}{row["rev"]}-{col}{row["trip"]}-{col}{row["rep"]}-{col}{row["run"]}-{col}{row["oh"]}'
    r += 1
    row['margin'] = r
    ws.cell(r, 1).value = 'Net margin'
    for col in 'BC':
        ws[f'{col}{r}'] = f'=IFERROR({col}{row["net"]}/{col}{row["rev"]},"")'
    for key in ('rev', 'trip', 'rep', 'run', 'oh', 'net'):
        rr = row[key]
        ws[f'D{rr}'] = f'=B{rr}-C{rr}'
        for col in 'BCD':
            ws[f'{col}{rr}'].number_format = MONEY
        if key in reasons and reasons[key]:
            note(rr, reasons[key], 5, W, size=9, italic=False)
            ws.cell(rr, 5).alignment = Alignment(wrap_text=True, vertical='center')
            if len(reasons[key]) > 62:
                ws.row_dimensions[rr].height = 26
    rr = row['margin']
    ws[f'D{rr}'] = f'=IF(OR(B{rr}="",C{rr}=""),"",(B{rr}-C{rr})*100)'
    for col in 'BC':
        ws[f'{col}{rr}'].number_format = PCT
    ws[f'D{rr}'].number_format = '+0.0" pts";-0.0" pts";-'
    for c in range(1, 5):
        ws.cell(row['net'], c).font = Font(name=FONT, bold=True)
        ws.cell(row['net'], c).fill = TOTAL_FILL
        ws.cell(row['net'], c).border = Border(top=THIN)
    green, red = Font(name=FONT, color='008000'), Font(name=FONT, color='C00000')
    for key, good_up in (('rev', True), ('trip', False), ('rep', False), ('run', False), ('oh', False), ('net', True), ('margin', True)):
        ref = f'D{row[key]}'
        ws.conditional_formatting.add(ref, CellIsRule(operator='greaterThan', formula=['0'], font=green if good_up else red))
        ws.conditional_formatting.add(ref, CellIsRule(operator='lessThan', formula=['0'], font=red if good_up else green))

    # ---------- activity and cash
    r += 2
    section(r, 'ACTIVITY AND CASH')
    heads(r + 1, ['', 'This week', 'Last week', 'Change'])
    r += 2
    ac = [('Trips completed', lambda s, e: f'=COUNTIFS({win(T, s, e)})', NUM),
          ('...of which with no value yet', lambda s, e: f'=COUNTIFS({win(T, s, e)},{rng(T, "G")},0)', NUM),
          ('Km logged', lambda s, e: f'=SUMIFS({rng(T, "I")},{win(T, s, e)})', NUM),
          ('Collected from customers', lambda s, e: f'=SUMIFS({rng(R, "C")},{win(R, s, e)})', MONEY),
          ('Cash spent', lambda s, e: f'={costs_sum(s, e, PAID)}', MONEY),
          ('Client fuel deducted by credit note', lambda s, e: f'={costs_sum(s, e, CREDIT)}', MONEY)]
    for label, f, fmt in ac:
        row[label] = r
        ws.cell(r, 1).value = label
        for col, (s, e) in WK.items():
            ws[f'{col}{r}'] = f(s, e)
            ws[f'{col}{r}'].number_format = fmt
        ws[f'D{r}'] = f'=B{r}-C{r}'
        ws[f'D{r}'].number_format = fmt
        r += 1
    row['cash'], row['owed'] = r, r + 1
    ws.cell(r, 1).value = 'Cash and bank balance at Sunday'
    ws[f'B{r}'] = cur['cash_total']
    ws[f'C{r}'] = f'=B{r}-B{row["Collected from customers"]}+B{row["Cash spent"]}'
    ws.cell(r + 1, 1).value = 'Owed by customers at Sunday'
    ws[f'B{r + 1}'] = cur['owed_total']
    ws[f'C{r + 1}'] = (f'=B{r + 1}-B{row["rev"]}+B{row["Collected from customers"]}'
                       f'+B{row["Client fuel deducted by credit note"]}')
    for rr in (r, r + 1):
        ws[f'B{rr}'].font = Font(name=FONT, color='0000FF')
        ws[f'D{rr}'] = f'=B{rr}-C{rr}'
        for col in 'BCD':
            ws[f'{col}{rr}'].number_format = MONEY
    note(r, 'This week: Cash & Banks total in the daily report at Sunday. Last week: worked back with this week\'s collections and cash spent.', 5, W)
    note(r + 1, 'This week: Customers total at Sunday. Last week: worked back with this week\'s revenue, collections and credit notes.', 5, W)
    for key, good_up in (('Trips completed', True), ('...of which with no value yet', False), ('Km logged', True),
                         ('Collected from customers', True), ('cash', True), ('owed', False)):
        ref = f'D{row[key]}'
        ws.conditional_formatting.add(ref, CellIsRule(operator='greaterThan', formula=['0'], font=green if good_up else red))
        ws.conditional_formatting.add(ref, CellIsRule(operator='lessThan', formula=['0'], font=red if good_up else green))
    r += 3

    # ---------- trucks
    section(r, 'TRUCKS (contribution = revenue less trip costs and repairs)')
    heads(r + 1, ['Truck', 'Trips', 'Revenue', 'Trip costs', 'Repairs', 'Contribution', 'Last week', 'Change'])
    r += 2
    status = {n: s for n, s, _ in cur['trucks']}
    active = [n for n, _, _ in cur['trucks'] if n in by_truck]
    idle = [f'{n} ({status[n]})' for n, _, _ in cur['trucks'] if n not in by_truck and n != 'Subcontracted Vehicles']

    def truck_contrib(n, k):
        rev = sum(t[6] for t in trips if t[2] == n and wk(t[0]) == k)
        cst = sum(c[3] for c in costs if c[1] == n and c[8] == 'Truck' and wk(c[0]) == k)
        return rev - cst
    active.sort(key=lambda n: -truck_contrib(n, 'W'))
    t0 = r
    for n in active:
        ws.cell(r, 1).value = n
        s, e = WK['B']
        ws[f'B{r}'] = f'=COUNTIFS({rng(T, "C")},$A{r},{win(T, s, e)})'
        ws[f'C{r}'] = f'=SUMIFS({rng(T, "G")},{rng(T, "C")},$A{r},{win(T, s, e)})'
        ws[f'D{r}'] = (f'=SUMIFS({rng(C, "D")},{rng(C, "B")},$A{r}{TRUCK},{win(C, s, e)})'
                       f'-E{r}')
        ws[f'E{r}'] = f'=SUMIFS({rng(C, "D")},{rng(C, "B")},$A{r}{TRUCK}{REP},{win(C, s, e)})'
        ws[f'F{r}'] = f'=C{r}-D{r}-E{r}'
        s, e = WK['C']
        ws[f'G{r}'] = (f'=SUMIFS({rng(T, "G")},{rng(T, "C")},$A{r},{win(T, s, e)})'
                       f'-SUMIFS({rng(C, "D")},{rng(C, "B")},$A{r}{TRUCK},{win(C, s, e)})')
        ws[f'H{r}'] = f'=F{r}-G{r}'
        ws[f'B{r}'].number_format = NUM
        for col in 'CDEFGH':
            ws[f'{col}{r}'].number_format = MONEY
        r += 1
    t1 = r - 1
    ws.cell(r, 1).value = 'All trucks'
    for col in 'BCDEFGH':
        ws[f'{col}{r}'] = f'=SUM({col}{t0}:{col}{t1})'
        ws[f'{col}{r}'].number_format = NUM if col == 'B' else MONEY
    for c in range(1, W + 1):
        ws.cell(r, c).font = Font(name=FONT, bold=True)
        ws.cell(r, c).fill = TOTAL_FILL
        ws.cell(r, c).border = Border(top=THIN)
    truck_total = r
    ws.conditional_formatting.add(f'F{t0}:F{t1}', CellIsRule(operator='lessThan', formula=['0'], font=red))
    ws.conditional_formatting.add(f'H{t0}:H{t1}', CellIsRule(operator='lessThan', formula=['0'], font=red))
    ws.conditional_formatting.add(f'H{t0}:H{t1}', CellIsRule(operator='greaterThan', formula=['0'], font=green))
    r += 1
    if idle:
        note(r, 'No trips or costs in either week: ' + ', '.join(idle) + '.', 1, W)
        r += 1
    r += 1

    # ---------- customers
    section(r, 'CUSTOMERS')
    heads(r + 1, ['Customer', 'Trips', 'Revenue', 'Last week', 'Collected', 'Last week', 'Owed at Sunday'])
    r += 2
    cust_names = set(by_cust) | {x[1] for x in receipts if x[1]} | {c[9] for c in costs if c[9]}
    cust_names = sorted(cust_names, key=lambda k: (-by_cust.get(k, {'W': 0})['W'], -by_cust.get(k, {'P': 0})['P'], k))
    c0 = r
    for n in cust_names:
        ws.cell(r, 1).value = n
        (s, e), (s2, e2) = WK['B'], WK['C']
        ws[f'B{r}'] = f'=COUNTIFS({rng(T, "D")},$A{r},{win(T, s, e)})'
        ws[f'C{r}'] = f'=SUMIFS({rng(T, "G")},{rng(T, "D")},$A{r},{win(T, s, e)})'
        ws[f'D{r}'] = f'=SUMIFS({rng(T, "G")},{rng(T, "D")},$A{r},{win(T, s2, e2)})'
        ws[f'E{r}'] = f'=SUMIFS({rng(R, "C")},{rng(R, "B")},$A{r},{win(R, s, e)})'
        ws[f'F{r}'] = f'=SUMIFS({rng(R, "C")},{rng(R, "B")},$A{r},{win(R, s2, e2)})'
        owed = cur['owed'].get(n)
        if owed is not None:
            ws[f'G{r}'] = owed
            ws[f'G{r}'].font = Font(name=FONT, color='0000FF')
        ws[f'B{r}'].number_format = NUM
        for col in 'CDEFG':
            ws[f'{col}{r}'].number_format = MONEY
        r += 1
    c1 = r - 1
    ws.cell(r, 1).value = 'All customers'
    for col in 'BCDEF':
        ws[f'{col}{r}'] = f'=SUM({col}{c0}:{col}{c1})'
        ws[f'{col}{r}'].number_format = NUM if col == 'B' else MONEY
    ws[f'G{r}'] = cur['owed_total']
    ws[f'G{r}'].number_format = MONEY
    for c in range(1, W + 1):
        ws.cell(r, c).font = Font(name=FONT, bold=True)
        ws.cell(r, c).fill = TOTAL_FILL
        ws.cell(r, c).border = Border(top=THIN)
    cust_total = r
    r += 1
    note(r, 'Owed at Sunday is from the Customers tab of the daily report; it includes every customer, so the total can exceed the rows shown.', 1, W)
    r += 2

    # ---------- biggest costs
    section(r, 'BIGGEST COSTS THIS WEEK (fuel excluded: it follows the trips)')
    heads(r + 1, ['Truck or overhead', 'Date', 'Amount', 'Category', '', 'Description'])
    ws.merge_cells(start_row=r + 1, start_column=4, end_row=r + 1, end_column=5)
    ws.merge_cells(start_row=r + 1, start_column=6, end_row=r + 1, end_column=W)
    for c in (2, 4, 6):
        ws.cell(r + 1, c).alignment = Alignment(horizontal='left', indent=1 if c == 4 else 0)
    r += 2
    big = [c for c in costs if wk(c[0]) == 'W' and 'fuel' not in (c[2] or '').lower()]
    for c in sorted(big, key=lambda x: -x[3])[:5]:
        ws.cell(r, 1).value = c[1]
        ws.cell(r, 2).value = datetime(c[0].year, c[0].month, c[0].day)
        ws.cell(r, 2).number_format = 'ddd d mmm'
        ws.cell(r, 2).alignment = Alignment(horizontal='left')
        ws.cell(r, 3).value = c[3]
        ws.cell(r, 3).number_format = MONEY
        ws.cell(r, 4).value = c[2]
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        ws.cell(r, 4).alignment = Alignment(indent=1)
        ws.cell(r, 6).value = c[5]
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=W)
        r += 1
    r += 1

    # ---------- needs attention
    section(r, 'NEEDS ATTENTION')
    r += 1
    excs = summary.get('exceptions', {})
    grouped = {}
    for kind, items in excs.items():
        hit = next((x for x in ATTENTION if x[0].lower() in kind.lower()), None)
        key = hit[1] if hit else None
        grouped.setdefault(key, {'hit': hit, 'items': []})['items'] += items
    n_lines = 0
    for match, label, action in ATTENTION:
        g = grouped.get(label)
        if not g:
            continue
        items = g['items']
        dated = [re.search(r'\d{4}-\d{2}-\d{2}', i) for i in items]
        new = sum(1 for m_ in dated if m_ and w_start <= to_date(m_.group(0)) <= week_end)
        amts = [float(m_.group(1).replace(',', '')) for m_ in (re.search(r'\$(-?[\d,]+\.\d{2})', i) for i in items) if m_]
        if match == 'Completed trip with no revenue':
            ex = []
            for i in items:
                t = all_trips.get(i.split()[0])
                ex.append(f'{i.split()[0]} {t[3]}, {t[8]:,.0f} km' if t else i.split()[0])
            detail = '; '.join(ex[:4]) + ('; ...' if len(ex) > 4 else '')
        else:
            detail = ''
        text = (f'{len(items)} {label}' + (f' ({new} this week)' if new else '')
                + (f', {fm(sum(amts))}' if amts else '') + f': {action}.' + (f' {detail}' if detail else ''))
        ws.cell(r, 1).value = '•  ' + text
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=W)
        ws.row_dimensions[r].height = 30 if len(text) > 125 else 16
        r += 1
        n_lines += 1
    other = sum(len(v['items']) for k, v in grouped.items() if k is None)
    if other:
        note(r, f'{other} other notes (reversals, re-filed categories, vehicles outside the fleet list) are in the daily report.', 1, W)
        r += 1
    if not n_lines and not other:
        ws.cell(r, 1).value = 'Nothing open.'
        r += 1
    r += 1

    # ---------- month to date
    section(r, f'MONTH TO DATE AT {week_end:%d %b %Y}'.upper())
    mtd = cur['mtd']
    want = [('Revenue', 'Revenue MTD'), ('Net profit', 'Net profit MTD'), ('Margin', 'Net margin'),
            ('Collected', 'Collected MTD'), ('Cash & bank', 'Cash & bank balance')]
    heads(r + 1, [''] + [w_[0] for w_ in want])
    ws.cell(r + 2, 1).value = f'1 to {week_end:%d %b}'
    for c in range(1, len(want) + 1):
        v = mtd.get(want[c - 1][1])
        cell = ws.cell(r + 2, c + 1)
        cell.value = v
        cell.number_format = PCT if 'margin' in want[c - 1][1].lower() else MONEY
        cell.font = Font(name=FONT, bold=True, color='0000FF')
    r += 3
    note(r, (f'From the Truck Summary of the daily report for {cur["month_start"]:%B %Y}, frozen at {week_end:%d %b}. '
             'Revenue counts on the trip completion date; running costs and fixed overheads accrue daily. '
             'Not included: anything not recorded in Wilogix or Vuma, depreciation until it is entered on the Assumptions tab, and Zoho Books.'),
         1, W, height=40)

    # ---------- page set-up
    set_widths(ws, [34, 12, 12, 12, 12, 13, 12, 13])
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f'A1:H{r}'
    ws.page_margins.left = ws.page_margins.right = 0.4
    for rr in range(1, r + 1):
        for c in range(1, W + 1):
            cell = ws.cell(rr, c)
            if cell.value is not None and cell.font.name != FONT:
                f = cell.font
                cell.font = Font(name=FONT, bold=f.bold, italic=f.italic, size=f.size, color=f.color)
            if cell.value is not None and cell.font.name == FONT and cell.font.size is None:
                pass

    # ---------- checks (hidden)
    ck = sheets['Checks']
    ck['A1'], ck['B1'] = 'Check', 'Difference (must be 0)'
    style_header(ck, 1, 1, 2)
    checks = [
        ('Trucks contribution total equals P&L revenue less trip costs and repairs',
         f"=ROUND('Weekly Report'!F{truck_total}-('Weekly Report'!B{row['rev']}-'Weekly Report'!B{row['trip']}-'Weekly Report'!B{row['rep']}),2)"),
        ('Trucks last-week contribution equals last week P&L',
         f"=ROUND('Weekly Report'!G{truck_total}-('Weekly Report'!C{row['rev']}-'Weekly Report'!C{row['trip']}-'Weekly Report'!C{row['rep']}),2)"),
        ('Customers revenue equals P&L revenue (this week)', f"=ROUND('Weekly Report'!C{cust_total}-'Weekly Report'!B{row['rev']},2)"),
        ('Customers revenue equals P&L revenue (last week)', f"=ROUND('Weekly Report'!D{cust_total}-'Weekly Report'!C{row['rev']},2)"),
        ('Customers collected equals Activity collected (this week)',
         f"=ROUND('Weekly Report'!E{cust_total}-'Weekly Report'!B{row['Collected from customers']},2)"),
        ('Net profit for the week equals the sum of daily net profit in the daily report',
         f"=ROUND('Weekly Report'!B{row['net']}-SUMIFS({rng(D, 'H')},{win(D, chr(39) + 'Weekly Report' + chr(39) + '!$B$3', chr(39) + 'Weekly Report' + chr(39) + '!$C$3')}),2)"),
    ]
    for i, (lab, f) in enumerate(checks, 2):
        ck.cell(i, 1).value, ck.cell(i, 2).value = lab, f
    set_widths(ck, [90, 22])
    for n in data_names:
        font_all(sheets[n])
        sheets[n].sheet_state = 'hidden'
    wb.active = 0
    wb.save(a.out)

    out = {
        'week': f'{w_start} to {week_end}', 'previous_week': f'{p_start} to {p_end}', 'output': a.out,
        'monthly_workbooks': [m['path'] for m in months],
        'headline': headline,
        'rows': {'trips': len(trips), 'costs': len(costs), 'receipts': len(receipts), 'days': len(daily)},
        'checks_sheet': 'Checks (hidden): every value in column B must be 0 after recalculation',
    }
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
    b.add_argument('--out', required=True)
    a = ap.parse_args()
    cmd_freeze(a) if a.cmd == 'freeze' else cmd_build(a)


if __name__ == '__main__':
    main()
