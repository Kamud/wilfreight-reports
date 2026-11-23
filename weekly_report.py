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

Recalculate the weekly workbook afterwards; it holds formulas over its data tabs.
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
def cmd_build(a):
    week_end = to_date(a.week_end)
    if week_end.weekday() != 6:
        raise SystemExit('--week-end must be a Sunday')
    w_start = week_end - timedelta(days=6)
    p_start, p_end = w_start - timedelta(days=7), w_start - timedelta(days=1)

    months = [read_monthly(p) for p in a.monthly]
    months.sort(key=lambda m: m['month_start'])
    cur = next((m for m in months if m['month_start'] <= week_end <= m['month_start'].replace(day=m['days_in_month'])), None)
    if cur is None:
        raise SystemExit('No monthly workbook covers the week-end date')
    if cur['as_at'] != week_end:
        raise SystemExit(f'{cur["path"]} is as at {cur["as_at"]}, not {week_end}: run the freeze step before recalculating')
    for d in (p_start, w_start, week_end):
        if not any(m['month_start'] <= d <= m['month_start'].replace(day=m['days_in_month']) for m in months):
            raise SystemExit(f'No monthly workbook covers {d}: pass the other month with a second --monthly')
    for m in months:
        if m is not cur and m['as_at'] < m['month_start'].replace(day=m['days_in_month']):
            raise SystemExit(f'{m["path"]} stops at {m["as_at"]}; the earlier month must run to its last day')

    inside = lambda d: d is not None and p_start <= d <= week_end
    trips = sorted([t for m in months for t in m['trips'] if inside(t[0])], key=lambda x: (x[0], x[1] or ''))
    costs = sorted([c for m in months for c in m['costs'] if inside(c[0])], key=lambda x: x[0])
    receipts = sorted([r for m in months for r in m['receipts'] if inside(r[0])], key=lambda x: x[0])
    daily = sorted([d for m in months for d in m['daily'] if inside(d[0])], key=lambda x: x[0])
    truck_days = []
    for m in months:
        for i in range(m['days_in_month']):
            d = m['month_start'] + timedelta(days=i)
            if inside(d) and d <= m['as_at']:
                for name, status, monthly in m['trucks']:
                    truck_days.append([d, name, round(monthly / m['days_in_month'], 6)])
    trucks = cur['trucks']
    summary = json.load(open(a.summary, encoding='utf-8')) if a.summary else {}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Week Summary'
    sheets = {n: wb.create_sheet(n) for n in ['By Truck', 'By Customer', 'By Day', 'Costs', 'Exceptions',
                                              'Trips data', 'Costs data', 'Receipts data', 'Daily data', 'Truck days']}

    # ---- data tabs (values copied from the daily-report workbooks)
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
    write_table(sheets['Truck days'], ['Date', 'Truck', 'Running cost accrued ($)'], truck_days, [DATE, None, MONEY])
    for n, w in (('Trips data', [16, 11, 11, 24, 34, 11, 12, 10, 8]), ('Costs data', [16, 14, 26, 11, 26, 50, 12, 12, 10, 18]),
                 ('Receipts data', [16, 26, 12, 36, 26]), ('Daily data', [16] + [13] * 9), ('Truck days', [16, 22, 14])):
        set_widths(sheets[n], w)

    T, C, R, D, TD = "'Trips data'", "'Costs data'", "'Receipts data'", "'Daily data'", "'Truck days'"
    rng = lambda sh, col: f'{sh}!${col}$2:${col}$5000'

    def win(sh, s, e):
        return f'{rng(sh, "A")},">="&{s},{rng(sh, "A")},"<="&{e}'

    # ---- Week Summary
    title(ws, 'Saraquel Logistics: weekly truck report',
          f'Week {w_start:%a %d %b} to {week_end:%a %d %b %Y}, compared with {p_start:%d %b} to {p_end:%d %b}. USD. '
          'Built from the daily report workbooks; same rules.')
    ws['A4'], ws['A5'] = 'This week', 'Last week'
    for r, (s, e) in ((4, (w_start, week_end)), (5, (p_start, p_end))):
        ws.cell(r, 2).value = datetime(s.year, s.month, s.day)
        ws.cell(r, 3).value = datetime(e.year, e.month, e.day)
        for c in (2, 3):
            ws.cell(r, c).number_format = DATE
            ws.cell(r, c).font = Font(name=FONT, color='0000FF')
    ws['D4'] = 'From (Mon) and to (Sun). Change these two dates to re-cut the tabs; data tabs hold the fortnight only.'
    ws['D4'].font = Font(name=FONT, italic=True, color='595959', size=9)
    TW, LW = ('$B$4', '$C$4'), ('$B$5', '$C$5')
    S = "'Week Summary'!"

    hdr = ['Measure', 'This week', 'Last week', 'Change', 'Change %', 'How it is worked out']
    for c, h in enumerate(hdr, 1):
        ws.cell(7, c).value = h
    style_header(ws, 7, 1, 6)

    def trips_f(col, s, e):
        return f'=SUMIFS({rng(T, col)},{win(T, s, e)})'

    def cost_f(s, e, extra=''):
        return f'SUMIFS({rng(C, "D")},{win(C, s, e)}{extra})'

    TRUCK = ',' + rng(C, 'I') + ',"Truck"'
    REPAIRS = ',' + rng(C, 'C') + ',"Repairs"'
    FUEL = ',' + rng(C, 'C') + ',"Fuel"'
    PAID = ',' + rng(C, 'E') + ',"<>"'
    CREDIT = ',' + rng(C, 'J') + ',"<>"'

    lines = [
        # label, formula maker(s, e) or None, fmt, note, kind
        ('ACTIVITY', None, None, None, 'section'),
        ('Trips completed', lambda s, e: f'=COUNTIFS({win(T, s, e)})', NUM, 'Trips with a completion date in the week', ''),
        ('Km logged', lambda s, e: trips_f('I', s, e), NUM, 'Wilogix mileage logs on those trips', ''),
        ('Trucks that completed a trip', None, NUM, 'Count of trucks with at least one trip (By Truck tab)', 'trucks'),
        ('PROFIT', None, None, None, 'section'),
        ('Revenue', lambda s, e: trips_f('G', s, e), MONEY, 'On trip completion date: invoice, else trip revenue, else 0.9 x km for Bakers Inn', ''),
        ('Trip costs (fuel, allowances, tolls, fines, compliance)', lambda s, e: f'={cost_f(s, e, TRUCK)}-{cost_f(s, e, TRUCK + REPAIRS)}', MONEY, 'Truck lines on the Expense Log except repairs, incl. client fuel deducted by credit note', ''),
        ('Repairs', lambda s, e: f'={cost_f(s, e, TRUCK + REPAIRS)}', MONEY, 'Truck repair lines', ''),
        ('Contribution after repairs', 'contrib', MONEY, 'Revenue less trip costs and repairs', 'calc'),
        ('Truck running costs (accrued)', lambda s, e: f'=SUMIFS({rng(D, "F")},{win(D, s, e)})', MONEY, 'Licence, insurance, Cartrack, depreciation accrued per day from the Assumptions tab', ''),
        ('Company overheads', lambda s, e: f'=SUMIFS({rng(D, "G")},{win(D, s, e)})', MONEY, 'Fixed overheads accrued per day plus overhead lines logged', ''),
        ('Net profit', 'net', MONEY, 'Contribution less running costs and overheads', 'calc'),
        ('Net margin', 'margin', PCT, 'Net profit / revenue', 'calc'),
        ('KEY RATIOS', None, None, None, 'section'),
        ('Fuel (all truck fuel)', lambda s, e: f'={cost_f(s, e, TRUCK + FUEL)}', MONEY, 'Cash fuel plus client fuel deducted by credit note', ''),
        ('Fuel as % of revenue', 'fuelpct', PCT, '', 'calc'),
        ('Revenue per km', 'rpk', MONEY, '', 'calc'),
        ('Average revenue per trip', 'rpt', MONEY, '', 'calc'),
        ('CASH', None, None, None, 'section'),
        ('Collected from customers', lambda s, e: f'=SUMIFS({rng(R, "C")},{win(R, s, e)})', MONEY, 'Vuma collections', ''),
        ('Cash spent (paid from an account)', lambda s, e: f'={cost_f(s, e, PAID)}', MONEY, 'Vuma expenses, including prepaid insurance', ''),
        ('Client fuel deducted by credit note', lambda s, e: f'={cost_f(s, e, CREDIT)}', MONEY, 'Reduces what the customer owes; no cash moves', ''),
        ('Cash and bank balance at end of week', 'cash', MONEY, 'This week: Cash & Banks total on Sunday. Last week: that, less this week\'s collections, plus this week\'s cash spent', 'bal'),
        ('Owed by customers at end of week', 'owed', MONEY, 'This week: Customers total on Sunday. Last week: that, less this week\'s revenue, plus collections and credit notes', 'bal'),
    ]
    rowof = {}
    r = 8
    for label, fm, fmt, note, kind in lines:
        ws.cell(r, 1).value = label
        if kind == 'section':
            for c in range(1, 7):
                ws.cell(r, c).fill = SUB_FILL
                ws.cell(r, c).font = Font(name=FONT, bold=True, color='1F3864')
            r += 1
            continue
        rowof[label] = r
        if note:
            ws.cell(r, 6).value = note
            ws.cell(r, 6).font = Font(name=FONT, size=9, color='595959')
        r += 1
    g = lambda label, col: f'{col}{rowof[label]}'
    for label, fm, fmt, note, kind in lines:
        if kind == 'section':
            continue
        rr = rowof[label]
        for col, (s, e) in (('B', TW), ('C', LW)):
            if kind == '':
                f = fm(s, e)
            elif kind == 'trucks':
                f = f"=COUNTIF('By Truck'!{'C' if col == 'B' else 'D'}$6:{'C' if col == 'B' else 'D'}$40,\">0\")"
            elif fm == 'contrib':
                f = f'={g("Revenue", col)}-{g("Trip costs (fuel, allowances, tolls, fines, compliance)", col)}-{g("Repairs", col)}'
            elif fm == 'net':
                f = f'={g("Contribution after repairs", col)}-{g("Truck running costs (accrued)", col)}-{g("Company overheads", col)}'
            elif fm == 'margin':
                f = f'=IFERROR({g("Net profit", col)}/{g("Revenue", col)},"")'
            elif fm == 'fuelpct':
                f = f'=IFERROR({g("Fuel (all truck fuel)", col)}/{g("Revenue", col)},"")'
            elif fm == 'rpk':
                f = f'=IFERROR({g("Revenue", col)}/{g("Km logged", col)},"")'
            elif fm == 'rpt':
                f = f'=IFERROR({g("Revenue", col)}/{g("Trips completed", col)},"")'
            elif fm == 'cash':
                f = '=$H$4' if col == 'B' else f'=B{rr}-{g("Collected from customers", "B")}+{g("Cash spent (paid from an account)", "B")}'
            elif fm == 'owed':
                f = '=$H$5' if col == 'B' else (f'=B{rr}-{g("Revenue", "B")}+{g("Collected from customers", "B")}'
                                                f'+{g("Client fuel deducted by credit note", "B")}')
            ws[f'{col}{rr}'] = f
            ws[f'{col}{rr}'].number_format = fmt
        ws[f'D{rr}'] = f'=IF(OR(B{rr}="",C{rr}=""),"",B{rr}-C{rr})'
        ws[f'D{rr}'].number_format = fmt
        if fmt == PCT:
            ws[f'D{rr}'].number_format = '+0.0" pts";-0.0" pts";-'
            ws[f'D{rr}'] = f'=IF(OR(B{rr}="",C{rr}=""),"",(B{rr}-C{rr})*100)'
            ws[f'E{rr}'] = ''
        else:
            ws[f'E{rr}'] = f'=IF(OR(B{rr}="",C{rr}="",N(C{rr})=0),"",(B{rr}-C{rr})/ABS(C{rr}))'
            ws[f'E{rr}'].number_format = '+0.0%;-0.0%;-'
    # Sunday balances (values from the daily workbook, documented)
    ws['G4'], ws['H4'] = 'Cash & Banks total on Sunday', cur['cash_total']
    ws['G5'], ws['H5'] = 'Customers owed total on Sunday', cur['owed_total']
    for c in ('H4', 'H5'):
        ws[c].number_format = MONEY
        ws[c].font = Font(name=FONT, color='0000FF')
    ws['G6'] = f'Source: daily report workbook as at {week_end:%d %b %Y} (Cash & Banks B, Customers B6).'
    ws['G6'].font = Font(name=FONT, size=9, italic=True, color='595959')
    # colour the change on cost lines the other way round
    good_up = ['Trips completed', 'Km logged', 'Trucks that completed a trip', 'Revenue', 'Contribution after repairs', 'Net profit', 'Net margin',
               'Revenue per km', 'Average revenue per trip', 'Collected from customers', 'Cash and bank balance at end of week']
    good_down = ['Trip costs (fuel, allowances, tolls, fines, compliance)', 'Repairs', 'Company overheads', 'Fuel as % of revenue', 'Owed by customers at end of week']
    green, red = Font(name=FONT, color='008000'), Font(name=FONT, color='C00000')
    for label in good_up + good_down:
        ref = f'D{rowof[label]}:E{rowof[label]}'
        up, dn = (green, red) if label in good_up else (red, green)
        ws.conditional_formatting.add(ref, CellIsRule(operator='greaterThan', formula=['0'], font=up))
        ws.conditional_formatting.add(ref, CellIsRule(operator='lessThan', formula=['0'], font=dn))
    for label in ('Net profit', 'Revenue'):
        for c in range(1, 6):
            ws.cell(rowof[label], c).font = Font(name=FONT, bold=True)
            ws.cell(rowof[label], c).fill = TOTAL_FILL

    # month-to-date block
    r = max(rowof.values()) + 2
    ws.cell(r, 1).value = f'MONTH TO DATE AT {week_end:%d %b %Y}'.upper()
    for c in range(1, 7):
        ws.cell(r, c).fill = SUB_FILL
        ws.cell(r, c).font = Font(name=FONT, bold=True, color='1F3864')
    r += 1
    for k, v in cur['mtd'].items():
        ws.cell(r, 1).value = k
        ws.cell(r, 2).value = v
        ws.cell(r, 2).number_format = PCT if 'margin' in (k or '').lower() else MONEY
        ws.cell(r, 2).font = Font(name=FONT, color='0000FF')
        r += 1
    ws.cell(r, 1).value = f'Source: Truck Summary row 5 of the daily report workbook for {cur["month_start"]:%B %Y}, frozen at {week_end:%d %b %Y}.'
    ws.cell(r, 1).font = Font(name=FONT, size=9, italic=True, color='595959')
    r += 2
    ws.cell(r, 1).value = ('Notes: revenue counts on the trip completion date. Running costs and fixed overheads accrue daily, so a quiet Sunday '
                           'still shows a small loss. The report sees only Wilogix and Vuma: anything not recorded there, depreciation not yet '
                           'on the Assumptions tab, and Zoho Books are not included.')
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 45
    set_widths(ws, [52, 16, 16, 14, 11, 70, 30, 14])
    ws.freeze_panes = 'B8'

    # ---- By Truck
    bt = sheets['By Truck']
    title(bt, 'By truck: this week against last week', 'Running costs are the daily accrual from the Assumptions tab. Company overheads are not split by truck.')
    groups = [('Trips', 'trips'), ('Revenue ($)', 'rev'), ('Trip costs ($)', 'cost'), ('Repairs ($)', 'rep'),
              ('Running costs ($)', 'run'), ('Truck net profit ($)', 'net')]
    bt['A5'], bt['B5'] = 'Truck', 'Status'
    col = 3
    gcols = {}
    for name, key in groups:
        bt.cell(4, col).value = name
        bt.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        bt.cell(5, col).value, bt.cell(5, col + 1).value = 'This week', 'Last week'
        gcols[key] = (get_column_letter(col), get_column_letter(col + 1))
        col += 2
    bt.cell(5, col).value = 'Change in net profit ($)'
    style_header(bt, 4, 1, col)
    style_header(bt, 5, 1, col)
    first = 6
    for i, (name, status, _) in enumerate(trucks):
        rr = first + i
        bt.cell(rr, 1).value, bt.cell(rr, 2).value = name, status
        for key, (ca, cb) in gcols.items():
            for c, (s, e) in ((ca, (S + TW[0], S + TW[1])), (cb, (S + LW[0], S + LW[1]))):
                if key == 'trips':
                    f = f'=COUNTIFS({rng(T, "C")},$A{rr},{win(T, s, e)})'
                elif key == 'rev':
                    f = f'=SUMIFS({rng(T, "G")},{rng(T, "C")},$A{rr},{win(T, s, e)})'
                elif key == 'cost':
                    f = (f'=SUMIFS({rng(C, "D")},{rng(C, "B")},$A{rr},{rng(C, "I")},"Truck",{win(C, s, e)})'
                         f'-SUMIFS({rng(C, "D")},{rng(C, "B")},$A{rr},{rng(C, "C")},"Repairs",{rng(C, "I")},"Truck",{win(C, s, e)})')
                elif key == 'rep':
                    f = f'=SUMIFS({rng(C, "D")},{rng(C, "B")},$A{rr},{rng(C, "C")},"Repairs",{rng(C, "I")},"Truck",{win(C, s, e)})'
                elif key == 'run':
                    f = f'=SUMIFS({rng(TD, "C")},{rng(TD, "B")},$A{rr},{win(TD, s, e)})'
                else:
                    x = 0 if c == ca else 1
                    rc, cc, pc, uc = (gcols[k][x] for k in ('rev', 'cost', 'rep', 'run'))
                    f = f'={rc}{rr}-{cc}{rr}-{pc}{rr}-{uc}{rr}'
                bt[f'{c}{rr}'] = f
                bt[f'{c}{rr}'].number_format = NUM if key == 'trips' else MONEY
        na, nb = gcols['net']
        bt.cell(rr, col).value = f'={na}{rr}-{nb}{rr}'
        bt.cell(rr, col).number_format = MONEY
    last = first + len(trucks) - 1
    tr = last + 1
    bt.cell(tr, 1).value = 'FLEET'
    for c in range(3, col + 1):
        L = get_column_letter(c)
        bt.cell(tr, c).value = f'=SUM({L}{first}:{L}{last})'
        bt.cell(tr, c).number_format = NUM if c in (3, 4) else MONEY
    for c in range(1, col + 1):
        bt.cell(tr, c).font = Font(name=FONT, bold=True)
        bt.cell(tr, c).fill = TOTAL_FILL
    bt.cell(tr + 1, 1).value = 'Check: fleet net profit before overheads equals Week Summary (should be 0)'
    bt.cell(tr + 1, col).value = (f'=ROUND({gcols["net"][0]}{tr}-({S}B{rowof["Net profit"]}+{S}B{rowof["Company overheads"]}),2)')
    bt.conditional_formatting.add(f'{get_column_letter(col)}{first}:{get_column_letter(col)}{last}',
                                  CellIsRule(operator='lessThan', formula=['0'], font=Font(name=FONT, color='C00000')))
    set_widths(bt, [22, 11] + [11] * 12 + [14])
    bt.freeze_panes = 'C6'

    # ---- By Customer
    bc = sheets['By Customer']
    title(bc, 'By customer: this week against last week', 'Sorted by revenue this week.')
    custs = {}
    for t in trips:
        custs.setdefault(t[3] or '(none)', [0, 0])
        custs[t[3] or '(none)'][0 if t[0] >= w_start else 1] += t[6]
    for rc_ in receipts:
        custs.setdefault(rc_[1] or '(none)', [0, 0])
    for c_ in costs:
        if c_[9]:
            custs.setdefault(c_[9], [0, 0])
    order = sorted(custs, key=lambda k: (-custs[k][0], -custs[k][1], k))
    cgroups = [('Trips', 'trips'), ('Revenue ($)', 'rev'), ('Collected ($)', 'col'), ('Fuel deducted by credit note ($)', 'cn')]
    bc['A5'] = 'Customer'
    col = 2
    ccols = {}
    for name, key in cgroups:
        bc.cell(4, col).value = name
        bc.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        bc.cell(5, col).value, bc.cell(5, col + 1).value = 'This week', 'Last week'
        ccols[key] = (get_column_letter(col), get_column_letter(col + 1))
        col += 2
    bc.cell(5, col).value = 'Change in revenue ($)'
    style_header(bc, 4, 1, col)
    style_header(bc, 5, 1, col)
    for i, name in enumerate(order):
        rr = 6 + i
        bc.cell(rr, 1).value = name
        for key, (ca, cb) in ccols.items():
            for c, (s, e) in ((ca, (S + TW[0], S + TW[1])), (cb, (S + LW[0], S + LW[1]))):
                if key == 'trips':
                    f = f'=COUNTIFS({rng(T, "D")},$A{rr},{win(T, s, e)})'
                elif key == 'rev':
                    f = f'=SUMIFS({rng(T, "G")},{rng(T, "D")},$A{rr},{win(T, s, e)})'
                elif key == 'col':
                    f = f'=SUMIFS({rng(R, "C")},{rng(R, "B")},$A{rr},{win(R, s, e)})'
                else:
                    f = f'=SUMIFS({rng(C, "D")},{rng(C, "J")},$A{rr},{win(C, s, e)})'
                bc[f'{c}{rr}'] = f
                bc[f'{c}{rr}'].number_format = NUM if key == 'trips' else MONEY
        bc.cell(rr, col).value = f'={ccols["rev"][0]}{rr}-{ccols["rev"][1]}{rr}'
        bc.cell(rr, col).number_format = MONEY
    tr = 6 + len(order)
    bc.cell(tr, 1).value = 'TOTAL'
    for c in range(2, col + 1):
        L = get_column_letter(c)
        bc.cell(tr, c).value = f'=SUM({L}6:{L}{tr - 1})'
        bc.cell(tr, c).number_format = NUM if c in (2, 3) else MONEY
    for c in range(1, col + 1):
        bc.cell(tr, c).font = Font(name=FONT, bold=True)
        bc.cell(tr, c).fill = TOTAL_FILL
    set_widths(bc, [30] + [12] * 8 + [14])
    bc.freeze_panes = 'B6'

    # ---- By Day
    bd = sheets['By Day']
    title(bd, 'Day by day: this week beside the same day last week', 'Net profit per day includes the daily accrual of running costs and fixed overheads.')
    heads = ['Day', 'Date', 'Revenue ($)', 'Trip costs and repairs ($)', 'Net profit ($)', 'Collected ($)',
             'Date', 'Revenue ($)', 'Trip costs and repairs ($)', 'Net profit ($)', 'Collected ($)', 'Revenue change ($)']
    bd.cell(4, 2).value, bd.cell(4, 7).value = 'THIS WEEK', 'LAST WEEK'
    bd.merge_cells('B4:F4')
    bd.merge_cells('G4:K4')
    style_header(bd, 4, 1, 12)
    for c, h in enumerate(heads, 1):
        bd.cell(5, c).value = h
    style_header(bd, 5, 1, 12)
    for i in range(7):
        rr = 6 + i
        bd.cell(rr, 1).value = (w_start + timedelta(days=i)).strftime('%A')
        for base, start in ((2, f'{S}$B$4'), (7, f'{S}$B$5')):
            dc = get_column_letter(base)
            bd.cell(rr, base).value = f'={start}+{i}'
            bd.cell(rr, base).number_format = 'd mmm'
            vals = [('B',), ('C', 'D'), ('H',), ('J',)]
            for k, cols_ in enumerate(vals):
                parts = '+'.join(f'SUMIFS({rng(D, x)},{rng(D, "A")},{dc}{rr})' for x in cols_)
                cell = bd.cell(rr, base + 1 + k)
                cell.value = f'={parts}'
                cell.number_format = MONEY
        bd.cell(rr, 12).value = f'=C{rr}-H{rr}'
        bd.cell(rr, 12).number_format = MONEY
    bd.cell(13, 1).value = 'WEEK'
    for c in (3, 4, 5, 6, 8, 9, 10, 11, 12):
        L = get_column_letter(c)
        bd.cell(13, c).value = f'=SUM({L}6:{L}12)'
        bd.cell(13, c).number_format = MONEY
    for c in range(1, 13):
        bd.cell(13, c).font = Font(name=FONT, bold=True)
        bd.cell(13, c).fill = TOTAL_FILL
    set_widths(bd, [12, 9, 12, 14, 12, 12, 9, 12, 14, 12, 12, 13])

    # ---- Costs
    co = sheets['Costs']
    title(co, 'Where the money went', 'Truck costs by category, then company overheads. Amounts counted in profit.')
    for c, h in enumerate(['Category', 'This week ($)', 'Last week ($)', 'Change ($)', 'Change %'], 1):
        co.cell(4, c).value = h
    style_header(co, 4, 1, 5)
    r = 5
    truck_cats = ['Fuel', 'Allowances', 'Remuneration', 'Law Enforcements', 'Compliancy', 'Tolls and Parking', 'Commission', 'Repairs']
    seen = [c_[2] for c_ in costs if c_[8] == 'Truck']
    truck_cats += sorted({x for x in seen if x not in truck_cats})

    def section(label):
        nonlocal r
        co.cell(r, 1).value = label
        for c in range(1, 6):
            co.cell(r, c).fill = SUB_FILL
            co.cell(r, c).font = Font(name=FONT, bold=True, color='1F3864')
        r += 1

    def line(label, fmaker):
        nonlocal r
        co.cell(r, 1).value = label
        for c, (s, e) in (('B', (S + TW[0], S + TW[1])), ('C', (S + LW[0], S + LW[1]))):
            co[f'{c}{r}'] = fmaker(s, e, c)
            co[f'{c}{r}'].number_format = MONEY
        co[f'D{r}'] = f'=B{r}-C{r}'
        co[f'D{r}'].number_format = MONEY
        co[f'E{r}'] = f'=IF(C{r}=0,"",(B{r}-C{r})/ABS(C{r}))'
        co[f'E{r}'].number_format = '+0.0%;-0.0%;-'
        r += 1

    section('TRUCK COSTS')
    t0 = r
    for cat in truck_cats:
        line(cat, lambda s, e, c, cat=cat: f'=SUMIFS({rng(C, "D")},{rng(C, "C")},"{cat}",{rng(C, "I")},"Truck",{win(C, s, e)})')
    t1 = r - 1
    line('Total truck costs', lambda s, e, c: f'=SUM({c}{t0}:{c}{t1})')
    for c in range(1, 6):
        co.cell(r - 1, c).font = Font(name=FONT, bold=True)
    section('COMPANY OVERHEADS')
    o0 = r
    line('Fixed overheads accrued (Assumptions tab)',
         lambda s, e, c: f'=SUMIFS({rng(D, "G")},{win(D, s, e)})-SUMIFS({rng(C, "H")},{rng(C, "I")},"Overhead",{win(C, s, e)})')
    for cat in sorted({c_[2] for c_ in costs if c_[8] == 'Overhead'}):
        line(cat, lambda s, e, c, cat=cat: f'=SUMIFS({rng(C, "H")},{rng(C, "C")},"{cat}",{rng(C, "I")},"Overhead",{win(C, s, e)})')
    o1 = r - 1
    line('Total company overheads', lambda s, e, c: f'=SUM({c}{o0}:{c}{o1})')
    for c in range(1, 6):
        co.cell(r - 1, c).font = Font(name=FONT, bold=True)
    co.cell(r, 1).value = 'Check: totals agree with Week Summary (should be 0)'
    co.cell(r, 2).value = (f'=ROUND(B{t1 + 1}+B{r - 1}-({S}B{rowof["Trip costs (fuel, allowances, tolls, fines, compliance)"]}'
                           f'+{S}B{rowof["Repairs"]}+{S}B{rowof["Company overheads"]}),2)')
    r += 2
    co.cell(r, 1).value = 'LARGEST COSTS THIS WEEK'
    for c in range(1, 6):
        co.cell(r, c).fill = SUB_FILL
        co.cell(r, c).font = Font(name=FONT, bold=True, color='1F3864')
    r += 1
    for c, h in enumerate(['Date', 'Applies to', 'Category', 'Amount ($)', 'Description'], 1):
        co.cell(r, c).value = h
    style_header(co, r, 1, 5)
    r += 1
    top = sorted([c_ for c_ in costs if c_[0] >= w_start], key=lambda x: -x[3])[:10]
    for c_ in top:
        co.cell(r, 1).value = datetime(c_[0].year, c_[0].month, c_[0].day)
        co.cell(r, 1).number_format = 'd mmm'
        co.cell(r, 2).value, co.cell(r, 3).value, co.cell(r, 4).value, co.cell(r, 5).value = c_[1], c_[2], c_[3], c_[5]
        co.cell(r, 4).number_format = MONEY
        r += 1
    set_widths(co, [40, 14, 14, 13, 60])

    # ---- Exceptions
    ex = sheets['Exceptions']
    title(ex, f'Open exceptions at {week_end:%d %b %Y}',
          'From the daily fill for the month. "This week" marks items dated in the week; the rest were already open.')
    for c, h in enumerate(['Type', 'When', 'Item', 'Amount ($)'], 1):
        ex.cell(4, c).value = h
    style_header(ex, 4, 1, 4)
    excs = summary.get('exceptions', {})

    def rank(kind):
        for i, k in enumerate(EXCEPTION_ORDER):
            if k.lower() in kind.lower():
                return i
        return len(EXCEPTION_ORDER)

    r = 5
    counts = []
    for kind in sorted(excs, key=rank):
        n_week = 0
        for item in excs[kind]:
            m = re.search(r'\d{4}-\d{2}-\d{2}', item)
            d = to_date(m.group(0)) if m else None
            when = 'This week' if d and w_start <= d <= week_end else ('Earlier' if d else '')
            n_week += when == 'This week'
            amt = re.search(r'\$(-?[\d,]+\.\d{2})', item)
            ex.cell(r, 1).value, ex.cell(r, 2).value, ex.cell(r, 3).value = kind, when, item
            if amt:
                ex.cell(r, 4).value = float(amt.group(1).replace(',', ''))
                ex.cell(r, 4).number_format = MONEY
            if when == 'This week':
                ex.cell(r, 2).font = Font(name=FONT, bold=True, color='C00000')
            r += 1
        counts.append((kind, len(excs[kind]), n_week))
    if not excs:
        ex.cell(5, 1).value = 'No exceptions (or no summary file was given).'
    set_widths(ex, [60, 11, 80, 12])
    for rr in range(5, r):
        for c in (1, 3):
            ex.cell(rr, c).alignment = Alignment(wrap_text=True, vertical='top')

    for sh in wb.worksheets:
        font_all(sh)
        sh.page_setup.orientation = 'landscape'
        sh.page_setup.fitToWidth, sh.page_setup.fitToHeight = 1, 0
        sh.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(a.out)

    out = {
        'week': f'{w_start} to {week_end}', 'previous_week': f'{p_start} to {p_end}', 'output': a.out,
        'monthly_workbooks': [m['path'] for m in months],
        'rows': {'trips': len(trips), 'costs': len(costs), 'receipts': len(receipts), 'days': len(daily), 'truck_days': len(truck_days)},
        'exception_counts': [{'type': k, 'open': n, 'this_week': w} for k, n, w in counts],
        'week_summary_rows': rowof,
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
