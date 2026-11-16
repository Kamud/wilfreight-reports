#!/usr/bin/env python3
"""Fill the daily truck report template from Wilogix and Vuma extracts.

Usage:
  python3 fill_report.py --data DIR --month-start 2026-11-01 --out report.xlsx
         [--fuel-rules fuel_rules.json] [--template template/daily_truck_report_template.xlsx]

DIR holds four pipe-delimited text files, one record per line (see README.md for the SQL):
  trips.txt     completed_on|trip_number|truck|client|route|invoice_no|invoiced_amount|trip_revenue|estimated_revenue|km|invoice_date|zoho_status
  fuel.txt      log_date|truck|client|trip_number|litres|total_cost|paid_through
  expenses.txt  expense_date|truck|category|amount|wallet|memo|request_no|ratification|check_result
  ledger.txt    entry_date|type|from_wallet|to_wallet|amount|counterparty|memo

Only the input cells of the four log tabs are written. Every summary tab keeps its own formulas.
Prints a JSON summary with totals and exceptions. Recalculate the workbook afterwards.
"""
import argparse, json, os, sys
from collections import OrderedDict
from datetime import date, datetime

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
FIRST_ROW, LAST_ROW = 6, 505
INPUT_COLS = {
    'Trips': [1, 2, 3, 4, 5, 6, 7, 9, 10, 11],
    'Expense Log': [1, 2, 3, 4, 5, 6, 7, 8, 12],
    'Receipts Log': [1, 2, 3, 4, 5, 6],
    'Transfers Log': [1, 2, 3, 4, 5, 6],
}


def read_rows(path, n_fields):
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, encoding='utf-8') as fh:
        for line_no, line in enumerate(fh, 1):
            line = line.rstrip('\n').rstrip('\r')
            if not line.strip():
                continue
            parts = line.split('|')
            if len(parts) < n_fields:
                parts += [''] * (n_fields - len(parts))
            if len(parts) > n_fields:
                raise SystemExit(f'{path} line {line_no}: expected {n_fields} fields, found {len(parts)}')
            rows.append([p.strip() for p in parts])
    return rows


def to_date(s):
    return datetime.strptime(s[:10], '%Y-%m-%d').date()


def to_num(s):
    return float(s) if s not in ('', None) else None


def column_values(ws, col, r1, r2):
    return [ws.cell(r, col).value for r in range(r1, r2 + 1)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--month-start', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--fuel-rules')
    ap.add_argument('--template', default=os.path.join(HERE, 'template', 'daily_truck_report_template.xlsx'))
    ap.add_argument('--config', default=os.path.join(HERE, 'config.json'))
    a = ap.parse_args()

    cfg = json.load(open(a.config, encoding='utf-8'))
    month_start = to_date(a.month_start)
    exc = OrderedDict()  # exception type -> list of strings

    def flag(kind, text):
        exc.setdefault(kind, []).append(text)

    fuel_rules = {k: v for k, v in cfg['fuel_treatment_fallback'].items() if not k.startswith('_')}
    fuel_rules_source = 'config.json fallback'
    if a.fuel_rules and os.path.exists(a.fuel_rules):
        loaded = json.load(open(a.fuel_rules, encoding='utf-8'))
        if loaded:
            fuel_rules, fuel_rules_source = loaded, 'Client rules sheet'

    wb = openpyxl.load_workbook(a.template)
    asm, trips_ws, exp_ws = wb['Assumptions'], wb['Trips'], wb['Expense Log']
    rec_ws, tr_ws, cust_ws, cash_ws = wb['Receipts Log'], wb['Transfers Log'], wb['Customers'], wb['Cash & Banks']

    # Lists the workbook itself validates against
    trucks = [v for v in column_values(asm, 1, 12, 23) if v]
    truck_cats = [v for v in column_values(asm, 11, 12, 19) if v]
    overhead_cats = [v for v in column_values(asm, 13, 12, 27) if v]

    def customers():
        return [v for v in column_values(cust_ws, 1, 7, 46) if v]

    def accounts():
        return [v for v in column_values(cash_ws, 1, 7, 22) if v]

    def add_to_list(ws, r1, r2, name):
        for r in range(r1, r2 + 1):
            if ws.cell(r, 1).value in (None, ''):
                ws.cell(r, 1).value = name
                return True
        return False

    def map_customer(name):
        if not name:
            return ''
        mapped = cfg['customer_map'].get(name, name)
        if mapped not in customers():
            if add_to_list(cust_ws, 7, 46, mapped):
                flag('New customer added to the Customers tab (enter its opening balance)', mapped)
            else:
                flag('Customer list is full', mapped)
        return mapped

    def map_wallet(name):
        if not name:
            return ''
        mapped = cfg['wallet_map'].get(name, name)
        if mapped not in accounts():
            if add_to_list(cash_ws, 7, 22, mapped):
                flag('New account added to the Cash & Banks tab (enter its opening balance)', mapped)
            else:
                flag('Account list is full', mapped)
        return mapped

    def map_truck(reg):
        return reg.replace('-', ' ').strip() if reg else ''

    # Empty the input cells (the template should already be empty; this makes reruns safe)
    for name, cols in INPUT_COLS.items():
        ws = wb[name]
        for r in range(FIRST_ROW, LAST_ROW + 1):
            for c in cols:
                if ws.cell(r, c).value is not None:
                    ws.cell(r, c).value = None

    def capacity_check(name, n):
        if n > LAST_ROW - FIRST_ROW + 1:
            raise SystemExit(f'{name}: {n} lines do not fit in the {LAST_ROW - FIRST_ROW + 1} rows of the log')

    # ---------------- Trips ----------------
    trip_rows = read_rows(os.path.join(a.data, 'trips.txt'), 12)
    capacity_check('Trips', len(trip_rows))
    rate = cfg['bakers_inn_rate_per_km']
    rate_customers = set(cfg['rate_per_km_customers'])
    revenue_basis = {'invoiced': 0.0, 'trip revenue': 0.0, 'estimate': 0.0, 'rate per km': 0.0}
    trips_no_value = 0
    r = FIRST_ROW
    for d, number, truck, client, route, inv_no, inv_amt, rev, est, km, inv_date, zoho in trip_rows:
        truck_name = map_truck(truck)
        if truck_name not in trucks:
            flag('Trip on a truck that is not on the Assumptions tab (its revenue will not show on Truck Summary)', f'{number} {truck_name}')
        cust = map_customer(client)
        inv_amt, rev, est, km = to_num(inv_amt), to_num(rev), to_num(est), to_num(km)
        trips_ws.cell(r, 1).value = to_date(d)
        trips_ws.cell(r, 2).value = number
        trips_ws.cell(r, 3).value = truck_name
        trips_ws.cell(r, 4).value = cust
        trips_ws.cell(r, 5).value = ' '.join(route.split())
        if km is not None:
            trips_ws.cell(r, 9).value = km
        if inv_no:
            trips_ws.cell(r, 6).value = inv_no
            trips_ws.cell(r, 7).value = inv_amt if inv_amt is not None else (rev or 0)
            revenue_basis['invoiced'] += trips_ws.cell(r, 7).value
            if inv_date:
                trips_ws.cell(r, 10).value = to_date(inv_date)
            trips_ws.cell(r, 11).value = 'Y' if zoho == 'synced' else 'N'
        elif rev:
            trips_ws.cell(r, 7).value = rev
            revenue_basis['trip revenue'] += rev
        elif est:
            trips_ws.cell(r, 7).value = est
            revenue_basis['estimate'] += est
        elif client in rate_customers and km:
            trips_ws.cell(r, 7).value = f'=ROUND(I{r}*{rate},2)'
            revenue_basis['rate per km'] += round(km * rate, 2)
        else:
            trips_ws.cell(r, 7).value = 0
            trips_no_value += 1
            flag('Completed trip with no revenue in Wilogix', f'{number} {truck_name} {cust} ({d})')
        r += 1

    # ---------------- Expense Log ----------------
    exp_rows = read_rows(os.path.join(a.data, 'expenses.txt'), 9)
    fuel_rows = read_rows(os.path.join(a.data, 'fuel.txt'), 7)
    lines = []  # (date, applies_to, category, amount, paid_from, description, ref, deducted_customer)
    cash_total = 0.0
    for d, truck, cat, amt, wallet, memo, req, rat, chk in exp_rows:
        amt = to_num(amt) or 0.0
        if amt < 0:
            flag('Correction: an earlier month\'s expense was reversed in Vuma this month (cash came back)', f'{d} {cat} ${amt:.2f}')
        if rat == 'declined':
            flag('Declined Vuma expense left out', f'{d} {cat} ${amt:.2f} {memo}')
            continue
        truck_name = map_truck(truck)
        is_truck = truck_name in trucks
        if truck_name and not is_truck:
            flag('Expense on a vehicle that is not a report truck, treated as overhead', f'{d} {truck_name} {cat} ${amt:.2f} {memo}')
        override = None
        for rule in cfg.get('memo_keyword_overrides', []):
            if rule['contains'].lower() in (memo or '').lower():
                override = rule
                break
        if is_truck:
            new_cat = cfg['truck_category_map'].get(cat)
            if override and new_cat != override['truck_category']:
                flag('Category changed because of the description (fix the category in Vuma)', f'{d} {truck_name} "{memo}": {cat} -> {override["truck_category"]}')
                new_cat = override['truck_category']
            if new_cat is None:
                new_cat = cat
            if new_cat not in truck_cats:
                flag('Truck expense category not on the Assumptions tab (shows as Check)', f'{d} {truck_name} {cat} ${amt:.2f}')
            applies = truck_name
        else:
            new_cat = cfg['overhead_category_map'].get(cat)
            if override and new_cat != override['overhead_category'] and new_cat != 'Truck insurance (prepaid)':
                if cat != 'Fuel and Mileage':
                    flag('Category changed because of the description (fix the category in Vuma)', f'{d} overhead "{memo}": {cat} -> {override["overhead_category"]}')
                new_cat = override['overhead_category']
            if new_cat is None:
                new_cat = cfg['overhead_category_default']
                flag('Overhead category not mapped, put under Other overheads', f'{d} {cat} ${amt:.2f} {memo}')
            if new_cat == 'Truck insurance (prepaid)':
                words = cfg.get('prepaid_insurance_keywords', [])
                if words and not any(w in (memo or '').lower() for w in words):
                    new_cat = cfg['overhead_category_default']
                    flag('Paid under licences and insurance but the description is not insurance, put under Other overheads (fix the category in Vuma)', f'{d} ${amt:.2f} {memo}')
            applies = 'Overhead'
        if chk and chk not in ('clean',):
            flag('Vuma expense not yet clean (' + chk.replace('_', ' ') + ')', f'{d} {applies} {cat} ${amt:.2f} {memo}')
        lines.append((to_date(d), applies, new_cat, amt, map_wallet(wallet), memo, req, ''))
        cash_total += amt

    deducted_total, deducted_by_customer = 0.0, {}
    for d, truck, client, trip, litres, cost, paid in fuel_rows:
        cost = to_num(cost) or 0.0
        if paid != 'client':
            continue  # fuel we paid for is already a Vuma cash expense
        treatment = fuel_rules.get(client, 'Not set')
        if treatment == 'Deduct by credit note':
            truck_name = map_truck(truck)
            cust = map_customer(client)
            if truck_name not in trucks:
                flag('Client fuel on a truck that is not on the Assumptions tab', f'{d} {truck_name} {client} ${cost:.2f}')
            lines.append((to_date(d), truck_name, 'Fuel', cost, '', f'Client fuel deducted from invoice ({cust}), {float(litres):g} L', trip, cust))
            deducted_total += cost
            deducted_by_customer[cust] = round(deducted_by_customer.get(cust, 0) + cost, 2)
        elif treatment == 'Client funds, ignore':
            pass
        elif treatment == 'Charged on invoice':
            flag('Fuel log says the client paid, but the client rule is "Charged on invoice" (not counted as a cost)', f'{d} {truck} {client} ${cost:.2f}')
        elif treatment == 'We fund':
            flag('Fuel log says the client paid, but the client rule is "We fund" (not counted as a cost)', f'{d} {truck} {client} ${cost:.2f}')
        else:
            flag('Client-paid fuel for a client with no fuel treatment set (not counted)', f'{d} {truck} {client or "no trip linked"} ${cost:.2f}')

    lines.sort(key=lambda x: (x[0], x[7] != '', x[1]))
    capacity_check('Expense Log', len(lines))
    r = FIRST_ROW
    for d, applies, cat, amt, paid_from, desc, ref, deducted in lines:
        exp_ws.cell(r, 1).value = d
        exp_ws.cell(r, 2).value = applies
        exp_ws.cell(r, 3).value = cat
        exp_ws.cell(r, 4).value = amt
        if paid_from:
            exp_ws.cell(r, 5).value = paid_from
        if desc:
            exp_ws.cell(r, 6).value = desc
        if ref:
            exp_ws.cell(r, 7).value = ref
        exp_ws.cell(r, 8).value = 'N'
        if deducted:
            exp_ws.cell(r, 12).value = deducted
        r += 1

    # ---------------- Receipts and Transfers ----------------
    led_rows = read_rows(os.path.join(a.data, 'ledger.txt'), 7)
    receipts, transfers = [], OrderedDict()
    receipts_total = 0.0
    for d, typ, fw, tw, amt, cp, memo in led_rows:
        amt = to_num(amt) or 0.0
        if typ == 'collection':
            cust = map_customer(cp) if cp else ''
            if not cust:
                flag('Receipt with no customer name in Vuma', f'{d} ${amt:.2f} {memo}')
            receipts.append((to_date(d), cust, amt, memo or 'Vuma collection', map_wallet(tw)))
            receipts_total += amt
        elif typ in ('transfer', 'disbursement'):
            key = (to_date(d), map_wallet(fw), map_wallet(tw), memo if typ == 'disbursement' else None, typ)
            if typ == 'disbursement' and key in transfers:
                transfers[key]['amt'] += amt
                transfers[key]['n'] += 1
            else:
                if typ == 'transfer':
                    key = key + (len(transfers),)  # keep each transfer as its own line
                transfers[key] = {'amt': amt, 'n': 1, 'memo': memo, 'typ': typ}
        else:
            flag('Vuma ledger entry of a type the report does not use', f'{d} {typ} ${amt:.2f} {fw} -> {tw} {memo}')

    capacity_check('Receipts Log', len(receipts))
    r = FIRST_ROW
    for d, cust, amt, ref, into in receipts:
        rec_ws.cell(r, 1).value = d
        rec_ws.cell(r, 2).value = cust
        rec_ws.cell(r, 3).value = amt
        rec_ws.cell(r, 4).value = ref
        rec_ws.cell(r, 5).value = into
        rec_ws.cell(r, 6).value = 'N'
        r += 1

    capacity_check('Transfers Log', len(transfers))
    transfers_total = 0.0
    r = FIRST_ROW
    for key, t in transfers.items():
        note = t['memo'] or ('Payout' if t['typ'] == 'disbursement' else 'Transfer')
        if t['n'] > 1:
            note += f' ({t["n"]} entries)'
        tr_ws.cell(r, 1).value = key[0]
        tr_ws.cell(r, 2).value = key[1]
        tr_ws.cell(r, 3).value = key[2]
        tr_ws.cell(r, 4).value = round(t['amt'], 2)
        tr_ws.cell(r, 5).value = note
        tr_ws.cell(r, 6).value = 'N'
        transfers_total += t['amt']
        r += 1

    # ---------------- Headings ----------------
    asm['B3'].value = datetime(month_start.year, month_start.month, 1)
    stamp = date.today().strftime('%d %b %Y')
    month_name = month_start.strftime('%B %Y')
    wb['Start Here']['A2'].value = f'{month_name}: filled automatically from Wilogix and Vuma on {stamp}. USD.'
    exp_ws['A3'].value = (f'Filled automatically on {stamp}: cash expenses from Vuma, client fuel from Wilogix fuel logs. '
                          'Declined and reversed Vuma expenses are left out.')

    wb.save(a.out)

    summary = OrderedDict([
        ('month', month_name),
        ('output', a.out),
        ('fuel_rules_source', fuel_rules_source),
        ('trips', len(trip_rows)),
        ('trips_with_no_value', trips_no_value),
        ('revenue_by_basis', {k: round(v, 2) for k, v in revenue_basis.items()}),
        ('revenue_total', round(sum(revenue_basis.values()), 2)),
        ('expense_lines', len(lines)),
        ('cash_expenses_total', round(cash_total, 2)),
        ('client_deducted_fuel_total', round(deducted_total, 2)),
        ('client_deducted_fuel_by_customer', deducted_by_customer),
        ('receipts_lines', len(receipts)),
        ('receipts_total', round(receipts_total, 2)),
        ('transfer_lines', len(transfers)),
        ('transfers_total', round(transfers_total, 2)),
        ('exceptions', exc),
    ])
    print(json.dumps(summary, indent=1, default=str))


if __name__ == '__main__':
    main()
