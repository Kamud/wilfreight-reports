# Willfreight daily truck report

Fills the daily truck profit workbook from Wilogix (trips, fuel logs) and Vuma (expenses, receipts, transfers).
Only the four log tabs are written. Truck Summary, Customers, Cash & Banks, Daily Trend and Charts calculate themselves.

## Files

| File | Purpose |
|---|---|
| `template/daily_truck_report_template.xlsx` | The workbook with empty logs. Assumptions, opening balances and customer opening balances are kept here. |
| `fill_report.py` | Writes the four logs from the extracts below and prints a JSON summary with exceptions. |
| `config.json` | Name mappings: customers, wallets to accounts, Vuma categories to report categories, fuel treatment fallback. |

## Daily run

Period: `:start` is the first day of the month being reported, `:end` is today (Harare date).
On the 1st of a month, report the whole previous month.

1. Run the four queries below with read-only SELECTs and save each result's `v` text exactly as returned
   into `data/trips.txt`, `data/fuel.txt`, `data/expenses.txt`, `data/ledger.txt` (one record per line).
2. Optional: write `data/fuel_rules.json` as `{"<Wilogix client name>": "<Fuel treatment>"}` from the
   Client rules tab of the "Willfreight Report Rules" sheet. Without it, `config.json` is used.
3. `python3 fill_report.py --data data --month-start :start --fuel-rules data/fuel_rules.json --out "out/Wilfreight Logistics Daily Truck Report YYYY-MM-DD.xlsx"`
4. Recalculate the output with the xlsx skill's `recalc.py`; it must report zero errors.
5. Check: the summary's `cash_expenses_total` equals the Vuma expense total for the period, `transfers_total`
   equals the Vuma transfer plus disbursement total, and Truck Summary revenue equals `revenue_total`.

### Wilogix (Supabase project `ameqpvqxjqgjisggoklo`)

```sql
with km as (select trip_id, sum(distance_km) km from mileage_logs where trip_id is not null group by 1),
t as (
  select (tr.actual_return at time zone 'Africa/Harare')::date d, tr.number, v.reg_number truck, coalesce(c.name, tr.client_name) client,
    regexp_replace(coalesce(tr.origin,'')||' to '||coalesce(tr.destination,''), '[|\n\r]+', ' ', 'g') route,
    case when i.voided_at is null then i.invoice_number end inv_no, tr.invoiced_amount, tr.revenue, tr.estimated_revenue, km.km,
    case when i.voided_at is null then i.invoice_date end inv_date, case when i.voided_at is null then i.zoho_sync_status end zs
  from trips tr left join vehicles v on v.id=tr.vehicle_id left join clients c on c.id=tr.client_id
    left join km on km.trip_id=tr.id left join invoices i on i.id=tr.invoice_id
  where tr.status='completed' and (tr.actual_return at time zone 'Africa/Harare')::date between :start and :end
),
f as (
  select fl.log_date d, v.reg_number truck, coalesce(c.name, t.client_name) client, t.number trip, fl.litres, fl.total_cost, fl.paid_through
  from fuel_logs fl left join vehicles v on v.id=fl.vehicle_id left join trips t on t.id=fl.trip_id left join clients c on c.id=t.client_id
  where fl.log_date between :start and :end
)
select 'TRIPS' k, string_agg(concat_ws('|', d, number, truck, client, route, coalesce(inv_no,''), coalesce(invoiced_amount::text,''),
    coalesce(revenue::text,''), coalesce(estimated_revenue::text,''), coalesce(round(km,1)::text,''), coalesce(inv_date::text,''), coalesce(zs,'')),
    E'\n' order by d, number) v from t
union all
select 'FUEL', string_agg(concat_ws('|', d, truck, coalesce(client,''), coalesce(trip,''), litres, round(total_cost,2), coalesce(paid_through,'')),
    E'\n' order by d, truck) from f;
```

### Vuma (Supabase project `tutwkppuyfduzcskhkve`)

```sql
with o as (select id from organisations where slug='wilfreight'),
e as (
  select e.expense_date d, coalesce(e.truck_reg,'') truck, coalesce(c.name,'') cat, (e.amount_cents/100.0)::numeric(12,2) amt, coalesce(w.name,'') wallet,
    regexp_replace(coalesce(e.memo,''), '[|\n\r]+', ' ', 'g') memo, coalesce(r.number,'') req, coalesce(e.ratification,'') rat, coalesce(e.check_result,'') chk, e.created_at
  from expenses e left join expense_categories c on c.id=e.category_id left join wallets w on w.id=e.wallet_id left join requests r on r.id=e.request_id
  where e.organisation_id=(select id from o) and e.expense_date between :start and :end
    and not exists (select 1 from ledger_entries x where x.reverses_entry_id=e.ledger_entry_id)
  union all
  -- an expense from an earlier month reversed in this period: the cash came back, so it goes in as a negative line
  select rv.entry_date, coalesce(e2.truck_reg,''), coalesce(c2.name,''), (-rv.amount_cents/100.0)::numeric(12,2), coalesce(tw.name,''),
    'Correction: expense dated '||to_char(orig.entry_date,'DD Mon YYYY')||' was reversed in Vuma, so the cash came back', '', 'not_needed', 'clean', rv.created_at
  from ledger_entries rv join ledger_entries orig on orig.id=rv.reverses_entry_id left join wallets tw on tw.id=rv.to_wallet_id
    left join expenses e2 on e2.ledger_entry_id=orig.id left join expense_categories c2 on c2.id=e2.category_id
  where rv.organisation_id=(select id from o) and rv.entry_date between :start and :end and orig.entry_date < :start
),
l as (
  select le.entry_date d, le.type, coalesce(fw.name,'') fw, coalesce(tw.name,'') tw, (le.amount_cents/100.0)::numeric(12,2) amt, coalesce(le.counterparty_name,'') cp,
    regexp_replace(coalesce(le.memo,'')||case when le.reference is not null and le.reference<>'' then ' ('||le.reference||')' else '' end, '[|\n\r]+', ' ', 'g') memo, le.seq
  from ledger_entries le left join wallets fw on fw.id=le.from_wallet_id left join wallets tw on tw.id=le.to_wallet_id
  where le.organisation_id=(select id from o) and le.entry_date between :start and :end and le.type <> 'expense'
)
select 'EXPENSES' k, string_agg(concat_ws('|', d, truck, cat, amt, wallet, memo, req, rat, chk), E'\n' order by d, created_at) v from e
union all
select 'LEDGER', string_agg(concat_ws('|', d, type, fw, tw, amt, cp, memo), E'\n' order by d, seq) from l;
```

## Truck by Day tab

Part of the template, so every run includes it; `fill_report.py` needs no change. One row per truck, one column per
day, with a month-to-date column. The yellow cell switches the measure between Revenue, Trip costs and repairs, and
Contribution (revenue less trip costs and repairs). It reads only the Trips and Expense Log tabs, counts days up to
the as-at date, and its "Check vs Truck Summary" cell must show 0 for all three measures. Running costs and company
overheads are monthly accruals and are not spread across days there; see Truck Summary and Daily Trend.

## Rules the fill applies

- Revenue counts on the trip completion date. Amount used, in order: invoiced amount, trip revenue, estimated revenue,
  then 0.9 x logged km for Bakers Inn. Anything else is written as 0 and listed as an exception.
- Cash expenses come from Vuma. Reversed and declined expenses are left out. An expense from an earlier month that is reversed
  in this period goes in as a negative line, because the cash came back this month.
- Fuel logs marked as paid by the client become Expense Log lines deducted from the customer only when the client's
  fuel treatment is "Deduct by credit note". Fuel we paid for is already a Vuma cash expense and is not added again.
- Receipts are Vuma collections. Transfers are Vuma transfers and payouts; payouts with the same date, accounts and note are combined.
- A customer or account missing from the workbook lists is added to the next empty row and reported, so its opening balance can be entered.

## What still needs a person

- Opening balances (Cash & Banks) and customer balances at month start (Customers), once a month in the template.
- Truck running costs and fixed overheads on the Assumptions tab.
- Anything not recorded in Wilogix or Vuma will not appear. Record it at the source rather than in the workbook.

## Weekly report (Monday morning)

Covers Monday to Sunday of the week just ended, compared with the week before. The workbook holds its own data:
the daily report's log tabs (Trips, Expense Log, Receipts Log, Transfers Log) from the 1st of the month (or from the
comparison Monday, if earlier) to Sunday, plus Cash & Banks, Customers and Assumptions. Every summary figure is a
formula over the logs, so a manual edit to a log line flows through. The only typed numbers outside the logs are
inputs: opening balances at the first log date (worked back from the daily report's Sunday balances) and Assumptions.

Dates: `:week_end` is yesterday (Sunday), `:week_start` is 6 days earlier, the comparison week is the 7 days before.
Logs start at the earlier of the comparison Monday and the 1st of `:week_end`'s month.

1. For each month the logs touch (one, or two when the comparison week starts in the previous month):
   - Extract with the two queries above. For the month holding `:week_end`, use `:start` = its first day and `:end` = `:week_end`.
     For an earlier month, use the whole month. Save into `data/<YYYY-MM>/` and write `fuel_rules.json` there as for the daily run.
   - `python3 fill_report.py --data data/<YYYY-MM> --month-start <YYYY-MM-01> --fuel-rules data/<YYYY-MM>/fuel_rules.json --out "out/weekly/month <YYYY-MM>.xlsx" > out/weekly/summary_<YYYY-MM>.json`
   - Freeze it BEFORE recalculating, so a Monday run does not accrue Monday's costs:
     `python3 weekly_report.py freeze --workbook "out/weekly/month <YYYY-MM>.xlsx" --as-at <:week_end, or month end for an earlier month>`
   - Recalculate it with the xlsx skill's `recalc.py` (zero errors).
2. `python3 weekly_report.py build --week-end <:week_end> --monthly "out/weekly/month <YYYY-MM>.xlsx" [--monthly "out/weekly/month <earlier YYYY-MM>.xlsx"] --summary out/weekly/summary_<YYYY-MM of :week_end>.json`
   It names the file `out/Wilfreight Logistics Weekly Report <period>.xlsx`, e.g. `16-22 November 2026` (or `27 October - 2 November 2026`), and prints the path.
3. Recalculate the weekly workbook with `recalc.py` (zero errors).
4. `python3 weekly_report.py verify --weekly "<the path build printed>" --monthly "out/weekly/month <YYYY-MM of :week_end>.xlsx"`
   must print `"status": "ok"`: the CHECKS block on Week Summary is all 0, and month to date, cash and customer balances
   equal the daily report.

Tabs: Week Summary (headline; activity, profit and loss, key ratios and cash for both weeks with change, change % and the
biggest change; trucks; top five customers; five biggest non-fuel costs; needs attention; month to date; checks),
By Truck, By Customer, By Day, Costs, Exceptions, the four logs, Cash & Banks, Customers, Assumptions.
The Monday in Week Summary B3 is the one input that sets both weeks.

Limit: when the comparison week falls in the previous month, the current Assumptions are used for it too.
