# Muckspreading App

Standalone web app for the office Raspberry Pi.

## Features

- Today's date filled in automatically
- Customer type-to-filter dropdown
- Farm name type-to-filter dropdown
- Field name type-to-filter dropdown
- Muck type type-to-filter dropdown
- New customer names are remembered for future entries
- New farm names are remembered for future entries
- New field names are remembered for future entries
- New muck types can be managed from `customer_master.xlsx` and are also remembered from saved jobs
- Field names are linked to each customer
- Total spreader tons
- Total Ops Center tons
- Recent jobs list
- CSV export
- Invoice creation and email for uninvoiced customer or farm jobs
- Optional weekly summary email
- Customer master spreadsheet import for address, email, and pricing

## Data files

The app keeps its own data in:

- `data/jobs.ndjson`
- `data/customers.json`
- `data/farms.json`
- `data/muck_types.json`
- `data/customer_fields.json`
- `data/email_config.json`
- `data/weekly_email_state.json`
- `data/invoice_ledger.json`
- `data/invoice_state.json`
- `data/invoices/`

The app reads a master spreadsheet file in the app folder:

- `customer_master.xlsx`

## Customer Master Spreadsheet

- create `customer_master.xlsx` in the app folder
- keep the first sheet as the customer master sheet
- keep these columns on the header row:

- `customer_name`
- `farm_name`
- `email`
- `address_line_1`
- `address_line_2`
- `town`
- `postcode`
- `rate_per_ton`
- `vat_rate`
- `active`
- `muck_type`

What it does:

- customer names from the spreadsheet appear in the customer suggestions
- farm names from the spreadsheet appear in the farm suggestions
- muck types from the spreadsheet appear in the muck type suggestions
- when a job matches a customer and farm from the spreadsheet, the app snapshots:
  - email
  - address
  - rate per ton
  - VAT rate
- those values are then stored with the saved job and included in the full CSV export

## Run

```bash
cd /home/pi/muckspreading-app
python3 app.py
```

## Development Auto Reload

For local editing, start it with:

```bash
MUCKSPREADING_APP_DEBUG=1 python3 app.py
```

That turns on Flask auto-reload so the app restarts after code edits.

## Summary Emails

The app can send automatic weekly and monthly jobs summary emails.

How it works:

- It sends a summary of the previous full week
- It sends a summary of the current month on the last day of the month
- Default schedule is Monday at `05:00`
- Monthly email default time is `05:00` on month end
- It includes total jobs, total spreader tons, total Ops Center tons, and attached `.xlsx` and `.pdf` summary files
- It will only send once for each weekly or monthly period
- If `weekly_summary_layout_template.xlsx` exists in the app folder, the generated attachment will follow that workbook's layout and styling

Setup:

1. Edit `email_settings.csv`
2. Put one `settings` row in it for the SMTP and schedule details
3. Add as many `recipient` rows as you need underneath
5. Set `enabled` to `1`
6. Set `monthly_enabled` to `1` if you want month-end email switched on
7. Keep the app running on the Pi service so the background worker can send it

Email settings CSV columns:

- `record_type`
- `email`
- `name`
- `enabled`
- `smtp_host`
- `smtp_port`
- `use_tls`
- `smtp_username`
- `smtp_password`
- `from_email`
- `to_emails`
- `send_weekday`
- `send_hour`
- `send_minute`
- `monthly_enabled`
- `monthly_send_hour`
- `monthly_send_minute`
- `subject_prefix`
- `active`

Use:

- `record_type`: `settings` for the main config row, `recipient` for recipient rows
- `email`: use this on recipient rows
- `name`: optional label for recipient rows
- `enabled`: `1` or `0`
- `use_tls`: `1` or `0`
- `send_weekday`: `0` for Monday through `6` for Sunday
- `monthly_enabled`: `1` or `0`
- `to_emails`: optional comma-separated emails on the `settings` row if you want to keep some addresses there too
- `active`: `1` or `0`

If you leave the monthly columns out, the app falls back to the weekly time settings.

If you already have `data/email_config.json`, the app will still accept it, but `email_settings.csv` is now the easiest way to manage the full email setup.

Useful endpoints:

```text
GET  /api/email/weekly/preview
POST /api/email/weekly/send-now
GET  /api/email/monthly/preview
POST /api/email/monthly/send-now
```

## Invoicing

The app can create and email an invoice for a customer or farm scope using jobs that have not already been invoiced.

How it works:

- choose a customer and optionally a farm from the main app
- the app pulls all uninvoiced jobs in that scope
- it uses the saved job snapshot for customer email, rate per ton, VAT, and address
- the rate per ton field defaults from the customer master and can be changed for that invoice
- you can add manual fee lines with separate description and amount boxes, with VAT fixed at `20%`
- it shows an on-screen preview with the customer email, automatic accounts recipients, editable subject/message text, and invoice lines before sending
- it creates numbered `.xlsx` and `.pdf` invoice files
- if `invoice_layout_template.xlsx` exists in the app folder, generated invoice workbooks follow that template layout
- it emails the `.pdf` invoice to the customer
- it emails both the `.pdf` and `.xlsx` invoice files to the automatic accounts recipients
- it stores the invoice record in `data/invoice_ledger.json`
- it archives both generated invoice files in `data/invoices/`

Notes:

- invoice numbers are consecutive and stored in `data/invoice_state.json`
- Office and Andrew are used for automatic invoice copies, while Owen is excluded from invoice emails
- you can restyle `invoice_layout_template.xlsx`, but it is best to keep the same basic row structure
- if no previous invoice exists for that scope, the first invoice uses all saved jobs in scope
- if a job has already been included in a previous invoice, it will not be included again
- if rate per ton or customer email is missing for a job scope, the app will stop and show an error instead of sending a bad invoice

Then open:

```text
http://127.0.0.1:8094
```

On the office network Pi at `192.168.1.19`, open:

```text
http://192.168.1.19:8094
```

## Tailscale access

The app binds to `0.0.0.0`, so once Tailscale is running on the office Pi you can also reach it remotely over the Pi's Tailscale IP or MagicDNS hostname on port `8094`.

Examples:

```text
http://100.x.x.x:8094
http://office-pi-name.tailnet-name.ts.net:8094
```

## Install As A Service On The Office Pi

Copy the app folder onto the Pi, for example:

```bash
/home/pi/muckspreading-app
```

Then install the included service:

```bash
cd /home/pi/muckspreading-app
sudo cp muckspreading-app.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now muckspreading-app.service
```

Useful checks:

```bash
sudo systemctl status muckspreading-app.service
journalctl -u muckspreading-app.service -n 100 --no-pager
```
