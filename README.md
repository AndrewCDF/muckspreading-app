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
- New muck types can be managed from `customer_master.csv` and are also remembered from saved jobs
- Field names are linked to each customer
- Total spreader tons
- Total Ops Center tons
- Recent jobs list
- CSV export
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

The app also reads a master spreadsheet file in the app folder:

- `customer_master.csv`

The GitHub repo includes a blank starter file:

- `customer_master.template.csv`

## Customer Master Spreadsheet

On a new install, copy `customer_master.template.csv` to `customer_master.csv`, then edit `customer_master.csv` directly in Excel or Numbers and keep these columns:

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
- when you save a job with a new customer, farm, or muck type combination, the app can add that combination back into `customer_master.csv`
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

## Weekly Summary Email

The app can send an automatic weekly jobs summary email.

How it works:

- It sends a summary of the previous full week
- Default schedule is Monday at `05:00`
- It includes total jobs, total spreader tons, total Ops Center tons, and an attached `.xlsx` summary file
- It will only send once for each weekly period
- If `weekly_summary_layout_template.xlsx` exists in the app folder, the generated attachment will follow that workbook's layout and styling

Setup:

1. Edit `email_settings.csv`
2. Put one `settings` row in it for the SMTP and schedule details
3. Add as many `recipient` rows as you need underneath
5. Set `enabled` to `1`
6. Keep the app running on the Pi service so the weekly worker can send it

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
- `subject_prefix`
- `active`

Use:

- `record_type`: `settings` for the main config row, `recipient` for recipient rows
- `email`: use this on recipient rows
- `name`: optional label for recipient rows
- `enabled`: `1` or `0`
- `use_tls`: `1` or `0`
- `send_weekday`: `0` for Monday through `6` for Sunday
- `to_emails`: optional comma-separated emails on the `settings` row if you want to keep some addresses there too
- `active`: `1` or `0`

If you already have `data/email_config.json`, the app will still accept it, but `email_settings.csv` is now the easiest way to manage the full email setup.

Useful endpoints:

```text
GET  /api/email/weekly/preview
POST /api/email/weekly/send-now
```

Then open:

```text
http://127.0.0.1:8093
```

On the office network Pi at `192.168.1.19`, open:

```text
http://192.168.1.19:8093
```

## Tailscale access

The app binds to `0.0.0.0`, so once Tailscale is running on the office Pi you can also reach it remotely over the Pi's Tailscale IP or MagicDNS hostname on port `8093`.

Examples:

```text
http://100.x.x.x:8093
http://office-pi-name.tailnet-name.ts.net:8093
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
