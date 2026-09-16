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
- Machinery maintenance records with service dates, work completed, parts, costs and next-service dates

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
- `data/straw_deliveries.sqlite3`

The app reads its settings from `settings.xlsx`. The customer/farm/field and email JSON files listed above are retained for legacy migration; after migration those settings are read and written in the workbook.

## One settings workbook

Edit **`settings.xlsx`** in the app folder. The app creates it once from the existing settings when it starts. It is the settings source after migration; the old CSV, JSON configuration files and `customer_master.xlsx` remain as backups.

| Tab | Controls |
| --- | --- |
| Customers | Customer/farm names, email, addresses, rate per ton, VAT and active status |
| Staff | Timesheet names, active status and a separate 4–8 digit PIN for each person; a blank PIN is created by the person on first use |
| Companies | Companies available on timesheets |
| Email Settings | SMTP credentials, weekly/monthly schedules and enable switches |
| Email Recipients | Recipient names and addresses; summary, invoice-option and monthly-timesheet switches |
| Invoice Settings | Sender, wording and payment terms |
| Farms / Fields | Farm suggestions and customer/farm/field links |
| Muck Types / Straw Crops | Material and crop options |
| Machinery | Machinery names offered in maintenance records; new names entered in the app are added automatically |
| Read Me | Editing instructions |

Keep tab names and column headings unchanged. Add rows to add names or recipients. Set `active` to `1` to show a row or `0` to hide it. Save and close Excel, then refresh the app. Workbook edits take effect without an app restart. Settings and customer edits made in the app save back to this workbook.

Email Settings uses `setting` and `value` columns. Boolean values are `1` or `0`; `send_weekday` uses Monday `0` through Sunday `6`; hours use `0`–`23`. On Email Recipients, `summary=1` includes an active recipient in summary emails, `invoice_option=1` makes them available for invoice copies, and `timesheet=1` sends them completed monthly timesheets. Existing invoice-copy selection rules continue to apply.

The Timesheet button opens a staff selector. If a staff member has no PIN, the app asks them to choose one the first time they open their name. They can later use Change PIN inside their timesheet. The app limits entries, exports, print reports and completed-month emails to the unlocked person. The calendar has controls below each month to export a complete Excel workbook, open a printable report that can be saved as PDF, and complete and email the month. Overnight hours are assigned to the calendar date on which they were worked, including shifts crossing month-end.

The Hay & Straw home screen has a **+** button for recording delivered lorry or trailer loads. The weight can be left blank initially, and saved deliveries can be reopened later to add or correct it.

### Moving data from the old Straw app

The old standalone app stores all field, crop, customer, stocktake, load, stock-movement, map and embedded photo data in `/home/pi/StrawApp/data/straw-records.json`. Copy that file to `/home/pi/muckspreading-app/data/straw-records.json` after pulling this app. The integrated app reads the file directly and migrates legacy loads into its editable Loads Out records. Old loads did not contain a customer field, so they appear as **Customer not recorded** until edited. The old app folder, mock spreadsheet and browser cache do not need to be copied when this JSON file exists.

```bash
cp /home/pi/StrawApp/data/straw-records.json /home/pi/muckspreading-app/data/straw-records.json
```

The workbook contains email credentials, so it is excluded from Git. App backups include it. Jobs, timesheets, invoices and saved straw records remain in their existing data stores.

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

1. Open `settings.xlsx`.
2. Configure the **Email Settings** and **Email Recipients** tabs.
3. Set `enabled` to `1` for weekly summaries and `monthly_enabled` to `1` for month-end summaries.
4. Save and close the workbook, and keep the app running for scheduled delivery.

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
