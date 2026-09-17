import csv
import json
import os
import tempfile
import unittest
from unittest.mock import patch

import app
from settings_workbook import Workbook


class SettingsWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        paths = {
            "SETTINGS_WORKBOOK_PATH": "settings.xlsx", "CUSTOMER_MASTER_XLSX_PATH": "customer_master.xlsx",
            "EMAIL_SETTINGS_CSV_PATH": "email_settings.csv", "EMAIL_CONFIG_PATH": "email.json",
            "APP_SETTINGS_PATH": "invoice.json", "CUSTOMERS_PATH": "customers.json",
            "FARMS_PATH": "farms.json", "FIELD_MAP_PATH": "fields.json", "MUCK_TYPES_PATH": "muck.json",
        }
        self.patch = patch.multiple(app, DATA_DIR=self.temp.name, **{key: os.path.join(self.temp.name, value) for key, value in paths.items()})
        self.patch.start()
        self.book = Workbook(app.SETTINGS_WORKBOOK_PATH)
        headers = app.CUSTOMER_MASTER_HEADERS
        customer = {"customer_name": "Example Customer", "farm_name": "Example Farm", "email": "customer@example.invalid", "active": "1", "rate_per_ton": "3.5", "vat_rate": "20"}
        Workbook(app.CUSTOMER_MASTER_XLSX_PATH).create({"Customers": [headers, [customer.get(key, "") for key in headers]]})
        settings = dict(app.DEFAULT_EMAIL_CONFIG, smtp_host="smtp.example.invalid", smtp_username="test@example.invalid", smtp_password="test password", from_email="test@example.invalid", enabled=False)
        with open(app.EMAIL_SETTINGS_CSV_PATH, "w", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["record_type", "name", "email", "active"] + list(settings))
            writer.writeheader()
            writer.writerow(dict(settings, record_type="settings", to_emails="extra@example.invalid"))
            writer.writerow({"record_type": "recipient", "name": "Office", "email": "office@example.invalid", "active": "1"})
            writer.writerow({"record_type": "recipient", "name": "Inactive", "email": "inactive@example.invalid", "active": "0"})
        for path, value in [(app.CUSTOMERS_PATH, ["Example Customer", "Remembered Customer"]), (app.FARMS_PATH, ["Example Farm"]), (app.MUCK_TYPES_PATH, ["Cattle"]), (app.FIELD_MAP_PATH, {"Example Customer": {"Example Farm": ["North Field"]}}), (app.APP_SETTINGS_PATH, {"invoice_default_payment_terms_days": "21"})]:
            with open(path, "w") as handle:
                json.dump(value, handle)

    def tearDown(self):
        self.patch.stop()
        self.temp.cleanup()

    def test_migration_preserves_effective_settings_and_original_files(self):
        with open(app.APP_SETTINGS_PATH, "w") as handle:
            json.dump({"invoice_customer_message_template": "Hello,\r\n\r\nInvoice attached.", "invoice_default_payment_terms_days": "21"}, handle)
        before = app.load_email_config()
        invoices = app.load_app_settings()
        recipients = app.load_email_recipient_options()
        fields = app.load_field_map()
        with open(app.CUSTOMER_MASTER_XLSX_PATH, "rb") as handle:
            original = handle.read()
        app.ensure_settings_workbook()
        self.assertIn("timesheet", self.book.rows("Email Recipients")[0])
        self.assertIn("pin", self.book.rows("Staff")[0])
        self.assertIn("email_2", self.book.rows("Customers")[0])
        self.assertEqual(self.book.rows("Machinery"), [["name", "active"]])
        self.assertEqual(app.load_email_config(), before)
        self.assertEqual(app.load_app_settings(), invoices)
        self.assertEqual(app.load_email_recipient_options(), recipients)
        self.assertEqual(app.load_field_map(), fields)
        self.assertEqual(app.load_customers(), ["Example Customer", "Remembered Customer"])
        self.assertEqual(app.settings_names("Staff", []), ["Owen Curl"])
        self.assertEqual(app.load_customer_master_rows()[0]["rate_per_ton"], "3.5")
        with open(app.CUSTOMER_MASTER_XLSX_PATH, "rb") as handle:
            self.assertEqual(handle.read(), original)
        self.book.set_rows("Staff", [["name", "active"], ["New Worker", "1"]])
        app.ensure_settings_workbook()
        self.assertEqual(app.settings_names("Staff", []), ["New Worker"])
        self.assertIn("pin", self.book.rows("Staff")[0])

    def test_workbook_edits_are_used_without_restart(self):
        app.ensure_settings_workbook()
        self.book.set_rows("Staff", [["name", "active", "pin"], ["Owen Curl", "1", "1234"], ["Second Worker", "1", "5678"], ["Hidden Worker", "0", "9999"]])
        self.book.set_rows("Companies", [["name", "active"], ["New Company", "1"]])
        with app.app.test_client() as client:
            with client.session_transaction() as session:
                session["timesheet_staff"] = "Second Worker"
            html = client.get("/timesheet/day/2026-09-15").get_data(as_text=True)
            self.assertIn('value="Second Worker"', html)
            self.assertNotIn("Hidden Worker", html)
            self.assertIn("New Company", html)
            saved = client.post("/api/timesheets", json={"name": "Second Worker", "date": "2026-09-15", "company": "New Company", "notes": "Test"})
            self.assertEqual(saved.status_code, 201)
        self.book.set_rows("Email Recipients", [["name", "email", "active", "summary", "invoice_option"]])
        self.assertEqual(app.load_email_config()["to_emails"], [])
        self.assertEqual(app.load_email_recipient_options(), [])
        app.update_settings_key_values("Email Settings", {"send_hour": "0", "monthly_send_hour": "0", "monthly_send_minute": "0"})
        self.assertEqual(app.load_email_config()["send_hour"], 0)
        self.assertEqual(app.load_email_config()["monthly_send_hour"], 0)

    def test_app_edits_update_workbook_and_preserve_other_tabs(self):
        app.ensure_settings_workbook()
        staff = self.book.rows("Staff")
        app.save_app_settings({"invoice_default_payment_terms_days": "30"})
        app.update_email_settings_csv({"smtp_host": "changed.example.invalid", "smtp_port": "587", "smtp_username": "new", "from_email": "new@example.invalid", "use_tls": "1", "smtp_password": ""})
        self.assertEqual(app.load_email_config()["smtp_password"], "test password")
        self.assertEqual(app.load_email_config()["smtp_host"], "changed.example.invalid")
        self.assertEqual(app.load_app_settings()["invoice_default_payment_terms_days"], "30")
        ok, _ = app.save_customer_master_customer_details("Example Customer", "updated@example.invalid", "4.25", row_number=2)
        self.assertTrue(ok)
        self.assertEqual(app.load_customer_master_rows()[0]["email"], "updated@example.invalid")
        app.save_customers(app.load_customers() + ["New Customer"])
        self.assertEqual(app.load_customer_master_rows()[0]["rate_per_ton"], "4.25")
        self.assertIn("New Customer", app.load_customers())
        app.save_farms(["Second Farm"])
        self.assertEqual(app.load_farms(), ["Second Farm"])
        app.save_muck_types(["Slurry"])
        self.assertIn("Slurry", app.load_muck_types())
        app.save_field_map({"New Customer": {"Second Farm": ["West Field"]}})
        self.assertEqual(app.load_field_map(), {"New Customer": {"Second Farm": ["West Field"]}})
        self.assertEqual(self.book.rows("Staff"), staff)

    def test_update_app_control_is_on_settings_page(self):
        self.assertNotIn(">Update App<", app.HTML)
        with patch.object(app, "app_version_label", return_value="abc123"):
            with app.app.test_client() as client:
                html = client.get("/settings").get_data(as_text=True)
        self.assertIn('action="/app/update"', html)
        self.assertIn(">Update App<", html)
        self.assertIn("Current version: abc123", html)


if __name__ == "__main__":
    unittest.main()
