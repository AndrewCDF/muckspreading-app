import tempfile
import unittest
import os
from unittest.mock import MagicMock, patch

import app


class TimesheetTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.data_patch = patch.object(app, "DATA_DIR", self.directory.name)
        self.data_patch.start()
        self.settings_patch = patch.object(app, "SETTINGS_WORKBOOK_PATH", self.directory.name + "/settings.xlsx")
        self.settings_patch.start()
        self.month_state_patch = patch.object(app, "TIMESHEET_MONTH_STATE_PATH", self.directory.name + "/timesheet_month_state.json")
        self.month_state_patch.start()
        self.client = app.app.test_client()
        self.authorize(self.client, "Owen Curl")
        self.entry = {
            "name": "Owen Curl", "date": "2026-09-15",
            "company": "Cherry Dene Farm Ltd.",
            "start": "08:00", "finish": "", "notes": "08:00 - Loaded straw",
        }

    def authorize(self, client, name):
        with client.session_transaction() as session:
            session["timesheet_staff"] = name

    def tearDown(self):
        self.data_patch.stop()
        self.settings_patch.stop()
        self.month_state_patch.stop()
        self.directory.cleanup()

    def test_save_reopen_and_continue_day(self):
        created = self.client.post("/api/timesheets", json=self.entry)
        self.assertEqual(created.status_code, 201)
        saved = created.get_json()["entry"]
        second_client = app.app.test_client()
        self.authorize(second_client, "Owen Curl")
        reopened = second_client.get("/api/timesheets").get_json()["entries"][0]
        self.assertEqual(reopened, saved)
        reopened.update(notes=reopened["notes"] + "\n11:00 - Delivered bales", finish="17:30")
        updated = self.client.put("/api/timesheets/%s" % saved["id"], json=reopened)
        self.assertEqual(updated.status_code, 200)
        rows = self.client.get("/api/timesheets").get_json()["entries"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["notes"], reopened["notes"])
        self.assertEqual(rows[0]["finish"], "17:30")
        stale = self.client.put("/api/timesheets/%s" % saved["id"], json=saved)
        self.assertEqual(stale.status_code, 409)
        self.assertEqual(self.client.get("/api/timesheets").get_json()["entries"], rows)

    def test_duplicate_day_and_separate_companies(self):
        self.entry["finish"] = "12:00"
        self.assertEqual(self.client.post("/api/timesheets", json=self.entry).status_code, 201)
        self.assertEqual(self.client.post("/api/timesheets", json=dict(self.entry, name="test worker")).status_code, 409)
        afternoon = dict(self.entry, company="A. Farrell Contracting Ltd.", start="12:30", finish="17:00", notes="Afternoon spreading")
        self.assertEqual(self.client.post("/api/timesheets", json=afternoon).status_code, 201)
        rows = self.client.get("/api/timesheets").get_json()["entries"]
        self.assertEqual(len(rows), 2)
        for expected in [self.entry, afternoon]:
            saved = next(row for row in rows if row["company"] == expected["company"])
            for field, value in expected.items():
                self.assertEqual(saved[field], value)

    def test_invalid_entries_are_not_saved(self):
        for changes in [{"company": "Unknown"}, {"date": "2026-02-30"}, {"start": "25:00"}, {"start": "08:07"}, {"notes": []}]:
            with self.subTest(changes=changes):
                self.assertEqual(self.client.post("/api/timesheets", json=dict(self.entry, **changes)).status_code, 400)
        self.assertEqual(self.client.get("/api/timesheets").get_json()["entries"], [])

    def test_form_and_assets(self):
        staff_page = self.client.get("/timesheet")
        self.assertEqual(staff_page.status_code, 200)
        self.assertIn(b'Choose your name', staff_page.data)
        self.assertIn(b'Owen Curl', staff_page.data)
        self.assertNotIn(b'id="month-calendar"', staff_page.data)
        self.authorize(self.client, "Owen Curl")
        calendar = self.client.get("/timesheet/calendar")
        self.assertIn(b'id="month-calendar"', calendar.data)
        self.assertIn(b'<strong aria-hidden="true">\xe2\x86\x90</strong>Back', calendar.data)
        self.assertNotIn(b'Switch Staff', calendar.data)
        page = self.client.get("/timesheet/day/2026-09-15")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b'data-date="2026-09-15"', page.data)
        self.assertIn(b'id="timesheet-form"', page.data)
        self.assertIn(b'<strong aria-hidden="true">\xe2\x86\x90</strong>Back', page.data)
        self.assertNotIn(b'id="month-calendar"', page.data)
        self.assertEqual(page.data.count(b'<option value="'), 4 + 2 * 96)
        self.assertIn(b'<option value="06:00">6:00 am</option>', page.data)
        self.assertIn(b'<option value="03:00">3:00 am</option>', page.data)
        self.assertEqual(self.client.get("/timesheet/day/2026-02-30").status_code, 404)
        for company in app.TIMESHEET_COMPANIES:
            self.assertIn(company.encode(), page.data)
        for path in ["/static/timesheet.js", "/static/timesheet.css"]:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)

    def test_overnight_shift_crosses_month_end(self):
        overnight = dict(self.entry, date="2026-09-30", start="06:00", finish="03:00")
        response = self.client.post("/api/timesheets", json=overnight)
        self.assertEqual(response.status_code, 201)
        entry = response.get_json()["entry"]
        self.assertTrue(entry["crosses_midnight"])
        self.assertEqual(entry["finish_date"], "2026-10-01")
        self.assertEqual(entry["duration_minutes"], 21 * 60)
        self.assertEqual(entry["duration_label"], "21 hours")
        self.assertEqual(entry["daily_segments"], [
            {"date": "2026-09-30", "start": "06:00", "finish": "00:00", "duration_minutes": 18 * 60, "duration_label": "18 hours"},
            {"date": "2026-10-01", "start": "00:00", "finish": "03:00", "duration_minutes": 3 * 60, "duration_label": "3 hours"},
        ])
        reopened = self.client.get("/api/timesheets").get_json()["entries"][0]
        self.assertEqual(reopened["finish_date"], "2026-10-01")
        self.assertEqual(reopened["duration_minutes"], 21 * 60)
        self.assertEqual(reopened["daily_segments"], entry["daily_segments"])

        september = app.timesheet_month_report("2026-09")
        october = app.timesheet_month_report("2026-10")
        self.assertEqual(september["total_minutes"], 18 * 60)
        self.assertEqual(september["rows"][0]["start"], "06:00")
        self.assertEqual(september["rows"][0]["finish"], "00:00")
        self.assertEqual(october["total_minutes"], 3 * 60)
        self.assertEqual(october["rows"][0]["start"], "00:00")
        self.assertEqual(october["rows"][0]["finish"], "03:00")

        export = self.client.get("/timesheet/month/2026-10/export.xlsx")
        self.assertEqual(export.status_code, 200)
        export_path = os.path.join(self.directory.name, "monthly.xlsx")
        with open(export_path, "wb") as handle:
            handle.write(export.data)
        exported_rows = app.xlsx_first_sheet_rows(export_path)
        self.assertTrue(any("Job Details / Notes" in row for row in exported_rows))
        self.assertTrue(any("08:00 - Loaded straw" in row for row in exported_rows))
        self.assertTrue(any("3" in row for row in exported_rows))
        printable = self.client.get("/timesheet/month/2026-10/print")
        self.assertEqual(printable.status_code, 200)
        self.assertIn(b"3 hours", printable.data)
        self.assertIn(b"Company Hours", printable.data)
        self.assertEqual(october["company_totals"], [{
            "company": "Cherry Dene Farm Ltd.", "minutes": 180, "hours": "3", "duration_label": "3 hours",
        }])

    def test_complete_month_emails_and_records_completion(self):
        self.assertEqual(self.client.post("/api/timesheets", json=dict(self.entry, finish="17:00")).status_code, 201)
        with patch.object(app, "send_timesheet_month_email", return_value=["office@example.invalid"]) as sender:
            response = self.client.post("/timesheet/month/complete", data={"month": "2026-09"})
        self.assertEqual(response.status_code, 302)
        sender.assert_called_once()
        state = app.read_json_file(app.TIMESHEET_MONTH_STATE_PATH, {})
        self.assertEqual(state["owen curl|2026-09"]["sent_to"], ["office@example.invalid"])
        self.assertEqual(state["owen curl|2026-09"]["total_minutes"], 9 * 60)

    def test_month_email_clearly_lists_each_company_total(self):
        cherry = dict(self.entry, start="08:00", finish="12:00")
        farrell = dict(self.entry, company="A. Farrell Contracting Ltd.", start="12:00", finish="17:00")
        self.assertEqual(self.client.post("/api/timesheets", json=cherry).status_code, 201)
        self.assertEqual(self.client.post("/api/timesheets", json=farrell).status_code, 201)
        report = app.timesheet_month_report("2026-09")
        smtp = MagicMock()
        server = smtp.return_value.__enter__.return_value
        config = {"smtp_host": "smtp.example.invalid", "smtp_port": 587, "use_tls": True, "smtp_username": "", "smtp_password": "", "from_email": "sender@example.invalid", "subject_prefix": "A. Farrell Contracting"}
        with patch.object(app, "load_timesheet_email_recipients", return_value=["office@example.invalid"]), patch.object(app, "load_email_config", return_value=config), patch.object(app.smtplib, "SMTP", smtp), patch.object(app, "convert_xlsx_bytes_to_pdf_bytes", return_value=b""):
            app.send_timesheet_month_email(report)
        message = server.send_message.call_args.args[0]
        body = message.get_body(preferencelist=("plain",)).get_content()
        self.assertIn("A. Farrell Contracting Ltd.: 5 hours", body)
        self.assertIn("Cherry Dene Farm Ltd.: 4 hours", body)
        self.assertIn("Overall total: 9 hours", body)
        self.assertEqual(len(list(message.iter_attachments())), 1)

    def test_pin_unlock_and_staff_records_are_isolated(self):
        from settings_workbook import Workbook
        Workbook(app.SETTINGS_WORKBOOK_PATH).create({
            "Staff": [["name", "active", "pin"], ["Owen Curl", "1", "1234"], ["Alice Smith", "1", "5678"]],
            "Companies": [["name", "active"]] + [[company, "1"] for company in app.TIMESHEET_COMPANIES],
        })
        anonymous = app.app.test_client()
        self.assertEqual(anonymous.get("/api/timesheets").status_code, 401)
        self.assertEqual(anonymous.get("/timesheet/calendar").status_code, 302)
        selector = anonymous.get("/timesheet")
        self.assertIn(b"Owen Curl", selector.data)
        self.assertIn(b"Alice Smith", selector.data)
        self.assertEqual(anonymous.post("/timesheet/unlock", data={"staff_name": "Alice Smith", "pin": "0000"}).status_code, 403)
        unlocked = anonymous.post("/timesheet/unlock", data={"staff_name": "Alice Smith", "pin": "5678"})
        self.assertEqual(unlocked.status_code, 302)
        created = anonymous.post("/api/timesheets", json=dict(self.entry, name="Owen Curl"))
        self.assertEqual(created.status_code, 201)
        self.assertEqual(created.get_json()["entry"]["name"], "Alice Smith")
        self.assertEqual([row["name"] for row in anonymous.get("/api/timesheets").get_json()["entries"]], ["Alice Smith"])

        with self.client.session_transaction() as current:
            current["timesheet_staff"] = "Owen Curl"
        alice_entry = created.get_json()["entry"]
        update = self.client.put("/api/timesheets/%s" % alice_entry["id"], json=dict(alice_entry, notes="Should be blocked"))
        self.assertEqual(update.status_code, 409)
        self.assertEqual(self.client.get("/api/timesheets").get_json()["entries"], [])

    def test_staff_can_set_first_pin_and_change_it_later(self):
        from settings_workbook import Workbook
        book = Workbook(app.SETTINGS_WORKBOOK_PATH)
        book.create({
            "Staff": [["name", "active", "pin"], ["Owen Curl", "1", ""]],
            "Companies": [["name", "active"]] + [[company, "1"] for company in app.TIMESHEET_COMPANIES],
        })
        client = app.app.test_client()
        setup_page = client.get("/timesheet/unlock/Owen%20Curl")
        self.assertIn(b"Set Your PIN", setup_page.data)
        mismatch = client.post("/timesheet/setup-pin", data={"staff_name": "Owen Curl", "new_pin": "2468", "confirm_pin": "1357"})
        self.assertEqual(mismatch.status_code, 400)
        created = client.post("/timesheet/setup-pin", data={"staff_name": "Owen Curl", "new_pin": "2468", "confirm_pin": "2468"})
        self.assertEqual(created.status_code, 302)
        self.assertEqual(book.records("Staff")[0]["pin"], "2468")
        calendar = client.get("/timesheet/calendar")
        self.assertIn(b"Change PIN", calendar.data)

        wrong = client.post("/timesheet/change-pin", data={"current_pin": "0000", "new_pin": "8642", "confirm_pin": "8642"})
        self.assertEqual(wrong.status_code, 403)
        changed = client.post("/timesheet/change-pin", data={"current_pin": "2468", "new_pin": "8642", "confirm_pin": "8642"})
        self.assertEqual(changed.status_code, 302)
        self.assertEqual(book.records("Staff")[0]["pin"], "8642")
        client.get("/timesheet")
        self.assertEqual(client.post("/timesheet/unlock", data={"staff_name": "Owen Curl", "pin": "2468"}).status_code, 403)
        self.assertEqual(client.post("/timesheet/unlock", data={"staff_name": "Owen Curl", "pin": "8642"}).status_code, 302)


if __name__ == "__main__":
    unittest.main()
