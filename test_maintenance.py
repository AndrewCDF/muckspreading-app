import tempfile
import unittest
from unittest.mock import patch

import app


class MaintenanceTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.data_patch = patch.object(app, "DATA_DIR", self.directory.name)
        self.settings_patch = patch.object(app, "SETTINGS_WORKBOOK_PATH", self.directory.name + "/settings.xlsx")
        self.data_patch.start()
        self.settings_patch.start()
        self.client = app.app.test_client()
        self.record = {
            "maintenance_date": "2026-09-16",
            "machinery": "John Deere 6215R",
            "company": "A. Farrell Contracting Ltd.",
            "machine_hours": "3456.7",
            "work_completed": "Changed engine oil and filters",
            "parts_used": "Engine oil, oil filter, fuel filter",
            "cost": "245.50",
            "completed_by": "Owen Curl",
            "next_service_date": "2027-03-16",
            "notes": "Checked tyre pressures",
        }

    def tearDown(self):
        self.settings_patch.stop()
        self.data_patch.stop()
        self.directory.cleanup()

    def test_home_and_maintenance_page(self):
        home = self.client.get("/")
        self.assertIn(b'href="/maintenance">Maintenance</a>', home.data)
        page = self.client.get("/maintenance")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Machinery Maintenance", page.data)
        self.assertIn(b"Notes / Job Done", page.data)
        self.assertIn(b"select or type a new one", page.data)
        self.assertNotIn(b"Parts / Materials Used", page.data)
        self.assertNotIn(b"Cost \xc2\xa3", page.data)
        self.assertNotIn(b"Additional Notes", page.data)
        self.assertNotIn(b"Completed By", page.data)
        self.assertNotIn(b"Next Service Due", page.data)
        self.assertNotIn(b"Company / Farm", page.data)
        self.assertLess(page.data.index(b'id="save-record"'), page.data.index(b'id="new-record"'))
        for path in ["/static/maintenance.js", "/static/maintenance.css"]:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)

    def test_create_reopen_and_update_record(self):
        created = self.client.post("/api/maintenance", json=self.record)
        self.assertEqual(created.status_code, 201)
        saved = created.get_json()["entry"]
        self.assertEqual(saved["cost"], "245.5")
        rows = self.client.get("/api/maintenance").get_json()["entries"]
        self.assertEqual(rows, [saved])
        updated = dict(saved, work_completed="Changed oil, filters and alternator belt", cost="310")
        response = self.client.put("/api/maintenance/%s" % saved["id"], json=updated)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["entry"]["version"], 2)
        stale = self.client.put("/api/maintenance/%s" % saved["id"], json=saved)
        self.assertEqual(stale.status_code, 409)

    def test_invalid_records_are_rejected(self):
        changes = [
            {"maintenance_date": "2026-02-30"}, {"machinery": ""},
            {"work_completed": ""}, {"machine_hours": "-1"},
            {"cost": "not a number"}, {"next_service_date": "2027-02-30"}, {"notes": []},
        ]
        for change in changes:
            with self.subTest(change=change):
                self.assertEqual(self.client.post("/api/maintenance", json=dict(self.record, **change)).status_code, 400)
        self.assertEqual(self.client.get("/api/maintenance").get_json()["entries"], [])

    def test_new_machinery_is_added_to_editable_dropdown(self):
        app.Workbook(app.SETTINGS_WORKBOOK_PATH).create({
            "Companies": [["name", "active"], ["A. Farrell Contracting Ltd.", "1"], ["Cherry Dene Farm Ltd.", "1"]],
            "Staff": [["name", "active"], ["Owen Curl", "1"]],
            "Machinery": [["name", "active"], ["Existing Tractor", "1"]],
        })
        response = self.client.post("/api/maintenance", json=self.record)
        self.assertEqual(response.status_code, 201)
        self.assertEqual(app.settings_names("Machinery", []), ["Existing Tractor", "John Deere 6215R"])
        page = self.client.get("/maintenance")
        self.assertIn(b'value="Existing Tractor"', page.data)
        self.assertIn(b'value="John Deere 6215R"', page.data)


if __name__ == "__main__":
    unittest.main()
