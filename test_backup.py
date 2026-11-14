import io
import os
import sqlite3
import tempfile
import unittest
import zipfile
from datetime import datetime
from unittest.mock import patch

import app


class WholeAppBackupTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        root = self.directory.name
        self.data = os.path.join(root, "data")
        os.makedirs(os.path.join(self.data, "gritting"))
        os.makedirs(os.path.join(self.data, "job_issue_photos"))
        with open(os.path.join(self.data, "jobs.ndjson"), "w") as handle:
            handle.write('{"customer": "Acme"}\n')
        with open(os.path.join(self.data, "job_issue_photos", "p1.jpg"), "wb") as handle:
            handle.write(b"jpg")
        db = sqlite3.connect(os.path.join(self.data, "timesheets.sqlite3"))
        db.execute("CREATE TABLE t (name TEXT)")
        db.execute("INSERT INTO t VALUES ('Owen')")
        db.commit()
        db.close()
        self.patches = [
            patch.object(app, "DATA_DIR", self.data),
            patch.object(app, "ISSUE_PHOTOS_DIR", os.path.join(self.data, "job_issue_photos")),
            patch.object(app, "INVOICE_ARCHIVE_DIR", os.path.join(self.data, "invoices")),
            patch.object(app, "BACKUP_DIR", os.path.join(root, "backups")),
        ]
        for item in self.patches:
            item.start()
        self.client = app.app.test_client()

    def tearDown(self):
        for item in self.patches:
            item.stop()
        self.directory.cleanup()

    def test_one_backup_holds_every_part_of_the_app(self):
        names = zipfile.ZipFile(io.BytesIO(self.client.get("/backup/export.zip").data)).namelist()
        for expected in ("app.py", "gritting.py", "weather.py", "data/jobs.ndjson", "data/timesheets.sqlite3",
                         "data/job_issue_photos/p1.jpg", "templates/straw_app.html", "static/af_brand.css"):
            self.assertIn(expected, names)
        self.assertFalse(any(name.startswith("backups/") or "__pycache__" in name for name in names))

    def test_databases_are_copied_safely_and_still_open(self):
        archive = zipfile.ZipFile(io.BytesIO(self.client.get("/backup/export.zip").data))
        copy_path = os.path.join(self.directory.name, "copy.sqlite3")
        with open(copy_path, "wb") as handle:
            handle.write(archive.read("data/timesheets.sqlite3"))
        self.assertEqual(sqlite3.connect(copy_path).execute("SELECT name FROM t").fetchone(), ("Owen",))

    def test_the_daily_backup_runs_once_a_day_after_2am_and_keeps_the_newest_14(self):
        self.assertIsNone(app.maybe_run_daily_backup(datetime(2026, 11, 14, 1, 30)))
        self.assertIsNotNone(app.maybe_run_daily_backup(datetime(2026, 11, 14, 2, 5)))
        self.assertIsNone(app.maybe_run_daily_backup(datetime(2026, 11, 14, 9, 0)))  # already done today
        for day in range(15, 31):
            app.maybe_run_daily_backup(datetime(2026, 11, day, 3, 0))
        backups = app.list_backups()
        self.assertEqual(len(backups), 14)
        self.assertEqual(backups[0]["made_at"].date(), datetime(2026, 11, 30).date())

    def test_settings_has_a_back_up_now_button_that_saves_a_backup(self):
        page = self.client.get("/settings").data
        self.assertIn(b"Back up now", page)
        response = self.client.post("/backup/now")
        self.assertEqual(response.status_code, 302)
        backups = app.list_backups()
        self.assertEqual(len(backups), 1)
        download = self.client.get("/backup/files/" + backups[0]["name"])
        self.assertEqual(download.status_code, 200)
        self.assertIn(backups[0]["label"].encode(), self.client.get("/settings").data)
        self.assertEqual(self.client.get("/backup/files/..%2Fapp.py").status_code, 404)

    def test_muck_spreading_buttons_and_saves_come_back_to_muck_spreading_not_the_app_home(self):
        for path in ("/invoice", "/invoice/history", "/admin"):
            page = self.client.get(path).get_data(as_text=True)
            self.assertIn('href="/muck">Back to muck spreading<', page, path)
        refused = self.client.post("/jobs/save", data={"customer": ""})  # a save that is refused still lands on muck spreading
        self.assertIn("/muck?", refused.headers["Location"])
        missing = self.client.post("/jobs/delete/999999")
        self.assertIn("/muck?", missing.headers["Location"])

    def test_the_parts_of_the_app_no_longer_have_their_own_backup_button(self):
        self.assertNotIn(b"Backup ZIP", self.client.get("/muck").data)


if __name__ == "__main__":
    unittest.main()
