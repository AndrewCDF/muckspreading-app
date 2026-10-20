import tempfile
import unittest
from unittest.mock import patch

import app


class BeetTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.data_patch = patch.object(app, "DATA_DIR", self.directory.name)
        self.settings_patch = patch.object(app, "SETTINGS_WORKBOOK_PATH", self.directory.name + "/settings.xlsx")
        self.data_patch.start()
        self.settings_patch.start()
        self.client = app.app.test_client()
        self.job = {"job_date": "2026-10-20", "customer": "Claxton Manor", "field": "Coleseeds", "area": "12.5", "diesel_used": "340", "diesel_price": "1.10"}

    def tearDown(self):
        self.settings_patch.stop()
        self.data_patch.stop()
        self.directory.cleanup()

    def test_home_link_and_page(self):
        self.assertIn(b"Sugar Beet</a>", self.client.get("/").data)
        page = self.client.get("/beet")
        self.assertEqual(page.status_code, 200)
        for text in (b"Diesel Used", b"Area (ha)", b"Customer", b"Field"):
            self.assertIn(text, page.data)
        for path in ["/static/beet.js", "/static/beet.css"]:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)

    def test_create_and_update_job(self):
        created = self.client.post("/api/beet", json=self.job)
        self.assertEqual(created.status_code, 201)
        saved = created.get_json()["entry"]
        self.assertEqual(saved["escalator"], "136.00")
        self.assertEqual(saved["base_price"], "0.7")
        self.assertEqual(saved["escalator_per_ha"], "10.88")
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"], [saved])
        response = self.client.put("/api/beet/%s" % saved["id"], json=dict(saved, diesel_used="355.5"))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["entry"]["version"], 2)
        self.assertEqual(response.get_json()["entry"]["escalator"], "142.20")
        self.assertEqual(self.client.put("/api/beet/%s" % saved["id"], json=saved).status_code, 409)

    def test_surcharge_only_after_trigger_rise(self):
        cases = [("0.75", "0.00"), ("0.769", "0.00"), ("0.77", "23.80"), ("1.10", "136.00")]
        for price, expected in cases:
            with self.subTest(price=price):
                saved = self.client.post("/api/beet", json=dict(self.job, diesel_price=price)).get_json()["entry"]
                self.assertEqual(saved["escalator"], expected)

    def test_settings_are_editable_and_snapshotted(self):
        self.assertEqual(self.client.get("/api/beet/settings").get_json()["settings"], {"base_price": 0.7, "trigger_pct": 10.0})
        old = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        saved = self.client.post("/api/beet/settings", json={"base_price": "0.80", "trigger_pct": "5"})
        self.assertEqual(saved.get_json()["settings"], {"base_price": 0.8, "trigger_pct": 5.0})
        new = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        self.assertEqual((new["base_price"], new["trigger_pct"], new["escalator"]), ("0.8", "5.0", "102.00"))
        rows = {row["id"]: row for row in self.client.get("/api/beet").get_json()["entries"]}
        self.assertEqual(rows[old["id"]]["escalator"], "136.00")
        for bad in ({"base_price": "0", "trigger_pct": "10"}, {"base_price": "x", "trigger_pct": "10"}, {"base_price": "0.7", "trigger_pct": "-1"}, {"base_price": "0.7"}):
            self.assertEqual(self.client.post("/api/beet/settings", json=bad).status_code, 400)

    def test_invalid_jobs_are_rejected(self):
        for change in [{"job_date": "2026-02-30"}, {"customer": ""}, {"field": ""}, {"area": "-1"}, {"area": ""}, {"diesel_used": "abc"}, {"diesel_price": ""}, {"diesel_price": "-1"}, {"field": []}]:
            with self.subTest(change=change):
                self.assertEqual(self.client.post("/api/beet", json=dict(self.job, **change)).status_code, 400)
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"], [])


if __name__ == "__main__":
    unittest.main()
