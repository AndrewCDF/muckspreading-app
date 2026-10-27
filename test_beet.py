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
        self.assertEqual(saved["escalator"], "112.20")  # 340 L x (£1.10 - 77p trigger)
        self.assertEqual(saved["base_price"], "0.7")
        self.assertEqual(saved["escalator_per_ha"], "8.98")
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"], [saved])
        response = self.client.put("/api/beet/%s" % saved["id"], json=dict(saved, diesel_used="355.5"))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["entry"]["version"], 2)
        self.assertEqual(response.get_json()["entry"]["escalator"], "117.32")
        self.assertEqual(self.client.put("/api/beet/%s" % saved["id"], json=saved).status_code, 409)

    def test_surcharge_only_after_trigger_rise(self):
        # only the part above 77p (base 70p + 10%) is charged; the first 10% is absorbed
        cases = [("0.75", "0.00"), ("0.769", "0.00"), ("0.77", "0.00"), ("0.771", "0.34"), ("1.10", "112.20")]
        for price, expected in cases:
            with self.subTest(price=price):
                saved = self.client.post("/api/beet", json=dict(self.job, diesel_price=price)).get_json()["entry"]
                self.assertEqual(saved["escalator"], expected)

    def test_half_penny_rounds_up_exactly(self):
        # 447 L x 25.5p above the trigger is exactly £113.985; binary floats gave £113.98
        saved = self.client.post("/api/beet", json=dict(self.job, diesel_used="447", diesel_price="1.025", area="10")).get_json()["entry"]
        self.assertEqual((saved["escalator"], saved["escalator_per_ha"]), ("113.99", "11.40"))
        self.assertEqual(str(app.beet_escalator("447", "1.025", "0.7", "10")), "113.99")

    def test_settings_are_editable_and_snapshotted(self):
        self.assertEqual(self.client.get("/api/beet/settings").get_json()["settings"], {"base_price": 0.7, "trigger_pct": 10.0})
        old = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        saved = self.client.post("/api/beet/settings", json={"base_price": "0.80", "trigger_pct": "5"})
        self.assertEqual(saved.get_json()["settings"], {"base_price": 0.8, "trigger_pct": 5.0})
        new = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        self.assertEqual((new["base_price"], new["trigger_pct"], new["escalator"]), ("0.8", "5.0", "88.40"))
        rows = {row["id"]: row for row in self.client.get("/api/beet").get_json()["entries"]}
        self.assertEqual(rows[old["id"]]["escalator"], "112.20")
        for bad in ({"base_price": "0", "trigger_pct": "10"}, {"base_price": "x", "trigger_pct": "10"}, {"base_price": "0.7", "trigger_pct": "-1"}, {"base_price": "0.7"}):
            self.assertEqual(self.client.post("/api/beet/settings", json=bad).status_code, 400)

    def test_office_adds_price_later(self):
        saved = self.client.post("/api/beet", json=dict(self.job, diesel_price="")).get_json()["entry"]
        self.assertEqual((saved["priced"], saved["escalator"], saved["escalator_per_ha"]), (False, "", ""))
        self.client.post("/api/beet/settings", json={"base_price": "0.80", "trigger_pct": "5"})
        reopened = self.client.get("/api/beet").get_json()["entries"][0]
        priced = self.client.put("/api/beet/%s" % saved["id"], json=dict(reopened, diesel_price="1.10"))
        self.assertEqual(priced.status_code, 200)
        entry = priced.get_json()["entry"]
        self.assertEqual((entry["priced"], entry["base_price"], entry["escalator"], entry["escalator_per_ha"]), (True, "0.8", "88.40", "7.07"))
        self.client.post("/api/beet/settings", json={"base_price": "0.50", "trigger_pct": "5"})
        again = self.client.put("/api/beet/%s" % saved["id"], json=dict(entry, diesel_price="1.20")).get_json()["entry"]
        self.assertEqual((again["base_price"], again["escalator"]), ("0.8", "122.40"))
        cleared = self.client.put("/api/beet/%s" % saved["id"], json=dict(again, diesel_price="")).get_json()["entry"]
        self.assertFalse(cleared["priced"])

    def test_invoiced_marker(self):
        unpriced = self.client.post("/api/beet", json=dict(self.job, diesel_price="")).get_json()["entry"]
        self.assertFalse(unpriced["invoiced"])
        blocked = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(unpriced, invoiced=True))
        self.assertEqual(blocked.status_code, 400)
        priced = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(unpriced, diesel_price="1.10")).get_json()["entry"]
        marked = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(priced, invoiced=True))
        entry = marked.get_json()["entry"]
        self.assertEqual((marked.status_code, entry["invoiced"]), (200, True))
        self.assertRegex(entry["invoiced_date"], r"^\d{4}-\d{2}-\d{2}$")
        kept = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(entry, area="13", invoiced=True)).get_json()["entry"]
        self.assertEqual(kept["invoiced_date"], entry["invoiced_date"])
        undone = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(kept, invoiced=False)).get_json()["entry"]
        self.assertEqual((undone["invoiced"], undone["invoiced_date"]), (False, ""))
        self.assertEqual(self.client.post("/api/beet", json=dict(self.job, invoiced="yes")).status_code, 400)
        created = self.client.post("/api/beet", json=dict(self.job, invoiced=True)).get_json()["entry"]
        self.assertTrue(created["invoiced"])

    def test_field_and_area_can_be_left_blank(self):
        saved = self.client.post("/api/beet", json=dict(self.job, field="", area="")).get_json()["entry"]
        self.assertEqual((saved["field"], saved["area"], saved["escalator"], saved["escalator_per_ha"]), ("", "", "112.20", ""))
        reopened = self.client.put("/api/beet/%s" % saved["id"], json=dict(saved, fields=["Three fields"], area="30"))
        self.assertEqual(reopened.get_json()["entry"]["escalator_per_ha"], "3.74")
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"][0]["field"], "Three fields")
        self.assertEqual(self.client.get("/beet").status_code, 200)

    def test_job_covers_a_list_of_fields(self):
        saved = self.client.post("/api/beet", json=dict(self.job, field="", fields=["Coleseeds", "", "Top Meadow"], area="30")).get_json()["entry"]
        self.assertEqual(saved["fields"], ["Coleseeds", "", "Top Meadow"])
        self.assertEqual(saved["field"], "Coleseeds, Field 2, Top Meadow")
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"][0]["fields"], ["Coleseeds", "", "Top Meadow"])
        blank = self.client.post("/api/beet", json=dict(self.job, fields=[""])).get_json()["entry"]
        self.assertEqual((blank["fields"], blank["field"]), ([], ""))
        numbered = self.client.post("/api/beet", json=dict(self.job, fields=["", "", ""])).get_json()["entry"]
        self.assertEqual(numbered["field"], "Field 1, Field 2, Field 3")
        for bad in (["x"] * 51, "Coleseeds", [1], ["x" * 161]):
            self.assertEqual(self.client.post("/api/beet", json=dict(self.job, fields=bad)).status_code, 400)
        names = self.client.get("/beet").data
        self.assertIn(b"Top Meadow", names)

    def test_tractor_diesel_is_added_to_harvester_diesel(self):
        saved = self.client.post("/api/beet", json=dict(self.job, tractor_diesel="160", area="30")).get_json()["entry"]
        self.assertEqual((saved["tractor_diesel"], saved["total_diesel"], saved["escalator"], saved["escalator_per_ha"]), ("160", "500", "165.00", "5.50"))
        none = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        self.assertEqual((none["tractor_diesel"], none["total_diesel"], none["escalator"]), ("", "340", "112.20"))
        below = self.client.post("/api/beet", json=dict(self.job, tractor_diesel="160", diesel_price="0.75")).get_json()["entry"]
        self.assertEqual(below["escalator"], "0.00")
        updated = self.client.put("/api/beet/%s" % saved["id"], json=dict(saved, tractor_diesel="")).get_json()["entry"]
        self.assertEqual((updated["tractor_diesel"], updated["total_diesel"]), ("", "340"))
        for bad in ("-5", "abc"):
            self.assertEqual(self.client.post("/api/beet", json=dict(self.job, tractor_diesel=bad)).status_code, 400)

    def test_figures_are_stored_when_the_price_is_entered(self):
        saved = self.client.post("/api/beet", json=self.job).get_json()["entry"]
        self.assertEqual((saved["escalator"], saved["escalator_per_ha"]), ("112.20", "8.98"))
        original = app.beet_escalator
        app.beet_escalator = lambda *args: app.Decimal("999.00")  # as if the rule or settings changed later
        try:
            self.assertEqual(self.client.get("/api/beet").get_json()["entries"][0]["escalator"], "112.20")
            marked = self.client.put("/api/beet/%s" % saved["id"], json=dict(saved, invoiced=True)).get_json()["entry"]
            self.assertEqual((marked["invoiced"], marked["escalator"]), (True, "112.20"))
            edited = self.client.put("/api/beet/%s" % saved["id"], json=dict(marked, diesel_used="400")).get_json()["entry"]
            self.assertEqual(edited["escalator"], "999.00")  # a changed figure recalculates and stores again
        finally:
            app.beet_escalator = original
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"][0]["escalator"], "999.00")
        unpriced = self.client.post("/api/beet", json=dict(self.job, diesel_price="")).get_json()["entry"]
        self.assertEqual(unpriced["escalator"], "")
        priced = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(unpriced, diesel_price="1.10")).get_json()["entry"]
        self.assertEqual(priced["escalator"], "112.20")
        cleared = self.client.put("/api/beet/%s" % unpriced["id"], json=dict(priced, diesel_price="")).get_json()["entry"]
        self.assertEqual((cleared["escalator"], cleared["escalator_per_ha"]), ("", ""))

    def test_existing_priced_jobs_are_backfilled_once(self):
        import sqlite3
        path = self.directory.name + "/beet.sqlite3"
        old = sqlite3.connect(path)
        old.execute("""CREATE TABLE beet_jobs (id INTEGER PRIMARY KEY, job_date TEXT NOT NULL, customer TEXT NOT NULL, field TEXT NOT NULL,
            area TEXT NOT NULL, diesel_used TEXT NOT NULL, version INTEGER NOT NULL, updated_at TEXT NOT NULL,
            diesel_price TEXT NOT NULL DEFAULT '1.1', base_price TEXT NOT NULL DEFAULT '0.7', trigger_pct TEXT NOT NULL DEFAULT '10.0',
            tractor_diesel TEXT NOT NULL DEFAULT '', fields_json TEXT NOT NULL DEFAULT '', invoiced_date TEXT NOT NULL DEFAULT '')""")
        old.execute("INSERT INTO beet_jobs (job_date, customer, field, area, diesel_used, version, updated_at, diesel_price) VALUES ('2026-10-01', 'Claxton Manor', 'A', '12.5', '340', 1, 'x', '1.10')")
        old.execute("INSERT INTO beet_jobs (job_date, customer, field, area, diesel_used, version, updated_at, diesel_price) VALUES ('2026-10-02', 'Claxton Manor', 'B', '5', '100', 1, 'x', '')")
        old.commit()
        old.close()
        rows = {row["field"]: row for row in self.client.get("/api/beet").get_json()["entries"]}
        self.assertEqual((rows["A"]["escalator"], rows["A"]["escalator_per_ha"]), ("112.20", "8.98"))
        self.assertEqual(rows["B"]["escalator"], "")
        connection = app.beet_connection()
        stored = connection.execute("SELECT surcharge_total FROM beet_jobs WHERE field='A'").fetchone()[0]
        connection.close()
        self.assertEqual(stored, "112.20")

    def test_old_jobs_with_one_field_still_load(self):
        connection = app.beet_connection()
        with connection:
            connection.execute("INSERT INTO beet_jobs (job_date, customer, field, area, diesel_used, diesel_price, base_price, trigger_pct, invoiced_date, version, updated_at) VALUES ('2026-10-01', 'Claxton Manor', 'Coleseeds', '5', '100', '', '0.7', '10.0', '', 1, 'x')")
        connection.close()
        entry = self.client.get("/api/beet").get_json()["entries"][0]
        self.assertEqual((entry["fields"], entry["field"]), (["Coleseeds"], "Coleseeds"))

    def test_invalid_jobs_are_rejected(self):
        for change in [{"job_date": "2026-02-30"}, {"customer": ""}, {"area": "-1"}, {"diesel_used": "abc"}, {"diesel_price": "-1"}, {"field": []}]:
            with self.subTest(change=change):
                self.assertEqual(self.client.post("/api/beet", json=dict(self.job, **change)).status_code, 400)
        self.assertEqual(self.client.get("/api/beet").get_json()["entries"], [])


if __name__ == "__main__":
    unittest.main()
