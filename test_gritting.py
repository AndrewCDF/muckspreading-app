import re
import tempfile
import unittest
from datetime import datetime
from unittest.mock import patch

import app
import gritting


class GrittingModuleTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.original = (gritting.DATA_ROOT, gritting.DATABASE_PATH, gritting.BUNDLED_CUSTOMER_XLSX, gritting.SEASON_ARCHIVE_DIR)
        gritting.configure(self.directory.name + "/gritting")
        gritting._initialised_database = None
        self.client = app.app.test_client()

    def tearDown(self):
        gritting.DATA_ROOT, gritting.DATABASE_PATH, gritting.BUNDLED_CUSTOMER_XLSX, gritting.SEASON_ARCHIVE_DIR = self.original
        gritting._initialised_database = None
        self.directory.cleanup()

    def test_home_links_to_gritting_and_gritting_links_home(self):
        self.assertIn(b'href="/gritting/">Gritting</a>', self.client.get("/").data)
        page = self.client.get("/gritting/")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b'aria-label="Home: back to the main app"', page.data)
        self.assertIn(b'aria-label="Back" title="Back" onclick="if (window.history.length &gt; 1)', page.data.replace(b"> 1)", b"&gt; 1)"))
        self.assertLess(page.data.index(b'aria-label="Back"'), page.data.index(b'aria-label="Home'))  # Back sits beside Home, before it
        self.assertIn(b"/static/gritting/gritting.css", page.data)
        self.assertIn(b"/gritting/live-jobs", page.data)

    def test_every_page_loads_with_an_empty_database(self):
        for path in ["/gritting/", "/gritting/live-jobs", "/gritting/new-callout", "/gritting/customers", "/gritting/settings",
                     "/gritting/callouts", "/gritting/monthly-summary", "/gritting/email-signature-preview",
                     "/gritting/monthly-summary/export?month=2026-01"]:
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 200)
                if "export" not in path:  # every page has Back beside Home
                    self.assertIn(b'aria-label="Back"', response.data)
                    self.assertIn(b'window.history.back()', response.data)
        sites = self.client.get("/gritting/sites")  # the old Sites page now forwards to Customers & Sites
        self.assertEqual((sites.status_code, sites.headers["Location"].endswith("/gritting/customers")), (302, True))
        for path in ["/static/gritting/gritting.css", "/static/gritting/brand/ngs-logo.png"]:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)

    def test_data_lives_under_the_configured_folder(self):
        self.client.get("/gritting/")
        self.assertTrue((gritting.DATABASE_PATH).exists())
        self.assertTrue(str(gritting.DATABASE_PATH).startswith(self.directory.name))

    def test_customer_site_callout_and_completion_flow(self):
        added = self.client.post("/gritting/customers", data={"name": "Acme Ltd", "contact_name": "Pat", "email": "pat@acme.test", "emergency_phone": "0123"})
        self.assertEqual(added.status_code, 302)
        self.assertTrue(added.headers["Location"].endswith("/gritting/customers"))
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer_id = db.execute("SELECT id FROM customers WHERE name='Acme Ltd'").fetchone()[0]
        db.close()
        site = self.client.post("/gritting/customers", data={
            "form_type": "site", "customer_id": customer_id, "name": "Yard", "address": "1 High St", "route_notes": "",
            "service_day_monday": "on", "pricing_label_1": "Weekday Night", "pricing_amount_1": "145",
        })
        self.assertEqual(site.status_code, 302)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        site_id = db.execute("SELECT id FROM sites WHERE name='Yard'").fetchone()[0]
        db.close()
        created = self.client.post("/gritting/new-callout", data={
            "site_id": site_id, "called_out_at": "2026-01-12T18:00", "weather_conditions": "Frost", "service_notes": "",
            "planned_attendance": "After midnight", "recipient": "pat@acme.test",
        })
        self.assertEqual(created.status_code, 302)
        detail_url = created.headers["Location"]
        self.assertRegex(detail_url, r"/gritting/callouts/\d+$")
        self.assertEqual(self.client.get(detail_url).status_code, 200)
        call_out_id = int(detail_url.rsplit("/", 1)[1])
        done = self.client.post("/gritting/callouts/%d/complete" % call_out_id, data={"completed_at": "2026-01-13T04:00", "service_notes": "Gritted"})
        self.assertEqual(done.status_code, 302)
        self.assertTrue(done.headers["Location"].endswith(detail_url[detail_url.index("/gritting"):]))
        summary = self.client.get("/gritting/monthly-summary?month=2026-01")
        self.assertEqual(summary.status_code, 200)
        self.assertIn(b"Acme Ltd", summary.data)
        export = self.client.get("/gritting/monthly-summary/export?month=2026-01")
        self.assertEqual(export.status_code, 200)
        self.assertIn(b"Acme Ltd", export.data)

    def test_backup_includes_gritting_data(self):
        import io
        import zipfile
        self.client.get("/gritting/")  # creates the database
        with patch.object(app, "DATA_DIR", self.directory.name):
            with patch.object(app, "ISSUE_PHOTOS_DIR", self.directory.name + "/photos"), patch.object(app, "INVOICE_ARCHIVE_DIR", self.directory.name + "/invoices"):
                backup = self.client.get("/backup/export.zip")
        names = zipfile.ZipFile(io.BytesIO(backup.data)).namelist()
        self.assertIn("gritting.py", names)
        self.assertIn("data/gritting/gritting.db", names)

    def make_sites(self, names=("Main yard", "Loading bay")):
        self.client.post("/gritting/customers", data={"name": "Acme Ltd", "contact_name": "Pat", "email": "pat@acme.test", "emergency_phone": "0123"})
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer_id = db.execute("SELECT id FROM customers WHERE name='Acme Ltd'").fetchone()[0]
        db.close()
        for name in names:
            self.client.post("/gritting/customers", data={
                "form_type": "site", "customer_id": customer_id, "name": name, "address": "1 High St", "route_notes": "",
                "service_day_monday": "on", "service_day_tuesday": "on", "service_day_wednesday": "on", "service_day_thursday": "on",
                "service_day_friday": "on", "service_day_saturday": "on", "service_day_sunday": "on",
                "pricing_label_1": "Weekday Night", "pricing_amount_1": "145",
            })
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        ids = [row[0] for row in db.execute("SELECT id FROM sites ORDER BY id")]
        db.close()
        return ids

    def bulk_callout(self, site_ids):
        data = {"called_out_at": "2026-01-12T18:00", "planned_attendance_from": "22:00", "planned_attendance_to": "04:00",
                "weather_conditions": "Frost", "service_notes": "", "subject": "", "message": "",
                "subject_override_enabled": "0", "message_override_enabled": "0", "selected_site_ids": [str(site_id) for site_id in site_ids]}
        for site_id in site_ids:
            data["pricing_label_%d" % site_id] = "Weekday Night"
        return self.client.post("/gritting/callouts/bulk", data=data)

    def test_bulk_callout_puts_every_job_straight_into_live_jobs(self):
        site_ids = self.make_sites()
        response = self.bulk_callout(site_ids)
        self.assertEqual(response.status_code, 302)
        live = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertEqual(live.count('class="job-card"'), 2)
        self.assertIn("Live jobs (2)", live)

    def test_cancelling_a_job_removes_it_from_live_and_the_totals(self):
        site_ids = self.make_sites()
        self.bulk_callout(site_ids)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        first, second = [row[0] for row in db.execute("SELECT id FROM call_outs ORDER BY id")]
        db.close()
        summary = lambda: self.client.get("/gritting/monthly-summary?month=2026-01").get_data(as_text=True)
        self.assertIn("£290.00", summary())  # both jobs carry the £145 weekday price
        cancelled = self.client.post("/gritting/callouts/%d/cancel" % first, data={"next": "live_jobs"})
        self.assertEqual(cancelled.status_code, 302)
        self.assertTrue(cancelled.headers["Location"].endswith("/gritting/live-jobs"))
        live = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertEqual(live.count('class="job-card"'), 1)
        self.assertIn("Live jobs (1)", live)
        self.assertNotIn("£290.00", summary())
        self.assertIn("£145.00", summary())
        self.assertIn("cancelled", summary())
        detail = self.client.get("/gritting/callouts/%d" % first).get_data(as_text=True)
        self.assertIn("did not need us to attend", detail)
        self.assertIn("Reopen this job", detail)
        # a cancelled job cannot be completed until it is reopened
        blocked = self.client.post("/gritting/callouts/%d/complete" % first, data={"completed_at": "2026-01-13T04:00", "service_notes": ""}, follow_redirects=True)
        self.assertIn("This job was cancelled", blocked.get_data(as_text=True))
        self.assertIn("This job was cancelled", self.client.post("/gritting/callouts/%d/complete-quick" % first, follow_redirects=True).get_data(as_text=True))
        # a completed job cannot be cancelled (completing needs the list finalised first)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        batch_id = db.execute("SELECT id FROM dispatch_batches").fetchone()[0]
        db.close()
        self.client.post("/gritting/batches/%d/finalise" % batch_id)
        self.client.post("/gritting/callouts/%d/complete-quick" % second)
        self.assertIn("A completed job cannot be cancelled", self.client.post("/gritting/callouts/%d/cancel" % second, follow_redirects=True).get_data(as_text=True))
        reopened = self.client.post("/gritting/callouts/%d/reopen" % first, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Job reopened", reopened)
        self.assertIn("Live jobs (1)", self.client.get("/gritting/live-jobs").get_data(as_text=True))

    def test_dashboard_live_count_ignores_cancelled_jobs(self):
        site_ids = self.make_sites(("Only site",))
        self.bulk_callout(site_ids)
        self.assertIn("Live jobs (1)", self.client.get("/gritting/").get_data(as_text=True))
        self.client.post("/gritting/callouts/1/cancel")
        dashboard = self.client.get("/gritting/").get_data(as_text=True)
        self.assertNotIn("Live jobs (", dashboard)
        self.assertIn("status-cancelled", self.client.get("/gritting/callouts").get_data(as_text=True))

    def test_bulk_call_out_lands_on_the_nights_list(self):
        site_ids = self.make_sites(("Main yard", "Loading bay", "Back gate"))
        response = self.bulk_callout(site_ids)
        self.assertTrue(response.headers["Location"].endswith("/gritting/live-jobs"))
        page = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn("Call-out list", page)
        self.assertIn(">Draft<", page)
        self.assertIn("<strong>3</strong> to grit · 0 not needed · 0 done", page)
        self.assertIn("Finalise the list (3 to grit)", page)
        self.assertIn("Draft list: you can open a job or mark it not needed.", page)
        self.assertIn("Not needed", page)
        self.assertNotIn("Mark complete", page)  # nothing can be completed until the list is finalised

    def test_cancel_then_finalise_then_print_the_final_list(self):
        site_ids = self.make_sites(("Main yard", "Loading bay", "Back gate"))
        self.bulk_callout(site_ids)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        call_outs = [row[0] for row in db.execute("SELECT id FROM call_outs ORDER BY id")]
        batch_id = db.execute("SELECT id FROM dispatch_batches").fetchone()[0]
        db.close()
        # the customer for the second site says they do not need us tonight
        cancelled = self.client.post("/gritting/callouts/%d/cancel" % call_outs[1], data={"next": "live_jobs"})
        self.assertTrue(cancelled.headers["Location"].endswith("/gritting/live-jobs"))
        page = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn("<strong>2</strong> to grit · 1 not needed · 0 done", page)
        self.assertIn("Put back on the list", page)
        self.assertIn("Finalise the list (2 to grit)", page)
        # finalise it
        finalised = self.client.post("/gritting/batches/%d/finalise" % batch_id, follow_redirects=True).get_data(as_text=True)
        self.assertIn("List finalised: 2 sites to grit, 1 not needed.", finalised)
        self.assertIn("Final ·", finalised)
        self.assertIn("Open the final list to print", finalised)
        self.assertNotIn("Draft list: you can open a job", finalised)
        self.assertIn("Mark complete", finalised)
        # once the list is final there is no Not needed button on it, nor on its jobs' pages
        self.assertNotIn("Not needed</button>", finalised)
        self.assertNotIn("Customer doesn&#39;t need us", self.client.get("/gritting/callouts/%d" % call_outs[0]).get_data(as_text=True).replace("doesn't", "doesn&#39;t"))
        # the final list: sites to grit, and who was cancelled
        sheet = self.client.get("/gritting/batches/%d/run-sheet" % batch_id).get_data(as_text=True)
        self.assertIn("Final list", sheet)
        self.assertIn("To grit (2)", sheet)
        self.assertIn("Not needed (1)", sheet)
        to_grit_part = sheet.split("Not needed (1)")[0]
        self.assertIn("Acme Ltd · Main yard", to_grit_part)
        self.assertIn("Acme Ltd · Back gate", to_grit_part)
        self.assertNotIn("Loading bay", to_grit_part)
        self.assertIn("Acme Ltd · Loading bay", sheet.split("Not needed (1)")[1])
        # a late cancellation after finalising says to reprint
        late = self.client.post("/gritting/callouts/%d/cancel" % call_outs[2], data={"next": "live_jobs"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("This list was already finalised, so print the final list again.", late)
        # reopening the list and putting a site back
        self.client.post("/gritting/batches/%d/reopen" % batch_id)
        draft = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn(">Draft<", draft)
        back = self.client.post("/gritting/callouts/%d/reopen" % call_outs[1], data={"next": "live_jobs"})
        self.assertTrue(back.headers["Location"].endswith("/gritting/live-jobs"))
        self.assertIn("<strong>2</strong> to grit · 1 not needed · 0 done", self.client.get("/gritting/live-jobs").get_data(as_text=True))

    def test_a_list_with_nothing_left_cannot_be_finalised_and_finished_lists_drop_off(self):
        site_ids = self.make_sites(("Main yard", "Loading bay"))
        self.bulk_callout(site_ids)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        call_outs = [row[0] for row in db.execute("SELECT id FROM call_outs ORDER BY id")]
        batch_id = db.execute("SELECT id FROM dispatch_batches").fetchone()[0]
        db.close()
        for call_out_id in call_outs:
            self.client.post("/gritting/callouts/%d/cancel" % call_out_id)
        refused = self.client.post("/gritting/batches/%d/finalise" % batch_id, follow_redirects=True).get_data(as_text=True)
        self.assertNotIn("List finalised", refused)
        self.assertIn("No live jobs right now.", refused)  # every site is cancelled, so the list has gone from Live
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        self.assertIsNone(db.execute("SELECT finalised_at FROM dispatch_batches").fetchone()[0])
        db.close()
        self.assertEqual(self.client.post("/gritting/batches/999/finalise").status_code, 302)
        self.assertEqual(self.client.get("/gritting/batches/999/run-sheet").status_code, 302)

    def test_jobs_on_a_draft_list_cannot_be_completed_until_it_is_finalised(self):
        site_ids = self.make_sites(("Main yard", "Loading bay"))
        self.bulk_callout(site_ids)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        first, second = [row[0] for row in db.execute("SELECT id FROM call_outs ORDER BY id")]
        batch_id = db.execute("SELECT id FROM dispatch_batches").fetchone()[0]
        db.close()
        status = lambda call_out_id: gritting.sqlite3.connect(gritting.DATABASE_PATH).execute("SELECT status FROM call_outs WHERE id=?", (call_out_id,)).fetchone()[0]
        draft = gritting.DRAFT_LIST_MESSAGE
        # the live list offers Open and Not needed only
        page = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertNotIn("Mark complete", page)
        self.assertEqual(page.count("Not needed</button>"), 2)
        # the quick button and the full form are both refused on the server
        quick = self.client.post("/gritting/callouts/%d/complete-quick" % first, data={"next": "live_jobs"}, follow_redirects=True)
        self.assertIn(draft, quick.get_data(as_text=True))
        form = self.client.post("/gritting/callouts/%d/complete" % first, data={"completed_at": "2026-01-13T04:00", "service_notes": "x"}, follow_redirects=True)
        self.assertIn(draft, form.get_data(as_text=True))
        self.assertEqual(status(first), "called_out")
        # the job page explains instead of offering the completion form, but cancelling is still there
        detail = self.client.get("/gritting/callouts/%d" % first).get_data(as_text=True)
        self.assertIn("On the draft list", detail)
        self.assertIn("cannot be completed yet", detail)
        self.assertNotIn("Complete job and email confirmation", detail)
        self.assertIn("Customer doesn&#39;t need us", detail.replace("doesn't", "doesn&#39;t"))
        # cancelling still works on a draft
        self.client.post("/gritting/callouts/%d/cancel" % second, data={"next": "live_jobs"})
        self.assertEqual(status(second), "cancelled")
        # finalise: now the list is the one to grit from and jobs can be completed through the night
        self.client.post("/gritting/batches/%d/finalise" % batch_id)
        final = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn("Mark complete", final)
        self.assertIn("Complete job and email confirmation", self.client.get("/gritting/callouts/%d" % first).get_data(as_text=True))
        done = self.client.post("/gritting/callouts/%d/complete-quick" % first, data={"next": "live_jobs"})
        self.assertEqual(done.status_code, 302)
        self.assertEqual(status(first), "completed")
        # once a job is done the list is being worked, so it cannot go back to draft
        self.client.post("/gritting/callouts/%d/reopen" % second)
        refused = self.client.post("/gritting/batches/%d/reopen" % batch_id, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Some jobs on this list are already done, so it cannot be reopened.", refused)
        self.assertNotIn("Reopen the list", refused)
        self.assertNotIn(">Draft<", refused)
        self.assertEqual(status(second), "called_out")
        self.client.post("/gritting/callouts/%d/complete-quick" % second)
        self.assertEqual(status(second), "completed")

    def test_lists_already_being_worked_are_final_when_drafts_are_introduced(self):
        import sqlite3
        path = self.directory.name + "/gritting/gritting.db"
        gritting.DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
        old = sqlite3.connect(path)
        old.executescript("""
            CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT);
            CREATE TABLE dispatch_batches (id INTEGER PRIMARY KEY AUTOINCREMENT, batch_key TEXT NOT NULL, called_out_at TEXT NOT NULL,
                weather_conditions TEXT NOT NULL, planned_attendance TEXT NOT NULL, service_notes TEXT, email_subject TEXT NOT NULL, created_at TEXT NOT NULL);
            CREATE TABLE call_outs (id INTEGER PRIMARY KEY AUTOINCREMENT, site_id INTEGER, dispatch_batch_id INTEGER, called_out_at TEXT, status TEXT);
            INSERT INTO dispatch_batches (batch_key, called_out_at, weather_conditions, planned_attendance, email_subject, created_at) VALUES ('a', '2026-01-01T18:00', 'f', 'x', 's', '2026-01-01T18:01');
            INSERT INTO dispatch_batches (batch_key, called_out_at, weather_conditions, planned_attendance, email_subject, created_at) VALUES ('b', '2026-01-02T18:00', 'f', 'x', 's', '2026-01-02T18:01');
            INSERT INTO call_outs (site_id, dispatch_batch_id, called_out_at, status) VALUES (1, 1, '2026-01-01T18:00', 'completed');
            INSERT INTO call_outs (site_id, dispatch_batch_id, called_out_at, status) VALUES (1, 2, '2026-01-02T18:00', 'called_out');
        """)
        old.close()
        gritting._initialised_database = None
        gritting.ensure_initialised()
        new = sqlite3.connect(path)
        finalised = dict(new.execute("SELECT id, finalised_at FROM dispatch_batches").fetchall())
        new.close()
        self.assertEqual(finalised[1], "2026-01-01T18:01")  # work had started: final
        self.assertIsNone(finalised[2])  # nothing done yet: a draft to finalise

    def test_confirmations_are_on_the_page_not_browser_pop_ups(self):
        # browser confirm() boxes are blocked in some phone browsers and made Finalise do nothing
        site_ids = self.make_sites(("Main yard", "Loading bay"))
        self.bulk_callout(site_ids)
        live = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertNotIn("confirm(", live.replace("data-confirm", ""))
        self.assertIn('data-confirm="Finalise this list? 2 to grit, 0 not needed."', live)
        self.assertEqual(live.count('data-confirm="Not needed tonight?'), 2)
        detail = self.client.get("/gritting/callouts/1").get_data(as_text=True)
        self.assertIn("data-confirm=", detail)
        self.assertNotIn("onsubmit", detail)
        self.assertIn('className = "inline-confirm"', detail)  # the script that draws the question lives in the page frame

    def test_single_call_outs_are_listed_separately_without_finalise(self):
        site_ids = self.make_sites(("Main yard",))
        created = self.client.post("/gritting/new-callout", data={
            "site_id": site_ids[0], "called_out_at": "2026-01-12T18:00", "weather_conditions": "Frost", "service_notes": "",
            "planned_attendance": "After midnight", "recipient": "pat@acme.test"})
        self.assertEqual(created.status_code, 302)
        page = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn("Single call-outs", page)
        self.assertNotIn("Finalise the list", page)
        self.assertIn("Mark complete", page)

    def test_one_form_adds_a_customer_together_with_a_site(self):
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertEqual(page.count("＋ Add a customer and site"), 1)
        self.assertNotIn("＋ Add a customer</span>", page)
        self.assertNotIn("＋ Add a site</span>", page)
        self.assertIn('name="form_type" value="customer_site"', page)
        form = {
            "form_type": "customer_site", "customer_choice": "new", "customer_name": "Beta Ltd", "contact_name": "Sam", "email": "SAM@beta.test",
            "emergency_phone": "0777", "site_name": "Depot", "address": "2 Low Rd", "route_notes": "Gate on the left",
            "service_day_monday": "on", "service_day_friday": "on", "pricing_label_1": "Weekday Night", "pricing_amount_1": "120",
            "pricing_label_2": "Friday Night", "pricing_amount_2": "140",
        }
        added = self.client.post("/gritting/customers", data=form, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Customer Beta Ltd and site Depot added.", added)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer = db.execute("SELECT id, contact_name, email, emergency_phone FROM customers WHERE name='Beta Ltd'").fetchone()
        site = db.execute("SELECT customer_id, name, address, route_notes, service_days, pricing_options FROM sites WHERE name='Depot'").fetchone()
        db.close()
        self.assertEqual(customer[1:], ("Sam", "sam@beta.test", "0777"))
        self.assertEqual((site[0], site[1], site[2], site[3]), (customer[0], "Depot", "2 Low Rd", "Gate on the left"))
        self.assertEqual(gritting.load_service_days(site[4]), ["monday", "friday"])
        self.assertEqual([(option["label"], option["amount"]) for option in gritting.load_pricing_options(site[5])], [("Weekday Night", 120.0), ("Friday Night", 140.0)])
        # the list offers the customer for another site, and a site for an existing customer reuses them
        self.assertIn("Add a site to Beta Ltd", self.client.get("/gritting/customers").get_data(as_text=True))
        more = dict(form, customer_choice=str(customer[0]), site_name="Yard", address="3 Low Rd")
        for key in ("customer_name", "contact_name", "email", "emergency_phone"):
            more.pop(key)
        self.assertIn("Beta Ltd: site Yard added.", self.client.post("/gritting/customers", data=more, follow_redirects=True).get_data(as_text=True))
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        self.assertEqual(db.execute("SELECT COUNT(*) FROM customers WHERE name='Beta Ltd'").fetchone()[0], 1)
        self.assertEqual(db.execute("SELECT COUNT(*) FROM sites WHERE customer_id=?", (customer[0],)).fetchone()[0], 2)
        yard_days = db.execute("SELECT service_days FROM sites WHERE name='Yard'").fetchone()[0]
        db.close()
        self.assertEqual(gritting.load_service_days(yard_days), ["monday", "friday"])

    def test_the_combined_form_needs_the_essentials_and_leaves_nothing_half_made(self):
        base = {"form_type": "customer_site", "customer_choice": "new", "customer_name": "Gamma Ltd", "contact_name": "Lee", "email": "lee@g.test",
                "site_name": "Yard", "address": "4 Mill Rd"}
        for change, message in [({"site_name": ""}, "Enter the site name and address."), ({"address": " "}, "Enter the site name and address."),
                                ({"customer_name": ""}, "Enter the customer name and contact name."), ({"contact_name": ""}, "Enter the customer name and contact name."),
                                ({"customer_choice": "999"}, "That customer could not be found.")]:
            with self.subTest(change=change):
                response = self.client.post("/gritting/customers", data=dict(base, **change), follow_redirects=True).get_data(as_text=True)
                self.assertIn(message, response)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        self.assertEqual((db.execute("SELECT COUNT(*) FROM customers").fetchone()[0], db.execute("SELECT COUNT(*) FROM sites").fetchone()[0]), (0, 0))
        db.close()
        # no days ticked means every day, as before
        self.client.post("/gritting/customers", data=base)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        days = db.execute("SELECT service_days FROM sites").fetchone()[0]
        db.close()
        self.assertEqual(len(gritting.load_service_days(days)), 7)

    # ---- extras on a site: charged per visit from a chosen night ---------------------------------------------

    def make_priced_site(self, price="100"):
        self.client.post("/gritting/customers", data={
            "form_type": "customer_site", "customer_choice": "new", "customer_name": "Delta Ltd", "contact_name": "Kim", "email": "kim@delta.test",
            "site_name": "Yard", "address": "5 Mill Rd", "pricing_label_1": "Weekday Night", "pricing_amount_1": price})
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        site_id = db.execute("SELECT id FROM sites WHERE name='Yard'").fetchone()[0]
        db.close()
        return site_id

    def call_out_night(self, site_id, night):
        response = self.client.post("/gritting/new-callout", data={
            "site_id": site_id, "called_out_at": night + "T18:00", "weather_conditions": "Frost", "service_notes": "",
            "planned_attendance": "After midnight", "recipient": "kim@delta.test", "pricing_label": "Weekday Night"})
        return int(response.headers["Location"].rsplit("/", 1)[1])

    def job(self, call_out_id):
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        db.row_factory = gritting.sqlite3.Row
        row = db.execute("SELECT * FROM call_outs WHERE id=?", (call_out_id,)).fetchone()
        db.close()
        return row

    def test_extra_added_on_night_11_is_charged_from_night_12(self):
        site_id = self.make_priced_site("100")
        nights = ["2026-01-%02d" % day for day in range(1, 11)]
        first_ten = [self.call_out_night(site_id, night) for night in nights]
        self.assertEqual([self.job(call_out_id)["pricing_amount"] for call_out_id in first_ten], [100.0] * 10)
        night_11 = self.call_out_night(site_id, "2026-01-11")
        # on the 11th night the customer adds an area worth 20 a visit, charged from the 12th night
        added = self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Rear loading bay", "amount": "20", "effective_from": "2026-01-12"}, follow_redirects=True)
        self.assertIn("Added Rear loading bay (+£20.00 per visit) from the night of", added.get_data(as_text=True))
        self.assertEqual(self.job(night_11)["pricing_amount"], 100.0)  # the 11th night is still the old price
        night_12, night_13 = self.call_out_night(site_id, "2026-01-12"), self.call_out_night(site_id, "2026-01-13")
        for call_out_id in (night_12, night_13):
            row = self.job(call_out_id)
            self.assertEqual((row["pricing_amount"], row["extras_amount"]), (120.0, 20.0))
            self.assertEqual(gritting.load_extras_items(row["extras_json"]), [{"description": "Rear loading bay", "amount": 20.0}])
        self.assertEqual(self.job(first_ten[0])["pricing_amount"], 100.0)  # earlier nights never change
        # the monthly summary totals the whole month: 11 nights at 100 and 2 at 120, and says what the extra was
        summary = self.client.get("/gritting/monthly-summary?month=2026-01").get_data(as_text=True)
        self.assertIn("£1340.00", summary)
        self.assertIn("incl. Rear loading bay +£20.00", summary)
        export = self.client.get("/gritting/monthly-summary/export?month=2026-01").get_data(as_text=True)
        self.assertIn("Weekday Night + Rear loading bay +£20.00", export)

    def test_extras_have_first_and_last_nights_and_reprice_only_open_jobs(self):
        site_id = self.make_priced_site("100")
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Side path", "amount": "20", "effective_from": "2026-02-10"})
        first_night, before = self.call_out_night(site_id, "2026-02-10"), self.call_out_night(site_id, "2026-02-09")
        self.assertEqual((self.job(before)["pricing_amount"], self.job(first_night)["pricing_amount"]), (100.0, 120.0))  # the first night counts
        done = self.call_out_night(site_id, "2026-02-11")
        self.client.post("/gritting/callouts/%d/complete-quick" % done)
        self.assertEqual(self.job(done)["status"], "completed")
        open_job = self.call_out_night(site_id, "2026-02-12")
        # a second extra and then the first one ending: open jobs follow, the finished job keeps its price
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Bin store", "amount": "5.50", "effective_from": "2026-02-11"})
        self.assertEqual((self.job(open_job)["pricing_amount"], self.job(done)["pricing_amount"]), (125.5, 120.0))
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        side_path = db.execute("SELECT id FROM site_extras WHERE description='Side path'").fetchone()[0]
        db.close()
        ended = self.client.post("/gritting/extras/%d/end" % side_path, data={"last_night": "2026-02-11"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Side path now ends after the night of", ended)
        self.assertIn("1 open job was repriced.", ended)
        self.assertEqual(self.job(open_job)["pricing_amount"], 105.5)  # 100 + the bin store; the side path stopped after the 11th
        self.assertEqual(self.job(first_night)["pricing_amount"], 120.0)
        self.assertEqual(self.job(done)["pricing_amount"], 120.0)
        # a job on a later night, made after the extra ended, never gets it
        self.assertEqual(self.job(self.call_out_night(site_id, "2026-02-20"))["pricing_amount"], 105.5)
        # removing an extra entered by mistake takes it off open jobs as well
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        bin_store = db.execute("SELECT id FROM site_extras WHERE description='Bin store'").fetchone()[0]
        db.close()
        self.assertIn("Removed Bin store.", self.client.post("/gritting/extras/%d/delete" % bin_store, follow_redirects=True).get_data(as_text=True))
        self.assertEqual(self.job(open_job)["pricing_amount"], 100.0)

    def test_completing_a_job_keeps_the_extras_in_its_price(self):
        site_id = self.make_priced_site("100")
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Rear bay", "amount": "20", "effective_from": "2026-03-01"})
        call_out_id = self.call_out_night(site_id, "2026-03-02")
        done = self.client.post("/gritting/callouts/%d/complete" % call_out_id, data={"completed_at": "2026-03-03T04:00", "service_notes": "ok", "pricing_label": "Weekday Night"})
        self.assertEqual(done.status_code, 302)
        row = self.job(call_out_id)
        self.assertEqual((row["status"], row["pricing_amount"], row["extras_amount"]), ("completed", 120.0, 20.0))
        detail = self.client.get("/gritting/callouts/%d" % call_out_id).get_data(as_text=True)
        self.assertIn("Site price £100.00 + extras: Rear bay +£20.00", detail)
        # a reduction is allowed with a minus; a job with no price chosen carries no extras
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Gate no longer needed", "amount": "-5", "effective_from": "2026-03-10"})
        self.assertEqual(self.job(self.call_out_night(site_id, "2026-03-10"))["pricing_amount"], 115.0)
        unpriced = self.client.post("/gritting/new-callout", data={
            "site_id": site_id, "called_out_at": "2026-03-11T18:00", "weather_conditions": "Frost", "service_notes": "",
            "planned_attendance": "x", "recipient": "kim@delta.test", "pricing_label": ""})
        self.assertIsNone(self.job(int(unpriced.headers["Location"].rsplit("/", 1)[1]))["pricing_amount"])

    def test_extras_show_on_the_live_list_and_the_final_list(self):
        site_id = self.make_priced_site("100")
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Rear bay", "amount": "20", "effective_from": "2026-04-01"})
        response = self.bulk_callout([site_id])
        self.assertEqual(response.status_code, 302)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        db.execute("UPDATE call_outs SET called_out_at='2026-04-02T18:00'")  # bulk_callout uses 12 January
        db.execute("UPDATE call_outs SET pricing_amount=100, extras_amount=0, extras_json=NULL")
        db.commit()
        batch_id = db.execute("SELECT id FROM dispatch_batches").fetchone()[0]
        db.close()
        site = self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Bin store", "amount": "5", "effective_from": "2026-04-02"}, follow_redirects=True)
        self.assertIn("1 open job was repriced.", site.get_data(as_text=True))
        live = self.client.get("/gritting/live-jobs").get_data(as_text=True)
        self.assertIn("Includes extras: Rear bay +£20.00, Bin store +£5.00", live)
        self.client.post("/gritting/batches/%d/finalise" % batch_id)
        sheet = self.client.get("/gritting/batches/%d/run-sheet" % batch_id).get_data(as_text=True)
        self.assertIn("Includes:", sheet)
        self.assertIn("Rear bay +£20.00, Bin store +£5.00", sheet)

    def test_the_customer_page_shows_and_manages_extras(self):
        site_id = self.make_priced_site("100")
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn("Extras on this site", page)
        self.assertIn('action="/gritting/sites/%d/extras"' % site_id, page)
        self.assertIn("First night charged", page)
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Rear bay", "amount": "20", "effective_from": "2020-01-01", "note": "Agreed by phone"})
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Future gate", "amount": "10", "effective_from": "2099-01-01"})
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn("+£20.00 per visit", page)
        self.assertIn("Agreed by phone", page)
        self.assertIn("Charged now", page)
        self.assertIn("Starts later", page)
        self.assertIn("End this extra", page)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        rear = db.execute("SELECT id FROM site_extras WHERE description='Rear bay'").fetchone()[0]
        db.close()
        self.client.post("/gritting/extras/%d/end" % rear, data={"last_night": "2020-06-30"})
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn("Ended", page)
        self.assertIn("to the night of", page)

    def test_extras_reject_bad_entries(self):
        site_id = self.make_priced_site("100")
        for data, message in [({"description": "", "amount": "20"}, "Describe what was added"), ({"description": "Bay", "amount": ""}, "Enter the extra price per visit"),
                              ({"description": "Bay", "amount": "0"}, "Enter the extra price per visit"), ({"description": "Bay", "amount": "x"}, "Enter the extra price per visit"),
                              ({"description": "Bay", "amount": "20", "effective_from": "not-a-date"}, "Choose a valid first night")]:
            with self.subTest(data=data):
                self.assertIn(message, self.client.post("/gritting/sites/%d/extras" % site_id, data=data, follow_redirects=True).get_data(as_text=True))
        self.assertIn("Site not found", self.client.post("/gritting/sites/999/extras", data={"description": "Bay", "amount": "5"}, follow_redirects=True).get_data(as_text=True))
        self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Bay", "amount": "5", "effective_from": "2026-05-10"})
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        extra_id = db.execute("SELECT id FROM site_extras").fetchone()[0]
        count = db.execute("SELECT COUNT(*) FROM site_extras").fetchone()[0]
        db.close()
        self.assertEqual(count, 1)  # nothing half-saved from the bad entries
        before_first = self.client.post("/gritting/extras/%d/end" % extra_id, data={"last_night": "2026-05-01"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("The last night cannot be before the first night", before_first)
        self.assertIn("Choose a valid last night", self.client.post("/gritting/extras/%d/end" % extra_id, data={"last_night": "xx"}, follow_redirects=True).get_data(as_text=True))
        self.assertIn("Extra not found", self.client.post("/gritting/extras/999/end", follow_redirects=True).get_data(as_text=True))
        self.assertIn("Extra not found", self.client.post("/gritting/extras/999/delete", follow_redirects=True).get_data(as_text=True))
        default_day = self.client.post("/gritting/sites/%d/extras" % site_id, data={"description": "Tomorrow's", "amount": "1"})
        self.assertEqual(default_day.status_code, 302)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        first_night = db.execute("SELECT effective_from FROM site_extras WHERE description=?", ("Tomorrow's",)).fetchone()[0]
        db.close()
        from datetime import datetime, timedelta
        self.assertEqual(first_night, (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d"))  # a blank first night means tomorrow

    # ---- several notification emails, several emergency contacts, a postcode ---------------------------------

    def contact_form(self, **changes):
        form = {
            "form_type": "customer_site", "customer_choice": "new", "customer_name": "Echo Ltd", "contact_name": "Ann",
            "email": ["Ops@Echo.test", "accounts@echo.test", "ops@echo.test"],  # a repeat in a different case is dropped
            "emergency_contact_name": ["Ann Lee", "Bob Ray"], "emergency_contact_phone": ["07700 900111", "07700 900222"],
            "site_name": "Yard", "address": "7 Quay St", "postcode": "nr151be", "site_emails": ["Manager@Echo.test"],
            "pricing_label_1": "Weekday Night", "pricing_amount_1": "100",
        }
        form.update(changes)
        return form

    def test_a_customer_has_several_emails_and_emergency_contacts_and_a_site_postcode(self):
        added = self.client.post("/gritting/customers", data=self.contact_form(), follow_redirects=True).get_data(as_text=True)
        self.assertIn("Customer Echo Ltd and site Yard added.", added)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer = db.execute("SELECT email, emergency_phone, emergency_contacts FROM customers WHERE name='Echo Ltd'").fetchone()
        site = db.execute("SELECT postcode, notify_emails FROM sites WHERE name='Yard'").fetchone()
        db.close()
        self.assertEqual(customer[0], "ops@echo.test, accounts@echo.test")
        self.assertEqual(customer[1], "07700 900111")  # the first contact's phone stays in the older single-phone field
        self.assertEqual(gritting.load_emergency_contacts(customer[2]), [{"name": "Ann Lee", "phone": "07700 900111"}, {"name": "Bob Ray", "phone": "07700 900222"}])
        self.assertEqual(site, ("NR15 1BE", "manager@echo.test"))  # the postcode is tidied into capitals with its space
        self.assertEqual(gritting.emergency_text(customer[2]), "Ann Lee 07700 900111 · Bob Ray 07700 900222")

    def test_contact_details_are_checked_before_anything_is_saved(self):
        cases = [
            ({"email": ["not-an-email"]}, "Check the email address: not-an-email."),
            ({"email": ["ok@echo.test", "two words@echo.test"]}, "Check the email address: two words@echo.test."),
            ({"site_emails": ["broken@"]}, "Check the email address: broken@."),
            ({"email": [""]}, "Enter at least one notification email for the customer."),
            ({"postcode": "ZZ"}, "That does not look like a UK postcode."),
            ({"emergency_contact_name": ["Ann"], "emergency_contact_phone": [""]}, "Add a phone number for the emergency contact Ann."),
            ({"emergency_contact_name": [""], "emergency_contact_phone": ["abc"]}, "abc does not look like a phone number."),
        ]
        for changes, message in cases:
            with self.subTest(changes=changes):
                response = self.client.post("/gritting/customers", data=self.contact_form(**changes), follow_redirects=True).get_data(as_text=True)
                self.assertIn(message, response)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        self.assertEqual((db.execute("SELECT COUNT(*) FROM customers").fetchone()[0], db.execute("SELECT COUNT(*) FROM sites").fetchone()[0]), (0, 0))
        db.close()
        # the postcode is optional, and so are the contacts and the site's own emails
        bare = self.contact_form(postcode="", site_emails=[""], emergency_contact_name=[""], emergency_contact_phone=[""])
        self.assertIn("Customer Echo Ltd and site Yard added.", self.client.post("/gritting/customers", data=bare, follow_redirects=True).get_data(as_text=True))

    def test_postcodes_are_tidied_and_checked(self):
        for typed, expected in [("nr151be", "NR15 1BE"), ("NR15  1BE", "NR15 1BE"), (" sw1a 1aa ", "SW1A 1AA"), ("M11AE", "M1 1AE"), ("", ""), (None, "")]:
            self.assertEqual(gritting.normalize_postcode(typed), expected)
        for bad in ("12345", "NR15", "HELLO WORLD", "NR15 1B"):
            self.assertIsNone(gritting.normalize_postcode(bad))

    def test_notifications_go_to_the_customer_emails_and_the_sites_own_emails(self):
        self.client.post("/gritting/customers", data=self.contact_form(site_emails=["Manager@Echo.test", "ops@echo.test"]))
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        site_id = db.execute("SELECT id FROM sites WHERE name='Yard'").fetchone()[0]
        db.close()
        sent = []
        with patch.object(gritting, "send_email_message", side_effect=lambda recipients, subject, body, attachments=None: sent.append(list(recipients)) or ("sent", "ok")):
            response = self.bulk_callout([site_id])
            self.assertEqual(response.status_code, 302)
            self.client.post("/gritting/batches/1/finalise")
            self.client.post("/gritting/callouts/1/complete-quick", data={"next": "live_jobs"})
        expected = ["ops@echo.test", "accounts@echo.test", "manager@echo.test"]  # customer list first, then the site's, no repeats
        self.assertEqual(sent, [expected, expected])  # the call-out email and the completion email
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        logged = db.execute("SELECT recipient FROM notification_logs ORDER BY id").fetchall()
        db.close()
        self.assertEqual([row[0] for row in logged], [", ".join(expected)] * 2)
        detail = self.client.get("/gritting/callouts/1").get_data(as_text=True)
        self.assertIn("ops@echo.test, accounts@echo.test, manager@echo.test", detail)
        self.assertIn("Ann Lee 07700 900111 · Bob Ray 07700 900222", detail)
        self.assertIn("7 Quay St, NR15 1BE", detail)

    def test_edit_forms_change_emails_contacts_and_postcode(self):
        self.client.post("/gritting/customers", data=self.contact_form())
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer_id = db.execute("SELECT id FROM customers").fetchone()[0]
        site_id = db.execute("SELECT id FROM sites").fetchone()[0]
        db.close()
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn('value="ops@echo.test"', page)
        self.assertIn('value="accounts@echo.test"', page)
        self.assertIn('value="Bob Ray"', page)
        self.assertIn('value="07700 900222"', page)
        self.assertIn('value="NR15 1BE"', page)
        self.assertIn('value="manager@echo.test"', page)
        self.assertGreaterEqual(page.count("Add another email"), 3)
        self.assertIn("Add another emergency contact", page)
        edited = self.client.post("/gritting/customers", data={
            "form_type": "edit_customer", "customer_id": customer_id, "name": "Echo Ltd", "contact_name": "Ann",
            "email": ["new@echo.test", "second@echo.test", "third@echo.test"],
            "emergency_contact_name": ["Cy Day"], "emergency_contact_phone": ["01603 555 000"]}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Customer updated.", edited)
        site = self.client.post("/gritting/customers", data={
            "form_type": "edit_site", "site_id": site_id, "name": "Yard", "address": "7 Quay St", "postcode": "NR3 1AB", "site_emails": [""],
            "route_notes": "", "service_day_monday": "on"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Site updated.", site)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        customer = db.execute("SELECT email, emergency_phone, emergency_contacts FROM customers").fetchone()
        site_row = db.execute("SELECT postcode, notify_emails FROM sites").fetchone()
        db.close()
        self.assertEqual(customer[0], "new@echo.test, second@echo.test, third@echo.test")
        self.assertEqual((customer[1], gritting.load_emergency_contacts(customer[2])), ("01603 555 000", [{"name": "Cy Day", "phone": "01603 555 000"}]))
        self.assertEqual(site_row, ("NR3 1AB", ""))  # clearing the site's extra emails works
        for bad, message in [({"email": [""]}, "Enter at least one notification email"), ({"email": ["bad"]}, "Check the email address: bad.")]:
            data = {"form_type": "edit_customer", "customer_id": customer_id, "name": "Echo Ltd", "contact_name": "Ann", "email": ["x@echo.test"], "emergency_contact_name": [""], "emergency_contact_phone": [""]}
            data.update(bad)
            self.assertIn(message, self.client.post("/gritting/customers", data=data, follow_redirects=True).get_data(as_text=True))
        self.assertIn("NR3 1AB", self.client.get("/gritting/customers").get_data(as_text=True))
        self.assertIn("does not look like a UK postcode", self.client.post("/gritting/customers", data={
            "form_type": "edit_site", "site_id": site_id, "name": "Yard", "address": "7 Quay St", "postcode": "nope", "route_notes": ""}, follow_redirects=True).get_data(as_text=True))

    def test_older_customers_with_one_phone_number_still_show_it_as_a_contact(self):
        self.client.get("/gritting/customers")
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        db.execute("INSERT INTO customers (name, contact_name, contact_method, contact_value, email, emergency_phone) VALUES ('Old Ltd', 'Al', 'Email', '07700 900999', 'al@old.test', '07700 900999')")
        db.commit()
        db.close()
        self.assertEqual(gritting.load_emergency_contacts(None, "07700 900999"), [{"name": "", "phone": "07700 900999"}])
        self.assertEqual(gritting.load_emergency_contacts("garbage", ""), [])
        self.assertEqual(gritting.emergency_text(None, "07700 900999"), "07700 900999")
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn('value="07700 900999"', page)

    def test_the_final_list_shows_the_postcode_and_every_emergency_contact(self):
        self.client.post("/gritting/customers", data=self.contact_form())
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        site_id = db.execute("SELECT id FROM sites").fetchone()[0]
        db.close()
        self.bulk_callout([site_id])
        self.client.post("/gritting/batches/1/finalise")
        sheet = self.client.get("/gritting/batches/1/run-sheet").get_data(as_text=True)
        self.assertIn("7 Quay St, NR15 1BE", sheet)
        self.assertIn("Emergency: Ann Lee 07700 900111 · Bob Ray 07700 900222", sheet)
        self.assertIn("7 Quay St, NR15 1BE", self.client.get("/gritting/live-jobs").get_data(as_text=True))

    def test_the_spreadsheet_import_reads_several_emails_a_contact_name_postcode_and_site_emails(self):
        gritting.ensure_initialised()
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        db.row_factory = gritting.sqlite3.Row
        rows = [{"customer_name": "Fox Ltd", "contact_name": "Fay", "email": "Fay@fox.test; ACCOUNTS@fox.test, fay@fox.test", "emergency_contact": "Fay Fox",
                 "emergency_phone": "07700 900333", "site_name": "Depot", "site_address": "1 Mill Rd", "postcode": "nr3 1ab", "site_emails": "depot@fox.test",
                 "service_days": "Monday, Tuesday"}]
        summary = gritting.import_customer_rows(db, rows)
        self.assertEqual((summary["customers_added"], summary["sites_added"]), (1, 1))
        customer = db.execute("SELECT email, emergency_contacts FROM customers WHERE name='Fox Ltd'").fetchone()
        site = db.execute("SELECT postcode, notify_emails FROM sites WHERE name='Depot'").fetchone()
        self.assertEqual(customer["email"], "fay@fox.test, accounts@fox.test")
        self.assertEqual(gritting.load_emergency_contacts(customer["emergency_contacts"]), [{"name": "Fay Fox", "phone": "07700 900333"}])
        self.assertEqual((site["postcode"], site["notify_emails"]), ("NR3 1AB", "depot@fox.test"))
        # importing again without those columns keeps what is already saved
        again = [{"customer_name": "Fox Ltd", "contact_name": "Fay", "email": "fay@fox.test", "site_name": "Depot", "site_address": "1 Mill Rd"}]
        gritting.import_customer_rows(db, again)
        site = db.execute("SELECT postcode, notify_emails FROM sites WHERE name='Depot'").fetchone()
        customer = db.execute("SELECT emergency_contacts FROM customers WHERE name='Fox Ltd'").fetchone()
        db.close()
        self.assertEqual((site["postcode"], site["notify_emails"]), ("NR3 1AB", "depot@fox.test"))
        self.assertEqual(len(gritting.load_emergency_contacts(customer["emergency_contacts"])), 1)

    # ---- prices that depend on the night: weekdays, 7-day sites and special dates (match days) ----------------

    def seven_day_form(self):
        form = {"form_type": "customer_site", "customer_choice": "new", "customer_name": "Stadium Ltd", "contact_name": "Sam", "email": ["ops@stadium.test"],
                "site_name": "Car park", "address": "1 Ground Rd"}
        for day in ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"):
            form["service_day_" + day] = "on"
        form.update({
            "pricing_label_1": "Weekday", "pricing_amount_1": "100", "pricing_days_1": ["monday", "tuesday", "wednesday", "thursday"],
            "pricing_label_2": "Friday and Saturday", "pricing_amount_2": "150", "pricing_days_2": ["friday", "saturday"],
            "pricing_label_3": "Sunday", "pricing_amount_3": "180", "pricing_days_3": ["sunday"],
            "pricing_label_4": "Match day", "pricing_amount_4": "250", "pricing_event_4": "Norwich home match",
        })
        return form

    def test_the_price_follows_the_night_for_a_site_that_runs_seven_days_a_week(self):
        self.client.post("/gritting/customers", data=self.seven_day_form())
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        stored = db.execute("SELECT pricing_options FROM sites").fetchone()[0]
        db.close()
        options = gritting.load_pricing_options(stored)
        self.assertEqual(len(options), 4)  # not limited to three prices
        self.assertEqual(options[1]["days"], ["friday", "saturday"])
        self.assertEqual(options[3]["event"], "Norwich home match")
        # ticked days are mornings: Friday night is Saturday morning, Saturday night is Sunday morning, Sunday night is Monday morning
        expected = {"2026-01-12": "Weekday", "2026-01-14": "Weekday", "2026-01-15": "Friday and Saturday", "2026-01-16": "Friday and Saturday", "2026-01-17": "Sunday", "2026-01-18": "Weekday"}
        for day, label in expected.items():
            with self.subTest(day=day):
                self.assertEqual(gritting.suggested_pricing_label(day + "T18:00", options), label)

    def test_a_special_date_overrides_the_weekday_price(self):
        self.client.post("/gritting/customers", data=self.seven_day_form())
        self.assertIn("Norwich home match added on 2026-01-17.", self.client.post("/gritting/settings", data={
            "form_type": "special_date_add", "event_date": "2026-01-17", "event_name": "Norwich home match"}, follow_redirects=True).get_data(as_text=True))
        self.assertIn("Enter a date and a name", self.client.post("/gritting/settings", data={"form_type": "special_date_add", "event_date": "nope", "event_name": "x"}, follow_redirects=True).get_data(as_text=True))
        # the picker on the call-out page chooses the match day price for that night only
        saturday = self.client.get("/gritting/new-callout?called_out_at=2026-01-17T18:00").get_data(as_text=True)
        self.assertRegex(saturday, r'<option value="Match day" selected>')
        friday = self.client.get("/gritting/new-callout?called_out_at=2026-01-16T18:00").get_data(as_text=True)
        self.assertRegex(friday, r'<option value="Friday and Saturday" selected>')
        self.assertIn('"2026-01-17": ["Norwich home match"]', saturday)
        # and the job keeps the price it was created at
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        site_id = db.execute("SELECT id FROM sites").fetchone()[0]
        db.close()
        data = {"called_out_at": "2026-01-17T18:00", "planned_attendance_from": "22:00", "planned_attendance_to": "04:00", "weather_conditions": "Frost",
                "service_notes": "", "subject": "", "message": "", "subject_override_enabled": "0", "message_override_enabled": "0",
                "selected_site_ids": [str(site_id)], "pricing_label_%d" % site_id: "Match day"}
        self.client.post("/gritting/callouts/bulk", data=data)
        db = gritting.sqlite3.connect(gritting.DATABASE_PATH)
        self.assertEqual(db.execute("SELECT pricing_label, pricing_amount FROM call_outs").fetchone(), ("Match day", 250.0))
        special_id = db.execute("SELECT id FROM special_dates").fetchone()[0]
        db.close()
        self.client.post("/gritting/settings", data={"form_type": "special_date_delete", "special_date_id": special_id})
        again = self.client.get("/gritting/new-callout?called_out_at=2026-01-17T18:00").get_data(as_text=True)
        self.assertRegex(again, r'<option value="Sunday" selected>')  # Saturday night is for Sunday morning

    def test_prices_set_up_the_old_way_still_pick_by_the_name(self):
        options = [{"label": "Weekday Night", "amount": 100.0}, {"label": "Friday Night", "amount": 120.0}, {"label": "Weekend", "amount": 150.0}]
        pick = lambda day: gritting.suggested_pricing_label(day + "T18:00", options)
        self.assertEqual((pick("2026-01-12"), pick("2026-01-16"), pick("2026-01-17"), pick("2026-01-18")), ("Weekday Night", "Friday Night", "Weekend", "Weekend"))
        self.assertEqual(gritting.suggested_pricing_label("2026-01-12T18:00", []), "")

    def test_the_site_form_offers_day_ticks_special_dates_and_more_rows(self):
        self.client.post("/gritting/customers", data=self.seven_day_form())
        self.client.post("/gritting/settings", data={"form_type": "special_date_add", "event_date": "2026-01-17", "event_name": "Norwich home match"})
        page = self.client.get("/gritting/customers").get_data(as_text=True)
        self.assertIn("Add another price", page)
        self.assertIn('name="pricing_days_4"', page)  # a fourth price shows on the existing site
        self.assertIn('<option value="Norwich home match" selected>', page)
        self.assertRegex(page, r'name="pricing_days_2" value="friday" checked')

    def test_service_days_are_the_mornings_a_site_needs_clearing(self):
        key = gritting.dispatch_day_key
        # Mon to Fri mornings means Sunday, Monday, Tuesday, Wednesday and Thursday nights
        nights = {"2026-01-11": "monday", "2026-01-12": "tuesday", "2026-01-15": "friday", "2026-01-16": "saturday", "2026-01-17": "sunday"}
        for night, morning in nights.items():
            with self.subTest(night=night):
                self.assertEqual(key(night + "T18:00"), morning)
        self.assertEqual(key("2026-01-17T01:30"), "saturday")  # sent after midnight, so it is already that morning
        form = {"form_type": "customer_site", "customer_choice": "new", "customer_name": "Office Ltd", "contact_name": "Olly", "email": ["o@office.test"],
                "site_name": "Yard", "address": "1 Road"}
        for day in ("monday", "tuesday", "wednesday", "thursday", "friday"):
            form["service_day_" + day] = "on"
        self.client.post("/gritting/customers", data=form)
        due = lambda stamp: "Office Ltd" in self.client.get("/gritting/new-callout?called_out_at=" + stamp).get_data(as_text=True).split("Skipped today")[0]
        self.assertTrue(due("2026-01-11T18:00"))   # Sunday night, for Monday morning
        self.assertTrue(due("2026-01-15T18:00"))   # Thursday night, for Friday morning
        self.assertFalse(due("2026-01-16T18:00"))  # Friday night, nobody in on Saturday
        self.assertFalse(due("2026-01-17T18:00"))  # Saturday night, nobody in on Sunday

    def test_urls_inside_pages_point_at_gritting(self):
        page = self.client.get("/gritting/customers").data.decode()
        for href in re.findall(r'href="(/[^"#]*)"', page):
            self.assertTrue(href == "/" or href.startswith("/gritting") or href.startswith("/static/"), href)


def raw_forecast(surface, air=None, dew=-5.0, rain=0.0, cloud=100, wind=10.0, start="2026-11-09T00:00", hours=96):
    """A hand-made Open-Meteo style response; surface/air are functions of the datetime."""
    from datetime import timedelta as delta
    begin = datetime.fromisoformat(start)
    stamps = [begin + delta(hours=index) for index in range(hours)]
    return {"hourly": {
        "time": [stamp.strftime("%Y-%m-%dT%H:%M") for stamp in stamps],
        "soil_temperature_0cm": [surface(stamp) for stamp in stamps],
        "temperature_2m": [(air or surface)(stamp) for stamp in stamps],
        "dew_point_2m": [dew] * hours, "precipitation": [rain] * hours, "snowfall": [0.0] * hours,
        "cloud_cover": [cloud] * hours, "wind_speed_10m": [wind] * hours,
    }}


class RoadForecastTests(unittest.TestCase):
    def test_hour_risk_rules_keep_a_safety_margin(self):
        def row(surface, air=8.0, dew=-5.0, rain=0.0, snow=0.0, cloud=100, wind=10):
            return {"surface": surface, "air": air, "dew": dew, "rain": rain, "snow": snow, "cloud": cloud, "wind": wind}
        cases = [
            (row(-1.0), "grit"), (row(0.5), "grit"), (row(0.6), "watch"), (row(2.0), "watch"), (row(2.1), "ok"), (row(3.0), "ok"),
            (row(1.9, rain=0.4), "grit"), (row(2.5, rain=0.4), "ok"), (row(1.4, dew=1.0), "grit"), (row(1.4, dew=-5.0), "watch"),
            (row(5.0, air=2.0, cloud=10, wind=3), "ok"),  # clear, calm and cold air no longer raises a warning by itself
        ]
        for case, expected in cases:
            with self.subTest(case=case):
                self.assertEqual(gritting.classify_road_hour(case)[0], expected)

    def test_three_models_use_the_middle_value(self):
        cold, mild = (lambda stamp: -1.0), (lambda stamp: 6.0)

        def three(first, second, third):
            hourly = {}
            for model, surface in (("ukmo_seamless", first), ("ecmwf_ifs025", second), ("icon_seamless", third)):
                part = raw_forecast(surface, hours=24 * 8)["hourly"]
                hourly.setdefault("time", part["time"])
                for key, values in part.items():
                    if key != "time":
                        hourly[f"{key}_{model}"] = values
            return {"hourly": hourly}
        now = datetime(2026, 11, 10, 19, 0)
        def first_night(*models, when=now):
            return gritting.road_forecast_from_open_meteo(three(*models), when)["nights"][0]
        self.assertEqual(first_night(cold, cold, cold)["level"], "grit")
        two_cold = first_night(cold, cold, mild)  # two of three say freezing
        self.assertEqual((two_cold["level"], two_cold["min_surface"]), ("grit", -1.0))
        self.assertEqual(two_cold["model_lows"], {"Met Office": -1.0, "ECMWF": -1.0, "ICON": 6.0})
        one_cold = first_night(cold, mild, mild)  # one model on its own is outvoted for the numbers but still worth a look
        self.assertEqual((one_cold["level"], one_cold["min_surface"], one_cold["models_disagree"]), ("watch", 6.0, True))
        self.assertIn("One forecast model", one_cold["reason"])
        nights = gritting.road_forecast_from_open_meteo(three(cold, mild, mild), now)["nights"]
        self.assertEqual([n["level"] for n in nights[:6]], ["watch"] * 4 + ["ok", "ok"])  # beyond three days a lone model is ignored
        self.assertEqual([n["outlook"] for n in nights[:6]], [False, False, False, False, True, True])
        self.assertEqual(len(gritting.split_models(three(cold, mild, mild))), 3)
        self.assertEqual(len(gritting.split_models(raw_forecast(mild))), 1)
        for model in ("ukmo_seamless", "ecmwf_ifs025", "icon_seamless"):
            self.assertIn(model, gritting.forecast_request_url())

    def test_nights_and_labels(self):
        freezing = lambda stamp: -1.0 if datetime(2026, 11, 10, 22) <= stamp < datetime(2026, 11, 11, 6) else 6.0
        data = gritting.road_forecast_from_open_meteo(raw_forecast(freezing, hours=24 * 12), datetime(2026, 11, 10, 20, 15))
        self.assertEqual([night["label"] for night in data["nights"]][:3], ["Tonight", "Tomorrow night", "Thursday night"])
        self.assertEqual(len(data["nights"]), 7)
        self.assertEqual(data["nights"][-1]["label"], "Monday night")
        self.assertEqual(len(data["week"]), 7 * 24)
        self.assertEqual(data["week"][0]["hour"], "20:00")
        self.assertEqual({"time", "hour", "day", "date", "rain", "surface", "air", "dew", "level"}, set(data["week"][0]))
        self.assertEqual(data["week"][0]["date"], "2026-11-10")
        self.assertNotIn("hours", data)
        # a gritting day for the hour blocks: 09:00 on the first night's date onwards
        self.assertEqual((data["hourly"][0]["date"], data["hourly"][0]["hour"]), ("2026-11-10", "09:00"))
        self.assertEqual(len([hour for hour in data["hourly"] if hour["date"] == "2026-11-10" or (hour["date"] == "2026-11-11" and hour["hour"] < "09:00")]), 24)
        tonight = data["nights"][0]
        self.assertEqual((tonight["level"], tonight["headline"], tonight["min_surface"]), ("grit", "GRITTING REQUIRED", -1.0))
        self.assertEqual((tonight["grit_from"], tonight["grit_until"]), ("22:00", "06:00"))
        self.assertEqual((tonight["min_surface_time"], tonight["min_surface_day"]), ("22:00", "Tue"))
        self.assertEqual(data["nights"][1]["level"], "ok")

    def test_rain_before_a_cold_spell_gives_a_grit_after_rain_time(self):
        def surface(stamp):
            return -1.0 if datetime(2026, 11, 11, 1) <= stamp < datetime(2026, 11, 11, 6) else 6.0
        raw = raw_forecast(surface, hours=24 * 12)
        times = raw["hourly"]["time"]
        for index, stamp in enumerate(times):  # rain 21:00-23:00 and a light 02:00 shower on the night of the 10th
            if stamp in ("2026-11-10T21:00", "2026-11-10T22:00"):
                raw["hourly"]["precipitation"][index] = 0.6
            if stamp == "2026-11-11T02:00":
                raw["hourly"]["precipitation"][index] = 0.2
        night = gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))["nights"][0]
        self.assertEqual(night["level"], "grit")
        self.assertEqual((night["rain_from"], night["rain_until"], night["rain_mm"]), ("21:00", "03:00", 1.4))
        self.assertEqual(night["grit_after_rain_at"], "03:00")
        self.assertEqual(night["rain_advice"], "Rain until 03:00: grit after it stops, from 03:00.")

    def test_rain_that_ends_before_the_frost_means_grit_when_the_frost_starts(self):
        def surface(stamp):
            return -1.0 if datetime(2026, 11, 11, 3) <= stamp < datetime(2026, 11, 11, 6) else 6.0
        raw = raw_forecast(surface, hours=24 * 12)
        for index, stamp in enumerate(raw["hourly"]["time"]):
            if stamp in ("2026-11-10T20:00", "2026-11-10T21:00"):
                raw["hourly"]["precipitation"][index] = 1.0
        night = gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))["nights"][0]
        self.assertEqual((night["rain_until"], night["grit_after_rain_at"]), ("22:00", "03:00"))
        self.assertIn("from 03:00", night["rain_advice"])

    def test_rain_through_the_cold_spell_warns_that_salt_may_wash_off(self):
        raw = raw_forecast(lambda stamp: -1.0 if datetime(2026, 11, 10, 22) <= stamp < datetime(2026, 11, 11, 2) else 6.0, hours=24 * 12)
        for index, stamp in enumerate(raw["hourly"]["time"]):
            if "2026-11-10T20:00" <= stamp <= "2026-11-11T04:00":
                raw["hourly"]["precipitation"][index] = 0.5
        night = gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))["nights"][0]
        self.assertIn("salt may wash off", night["rain_advice"])
        self.assertEqual(night["grit_after_rain_at"], "")

    def test_dry_nights_have_no_rain_advice_and_the_next_rain_spell_is_found(self):
        raw = raw_forecast(lambda stamp: 7.0, hours=24 * 12)
        data = gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))
        self.assertEqual((data["nights"][0]["rain_mm"], data["nights"][0]["rain_advice"], data["next_rain"]), (0.0, "", None))
        for index, stamp in enumerate(raw["hourly"]["time"]):
            if stamp in ("2026-11-12T03:00", "2026-11-12T04:00", "2026-11-12T09:00"):
                raw["hourly"]["precipitation"][index] = 0.5
        spell = gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))["next_rain"]
        self.assertEqual(spell, {"from": "Thu 03:00", "until": "Thu 05:00", "mm": 1.0})
        self.assertTrue(all("rain" in hour for hour in gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 19, 0))["week"]))

    def test_a_night_the_forecast_only_partly_covers_is_flagged(self):
        data = gritting.road_forecast_from_open_meteo(raw_forecast(lambda stamp: 7.0, hours=24 * 5), datetime(2026, 11, 10, 20, 0))
        coverage = [night["hours_covered"] for night in data["nights"]]
        self.assertEqual(coverage[:3], [12, 12, 12])
        self.assertLess(coverage[-1], 12)

    def test_after_midnight_the_night_in_progress_is_still_tonight(self):
        mild = lambda stamp: 7.0
        data = gritting.road_forecast_from_open_meteo(raw_forecast(mild), datetime(2026, 11, 11, 3, 0))
        self.assertEqual(data["nights"][0]["label"], "Tonight")
        self.assertEqual(data["nights"][0]["date"], "2026-11-10")
        self.assertEqual(data["nights"][0]["level"], "ok")
        self.assertEqual((data["hourly"][0]["date"], data["hourly"][0]["hour"]), ("2026-11-10", "09:00"))  # the night began yesterday evening

    def test_missing_values_are_skipped_and_empty_forecasts_fail(self):
        raw = raw_forecast(lambda stamp: 7.0)
        raw["hourly"]["soil_temperature_0cm"][30] = None
        self.assertTrue(gritting.road_forecast_from_open_meteo(raw, datetime(2026, 11, 10, 20, 0))["nights"])
        empty = raw_forecast(lambda stamp: 7.0)
        empty["hourly"]["soil_temperature_0cm"] = [None] * len(empty["hourly"]["time"])
        with self.assertRaises(ValueError):
            gritting.road_forecast_from_open_meteo(empty, datetime(2026, 11, 10, 20, 0))

    def test_forecast_is_cached_and_falls_back_to_an_older_copy(self):
        import io
        import json
        gritting._forecast_cache.update(at=0.0, data=None)
        raw = raw_forecast(lambda stamp: 7.0)
        calls = []

        def fake_open(request_obj, timeout=None):
            calls.append(request_obj.full_url)
            return io.BytesIO(json.dumps(raw).encode())
        with patch.object(gritting.urllib.request, "urlopen", fake_open):
            first, stale = gritting.fetch_open_meteo_forecast()
            gritting.fetch_open_meteo_forecast()
        self.assertEqual((len(calls), stale), (1, False))
        self.assertIn("models=ukmo_seamless", calls[0])
        self.assertIn("forecast_days=8", calls[0])
        self.assertIn("soil_temperature_0cm", calls[0])
        gritting._forecast_cache["at"] -= gritting.FORECAST_CACHE_SECONDS + 1  # the copy is now due a refresh
        def failing(request_obj, timeout=None):
            raise OSError("offline")
        with patch.object(gritting.urllib.request, "urlopen", failing):
            again, stale = gritting.fetch_open_meteo_forecast()
        self.assertEqual((again, stale), (first, True))
        gritting._forecast_cache["at"] -= gritting.FORECAST_STALE_LIMIT_SECONDS
        with patch.object(gritting.urllib.request, "urlopen", failing):
            with self.assertRaises(OSError):
                gritting.fetch_open_meteo_forecast()
        gritting._forecast_cache.update(at=0.0, data=None)

    def test_a_paid_open_meteo_key_switches_to_the_commercial_endpoint(self):
        self.assertTrue(gritting.forecast_request_url().startswith("https://api.open-meteo.com/v1/forecast?"))
        with patch.dict(gritting.os.environ, {"OPEN_METEO_API_KEY": "abc123"}):
            url = gritting.forecast_request_url()
        self.assertTrue(url.startswith("https://customer-api.open-meteo.com/v1/forecast?"))
        self.assertTrue(url.endswith("&apikey=abc123"))

    def test_forecast_endpoint_and_dashboard_card(self):
        client = app.app.test_client()
        directory = tempfile.TemporaryDirectory()
        original = (gritting.DATA_ROOT, gritting.DATABASE_PATH, gritting.BUNDLED_CUSTOMER_XLSX, gritting.SEASON_ARCHIVE_DIR)
        gritting.configure(directory.name + "/gritting")
        gritting._initialised_database = None
        try:
            raw = raw_forecast(lambda stamp: -1.0)
            with patch.object(gritting, "fetch_open_meteo_forecast", return_value=(raw, False)), patch.object(gritting, "local_now", return_value=datetime(2026, 11, 10, 20, 0)):
                payload = client.get("/gritting/road-forecast.json").get_json()
            self.assertTrue(payload["ok"])
            self.assertEqual(payload["nights"][0]["level"], "grit")
            self.assertEqual(payload["nights"][0]["headline"], "GRITTING REQUIRED")
            with patch.object(gritting, "fetch_open_meteo_forecast", side_effect=OSError("offline")):
                failed = client.get("/gritting/road-forecast.json")
            self.assertEqual(failed.status_code, 200)
            self.assertFalse(failed.get_json()["ok"])
            page = client.get("/gritting/").get_data(as_text=True)
            self.assertIn('id="road-forecast"', page)
            self.assertIn("/gritting/road-forecast.json", page)
            self.assertIn("not a road sensor reading", page)
        finally:
            gritting.DATA_ROOT, gritting.DATABASE_PATH, gritting.BUNDLED_CUSTOMER_XLSX, gritting.SEASON_ARCHIVE_DIR = original
            gritting._initialised_database = None
            directory.cleanup()


if __name__ == "__main__":
    unittest.main()
