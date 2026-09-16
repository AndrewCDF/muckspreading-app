import json
import io
import os
import tempfile
import unittest
import zipfile
from unittest.mock import patch

import app


class InvoiceHistoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.ledger_path = os.path.join(self.temp.name, "invoice_ledger.json")
        self.archive_dir = os.path.join(self.temp.name, "invoices")
        self.paths = patch.multiple(
            app,
            INVOICE_LEDGER_PATH=self.ledger_path,
            INVOICE_ARCHIVE_DIR=self.archive_dir,
        )
        self.paths.start()

    def tearDown(self):
        self.paths.stop()
        self.temp.cleanup()

    def test_history_has_view_and_download_actions(self):
        row = {
            "ledger_index": 0,
            "reference_label": "Invoice 12",
            "type_label": "Invoice",
            "customer": "Example Customer",
            "farm_name": "Example Farm",
            "period_label": "01/09/2026 to 02/09/2026",
            "job_count": 2,
            "grand_total_label": "£120.00",
            "created_label": "02/09/2026 12:00",
            "note": "",
            "can_edit": True,
        }
        with patch.object(app, "ensure_data_dir"), patch.object(app, "invoice_history_rows", return_value=[row]):
            with app.app.test_client() as client:
                html = client.get("/invoice/history").get_data(as_text=True)
        self.assertIn("View PDF", html)
        self.assertIn("Save PDF", html)
        self.assertIn("Save Excel", html)
        self.assertIn("/invoice/history/0/download.pdf", html)
        self.assertIn('data-invoice-download="1"', html)
        self.assertIn("navigator.share", html)
        self.assertIn("Previous Invoices", app.HTML)

    def test_missing_history_files_are_rebuilt_and_archived(self):
        history_row = {"invoice_number": 12, "job_ids": [1]}
        with open(self.ledger_path, "w") as handle:
            json.dump([history_row], handle)
        invoice = {"filename": "invoice_12.xlsx"}
        with patch.object(app, "build_invoice_from_history_entry", return_value=(invoice, "")), patch.object(
            app, "build_invoice_xlsx_bytes", return_value=b"xlsx-data"
        ), patch.object(app, "build_invoice_pdf_bytes", return_value=b"pdf-data"):
            rebuilt, xlsx_bytes, pdf_bytes, error = app.rebuild_and_archive_history_invoice(history_row, 0)

        self.assertEqual(error, "")
        self.assertEqual(rebuilt, invoice)
        self.assertEqual(xlsx_bytes, b"xlsx-data")
        self.assertEqual(pdf_bytes, b"pdf-data")
        with open(os.path.join(self.archive_dir, "invoice_12.xlsx"), "rb") as handle:
            self.assertEqual(handle.read(), b"xlsx-data")
        with open(os.path.join(self.archive_dir, "invoice_12.pdf"), "rb") as handle:
            self.assertEqual(handle.read(), b"pdf-data")
        with open(self.ledger_path, "r") as handle:
            ledger = json.load(handle)
        self.assertEqual(ledger[0]["xlsx_filename"], "invoice_12.xlsx")
        self.assertEqual(ledger[0]["pdf_filename"], "invoice_12.pdf")

    def test_pdf_download_uses_saved_archive(self):
        os.makedirs(self.archive_dir)
        with open(os.path.join(self.archive_dir, "invoice_12.pdf"), "wb") as handle:
            handle.write(b"%PDF-test")
        row = {"invoice_number": 12, "pdf_filename": "invoice_12.pdf"}
        with patch.object(app, "ensure_data_dir"), patch.object(
            app, "invoice_history_row_at", return_value=(row, 0, [row])
        ):
            with app.app.test_client() as client:
                response = client.get("/invoice/history/0/download.pdf")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"%PDF-test")
        self.assertEqual(response.mimetype, "application/pdf")
        self.assertIn("attachment", response.headers.get("Content-Disposition", ""))

    def test_generated_invoice_removes_template_external_links(self):
        template_path = app.resolve_invoice_template_path()
        self.assertTrue(template_path)
        with open(template_path, "rb") as handle:
            cleaned = app.enforce_invoice_single_page_print_settings(handle.read())
        with zipfile.ZipFile(io.BytesIO(cleaned)) as archive:
            names = archive.namelist()
            workbook_xml = archive.read("xl/workbook.xml")
            workbook_rels = archive.read("xl/_rels/workbook.xml.rels")
        self.assertFalse(any(name.startswith("xl/externalLinks/") for name in names))
        self.assertNotIn(b"externalReferences", workbook_xml)
        self.assertNotIn(b"externalLink", workbook_rels)
        self.assertNotIn(b"[1]!Customers", workbook_xml)

    def test_saved_workbook_is_repaired_when_downloaded(self):
        os.makedirs(self.archive_dir)
        template_path = app.resolve_invoice_template_path()
        archived_path = os.path.join(self.archive_dir, "invoice_12.xlsx")
        with open(template_path, "rb") as source, open(archived_path, "wb") as target:
            target.write(source.read())
        response = app.invoice_archive_response(
            "invoice_12.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            download=True,
        )
        with zipfile.ZipFile(io.BytesIO(response.get_data())) as archive:
            self.assertFalse(any(name.startswith("xl/externalLinks/") for name in archive.namelist()))
        with zipfile.ZipFile(archived_path) as archive:
            self.assertFalse(any(name.startswith("xl/externalLinks/") for name in archive.namelist()))


if __name__ == "__main__":
    unittest.main()
