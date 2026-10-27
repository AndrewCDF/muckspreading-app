import io
import re
import unittest
import zipfile

import app


def sample_invoice(line_count):
    lines = [
        {"job_date_label": "01/10/2026", "farm_name": "Home farm", "field_name": "Field %d" % (index + 1), "muck_type": "FYM",
         "tons": "20", "rate_per_ton": "3.25", "line_total": 65.0, "vat_rate": 20}
        for index in range(line_count)
    ]
    return {
        "invoice_number_label": "5390", "invoice_date_label": "27/10/2026", "invoice_date": "2026-10-27", "payment_terms_days": "14",
        "customer_name": "MR P Thain", "address_line_1": "Heath Road", "town": "Gt. Yarmouth", "postcode": "NR25 5HJ",
        "line_rows": lines, "rate_override_label": "£3.25", "filename": "invoice-5390.xlsx", "total_tons": 20 * line_count,
    }


class InvoiceFooterTests(unittest.TestCase):
    def sheet(self, line_count):
        data = app.build_invoice_layout_xlsx_bytes(sample_invoice(line_count))
        return zipfile.ZipFile(io.BytesIO(data)).read("xl/worksheets/sheet1.xml").decode()

    def test_company_details_are_the_page_footer_not_rows_in_the_body(self):
        for line_count in (1, 24, 30):
            with self.subTest(lines=line_count):
                sheet = self.sheet(line_count)
                footer = re.search(r"<oddFooter>(.*?)</oddFooter>", sheet, re.S).group(1)
                for text in ("Company Details", "Company No.: 4761116", "VAT No.: 731972228", "Payment Details", "Sort Code: 53-61-24"):
                    self.assertIn(text, footer)
                body = "".join(re.findall(r"<t[^>]*>(.*?)</t>", sheet, re.S))
                self.assertNotIn("Company Details", body)
                self.assertNotIn("731972228", body)
                self.assertEqual(re.findall(r"<(pageMargins|pageSetup|headerFooter|drawing)\b", sheet), ["pageMargins", "pageSetup", "headerFooter", "drawing"])
                self.assertEqual(len(re.findall(r"<mergeCell ref", sheet)), int(re.search(r'<mergeCells count="(\d+)"', sheet).group(1)))

    def test_footer_text_escapes_ampersands_and_numbers(self):
        text = app.invoice_footer_text([["Company Details", "07952 683364", "Name: A & B"], ["Payment Details"]])
        self.assertIn("A && B", text)
        self.assertIn("&8 07952 683364", text)  # a space keeps the digits out of the font-size code
        self.assertTrue(text.startswith("&L"))
        self.assertIn("&R", text)

    def test_fallback_pdf_has_the_details_footer_on_every_page(self):
        original = app.convert_xlsx_bytes_to_pdf_bytes
        app.convert_xlsx_bytes_to_pdf_bytes = lambda *args, **kwargs: b""
        try:
            pdf = app.build_invoice_pdf_bytes(sample_invoice(20))
        finally:
            app.convert_xlsx_bytes_to_pdf_bytes = original
        self.assertEqual(pdf.count(b"(Company Details)"), 2)
        self.assertEqual(pdf.count(b"(Payment Details)"), 2)


if __name__ == "__main__":
    unittest.main()
