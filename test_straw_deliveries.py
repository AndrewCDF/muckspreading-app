import json
import os
import tempfile
import unittest
from unittest.mock import patch

import app


class StrawDeliveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.data_patch = patch.object(app, "DATA_DIR", self.temp.name)
        self.data_patch.start()
        self.state_patch = patch.object(app, "STRAW_STATE_PATH", os.path.join(self.temp.name, "straw-records.json"))
        self.state_patch.start()
        self.client = app.app.test_client()
        self.delivery = {
            "customer": "Example Farm",
            "registration": "ab12 cde",
            "delivery_date": "2026-09-16",
            "delivery_time": "14:35",
            "bale_total": "42",
            "weight_total": "",
        }

    def tearDown(self):
        self.data_patch.stop()
        self.state_patch.stop()
        self.temp.cleanup()

    def test_create_then_add_weight_later(self):
        created = self.client.post("/api/straw/deliveries", json=self.delivery)
        self.assertEqual(created.status_code, 201)
        saved = created.get_json()["delivery"]
        self.assertEqual(saved["registration"], "AB12 CDE")
        self.assertEqual(saved["bale_total"], 42)
        self.assertEqual(saved["weight_total"], "")
        self.assertEqual(self.client.get("/api/straw/customers").get_json()["customers"], ["Example Farm"])

        saved["weight_total"] = "18.50"
        updated = self.client.put("/api/straw/deliveries/%s" % saved["id"], json=saved)
        self.assertEqual(updated.status_code, 200)
        current = updated.get_json()["delivery"]
        self.assertEqual(current["weight_total"], "18.5")
        self.assertEqual(current["version"], 2)
        self.assertEqual(self.client.get("/api/straw/deliveries").get_json()["deliveries"], [current])

        self.assertEqual(self.client.put("/api/straw/deliveries/%s" % saved["id"], json=saved).status_code, 409)

    def test_invalid_delivery_is_rejected(self):
        invalid_changes = [
            {"customer": ""}, {"registration": ""}, {"delivery_date": "2026-02-30"},
            {"delivery_time": "25:00"}, {"bale_total": "2.5"}, {"bale_total": "-1"}, {"weight_total": "unknown"},
        ]
        for changes in invalid_changes:
            with self.subTest(changes=changes):
                response = self.client.post("/api/straw/deliveries", json=dict(self.delivery, **changes))
                self.assertEqual(response.status_code, 400)
        self.assertEqual(self.client.get("/api/straw/deliveries").get_json()["deliveries"], [])

    def test_straw_page_has_home_delivery_action_and_no_stock_load_button(self):
        page = self.client.get("/straw")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b'id="addDeliveryButton"', page.data)
        self.assertIn(b'id="deliveryForm"', page.data)
        self.assertIn(b'id="recentDeliveries"', page.data)
        self.assertIn(b'id="cropTotals"', page.data)
        self.assertIn(b'id="customerTotals"', page.data)
        self.assertIn(b'Recent Loads Out', page.data)
        self.assertIn(b'Log Load Out', page.data)
        self.assertIn(b'id="deliveryDateTime" type="datetime-local"', page.data)
        self.assertIn(b'id="straw-customers"', page.data)
        self.assertNotIn(b'id="settings-customers"', page.data)
        self.assertNotIn(b'id="delivery-customers"', page.data)
        self.assertNotIn(b'Add load removed', page.data)
        for element_id in (b'addStocktakeButton', b'stocktakeDialog', b'stocktakeHistory', b'completedLoads', b'addStockMovementButton', b'stockMovementsList'):
            self.assertIn(b'id="' + element_id + b'"', page.data)

    def test_straw_customer_list_starts_separate_and_remembers_names(self):
        self.assertEqual(self.client.get("/api/straw/customers").get_json()["customers"], [])
        first = self.client.post("/api/straw/customers", json={"customer": "New Hay Customer"})
        self.assertEqual(first.status_code, 201)
        self.client.post("/api/straw/customers", json={"customer": "new hay customer"})
        self.assertEqual(self.client.get("/api/straw/customers").get_json()["customers"], ["new hay customer"])

    def test_old_straw_records_file_is_loaded_and_legacy_loads_are_migrated_once(self):
        legacy = {
            "fields": [{"id": "field-1", "customer": "Old Straw Customer", "crop": "Wheat", "bales": 120}],
            "stocktakes": [{"id": "stock-1", "bales": 120}],
            "loads": [{"id": "load-1", "date": "2026-08-14T13:45:00.000Z", "bales": 36, "vehicleReg": "ab12 cde", "weight": "15.75"}],
            "stockMovements": [{"id": "movement-1", "type": "bought-in", "bales": 10}],
        }
        with open(app.STRAW_STATE_PATH, "w") as handle:
            json.dump(legacy, handle)
        state = self.client.get("/api/straw/state").get_json()
        self.assertEqual(state, legacy)
        deliveries = self.client.get("/api/straw/deliveries").get_json()["deliveries"]
        self.assertEqual(len(deliveries), 1)
        self.assertEqual(deliveries[0]["registration"], "AB12 CDE")
        self.assertEqual(deliveries[0]["bale_total"], 36)
        self.assertEqual(deliveries[0]["weight_total"], "15.75")
        self.assertEqual(deliveries[0]["customer"], "Customer not recorded")
        self.assertEqual(len(self.client.get("/api/straw/deliveries").get_json()["deliveries"]), 1)

        changed = dict(legacy)
        changed["fields"] = legacy["fields"] + [{"id": "field-2", "customer": "Second", "crop": "Hay", "bales": 20}]
        changed["loads"] = []
        self.assertEqual(self.client.put("/api/straw/state", json=changed).status_code, 200)
        saved = self.client.get("/api/straw/state").get_json()
        self.assertEqual(len(saved["fields"]), 2)
        self.assertEqual(saved["loads"], legacy["loads"])


if __name__ == "__main__":
    unittest.main()
