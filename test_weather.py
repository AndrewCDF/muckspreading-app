import io
import json
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

import app
import gritting
import weather


def raw_weather(surface=lambda stamp: 6.0, start="2026-11-10T00:00", days=14, rain_at=()):
    """A hand-made Open-Meteo style response for the weather page."""
    begin = datetime.fromisoformat(start)
    stamps = [begin + timedelta(hours=index) for index in range(days * 24)]
    count = len(stamps)
    wet = {stamp for stamp in rain_at}
    hourly = {"time": [stamp.strftime("%Y-%m-%dT%H:%M") for stamp in stamps]}
    for name in weather.HOURLY_FIELDS:
        hourly[name] = [10.0] * count
    hourly["soil_temperature_0cm"] = [surface(stamp) for stamp in stamps]
    hourly["temperature_2m"] = [surface(stamp) + 1 for stamp in stamps]
    hourly["dew_point_2m"] = [-5.0] * count
    hourly["precipitation"] = [0.6 if stamp.strftime("%Y-%m-%dT%H:%M") in wet else 0.0 for stamp in stamps]
    hourly["snowfall"] = [0.0] * count
    hourly["weather_code"] = [3] * count
    hourly["visibility"] = [16800.0] * count
    hourly["cloud_cover"] = [100] * count
    hourly["wind_speed_10m"] = [10.0] * count
    hourly["wind_direction_10m"] = [225] * count
    hourly["is_day"] = [1] * count
    dates = [(begin + timedelta(days=index)).strftime("%Y-%m-%d") for index in range(days)]
    daily = {"time": dates}
    for name in weather.DAILY_FIELDS:
        daily[name] = [1.0] * days
    daily["weather_code"] = [61] * days
    daily["sunrise"] = [date + "T07:13" for date in dates]
    daily["sunset"] = [date + "T16:04" for date in dates]
    daily["daylight_duration"] = [31680.0] * days
    daily["sunshine_duration"] = [18000.0] * days
    daily["wind_direction_10m_dominant"] = [270] * days
    return {"hourly": hourly, "daily": daily, "current": {
        "temperature_2m": 8.6, "apparent_temperature": 5.8, "relative_humidity_2m": 77, "precipitation": 0.0, "weather_code": 3,
        "cloud_cover": 100, "pressure_msl": 1008.7, "wind_speed_10m": 6.3, "wind_direction_10m": 72, "wind_gusts_10m": 13.4, "is_day": 0}}


class WeatherDataTests(unittest.TestCase):
    NOW = datetime(2026, 11, 10, 19, 20)

    def test_codes_and_compass(self):
        self.assertEqual(weather.describe_weather(0), ("Clear sky", "☀️"))
        self.assertEqual(weather.describe_weather(63)[0], "Rain")
        self.assertEqual(weather.describe_weather(75)[0], "Heavy snow")
        self.assertEqual(weather.describe_weather(None)[0], "Unknown")
        self.assertEqual([weather.compass_point(value) for value in (0, 72, 90, 180, 225, 270, 348.75, 359)], ["N", "ENE", "E", "S", "SW", "W", "N", "N"])
        self.assertEqual(weather.compass_point(None), "")

    def test_wind_arrows_point_the_way_the_wind_blows(self):
        # the letters give where the wind comes from; the arrow points the other way
        cases = [(0, "N ↓"), (45, "NE ↙"), (72, "ENE ←"), (90, "E ←"), (135, "SE ↖"), (180, "S ↑"), (225, "SW ↗"), (270, "W →"), (315, "NW ↘"), (359, "N ↓")]
        for degrees, expected in cases:
            with self.subTest(degrees=degrees):
                self.assertEqual(weather.wind_direction(degrees), expected)
        self.assertEqual((weather.wind_direction(None), weather.wind_arrow("x")), ("", ""))

    def test_fourteen_days_each_with_their_hours_and_current_conditions(self):
        data = weather.build_weather(raw_weather(), self.NOW)
        self.assertEqual(len(data["days"]), 14)
        self.assertEqual(data["hours"][0]["t"], "2026-11-10T19:00")  # past hours of today are dropped
        self.assertEqual(len(data["hours"]), 14 * 24 - 19)
        first, second = data["days"][0], data["days"][1]
        self.assertEqual((first["start"], first["end"]), (0, 5))
        self.assertEqual((second["start"], second["end"]), (5, 29))
        self.assertEqual(first["label"], "Tue 10 Nov")
        self.assertEqual((first["sunrise"], first["sunset"], first["daylight_hours"], first["sunshine_hours"]), ("07:13", "16:04", 8.8, 5.0))
        self.assertEqual((first["description"], first["wind_compass"]), ("Light rain", "W →"))
        self.assertEqual((first["rain_mm"], first["rain_hours"], first["prob_max"]), (1.0, 1, 1))  # the whole day's rain, not per hour
        current = data["current"]
        self.assertEqual((current["temp"], current["feels"], current["description"], current["compass"]), (8.6, 5.8, "Overcast", "ENE ←"))
        self.assertEqual(current["surface"], 6.0)
        hour = data["hours"][0]
        self.assertEqual((hour["compass"], hour["visibility_km"], hour["icon"], hour["level"]), ("SW ↗", 16.8, "☁️", "ok"))

    def test_high_and_low_come_with_the_time_they_happen(self):
        def surface(stamp):
            return 10.0 if stamp.hour == 15 else (0.0 if stamp.hour == 4 else 5.0)
        data = weather.build_weather(raw_weather(surface), self.NOW)  # 19:20, so the 15:00 high is already past
        today, tomorrow = data["days"][0], data["days"][1]
        self.assertEqual((today["temp_max_time"], today["temp_min_time"]), ("15:00", "04:00"))
        self.assertEqual((tomorrow["temp_max_time"], tomorrow["temp_min_time"]), ("15:00", "04:00"))
        self.assertEqual((data["current"]["high_time"], data["current"]["low_time"]), ("15:00", "04:00"))
        self.assertEqual((today["feels_max_time"], today["feels_min_time"]), ("00:00", "00:00"))  # feels-like is flat in the fixture: earliest hour
        raw = raw_weather(surface)
        raw["hourly"]["temperature_2m"][10] = None
        self.assertTrue(weather.build_weather(raw, self.NOW)["days"][0]["temp_max_time"])

    def test_right_now_has_the_extra_detail(self):
        data = weather.build_weather(raw_weather(), self.NOW)
        now = data["current"]
        self.assertEqual((now["high"], now["low"], now["sunrise"], now["sunset"], now["uv_max"]), (1.0, 1.0, "07:13", "16:04", 1.0))
        self.assertEqual((now["pressure_trend"], now["pressure_change"], now["ground6"]), ("Steady", 0.0, 10.0))
        self.assertEqual((now["next_rain"], now["rain_next_6h"], now["rain_today"]), (None, 0.0, 1.0))
        self.assertEqual((now["tonight"]["level"], now["tonight"]["min_surface"]), ("ok", 6.0))
        raw = raw_weather(rain_at=["2026-11-10T21:00", "2026-11-10T22:00"])
        wet = weather.build_weather(raw, self.NOW)["current"]
        self.assertEqual(wet["next_rain"], {"from": "Tue 21:00", "until": "Tue 23:00", "mm": 1.2, "starts_in_hours": 2})
        self.assertEqual(wet["rain_next_6h"], 1.2)
        raining = weather.build_weather(raw_weather(rain_at=["2026-11-10T19:00", "2026-11-10T20:00"]), self.NOW)["current"]
        self.assertEqual(raining["next_rain"]["starts_in_hours"], 0)
        self.assertEqual(raining["next_rain"]["until"], "Tue 21:00")

    def test_pressure_trend(self):
        hours = [{"pressure": 1000.0}, {"pressure": 1000.5}, {"pressure": 1001.0}, {"pressure": 1002.0}]
        self.assertEqual(weather.pressure_change(hours), {"trend": "Rising", "change": 2.0})
        self.assertEqual(weather.pressure_change([{"pressure": 1002.0}, {}, {}, {"pressure": 1000.5}])["trend"], "Falling")
        self.assertEqual(weather.pressure_change([{"pressure": 1002.0}, {}, {}, {"pressure": 1002.4}])["trend"], "Steady")
        self.assertEqual(weather.pressure_change([{"pressure": None}]), {"trend": "", "change": None})

    def test_the_weather_page_uses_the_gritting_pages_forecast_where_it_has_one(self):
        import gritting
        from test_gritting import raw_forecast
        mild_here = raw_weather(lambda stamp: -3.0)  # this page's own forecast says a hard frost every night
        road_raw = {"hourly": {}}
        parts = {}
        for model in ("ukmo_seamless", "ecmwf_ifs025", "icon_seamless"):
            part = raw_forecast(lambda stamp: 6.0, start="2026-11-10T00:00", hours=24 * 8)["hourly"]
            parts.setdefault("time", part["time"])
            for key, values in part.items():
                if key != "time":
                    parts[f"{key}_{model}"] = values
        road = gritting.road_forecast_from_open_meteo({"hourly": parts}, self.NOW)
        alone = {day["date"]: day for day in weather.build_weather(mild_here, self.NOW)["days"]}
        shared = {day["date"]: day for day in weather.build_weather(mild_here, self.NOW, road=road)["days"]}
        self.assertEqual(alone["2026-11-11"]["night"]["level"], "grit")
        self.assertEqual(shared["2026-11-11"]["night"]["level"], "ok")  # the three models say mild, so both pages agree
        self.assertEqual(shared["2026-11-11"]["night"]["min_surface"], 6.0)
        self.assertEqual(shared["2026-11-21"]["night"]["level"], "grit")  # beyond the gritting forecast it keeps its own

    def test_road_surface_and_gritting_verdicts_per_day(self):
        freezing = lambda stamp: -1.0 if datetime(2026, 11, 13, 0) <= stamp < datetime(2026, 11, 13, 6) else 6.0
        data = weather.build_weather(raw_weather(freezing), self.NOW)
        by_date = {day["date"]: day for day in data["days"]}
        thursday_night = by_date["2026-11-12"]["night"]  # 18:00 Thursday to 06:00 Friday
        self.assertEqual((thursday_night["level"], thursday_night["min_surface"]), ("grit", -1.0))
        # the cold hours (00:00 to 06:00 Friday) belong to Thursday's night, which runs 18:00 to 06:00
        self.assertEqual((thursday_night["min_surface_day"], thursday_night["min_surface_time"]), ("Fri", "00:00"))
        self.assertTrue(by_date["2026-11-12"]["ground_frost"])
        friday = by_date["2026-11-13"]
        self.assertEqual(friday["min_surface"], -1.0)  # midnight to midnight still sees them
        self.assertFalse(friday["ground_frost"])  # but Friday night itself is mild
        self.assertTrue(any(hour["level"] == "grit" for hour in data["hours"][friday["start"]:friday["end"]]))
        self.assertEqual(by_date["2026-11-11"]["night"]["level"], "ok")
        self.assertFalse(by_date["2026-11-11"]["ground_frost"])

    def test_overnight_low_covers_the_whole_night_not_midnight_to_midnight(self):
        # warm evening, coldest at 05:00 the next morning: the night's low must be that 05:00 figure
        cold_morning = lambda stamp: -2.0 if stamp == datetime(2026, 11, 12, 5) else (4.0 if stamp.hour < 6 or stamp.hour >= 18 else 8.0)
        data = weather.build_weather(raw_weather(cold_morning), self.NOW)
        wednesday = {day["date"]: day for day in data["days"]}["2026-11-11"]
        night = wednesday["night"]
        self.assertEqual((night["min_surface"], night["min_surface_day"], night["min_surface_time"]), (-2.0, "Thu", "05:00"))
        self.assertEqual(wednesday["min_surface"], 4.0)  # the calendar day alone would have hidden the frost
        self.assertTrue(wednesday["ground_frost"])

    def test_after_midnight_tonight_is_the_night_already_in_progress(self):
        def surface(stamp):
            return -1.0 if datetime(2026, 11, 11, 1) <= stamp < datetime(2026, 11, 11, 6) else 7.0
        early = weather.build_weather(raw_weather(surface, start="2026-11-10T00:00"), datetime(2026, 11, 11, 2, 10))["current"]["tonight"]
        self.assertEqual((early["label"], early["level"], early["min_surface"], early["min_surface_day"]), ("This morning, until 06:00", "grit", -1.0, "Wed"))
        evening = weather.build_weather(raw_weather(surface, start="2026-11-11T00:00"), datetime(2026, 11, 11, 19, 10))["current"]["tonight"]
        self.assertEqual((evening["label"], evening["level"]), ("Tonight, 18:00 to 06:00", "ok"))

    def test_rain_advice_flows_through_to_the_day(self):
        freezing = lambda stamp: -1.0 if datetime(2026, 11, 11, 1) <= stamp < datetime(2026, 11, 11, 6) else 6.0
        data = weather.build_weather(raw_weather(freezing, rain_at=["2026-11-10T21:00", "2026-11-10T22:00"]), self.NOW)
        night = data["days"][0]["night"]
        self.assertEqual(night["level"], "grit")
        self.assertIn("grit after it stops", night["rain_advice"])

    def test_missing_values_do_not_break_the_page_data(self):
        raw = raw_weather()
        raw["hourly"]["wind_gusts_10m"] = None
        raw["hourly"]["soil_temperature_6cm"][30] = None
        raw["daily"]["uv_index_max"] = []
        raw["current"] = {}
        data = weather.build_weather(raw, self.NOW)
        self.assertIsNone(data["days"][0]["uv_max"])
        self.assertIsNone(data["hours"][0]["gust"])
        self.assertEqual(data["current"]["description"], "Unknown")

    def test_forecast_is_for_postcode_nr15_1be(self):
        url = weather.weather_url()
        self.assertIn("latitude=52.5693", url)
        self.assertIn("longitude=1.4067", url)
        self.assertEqual(weather.WEATHER_PLACE, "NR15 1BE, Bergh Apton")
        # the gritting forecast stays on Norwich
        self.assertEqual(gritting.FORECAST_PLACE, "Norwich")
        self.assertIn("latitude=52.63&", gritting.forecast_request_url())
        self.assertIn("longitude=1.3&", gritting.forecast_request_url())

    def test_forecast_is_requested_for_fourteen_days_with_everything(self):
        url = weather.weather_url()
        for text in ("forecast_days=14", "soil_temperature_0cm", "apparent_temperature", "wind_gusts_10m", "pressure_msl", "sunrise", "uv_index_max", "current="):
            self.assertIn(text, url)
        with patch.dict(weather.os.environ if hasattr(weather, "os") else __import__("os").environ, {"OPEN_METEO_API_KEY": "abc123"}):
            self.assertTrue(weather.weather_url().startswith("https://customer-api.open-meteo.com/"))

    def test_cache_and_stale_fallback(self):
        weather._cache.update(at=0.0, data=None)
        raw = raw_weather()
        calls = []

        def fake_open(request_obj, timeout=None):
            calls.append(request_obj.full_url)
            return io.BytesIO(json.dumps(raw).encode())
        with patch.object(weather.urllib.request, "urlopen", fake_open):
            first, stale = weather.fetch_weather()
            weather.fetch_weather()
        self.assertEqual((len(calls), stale), (1, False))
        weather._cache["at"] -= weather.WEATHER_CACHE_SECONDS + 1

        def failing(request_obj, timeout=None):
            raise OSError("offline")
        with patch.object(weather.urllib.request, "urlopen", failing):
            again, stale = weather.fetch_weather()
        self.assertEqual((again, stale), (first, True))
        weather._cache["at"] -= weather.WEATHER_STALE_LIMIT_SECONDS
        with patch.object(weather.urllib.request, "urlopen", failing):
            with self.assertRaises(OSError):
                weather.fetch_weather()
        weather._cache.update(at=0.0, data=None)


class WeatherPageTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()

    def test_home_button_and_page(self):
        self.assertIn(b'href="/weather/">Weather</a>', self.client.get("/").data)
        page = self.client.get("/weather/")
        self.assertEqual(page.status_code, 200)
        for text in (b"Weather", b"Daily forecast", b"class=\"wx-rows\"", b"/weather/forecast.json"):
            self.assertIn(text, page.data)
        self.assertIn(b"window.history.back()", page.data)
        self.assertNotIn(b"About this forecast", page.data)
        self.assertNotIn(b"Gritting outlook", page.data)
        self.assertNotIn(b"Day by day", page.data)
        for path in ("/static/weather.js", "/static/weather.css"):
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200)

    def test_forecast_endpoint(self):
        with patch.object(weather, "fetch_weather", return_value=(raw_weather(), False)), patch.object(gritting, "local_now", return_value=datetime(2026, 11, 10, 19, 20)):
            payload = self.client.get("/weather/forecast.json").get_json()
        self.assertTrue(payload["ok"])
        self.assertEqual((len(payload["days"]), payload["place"]), (14, "NR15 1BE, Bergh Apton"))
        with patch.object(weather, "fetch_weather", side_effect=OSError("offline")):
            failed = self.client.get("/weather/forecast.json")
        self.assertEqual(failed.status_code, 200)
        self.assertFalse(failed.get_json()["ok"])

    def test_backup_includes_the_weather_module(self):
        import tempfile
        import zipfile
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(app, "DATA_DIR", directory), patch.object(app, "ISSUE_PHOTOS_DIR", directory + "/photos"), patch.object(app, "INVOICE_ARCHIVE_DIR", directory + "/invoices"):
                backup = self.client.get("/backup/export.zip")
        self.assertIn("weather.py", zipfile.ZipFile(io.BytesIO(backup.data)).namelist())


if __name__ == "__main__":
    unittest.main()
