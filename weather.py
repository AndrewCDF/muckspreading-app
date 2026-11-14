"""Detailed 14-day weather for the farm, including model road surface temperatures.

Data is the hourly and daily forecast from Open-Meteo (Met Office UK model for the first days, then other
global models). The road surface figure is the model's surface (0 cm) temperature, so it is a guide and not a
road sensor reading. The gritting risk rules are shared with the gritting module so both pages agree."""
from __future__ import annotations

import json
import threading
import time
import urllib.request
from datetime import datetime, timedelta
from typing import Any

from flask import Blueprint, jsonify, render_template

import gritting

bp = Blueprint("weather", __name__, url_prefix="/weather")

# The weather page is for the farm; the gritting forecast has its own location (Norwich) in gritting.py.
WEATHER_PLACE = "NR15 1BE, Bergh Apton"
WEATHER_LATITUDE = 52.5693  # postcode NR15 1BE
WEATHER_LONGITUDE = 1.4067
WEATHER_CACHE_SECONDS = 20 * 60
WEATHER_STALE_LIMIT_SECONDS = 6 * 60 * 60
FORECAST_DAYS = 14
_cache: dict[str, Any] = {"at": 0.0, "data": None}
_lock = threading.Lock()

HOURLY_FIELDS = [
    "temperature_2m", "relative_humidity_2m", "dew_point_2m", "apparent_temperature", "precipitation_probability",
    "precipitation", "snowfall", "snow_depth", "weather_code", "pressure_msl", "cloud_cover", "visibility",
    "wind_speed_10m", "wind_direction_10m", "wind_gusts_10m", "soil_temperature_0cm", "soil_temperature_6cm",
    "uv_index", "is_day", "sunshine_duration",
]
DAILY_FIELDS = [
    "weather_code", "temperature_2m_max", "temperature_2m_min", "apparent_temperature_max", "apparent_temperature_min",
    "sunrise", "sunset", "daylight_duration", "sunshine_duration", "uv_index_max", "precipitation_sum", "rain_sum",
    "showers_sum", "snowfall_sum", "precipitation_hours", "precipitation_probability_max", "wind_speed_10m_max",
    "wind_gusts_10m_max", "wind_direction_10m_dominant",
]
CURRENT_FIELDS = [
    "temperature_2m", "apparent_temperature", "relative_humidity_2m", "precipitation", "weather_code", "cloud_cover",
    "pressure_msl", "wind_speed_10m", "wind_direction_10m", "wind_gusts_10m", "is_day",
]

# WMO weather codes: description and a simple icon.
WEATHER_CODES = {
    0: ("Clear sky", "☀️"), 1: ("Mainly clear", "🌤️"), 2: ("Partly cloudy", "⛅"), 3: ("Overcast", "☁️"),
    45: ("Fog", "🌫️"), 48: ("Freezing fog", "🌫️"),
    51: ("Light drizzle", "🌦️"), 53: ("Drizzle", "🌦️"), 55: ("Heavy drizzle", "🌧️"),
    56: ("Light freezing drizzle", "🌧️"), 57: ("Freezing drizzle", "🌧️"),
    61: ("Light rain", "🌦️"), 63: ("Rain", "🌧️"), 65: ("Heavy rain", "🌧️"),
    66: ("Light freezing rain", "🌧️"), 67: ("Freezing rain", "🌧️"),
    71: ("Light snow", "🌨️"), 73: ("Snow", "❄️"), 75: ("Heavy snow", "❄️"), 77: ("Snow grains", "🌨️"),
    80: ("Light showers", "🌦️"), 81: ("Showers", "🌧️"), 82: ("Heavy showers", "🌧️"),
    85: ("Snow showers", "🌨️"), 86: ("Heavy snow showers", "❄️"),
    95: ("Thunderstorm", "⛈️"), 96: ("Thunderstorm with hail", "⛈️"), 99: ("Thunderstorm with heavy hail", "⛈️"),
}
COMPASS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]


def describe_weather(code: Any) -> tuple[str, str]:
    try:
        return WEATHER_CODES.get(int(code), ("Unknown", "·"))
    except (TypeError, ValueError):
        return ("Unknown", "·")


def compass_point(degrees: Any) -> str:
    try:
        return COMPASS[int(round(float(degrees) / 22.5)) % 16]
    except (TypeError, ValueError):
        return ""


def wind_arrow(degrees: Any) -> str:
    """An arrow for the way the wind is blowing (the opposite of the direction it comes from)."""
    try:
        towards = (float(degrees) + 180.0) % 360.0
    except (TypeError, ValueError):
        return ""
    return ["↑", "↗", "→", "↘", "↓", "↙", "←", "↖"][int(round(towards / 45.0)) % 8]


def wind_direction(degrees: Any) -> str:
    """The compass abbreviation with its arrow, such as "NE ↙" for a wind from the north east."""
    point = compass_point(degrees)
    return f"{point} {wind_arrow(degrees)}" if point else ""


def weather_url() -> str:
    import os

    query = (
        f"latitude={WEATHER_LATITUDE}&longitude={WEATHER_LONGITUDE}"
        f"&current={','.join(CURRENT_FIELDS)}&hourly={','.join(HOURLY_FIELDS)}&daily={','.join(DAILY_FIELDS)}"
        f"&forecast_days={FORECAST_DAYS}&timezone=Europe%2FLondon&wind_speed_unit=mph"
    )
    api_key = os.environ.get("OPEN_METEO_API_KEY", "").strip()
    if api_key:  # a paid Open-Meteo plan, for commercial use
        return f"{gritting.FORECAST_COMMERCIAL_URL}?{query}&apikey={api_key}"
    return f"{gritting.FORECAST_URL}?{query}"


def fetch_weather(timeout: float = 15.0) -> tuple[dict[str, Any], bool]:
    """Return (raw forecast, is_stale). Cached for a few minutes; an older copy is used if the service is down."""
    with _lock:
        age = time.time() - _cache["at"]
        if _cache["data"] is not None and age < WEATHER_CACHE_SECONDS:
            return _cache["data"], False
        try:
            request_obj = urllib.request.Request(weather_url(), headers={"User-Agent": "farm-weather-app"})
            with urllib.request.urlopen(request_obj, timeout=timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
            if not isinstance(data.get("hourly"), dict) or not isinstance(data.get("daily"), dict):
                raise ValueError("The forecast response was incomplete.")
        except Exception:
            if _cache["data"] is not None and age < WEATHER_STALE_LIMIT_SECONDS:
                return _cache["data"], True
            raise
        _cache["at"] = time.time()
        _cache["data"] = data
        return data, False


def _number(values: list[Any], index: int, digits: int = 1) -> float | None:
    if index >= len(values) or values[index] is None:
        return None
    try:
        return round(float(values[index]), digits)
    except (TypeError, ValueError):
        return None


def _clock(stamp: Any) -> str:
    try:
        return datetime.fromisoformat(str(stamp)).strftime("%H:%M")
    except ValueError:
        return ""


def next_wet_spell(hours: list[dict[str, Any]]) -> dict[str, Any] | None:
    """The first run of wet hours from now, or None if the whole forecast is dry."""
    spell: list[dict[str, Any]] = []
    for hour in hours:
        wet = (hour["rain"] or 0) >= 0.1 or (hour["snow"] or 0) > 0
        if wet:
            if spell and datetime.fromisoformat(hour["t"]) - datetime.fromisoformat(spell[-1]["t"]) > timedelta(hours=1):
                break
            spell.append(hour)
        elif spell:
            break
    if not spell:
        return None
    end = datetime.fromisoformat(spell[-1]["t"]) + timedelta(hours=1)
    return {"from": f'{spell[0]["day"]} {spell[0]["hour"]}', "until": f'{end.strftime("%a %H:%M")}',
            "mm": round(sum(float(hour["rain"] or 0) + float(hour["snow"] or 0) * 10 for hour in spell), 1),
            "starts_in_hours": 0 if spell[0] is hours[0] else hours.index(spell[0])}


def pressure_change(hours: list[dict[str, Any]], ahead: int = 3) -> dict[str, Any]:
    """How pressure is expected to move over the next few hours."""
    if len(hours) <= ahead or hours[0]["pressure"] is None or hours[ahead]["pressure"] is None:
        return {"trend": "", "change": None}
    change = round(hours[ahead]["pressure"] - hours[0]["pressure"], 1)
    trend = "Rising" if change >= 1 else "Falling" if change <= -1 else "Steady"
    return {"trend": trend, "change": change}


def build_weather(raw: dict[str, Any], now: datetime, bias: float = 0.0, road: dict[str, Any] | None = None) -> dict[str, Any]:
    hourly, daily = raw["hourly"], raw["daily"]
    times = hourly.get("time") or []
    this_hour = now.replace(minute=0, second=0, microsecond=0)

    def series(name: str) -> list[Any]:
        values = hourly.get(name) or []
        return list(values) + [None] * (len(times) - len(values))

    columns = {name: series(name) for name in HOURLY_FIELDS}

    # When each day's highest and lowest temperature happen, from every hour of the day (including hours already past).
    extremes: dict[str, dict[str, tuple[float, str]]] = {}
    for index, stamp in enumerate(times):
        for key, name in (("air", "temperature_2m"), ("feels", "apparent_temperature")):
            value = _number(columns[name], index)
            if value is None:
                continue
            slot = extremes.setdefault(stamp[:10], {})
            clock = stamp[11:16]
            if key + "_max" not in slot or value > slot[key + "_max"][0]:
                slot[key + "_max"] = (value, clock)
            if key + "_min" not in slot or value < slot[key + "_min"][0]:
                slot[key + "_min"] = (value, clock)
    hours: list[dict[str, Any]] = []
    for index, stamp in enumerate(times):
        try:
            moment = datetime.fromisoformat(stamp)
        except ValueError:
            continue
        if moment < this_hour:
            continue
        surface = _number(columns["soil_temperature_0cm"], index)
        if surface is not None:
            surface = round(surface + bias, 1)  # corrected by what your own road readings show
        air = _number(columns["temperature_2m"], index)
        dew = _number(columns["dew_point_2m"], index)
        rain = _number(columns["precipitation"], index) or 0.0
        snow = _number(columns["snowfall"], index) or 0.0
        level = ""
        if surface is not None and air is not None:
            level = gritting.classify_road_hour({
                "surface": surface, "air": air, "dew": dew, "rain": rain, "snow": snow,
                "cloud": _number(columns["cloud_cover"], index, 0), "wind": _number(columns["wind_speed_10m"], index),
            })[0]
        icon = describe_weather(columns["weather_code"][index])[1]
        visibility = _number(columns["visibility"], index, 0)
        hours.append({
            "t": stamp, "hour": moment.strftime("%H:%M"), "day": moment.strftime("%a"), "date": moment.strftime("%Y-%m-%d"),
            "temp": air, "feels": _number(columns["apparent_temperature"], index), "dew": dew,
            "humidity": _number(columns["relative_humidity_2m"], index, 0), "rain": rain, "snow": snow,
            "prob": _number(columns["precipitation_probability"], index, 0), "cloud": _number(columns["cloud_cover"], index, 0),
            "wind": _number(columns["wind_speed_10m"], index), "gust": _number(columns["wind_gusts_10m"], index),
            "compass": wind_direction(columns["wind_direction_10m"][index]),
            "pressure": _number(columns["pressure_msl"], index), "visibility_km": None if visibility is None else round(visibility / 1000, 1),
            "surface": surface, "surface6": _number(columns["soil_temperature_6cm"], index), "uv": _number(columns["uv_index"], index),
            "icon": icon, "level": level,
        })

    # Where the gritting page has its own three-model road forecast, use it here too, so both pages give the same answer.
    if road:
        road_hours = {(item["date"], item["hour"]): item for item in road.get("hourly", [])}
        for hour in hours:
            match = road_hours.get((hour["date"], hour["hour"]))
            if match:
                hour["surface"], hour["level"] = match["surface"], match["level"]

    # Gritting verdict per night (18:00 to 06:00), the same rules as the gritting page.
    night_forecast = gritting.road_forecast_from_open_meteo(raw, now, night_count=FORECAST_DAYS, bias=bias)
    nights_by_date = {night["date"]: night for night in night_forecast["nights"]}
    for night in (road or {}).get("nights", []):
        nights_by_date[night["date"]] = night

    days = []
    for index, date in enumerate(daily.get("time") or []):
        indexes = [position for position, hour in enumerate(hours) if hour["date"] == date]
        day_hours = [hours[position] for position in indexes]
        surfaces = [hour["surface"] for hour in day_hours if hour["surface"] is not None]
        coldest = min((hour for hour in day_hours if hour["surface"] is not None), key=lambda hour: hour["surface"], default=None)
        description, icon = describe_weather(_number(daily.get("weather_code") or [], index, 0))
        column = lambda name, digits=1: _number(daily.get(name) or [], index, digits)
        sunshine, daylight = column("sunshine_duration", 0), column("daylight_duration", 0)
        moment = datetime.fromisoformat(date)
        days.append({
            "date": date, "label": moment.strftime("%a %d %b"), "short": moment.strftime("%a"),
            "start": indexes[0] if indexes else None, "end": indexes[-1] + 1 if indexes else None,
            "description": description, "icon": icon, "temp_max": column("temperature_2m_max"), "temp_min": column("temperature_2m_min"),
            "feels_max": column("apparent_temperature_max"), "feels_min": column("apparent_temperature_min"),
            "temp_max_time": extremes.get(date, {}).get("air_max", (None, ""))[1], "temp_min_time": extremes.get(date, {}).get("air_min", (None, ""))[1],
            "feels_max_time": extremes.get(date, {}).get("feels_max", (None, ""))[1], "feels_min_time": extremes.get(date, {}).get("feels_min", (None, ""))[1],
            "sunrise": _clock((daily.get("sunrise") or [None] * (index + 1))[index]), "sunset": _clock((daily.get("sunset") or [None] * (index + 1))[index]),
            "daylight_hours": None if daylight is None else round(daylight / 3600, 1),
            "sunshine_hours": None if sunshine is None else round(sunshine / 3600, 1),
            "uv_max": column("uv_index_max"), "rain_mm": column("precipitation_sum"), "showers_mm": column("showers_sum"),
            "snow_cm": column("snowfall_sum"), "rain_hours": column("precipitation_hours", 0), "prob_max": column("precipitation_probability_max", 0),
            "wind_max": column("wind_speed_10m_max"), "gust_max": column("wind_gusts_10m_max"),
            "wind_compass": wind_direction((daily.get("wind_direction_10m_dominant") or [None] * (index + 1))[index]),
            "min_surface": None if not surfaces else round(min(surfaces), 1),
            "min_surface_time": coldest["hour"] if coldest else "",
            "night": nights_by_date.get(date),
        })
        # Gritting is overnight for the next morning, so frost is judged on the whole night (18:00 to 06:00),
        # not on midnight to midnight.
        night = days[-1]["night"]
        days[-1]["ground_frost"] = bool(night) and night["min_surface"] <= gritting.GRIT_TRIGGER_C
    current = raw.get("current") or {}
    description, icon = describe_weather(current.get("weather_code"))
    now_hour = hours[0] if hours else {}
    today = days[0] if days else {}
    pressure = pressure_change(hours)
    # After midnight the night already in progress is the one that matters, so show that until 06:00.
    in_progress = now.hour < 6
    tonight = (nights_by_date.get((now - timedelta(days=1)).strftime("%Y-%m-%d")) if in_progress else today.get("night")) or {}
    return {
        "place": WEATHER_PLACE,
        "current": {
            "temp": current.get("temperature_2m"), "feels": current.get("apparent_temperature"), "humidity": current.get("relative_humidity_2m"),
            "rain": current.get("precipitation"), "cloud": current.get("cloud_cover"), "pressure": current.get("pressure_msl"),
            "wind": current.get("wind_speed_10m"), "gust": current.get("wind_gusts_10m"), "compass": wind_direction(current.get("wind_direction_10m")),
            "description": description, "icon": icon, "is_day": bool(current.get("is_day", 1)),
            "surface": now_hour.get("surface"), "dew": now_hour.get("dew"), "level": now_hour.get("level", ""),
            "visibility_km": now_hour.get("visibility_km"), "uv": now_hour.get("uv"),
            "ground6": now_hour.get("surface6"), "pressure_trend": pressure["trend"], "pressure_change": pressure["change"],
            "high": today.get("temp_max"), "low": today.get("temp_min"), "high_time": today.get("temp_max_time"), "low_time": today.get("temp_min_time"), "sunrise": today.get("sunrise"), "sunset": today.get("sunset"),
            "uv_max": today.get("uv_max"), "rain_today": today.get("rain_mm"), "rain_next_6h": round(sum(float(hour["rain"] or 0) for hour in hours[:6]), 1),
            "prob_next_6h": max((hour["prob"] or 0 for hour in hours[:6]), default=None), "next_rain": next_wet_spell(hours),
            "tonight": {"label": "This morning, until 06:00" if in_progress else "Tonight, 18:00 to 06:00", "level": tonight.get("level"),
                        "min_surface": tonight.get("min_surface"), "min_surface_time": tonight.get("min_surface_time"),
                        "min_surface_day": tonight.get("min_surface_day"), "rain_advice": tonight.get("rain_advice", "")} if tonight else None,
        },
        "days": days,
        "hours": hours,
    }


@bp.route("/")
def weather_home() -> str:
    return render_template("weather.html", place=WEATHER_PLACE)


@bp.get("/forecast.json")
def forecast_json() -> Any:
    try:
        raw, stale = fetch_weather()
        bias = gritting.current_road_bias()
        try:
            road_raw, _ = gritting.fetch_open_meteo_forecast()
            road = gritting.road_forecast_from_open_meteo(road_raw, gritting.local_now(), bias=bias)
        except Exception:
            road = None  # the weather page still works on its own forecast
        payload = build_weather(raw, gritting.local_now(), bias=bias, road=road)
        payload.update(ok=True, stale=stale, updated=datetime.now().strftime("%H:%M"))
    except Exception:
        payload = {"ok": False, "message": "The forecast could not be loaded just now. Check the internet connection and try again."}
    response = jsonify(payload)
    response.headers["Cache-Control"] = "no-store"
    return response


def register(flask_app: Any) -> None:
    flask_app.register_blueprint(bp)
