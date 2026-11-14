from __future__ import annotations

import csv
import html
import json
import os
import re
import sqlite3
import smtplib
import threading
import time
import urllib.request
import uuid
from calendar import monthrange
from datetime import datetime, timedelta
from email.message import EmailMessage
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any
import zipfile
from xml.etree import ElementTree as ET

from flask import Blueprint, Response, flash, g, jsonify, redirect, render_template, request, url_for


BASE_DIR = Path(__file__).resolve().parent
# Everything this module stores sits under data/gritting, laid out as the standalone app's folder was
# (gritting.db, data/customer_database.xlsx, exports/season_archives), so those can be copied straight in.
DATA_ROOT = BASE_DIR / "data" / "gritting"
DATABASE_PATH = DATA_ROOT / "gritting.db"
BUNDLED_CUSTOMER_XLSX = DATA_ROOT / "data" / "customer_database.xlsx"
EMAIL_LOGO_PATH = BASE_DIR / "static" / "gritting" / "brand" / "ngs-logo.png"
SEASON_ARCHIVE_DIR = DATA_ROOT / "exports" / "season_archives"


def configure(data_root: Any) -> None:
    """Point the module's storage at another folder (the muck app's data folder, or a test's temporary one)."""
    global DATA_ROOT, DATABASE_PATH, BUNDLED_CUSTOMER_XLSX, SEASON_ARCHIVE_DIR
    DATA_ROOT = Path(data_root)
    DATABASE_PATH = DATA_ROOT / "gritting.db"
    BUNDLED_CUSTOMER_XLSX = DATA_ROOT / "data" / "customer_database.xlsx"
    SEASON_ARCHIVE_DIR = DATA_ROOT / "exports" / "season_archives"
VAT_RATE = 0.20
COMPANY_NAME = "Norwich Gritting Services"
DEFAULT_EMAIL_SIGNATURE = {
    "email_signature_html": (
        '<div style="margin-top:24px;padding-top:16px;border-top:1px solid #dbe5fb;'
        'font-family:Arial,sans-serif;color:#11275d;">'
        '<div style="font-size:14px;margin-bottom:14px;">Regards,</div>'
        '<div style="font-size:15px;font-weight:700;margin-bottom:14px;">Andrew Farrell</div>'
        '<div style="font-size:14px;line-height:1.75;">'
        '<div style="font-weight:700;">{company_name}</div>'
        '<div>E: <a href="mailto:info@norwichgrittingservices.co.uk" '
        'style="color:#24418f;text-decoration:none;">info@norwichgrittingservices.co.uk</a></div>'
        '<div>W: <a href="https://www.norwichgrittingservices.co.uk" '
        'style="color:#24418f;text-decoration:none;">www.norwichgrittingservices.co.uk</a></div>'
        '<div>T: <a href="tel:07952683364" style="color:#24418f;text-decoration:none;">07952 683364</a></div>'
        "</div>"
        '<img src="cid:ngs-logo" alt="Norwich Gritting Services logo" '
        'style="display:block;max-width:220px;height:auto;margin-top:16px;">'
        "</div>"
    ),
    "email_signature_text": (
        "Regards,\n\n"
        "Andrew Farrell\n\n"
        "{company_name}\n"
        "E: info@norwichgrittingservices.co.uk\n"
        "W: www.norwichgrittingservices.co.uk\n"
        "T: 07952 683364"
    ),
}
DEFAULT_REPORT_SETTINGS = {
    "monthly_summary_recipients": "",
    "monthly_summary_auto_send": "false",
    "monthly_summary_auto_day": "1",
    "monthly_summary_auto_time": "06:00",
}
DEFAULT_REPORT_STATE = {
    "monthly_summary_last_attempt_month": "",
    "monthly_summary_last_sent_month": "",
    "monthly_summary_last_attempt_at": "",
    "monthly_summary_last_status": "",
}
DEFAULT_SEASON_SETTINGS = {
    "active_season_start_year": "",
}
DEFAULT_SEASON_STATE = {
    "last_archived_season_label": "",
    "last_archived_season_path": "",
    "last_archived_at": "",
}
DEFAULT_SMTP_SETTINGS = {
    "smtp_host": "",
    "smtp_port": "587",
    "smtp_from_email": "",
    "smtp_username": "",
    "smtp_password": "",
    "smtp_use_tls": "true",
    "smtp_use_ssl": "false",
}
SMTP_ENVIRONMENT_MAP = {
    "smtp_host": "SMTP_HOST",
    "smtp_port": "SMTP_PORT",
    "smtp_from_email": "SMTP_FROM_EMAIL",
    "smtp_username": "SMTP_USERNAME",
    "smtp_password": "SMTP_PASSWORD",
    "smtp_use_tls": "SMTP_USE_TLS",
    "smtp_use_ssl": "SMTP_USE_SSL",
}
LEGACY_EMAIL_TEMPLATE_VALUES = {
    "callout_subject": f"{COMPANY_NAME} call-out for {{site_name}}",
    "callout_body": (
        "Dear {contact_name},\n\n"
        f"{COMPANY_NAME} has scheduled a gritting call-out for your site.\n\n"
        "Customer: {customer_name}\n"
        "Site: {site_name}\n"
        "Weather conditions: {weather_conditions}\n"
        "Rough attendance time: {planned_attendance}\n"
        "Call-out issued at: {called_out_at}\n"
        "Additional notes: {service_notes}\n\n"
        "Regards,\n"
        f"{COMPANY_NAME}"
    ),
}
DEFAULT_EMAIL_TEMPLATES = {
    "callout_subject": f"{COMPANY_NAME} overnight gritting call-out for {{site_name}}",
    "callout_body": (
        "Dear {contact_name},\n\n"
        f"{COMPANY_NAME} has scheduled overnight gritting for your site.\n\n"
        "Customer: {customer_name}\n"
        "Site: {site_name}\n"
        "Weather conditions: {weather_conditions}\n"
        "Service window: {service_window}\n"
        "Timing may change through the night depending on live weather conditions.\n"
        "Call-out issued at: {called_out_at}\n"
        "Additional notes: {service_notes}"
    ),
    "completion_subject": f"{COMPANY_NAME} job completed for {{site_name}}",
    "completion_body": (
        "Dear {contact_name},\n\n"
        f"This is a confirmation that {COMPANY_NAME} attended and completed the call-out for your site.\n\n"
        "Customer: {customer_name}\n"
        "Site: {site_name}\n"
        "Call-out issued at: {called_out_at}\n"
        "Job notes: {service_notes}"
    ),
    "no_access_subject": f"{COMPANY_NAME} could not access {{site_name}}",
    "no_access_body": (
        "Dear {contact_name},\n\n"
        f"{COMPANY_NAME} attended your site on the night of {{called_out_at}} to carry out the gritting but was unable to gain access.\n\n"
        "Customer: {customer_name}\n"
        "Site: {site_name}\n"
        "Access problem: {access_note}\n\n"
        "Please let us know if the access can be sorted so we can attend on the next call-out."
    ),
}
DAY_KEYS = [
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday",
    "saturday",
    "sunday",
]
DAY_LABELS = {
    "monday": "Mon",
    "tuesday": "Tue",
    "wednesday": "Wed",
    "thursday": "Thu",
    "friday": "Fri",
    "saturday": "Sat",
    "sunday": "Sun",
}

bp = Blueprint("gritting", __name__, url_prefix="/gritting")
_flask_app = None
_initialised_database: Any = None
_initialise_lock = threading.Lock()
TEMPLATE_FILTERS: dict[str, Any] = {}

AUTO_MONTHLY_SUMMARY_CHECK_INTERVAL_SECONDS = 30
_auto_monthly_summary_thread_started = False
_auto_monthly_summary_thread_lock = threading.Lock()


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        connection = sqlite3.connect(DATABASE_PATH)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        g.db = connection
    return g.db


def close_db(_: BaseException | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def now_iso() -> str:
    return datetime.now().replace(second=0, microsecond=0).isoformat(timespec="minutes")


def format_datetime(value: str | None) -> str:
    if not value:
        return "Not set"
    parsed = datetime.fromisoformat(value)
    return parsed.strftime("%d %b %Y, %H:%M")


def format_date_from_iso(value: str | None) -> str:
    if not value:
        return "Not set"
    parsed = datetime.fromisoformat(value)
    return parsed.strftime("%-d %b %Y")


def format_month_label(month_value: str) -> str:
    return datetime.strptime(month_value, "%Y-%m").strftime("%B %Y")


def previous_month_value(reference: datetime | None = None) -> str:
    current = reference or datetime.now()
    year = current.year
    month = current.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{year:04d}-{month:02d}"


def format_date_label(value: str) -> str:
    return datetime.strptime(value, "%Y-%m-%d").strftime("%-d %b %Y")


def monthly_day_options() -> list[str]:
    return [str(day) for day in range(1, 32)]


def schedule_time_options() -> list[str]:
    options: list[str] = []
    for hour in range(24):
        for minute in (0, 30):
            options.append(f"{hour:02d}:{minute:02d}")
    return options


def format_bool(value: int | bool | None) -> str:
    return "Yes" if value else "No"


def format_currency(value: float | int | None) -> str:
    if value is None or value == "":
        return "-"
    return f"£{float(value):.2f}"


def vat_amount(value: float | int | None) -> float | None:
    if value is None or value == "":
        return None
    return round(float(value) * VAT_RATE, 2)


def gross_amount(value: float | int | None) -> float | None:
    if value is None or value == "":
        return None
    return round(float(value) + (vat_amount(value) or 0), 2)


def format_price_plus_vat(value: float | int | None) -> str:
    if value is None or value == "":
        return "-"
    return format_currency(value)


class SafeTemplateDict(dict):
    def __missing__(self, key: str) -> str:
        return "{" + key + "}"


def parse_bool(value: str | None, default: bool = True) -> bool:
    if value is None or not value.strip():
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y"}:
        return True
    if normalized in {"0", "false", "no", "n"}:
        return False
    return default


def dispatch_day_key(called_out_at: str) -> str:
    return datetime.fromisoformat(called_out_at).strftime("%A").lower()


def normalize_email(value: str) -> str:
    return value.strip().lower()


EMAIL_PATTERN = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")
POSTCODE_PATTERN = re.compile(r"^[A-Z]{1,2}[0-9][A-Z0-9]? ?[0-9][A-Z]{2}$")
# Notification addresses for a job: the customer's own list plus any extra ones set on the site, one list for sending.
NOTIFY_EMAILS_SQL = (
    "TRIM(COALESCE(customers.email, '') || CASE WHEN COALESCE(sites.notify_emails, '') <> '' "
    "THEN ', ' || sites.notify_emails ELSE '' END, ', ') AS email"
)


def split_emails(raw_value: str | None) -> list[str]:
    """Addresses from text separated by commas, semicolons or new lines, lower case and without repeats."""
    found: list[str] = []
    for token in (raw_value or "").replace(";", "\n").replace(",", "\n").splitlines():
        email = normalize_email(token)
        if email and email not in found:
            found.append(email)
    return found


def emails_from_form(form: Any, key: str) -> tuple[list[str], list[str]]:
    """The addresses typed into every box named key: (valid ones, ones that do not look like an email address)."""
    valid: list[str] = []
    invalid: list[str] = []
    for value in form.getlist(key):
        for email in split_emails(value):
            if EMAIL_PATTERN.match(email):
                if email not in valid:
                    valid.append(email)
            elif email not in invalid:
                invalid.append(email)
    return valid, invalid


PHONE_PATTERN = re.compile(r"^[0-9 +()./-]{5,30}$")


def load_emergency_contacts(raw: str | None, fallback_phone: str | None = "") -> list[dict[str, str]]:
    """The customer's emergency contacts, each with a name and a phone number. Older records only had a phone."""
    contacts: list[dict[str, str]] = []
    try:
        data = json.loads(raw or "[]")
    except ValueError:
        data = []
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and str(item.get("phone", "")).strip():
                contacts.append({"name": str(item.get("name", "")).strip(), "phone": str(item["phone"]).strip()})
    if not contacts and (fallback_phone or "").strip():
        contacts.append({"name": "", "phone": fallback_phone.strip()})
    return contacts


def emergency_contacts_from_form(form: Any) -> tuple[list[dict[str, str]], str]:
    """Contacts typed into the paired name and phone boxes: (contacts, error message or "")."""
    names = form.getlist("emergency_contact_name")
    phones = form.getlist("emergency_contact_phone")
    contacts: list[dict[str, str]] = []
    for index in range(max(len(names), len(phones))):
        name = (names[index] if index < len(names) else "").strip()
        phone = (phones[index] if index < len(phones) else "").strip()
        if not name and not phone:
            continue
        if not phone:
            return [], f"Add a phone number for the emergency contact {name}."
        if not PHONE_PATTERN.match(phone) or not re.search(r"\d{3}", phone):
            return [], f"{phone} does not look like a phone number."
        contacts.append({"name": name[:120], "phone": phone})
    if not contacts and form.get("emergency_phone", "").strip():  # a single phone number sent by an older form
        contacts.append({"name": "", "phone": form["emergency_phone"].strip()})
    return contacts, ""


def emergency_text(raw: str | None, fallback_phone: str | None = "") -> str:
    """Emergency contacts in one line: "Pat Shah 07700 900101 · Sam 07700 900102"."""
    return " · ".join(f"{item['name']} {item['phone']}".strip() for item in load_emergency_contacts(raw, fallback_phone))


TEMPLATE_FILTERS["emergency_text"] = emergency_text


def normalize_postcode(value: str | None) -> str | None:
    """A UK postcode in capitals with its space (NR151BE becomes NR15 1BE); "" if blank; None if it is not a postcode."""
    text = re.sub(r"\s+", "", (value or "").upper())
    if not text:
        return ""
    candidate = f"{text[:-3]} {text[-3:]}"
    return candidate if POSTCODE_PATTERN.match(candidate) else None


def preferred_email(row: dict[str, str]) -> str:
    """The customer's notification addresses as one list (a sheet cell may hold several separated by commas or semicolons)."""
    if row.get("email", "").strip():
        return ", ".join(split_emails(row["email"]))
    if "@" in row.get("contact_value", ""):
        return ", ".join(split_emails(row["contact_value"]))
    return ""


def preferred_emergency_phone(row: dict[str, str]) -> str:
    if row.get("emergency_phone", "").strip():
        return row["emergency_phone"].strip()
    if row.get("contact_value", "").strip() and "@" not in row["contact_value"]:
        return row["contact_value"].strip()
    return ""


def imported_service_days(row: dict[str, str], source_row: dict[str, str] | None = None) -> list[str]:
    raw_row = source_row or {}
    explicit_days = [
        day
        for day in DAY_KEYS
        if parse_bool(raw_row.get(day, ""), default=False)
        or parse_bool(raw_row.get(f"service_day_{day}", ""), default=False)
    ]
    if explicit_days:
        return explicit_days
    if row.get("service_days", "").strip():
        return parse_service_days(row["service_days"])
    if row.get("weekend_service_required", "").strip():
        if parse_bool(row["weekend_service_required"], default=True):
            return DAY_KEYS.copy()
        return ["monday", "tuesday", "wednesday", "thursday", "friday"]
    return DAY_KEYS.copy()


def parse_service_days(raw_value: str | None) -> list[str]:
    if not raw_value or not raw_value.strip():
        return DAY_KEYS.copy()

    tokens = []
    normalized = raw_value.replace("\n", ",").replace(";", ",").replace("|", ",")
    for token in normalized.split(","):
        candidate = token.strip().lower()
        if candidate:
            tokens.append(candidate)

    matched: list[str] = []
    aliases = {
        "mon": "monday",
        "monday": "monday",
        "tue": "tuesday",
        "tues": "tuesday",
        "tuesday": "tuesday",
        "wed": "wednesday",
        "wednesday": "wednesday",
        "thu": "thursday",
        "thur": "thursday",
        "thurs": "thursday",
        "thursday": "thursday",
        "fri": "friday",
        "friday": "friday",
        "sat": "saturday",
        "saturday": "saturday",
        "sun": "sunday",
        "sunday": "sunday",
    }
    for token in tokens:
        mapped = aliases.get(token)
        if mapped and mapped not in matched:
            matched.append(mapped)
    return matched


def serialize_service_days(days: list[str]) -> str:
    return json.dumps([day for day in DAY_KEYS if day in days])


def load_service_days(raw_value: str | None) -> list[str]:
    if not raw_value or not raw_value.strip():
        return DAY_KEYS.copy()
    try:
        loaded = json.loads(raw_value)
    except json.JSONDecodeError:
        return parse_service_days(raw_value)
    if not isinstance(loaded, list):
        return DAY_KEYS.copy()
    cleaned = []
    for item in loaded:
        day = str(item).strip().lower()
        if day in DAY_KEYS and day not in cleaned:
            cleaned.append(day)
    return cleaned or DAY_KEYS.copy()


def service_days_display(raw_value: str | None) -> str:
    days = load_service_days(raw_value)
    return ", ".join(DAY_LABELS[day] for day in DAY_KEYS if day in days)


def parse_pricing_options(raw_value: str | None) -> list[dict[str, Any]]:
    if not raw_value or not raw_value.strip():
        return []

    parsed_options: list[dict[str, Any]] = []
    for line in raw_value.splitlines():
        candidate = line.strip()
        if not candidate:
            continue

        label = ""
        amount_text = ""
        if "|" in candidate:
            label, amount_text = [part.strip() for part in candidate.split("|", 1)]
        elif "=" in candidate:
            label, amount_text = [part.strip() for part in candidate.split("=", 1)]
        elif "," in candidate:
            label, amount_text = [part.strip() for part in candidate.rsplit(",", 1)]
        else:
            continue

        amount_text = amount_text.replace("GBP", "").replace("gbp", "").replace("PS", "").replace("ps", "").replace("£", "").strip()
        try:
            amount = round(float(amount_text), 2)
        except ValueError:
            continue
        if label:
            parsed_options.append({"label": label, "amount": amount})

    return parsed_options


def imported_pricing_options(row: dict[str, str], source_row: dict[str, str] | None = None) -> list[dict[str, Any]]:
    raw_row = source_row or {}
    structured_options: list[dict[str, Any]] = []
    for index in range(1, 4):
        label = (
            raw_row.get(f"pricing_option_{index}", "").strip()
            or raw_row.get(f"pricing_label_{index}", "").strip()
            or raw_row.get(f"option_{index}", "").strip()
        )
        amount_text = (
            raw_row.get(f"pricing_amount_{index}", "").strip()
            or raw_row.get(f"price_{index}_ex_vat", "").strip()
            or raw_row.get(f"option_{index}_price", "").strip()
        )
        if not label and not amount_text:
            continue
        if not label or not amount_text:
            continue
        cleaned_amount = (
            amount_text.replace("GBP", "")
            .replace("gbp", "")
            .replace("PS", "")
            .replace("ps", "")
            .replace("£", "")
            .strip()
        )
        try:
            amount = round(float(cleaned_amount), 2)
        except ValueError:
            continue
        structured_options.append({"label": label, "amount": amount})
    if structured_options:
        return structured_options
    return parse_pricing_options(row.get("pricing_options", ""))


MAX_PRICE_ROWS = 30


def pricing_options_from_form(form: Any) -> list[dict[str, Any]]:
    """Every price row on the form: a name, a price ex VAT, the days it applies on and an optional special-date event."""
    structured_options: list[dict[str, Any]] = []
    for index in range(1, MAX_PRICE_ROWS + 1):
        label = form.get(f"pricing_label_{index}", "").strip()
        amount_text = form.get(f"pricing_amount_{index}", "").strip()
        if not label or not amount_text:
            continue
        try:
            amount = round(float(amount_text), 2)
        except ValueError:
            continue
        option: dict[str, Any] = {"label": label, "amount": amount}
        days = [day for day in form.getlist(f"pricing_days_{index}") if day in DAY_KEYS]
        if days:
            option["days"] = [day for day in DAY_KEYS if day in days]
        event = form.get(f"pricing_event_{index}", "").strip()
        if event:
            option["event"] = event
        structured_options.append(option)

    if structured_options:
        return structured_options
    return parse_pricing_options(form.get("pricing_options", ""))


def serialize_pricing_options(options: list[dict[str, Any]]) -> str:
    return json.dumps(options)


def load_pricing_options(raw_value: str | None) -> list[dict[str, Any]]:
    if not raw_value or not raw_value.strip():
        return []
    try:
        loaded = json.loads(raw_value)
    except json.JSONDecodeError:
        return []
    if not isinstance(loaded, list):
        return []
    normalized: list[dict[str, Any]] = []
    for item in loaded:
        if not isinstance(item, dict):
            continue
        label = str(item.get("label", "")).strip()
        try:
            amount = round(float(item.get("amount")), 2)
        except (TypeError, ValueError):
            continue
        if label:
            option: dict[str, Any] = {"label": label, "amount": amount}
            days = [day for day in DAY_KEYS if day in (item.get("days") or [])]
            if days:
                option["days"] = days
            event = str(item.get("event", "") or "").strip()
            if event:
                option["event"] = event
            normalized.append(option)
    return normalized


def pricing_options_display(raw_value: str | None) -> str:
    options = load_pricing_options(raw_value)
    if not options:
        return "No pricing options set"
    return ", ".join(
        f"{option['label']} ({format_price_plus_vat(option['amount'])})"
        for option in options
    )


def special_dates_by_day() -> dict[str, list[str]]:
    """Every special date, as {date: [names]}, for the call-out page to price by the chosen night."""
    found: dict[str, list[str]] = {}
    rows = get_db().execute("SELECT event_date, name FROM special_dates ORDER BY event_date, name").fetchall()
    for row in rows:
        found.setdefault(row["event_date"], []).append(row["name"])
    return found


def special_event_names(date_text: str) -> list[str]:
    """Names of the special dates (football matches and the like) on this date."""
    rows = get_db().execute("SELECT name FROM special_dates WHERE event_date = ? ORDER BY name", ((date_text or "")[:10],)).fetchall()
    return [row["name"] for row in rows]


def suggested_pricing_label(called_out_at: str, options: list[dict[str, Any]], event_names: Any = ()) -> str:
    """The price for the night: a price tied to a special date that falls on it, then one ticked for that weekday,
    then (for prices set up the old way) one whose name mentions the day, and finally the first ordinary price."""
    if not options:
        return ""

    for option in options:
        if option.get("event") and option["event"] in event_names:
            return option["label"]

    ordinary = [option for option in options if not option.get("event")] or options
    weekday_name = datetime.fromisoformat(called_out_at).strftime("%A").lower()
    for option in ordinary:
        if weekday_name in option.get("days", []):
            return option["label"]

    undated = [option for option in ordinary if not option.get("days")] or ordinary
    for option in undated:
        if weekday_name in option["label"].lower():
            return option["label"]

    if weekday_name in {"saturday", "sunday"}:
        for option in undated:
            if "weekend" in option["label"].lower():
                return option["label"]

    for option in undated:
        if "weekday" in option["label"].lower():
            return option["label"]

    return undated[0]["label"]


def selected_pricing_details(
    options: list[dict[str, Any]],
    selected_label: str | None,
) -> tuple[str | None, float | None]:
    cleaned_label = (selected_label or "").strip()
    if not cleaned_label:
        return None, None
    matched = next((option for option in options if option["label"] == cleaned_label), None)
    if matched:
        return cleaned_label, matched["amount"]
    return None, None


def decorate_site_for_callout(site: sqlite3.Row, called_out_at: str) -> dict[str, Any]:
    pricing_options = load_pricing_options(site["pricing_options"])
    service_days = load_service_days(site["service_days"])
    return {
        **dict(site),
        "service_days_list": service_days,
        "pricing_options_list": pricing_options,
        "pricing_options_json": json.dumps(pricing_options),
        "suggested_pricing_label": suggested_pricing_label(called_out_at, pricing_options, special_event_names(called_out_at)),
    }


def render_email_template(template_text: str, context: dict[str, Any]) -> str:
    normalized_context = {
        key: "" if value is None else str(value)
        for key, value in context.items()
    }
    return template_text.format_map(SafeTemplateDict(normalized_context))


def fetch_email_templates() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_EMAIL_TEMPLATES)})
        """,
        tuple(DEFAULT_EMAIL_TEMPLATES.keys()),
    ).fetchall()
    templates = DEFAULT_EMAIL_TEMPLATES.copy()
    for row in rows:
        templates[row["key"]] = row["value"]
    return templates


def fetch_email_signature() -> dict[str, str]:
    rows = get_db().execute(
        """
        SELECT key, value
        FROM app_settings
        WHERE key IN (?, ?)
        """,
        tuple(DEFAULT_EMAIL_SIGNATURE.keys()),
    ).fetchall()
    signature = DEFAULT_EMAIL_SIGNATURE.copy()
    for row in rows:
        signature[row["key"]] = row["value"]
    return signature


def fetch_report_settings() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_REPORT_SETTINGS)})
        """,
        tuple(DEFAULT_REPORT_SETTINGS.keys()),
    ).fetchall()
    settings = DEFAULT_REPORT_SETTINGS.copy()
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def fetch_report_state() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_REPORT_STATE)})
        """,
        tuple(DEFAULT_REPORT_STATE.keys()),
    ).fetchall()
    settings = DEFAULT_REPORT_STATE.copy()
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def fetch_season_settings() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_SEASON_SETTINGS)})
        """,
        tuple(DEFAULT_SEASON_SETTINGS.keys()),
    ).fetchall()
    settings = DEFAULT_SEASON_SETTINGS.copy()
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def fetch_season_state() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_SEASON_STATE)})
        """,
        tuple(DEFAULT_SEASON_STATE.keys()),
    ).fetchall()
    settings = DEFAULT_SEASON_STATE.copy()
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def fetch_smtp_settings() -> dict[str, str]:
    rows = get_db().execute(
        f"""
        SELECT key, value
        FROM app_settings
        WHERE key IN ({", ".join("?" for _ in DEFAULT_SMTP_SETTINGS)})
        """,
        tuple(DEFAULT_SMTP_SETTINGS.keys()),
    ).fetchall()
    settings = DEFAULT_SMTP_SETTINGS.copy()
    for key, env_name in SMTP_ENVIRONMENT_MAP.items():
        env_value = os.getenv(env_name, "").strip()
        if env_value:
            settings[key] = env_value
    for row in rows:
        settings[row["key"]] = row["value"]
    return settings


def fetch_email_settings() -> dict[str, str]:
    return {
        **fetch_email_templates(),
        **fetch_email_signature(),
        **fetch_report_settings(),
    }


def save_email_templates(db: sqlite3.Connection, form: Any) -> None:
    for key, default_value in DEFAULT_EMAIL_TEMPLATES.items():
        value = form.get(key, "").strip() or default_value
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )
    for key, default_value in DEFAULT_EMAIL_SIGNATURE.items():
        value = form.get(key, "").strip() or default_value
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def save_report_settings(db: sqlite3.Connection, form: Any) -> None:
    for key, default_value in DEFAULT_REPORT_SETTINGS.items():
        value = form.get(key, "").strip()
        if key == "monthly_summary_auto_send":
            value = "true" if value.lower() == "true" else "false"
        elif key == "monthly_summary_auto_day":
            try:
                parsed_day = int(value)
            except ValueError:
                parsed_day = int(default_value)
            value = str(min(max(parsed_day, 1), 31))
        elif key == "monthly_summary_auto_time":
            value = value if value in schedule_time_options() else default_value
        elif not value:
            value = default_value
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def save_report_state(db: sqlite3.Connection, updates: dict[str, str]) -> None:
    for key, default_value in DEFAULT_REPORT_STATE.items():
        if key not in updates:
            continue
        value = str(updates.get(key, default_value) or default_value)
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def save_season_settings(db: sqlite3.Connection, updates: dict[str, str]) -> None:
    for key, default_value in DEFAULT_SEASON_SETTINGS.items():
        if key not in updates:
            continue
        value = str(updates.get(key, default_value) or default_value)
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def save_season_state(db: sqlite3.Connection, updates: dict[str, str]) -> None:
    for key, default_value in DEFAULT_SEASON_STATE.items():
        if key not in updates:
            continue
        value = str(updates.get(key, default_value) or default_value)
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def save_smtp_settings(db: sqlite3.Connection, form: Any) -> None:
    for key, default_value in DEFAULT_SMTP_SETTINGS.items():
        value = form.get(key, "").strip()
        if key in {"smtp_use_tls", "smtp_use_ssl"}:
            value = "true" if value.lower() == "true" else "false"
        elif not value:
            value = default_value
        db.execute(
            """
            INSERT INTO app_settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )


def text_body_to_html(body_text: str) -> str:
    paragraphs = []
    for block in body_text.strip().split("\n\n"):
        cleaned = block.strip()
        if not cleaned:
            continue
        paragraphs.append(
            f'<p style="margin:0 0 14px;">{html.escape(cleaned).replace(chr(10), "<br>")}</p>'
        )
    return "".join(paragraphs)


def build_email_content(body_text: str) -> tuple[str, str]:
    signature = fetch_email_signature()
    signature_context = {
        "company_name": COMPANY_NAME,
    }
    signature_text = render_email_template(signature["email_signature_text"], signature_context).strip()
    signature_html = render_email_template(signature["email_signature_html"], signature_context).strip()

    text_content = body_text.strip()
    if signature_text:
        text_content = f"{text_content}\n\n{signature_text}"

    html_content = text_body_to_html(body_text)
    if signature_html:
        html_content = f'<div style="font-family:Arial,sans-serif;color:#11275d;font-size:14px;line-height:1.6;">{html_content}{signature_html}</div>'

    return text_content, html_content


def build_callout_email_context(
    site: sqlite3.Row | dict[str, Any],
    *,
    called_out_at: str,
    weather_conditions: str,
    planned_attendance: str,
    service_notes: str,
) -> dict[str, Any]:
    site_data = dict(site)
    service_window = planned_attendance.strip() or "18:00 to 06:00"
    return {
        "company_name": COMPANY_NAME,
        "contact_name": site_data.get("contact_name", ""),
        "customer_name": site_data.get("customer_name", ""),
        "site_name": site_data.get("site_name") or site_data.get("name", ""),
        "weather_conditions": weather_conditions.strip(),
        "service_window": service_window,
        "planned_attendance": planned_attendance.strip(),
        "called_out_at": format_datetime(called_out_at),
        "service_notes": service_notes.strip(),
    }


def pricing_summary_text(pricing_label: str | None, pricing_amount: float | None) -> str:
    if pricing_label and pricing_amount is not None:
        return f"{pricing_label} at {format_price_plus_vat(pricing_amount)}"
    return "Not selected"


def overnight_time_options() -> list[str]:
    options: list[str] = []
    hour = 18
    minute = 0
    while True:
        options.append(f"{hour:02d}:{minute:02d}")
        if hour == 6 and minute == 0:
            break
        minute += 30
        if minute == 60:
            minute = 0
            hour = (hour + 1) % 24
    return options


def planned_attendance_from_form(form: Any) -> str:
    start_time = form.get("planned_attendance_from", "").strip()
    end_time = form.get("planned_attendance_to", "").strip()
    if start_time and end_time:
        if start_time == end_time:
            return start_time
        return f"Between {start_time} and {end_time}"
    return form.get("planned_attendance", "").strip()


def current_gritting_season(reference: datetime | None = None) -> dict[str, Any]:
    today = reference or datetime.now()
    override_value = fetch_season_settings().get("active_season_start_year", "").strip()
    if override_value:
        try:
            start_year = int(override_value)
        except ValueError:
            start_year = today.year if today.month >= 10 else today.year - 1
        end_year = start_year + 1
        is_active = True
        status_label = "Current gritting season"
    elif today.month >= 10:
        start_year = today.year
        end_year = today.year + 1
        is_active = True
        status_label = "Current gritting season"
    elif today.month <= 4:
        start_year = today.year - 1
        end_year = today.year
        is_active = True
        status_label = "Current gritting season"
    else:
        start_year = today.year - 1
        end_year = today.year
        is_active = False
        status_label = "Last completed gritting season"

    start_date = f"{start_year}-10-01"
    end_date = f"{end_year}-04-30"
    return {
        "start_year": start_year,
        "end_year": end_year,
        "start_date": start_date,
        "end_date": end_date,
        "is_active": is_active,
        "label": f"{start_year}/{str(end_year)[-2:]} Season",
        "date_range_label": f"{format_date_label(start_date)} to {format_date_label(end_date)}",
        "status_label": status_label,
    }


TEMPLATE_FILTERS["datetime_display"] = format_datetime
TEMPLATE_FILTERS["date_display"] = format_date_from_iso
TEMPLATE_FILTERS["month_label"] = format_month_label
TEMPLATE_FILTERS["bool_display"] = format_bool
TEMPLATE_FILTERS["currency_display"] = format_currency
TEMPLATE_FILTERS["vat_amount"] = vat_amount
TEMPLATE_FILTERS["gross_amount"] = gross_amount
TEMPLATE_FILTERS["price_plus_vat_display"] = format_price_plus_vat
TEMPLATE_FILTERS["pricing_options_display"] = pricing_options_display
TEMPLATE_FILTERS["service_days_display"] = service_days_display


def init_db() -> None:
    db = sqlite3.connect(DATABASE_PATH)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys = ON")
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            contact_name TEXT NOT NULL,
            contact_method TEXT NOT NULL,
            contact_value TEXT NOT NULL,
            email TEXT,
            emergency_phone TEXT,
            weekend_service_required INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            address TEXT NOT NULL,
            route_notes TEXT,
            service_days TEXT,
            pricing_options TEXT,
            FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS call_outs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            dispatch_batch_id INTEGER,
            called_out_at TEXT NOT NULL,
            weather_conditions TEXT NOT NULL,
            planned_attendance TEXT,
            service_notes TEXT,
            status TEXT NOT NULL DEFAULT 'called_out',
            pricing_label TEXT,
            pricing_amount REAL,
            salt_used_tonnes REAL,
            completed_at TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (site_id) REFERENCES sites(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS dispatch_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_key TEXT NOT NULL,
            called_out_at TEXT NOT NULL,
            weather_conditions TEXT NOT NULL,
            planned_attendance TEXT NOT NULL,
            service_notes TEXT,
            email_subject TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS notification_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            call_out_id INTEGER NOT NULL,
            channel TEXT NOT NULL,
            recipient TEXT NOT NULL,
            subject TEXT,
            message TEXT NOT NULL,
            sent_at TEXT NOT NULL,
            delivery_status TEXT NOT NULL DEFAULT 'logged',
            delivery_details TEXT,
            FOREIGN KEY (call_out_id) REFERENCES call_outs(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS attendance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            call_out_id INTEGER NOT NULL,
            staff_name TEXT NOT NULL,
            arrived_at TEXT NOT NULL,
            departed_at TEXT,
            arrival_notes TEXT,
            departure_notes TEXT,
            FOREIGN KEY (call_out_id) REFERENCES call_outs(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )

    def ensure_column(table: str, column: str, definition: str) -> None:
        columns = {row[1] for row in db.execute(f"PRAGMA table_info({table})").fetchall()}
        if column not in columns:
            try:
                db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
            except sqlite3.OperationalError as exc:
                if "duplicate column name" not in str(exc).lower():
                    raise

    ensure_column("customers", "email", "TEXT")
    ensure_column("customers", "emergency_phone", "TEXT")
    ensure_column("customers", "emergency_contacts", "TEXT")
    ensure_column("sites", "postcode", "TEXT")
    ensure_column("sites", "notify_emails", "TEXT")
    ensure_column("customers", "weekend_service_required", "INTEGER NOT NULL DEFAULT 1")
    ensure_column("sites", "service_days", "TEXT")
    ensure_column("sites", "pricing_options", "TEXT")
    ensure_column("call_outs", "dispatch_batch_id", "INTEGER")
    ensure_column("call_outs", "planned_attendance", "TEXT")
    ensure_column("call_outs", "pricing_label", "TEXT")
    ensure_column("call_outs", "pricing_amount", "REAL")
    ensure_column("call_outs", "cancelled_at", "TEXT")
    ensure_column("call_outs", "extras_amount", "REAL NOT NULL DEFAULT 0")
    ensure_column("call_outs", "extras_json", "TEXT")
    ensure_column("call_outs", "no_access_at", "TEXT")
    ensure_column("call_outs", "access_note", "TEXT")
    had_job_numbers = "job_number" in {row[1] for row in db.execute("PRAGMA table_info(call_outs)").fetchall()}
    ensure_column("call_outs", "job_number", "INTEGER")
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS road_readings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            call_out_id INTEGER,
            site_id INTEGER,
            read_at TEXT NOT NULL,
            reading REAL NOT NULL,
            forecast_surface REAL,
            forecast_air REAL,
            note TEXT
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS special_dates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_date TEXT NOT NULL,
            name TEXT NOT NULL,
            UNIQUE (event_date, name)
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS site_extras (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            description TEXT NOT NULL,
            amount REAL NOT NULL,
            effective_from TEXT NOT NULL,
            effective_to TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    had_finalised = "finalised_at" in {row[1] for row in db.execute("PRAGMA table_info(dispatch_batches)").fetchall()}
    ensure_column("dispatch_batches", "finalised_at", "TEXT")
    if not had_finalised:
        # Lists that existed before drafts were introduced and already have finished jobs are being worked: treat them as final.
        db.execute(
            """
            UPDATE dispatch_batches SET finalised_at = created_at
            WHERE id IN (SELECT DISTINCT dispatch_batch_id FROM call_outs WHERE status = 'completed' AND dispatch_batch_id IS NOT NULL)
            """
        )
    if not had_job_numbers:
        # Jobs that already exist keep the number they showed (their id). Cancelled jobs and jobs still on a draft list have none.
        db.execute(
            """
            UPDATE call_outs SET job_number = id
            WHERE status <> 'cancelled'
              AND (dispatch_batch_id IS NULL OR dispatch_batch_id IN (SELECT id FROM dispatch_batches WHERE finalised_at IS NOT NULL))
            """
        )
    ensure_column("notification_logs", "subject", "TEXT")
    ensure_column("notification_logs", "delivery_status", "TEXT NOT NULL DEFAULT 'logged'")
    ensure_column("notification_logs", "delivery_details", "TEXT")
    ensure_column("notification_logs", "kind", "TEXT")
    for key, value in DEFAULT_EMAIL_TEMPLATES.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
    for key, legacy_value in LEGACY_EMAIL_TEMPLATE_VALUES.items():
        db.execute(
            "UPDATE app_settings SET value = ? WHERE key = ? AND value = ?",
            (DEFAULT_EMAIL_TEMPLATES[key], key, legacy_value),
        )
    for key, value in DEFAULT_EMAIL_SIGNATURE.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
    for key, value in DEFAULT_REPORT_SETTINGS.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
    for key, value in DEFAULT_REPORT_STATE.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
    for key, value in DEFAULT_SEASON_SETTINGS.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
    for key, value in DEFAULT_SEASON_STATE.items():
        db.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )

    callout_subject_row = db.execute(
        "SELECT value FROM app_settings WHERE key = ?",
        ("callout_subject",),
    ).fetchone()
    if callout_subject_row:
        current_subject = callout_subject_row["value"]
        if current_subject == LEGACY_EMAIL_TEMPLATE_VALUES["callout_subject"]:
            db.execute(
                "UPDATE app_settings SET value = ? WHERE key = ?",
                (DEFAULT_EMAIL_TEMPLATES["callout_subject"], "callout_subject"),
            )

    callout_body_row = db.execute(
        "SELECT value FROM app_settings WHERE key = ?",
        ("callout_body",),
    ).fetchone()
    if callout_body_row:
        updated_body = callout_body_row["value"]
        if "scheduled a gritting call-out for your site." in updated_body:
            updated_body = updated_body.replace(
                "scheduled a gritting call-out for your site.",
                "scheduled overnight gritting for your site.",
            )
        if (
            "Rough attendance time: {planned_attendance}" in updated_body
            and "Service window: 18:00 to 06:00" not in updated_body
        ):
            updated_body = updated_body.replace(
                "Rough attendance time: {planned_attendance}",
                "Service window: {service_window}\n"
                "Timing may change through the night depending on live weather conditions.",
            )
        if "Service window: 18:00 to 06:00" in updated_body:
            updated_body = updated_body.replace(
                "Service window: 18:00 to 06:00",
                "Service window: {service_window}",
            )
        if "Expected attendance timing: {planned_attendance}\n" in updated_body:
            updated_body = updated_body.replace(
                "Expected attendance timing: {planned_attendance}\n",
                "",
            )
        if updated_body != callout_body_row["value"]:
            db.execute(
                "UPDATE app_settings SET value = ? WHERE key = ?",
                (updated_body, "callout_body"),
            )

    customer_count = db.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    if customer_count != 0:
        db.execute(
            """
            UPDATE customers
            SET email = CASE
                WHEN email IS NULL OR trim(email) = '' THEN CASE
                    WHEN instr(contact_value, '@') > 0 THEN lower(contact_value)
                    ELSE email
                END
                ELSE lower(email)
            END
            """
        )
        db.execute(
            """
            UPDATE customers
            SET emergency_phone = CASE
                WHEN emergency_phone IS NULL OR trim(emergency_phone) = '' THEN CASE
                    WHEN instr(contact_value, '@') = 0 THEN contact_value
                    ELSE emergency_phone
                END
                ELSE emergency_phone
            END
            """
        )
        db.execute("UPDATE customers SET contact_method = 'Email'")
        db.execute(
            """
            UPDATE sites
            SET service_days = CASE
                WHEN service_days IS NULL OR trim(service_days) = '' THEN CASE
                    WHEN customer_id IN (
                        SELECT id FROM customers WHERE weekend_service_required = 0
                    ) THEN ?
                    ELSE ?
                END
                ELSE service_days
            END
            """,
            (
                serialize_service_days(["monday", "tuesday", "wednesday", "thursday", "friday"]),
                serialize_service_days(DAY_KEYS),
            ),
        )

    try:
        sync_bundled_customer_workbook(db)
    except (ValueError, zipfile.BadZipFile, ET.ParseError):
        # If the bundled workbook is missing or malformed, keep the app usable
        # and fall back to the existing database contents.
        pass

    db.commit()
    db.close()


def fetch_customers() -> list[sqlite3.Row]:
    return get_db().execute(
        "SELECT * FROM customers ORDER BY name COLLATE NOCASE"
    ).fetchall()


def normalize_header(value: str) -> str:
    return value.strip().lower().replace(" ", "_")


def column_index_from_ref(cell_ref: str) -> int:
    letters = "".join(character for character in cell_ref if character.isalpha())
    index = 0
    for character in letters:
        index = index * 26 + (ord(character.upper()) - 64)
    return max(index - 1, 0)


def read_xlsx_rows(file_bytes: bytes) -> list[dict[str, str]]:
    namespace = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    with zipfile.ZipFile(BytesIO(file_bytes)) as workbook:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for string_item in shared_root.findall("main:si", namespace):
                parts = [node.text or "" for node in string_item.findall(".//main:t", namespace)]
                shared_strings.append("".join(parts))

        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        sheets = workbook_root.find("main:sheets", namespace)
        if sheets is None or len(sheets) == 0:
            return []

        first_sheet = sheets[0]
        relation_id = first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        relations_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        target = None
        for relation in relations_root.findall("rel:Relationship", namespace):
            if relation.attrib.get("Id") == relation_id:
                target = relation.attrib.get("Target")
                break

        if not target:
            return []

        sheet_path = f"xl/{target}" if not target.startswith("xl/") else target
        sheet_root = ET.fromstring(workbook.read(sheet_path))
        sheet_data = sheet_root.find("main:sheetData", namespace)
        if sheet_data is None:
            return []

        rows: list[list[str]] = []
        for row in sheet_data.findall("main:row", namespace):
            values: list[str] = []
            for cell in row.findall("main:c", namespace):
                index = column_index_from_ref(cell.attrib.get("r", "A1"))
                while len(values) <= index:
                    values.append("")

                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    text_parts = [node.text or "" for node in cell.findall(".//main:t", namespace)]
                    value = "".join(text_parts)
                else:
                    raw_value = cell.findtext("main:v", default="", namespaces=namespace)
                    if cell_type == "s" and raw_value:
                        value = shared_strings[int(raw_value)]
                    else:
                        value = raw_value
                values[index] = value.strip()
            rows.append(values)

    if not rows:
        return []

    headers = [normalize_header(header) for header in rows[0]]
    normalized_rows: list[dict[str, str]] = []
    for raw_row in rows[1:]:
        row_map: dict[str, str] = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            row_map[header] = raw_row[index].strip() if index < len(raw_row) else ""
        if any(value for value in row_map.values()):
            normalized_rows.append(row_map)
    return normalized_rows


def map_import_row(row: dict[str, str]) -> dict[str, str]:
    aliases = {
        "customer_name": ["customer_name", "customer", "client_name", "client"],
        "contact_name": ["contact_name", "contact", "contact_person"],
        "contact_method": ["contact_method", "notification_method", "preferred_contact_method"],
        "contact_value": ["contact_value", "contact_details", "phone", "email", "mobile"],
        "email": ["email", "client_email", "notification_email"],
        "emergency_phone": ["emergency_phone", "phone", "mobile", "emergency_contact_number"],
        "emergency_contact_name": ["emergency_contact_name", "emergency_contact", "emergency_contact_person"],
        "postcode": ["postcode", "post_code", "site_postcode"],
        "site_emails": ["site_emails", "site_notification_emails", "site_email"],
        "weekend_service_required": [
            "weekend_service_required",
            "weekend_required",
            "requires_weekend_gritting",
            "weekend_cover",
        ],
        "site_name": ["site_name", "site", "location_name"],
        "site_address": ["site_address", "address", "location_address"],
        "route_notes": ["route_notes", "notes", "site_notes", "route_instructions"],
        "service_days": ["service_days", "gritting_days", "active_days", "service_schedule"],
        "pricing_options": ["pricing_options", "pricing", "site_pricing", "pricing_dropdown"],
    }
    mapped: dict[str, str] = {}
    for target_key, options in aliases.items():
        mapped[target_key] = next((row.get(option, "").strip() for option in options if row.get(option, "").strip()), "")
    return mapped


def import_customer_rows(db: sqlite3.Connection, rows: list[dict[str, str]]) -> dict[str, int]:
    if not rows:
        raise ValueError("The workbook is empty or the first sheet could not be read.")

    usable_rows = 0
    customers_added = 0
    customers_updated = 0
    sites_added = 0
    sites_updated = 0

    for source_row in rows:
        row = map_import_row(source_row)
        required = ["customer_name", "contact_name"]
        missing = [field for field in required if not row[field]]
        email_value = preferred_email(row)
        if missing or not email_value:
            continue
        usable_rows += 1

        existing_customer = db.execute(
            "SELECT id FROM customers WHERE lower(name) = lower(?)",
            (row["customer_name"],),
        ).fetchone()
        imported_phone = preferred_emergency_phone(row)
        imported_contacts_json = (
            json.dumps([{"name": row["emergency_contact_name"], "phone": imported_phone}]) if imported_phone else ""
        )
        imported_postcode = normalize_postcode(row["postcode"])
        imported_postcode = row["postcode"].strip().upper() if imported_postcode is None else imported_postcode
        imported_site_emails = ", ".join(split_emails(row["site_emails"]))

        if existing_customer:
            customer_id = existing_customer["id"]
            db.execute(
                """
                UPDATE customers
                SET contact_name = ?, contact_method = ?, contact_value = ?, email = ?, emergency_phone = ?,
                    emergency_contacts = CASE WHEN ? <> '' THEN ? ELSE emergency_contacts END
                WHERE id = ?
                """,
                (
                    row["contact_name"],
                    "Email",
                    preferred_emergency_phone(row),
                    email_value,
                    preferred_emergency_phone(row),
                    imported_contacts_json,
                    imported_contacts_json,
                    customer_id,
                ),
            )
            customers_updated += 1
        else:
            cursor = db.execute(
                """
                INSERT INTO customers (
                    name,
                    contact_name,
                    contact_method,
                    contact_value,
                    email,
                    emergency_phone,
                    emergency_contacts
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    row["customer_name"],
                    row["contact_name"],
                    "Email",
                    preferred_emergency_phone(row),
                    email_value,
                    preferred_emergency_phone(row),
                    imported_contacts_json or "[]",
                ),
            )
            customer_id = cursor.lastrowid
            customers_added += 1

        if row["site_name"]:
            parsed_site_pricing = imported_pricing_options(row, source_row)
            existing_site = db.execute(
                """
                SELECT id
                FROM sites
                WHERE customer_id = ? AND lower(name) = lower(?)
                """,
                (customer_id, row["site_name"]),
            ).fetchone()

            if existing_site:
                db.execute(
                    """
                    UPDATE sites
                    SET address = ?, route_notes = ?, service_days = ?, pricing_options = ?,
                        postcode = CASE WHEN ? <> '' THEN ? ELSE postcode END,
                        notify_emails = CASE WHEN ? <> '' THEN ? ELSE notify_emails END
                    WHERE id = ?
                    """,
                    (
                        row["site_address"] or "Address not supplied",
                        row["route_notes"],
                        serialize_service_days(imported_service_days(row, source_row)),
                        serialize_pricing_options(parsed_site_pricing)
                        if parsed_site_pricing
                        else db.execute("SELECT pricing_options FROM sites WHERE id = ?", (existing_site["id"],)).fetchone()[0],
                        imported_postcode,
                        imported_postcode,
                        imported_site_emails,
                        imported_site_emails,
                        existing_site["id"],
                    ),
                )
                sites_updated += 1
            else:
                db.execute(
                    """
                    INSERT INTO sites (customer_id, name, address, postcode, notify_emails, route_notes, service_days, pricing_options)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        customer_id,
                        row["site_name"],
                        row["site_address"] or "Address not supplied",
                        imported_postcode,
                        imported_site_emails,
                        row["route_notes"],
                        serialize_service_days(imported_service_days(row, source_row)),
                        serialize_pricing_options(parsed_site_pricing),
                    ),
                )
                sites_added += 1

    if usable_rows == 0:
        raise ValueError("No usable rows were found in the workbook.")

    return {
        "customers_added": customers_added,
        "customers_updated": customers_updated,
        "sites_added": sites_added,
        "sites_updated": sites_updated,
    }


def import_customer_workbook(file_bytes: bytes) -> dict[str, int]:
    rows = read_xlsx_rows(file_bytes)
    summary = import_customer_rows(get_db(), rows)
    get_db().commit()
    return summary


def sync_bundled_customer_workbook(db: sqlite3.Connection) -> dict[str, int] | None:
    if not BUNDLED_CUSTOMER_XLSX.exists():
        return None

    rows = read_xlsx_rows(BUNDLED_CUSTOMER_XLSX.read_bytes())
    if not rows:
        return None
    return import_customer_rows(db, rows)


def fetch_sites() -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT
            sites.*,
            customers.name AS customer_name,
            customers.contact_name,
            customers.contact_method,
            customers.contact_value,
            NOTIFY_EMAILS_SQL,
            customers.emergency_phone,
            customers.emergency_contacts
        FROM sites
        JOIN customers ON customers.id = sites.customer_id
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """.replace("NOTIFY_EMAILS_SQL", NOTIFY_EMAILS_SQL)
    ).fetchall()


def fetch_customer_directory_rows() -> list[sqlite3.Row]:
    rows = get_db().execute(
        """
        SELECT
            customers.id AS customer_id,
            customers.name AS customer_name,
            customers.contact_name,
            customers.email,
            customers.emergency_phone,
            customers.emergency_contacts,
            sites.id AS site_id,
            sites.name AS site_name,
            sites.address,
            sites.postcode,
            sites.notify_emails AS site_notify_emails,
            sites.route_notes,
            sites.service_days,
            sites.pricing_options
        FROM customers
        LEFT JOIN sites ON sites.customer_id = customers.id
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """
    ).fetchall()
    decorated_rows: list[dict[str, Any]] = []
    for row in rows:
        row_data = dict(row)
        pricing_options_list = load_pricing_options(row_data.get("pricing_options"))
        pricing_fields = []
        for index in range(max(3, len(pricing_options_list))):
            option = pricing_options_list[index] if index < len(pricing_options_list) else {}
            pricing_fields.append(
                {
                    "index": index + 1,
                    "label": option.get("label", ""),
                    "amount": option.get("amount", ""),
                    "days": option.get("days", []),
                    "event": option.get("event", ""),
                }
            )
        row_data["contacts_list"] = load_emergency_contacts(row_data.get("emergency_contacts"), row_data.get("emergency_phone"))
        row_data["emails_list"] = split_emails(row_data.get("email"))
        row_data["site_emails_list"] = split_emails(row_data.get("site_notify_emails"))
        row_data["service_days_list"] = load_service_days(row_data.get("service_days"))
        row_data["pricing_options_list"] = pricing_options_list
        row_data["pricing_fields"] = pricing_fields
        decorated_rows.append(row_data)
    return decorated_rows


def fetch_site_notification_row(site_id: int) -> sqlite3.Row | None:
    return get_db().execute(
        """
        SELECT
            sites.id,
            sites.name AS site_name,
            customers.name AS customer_name,
            customers.contact_name,
            NOTIFY_EMAILS_SQL
        FROM sites
        JOIN customers ON customers.id = sites.customer_id
        WHERE sites.id = ?
        """.replace("NOTIFY_EMAILS_SQL", NOTIFY_EMAILS_SQL),
        (site_id,),
    ).fetchone()


def fetch_dispatch_batches(limit: int = 8) -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT
            dispatch_batches.id,
            dispatch_batches.called_out_at,
            dispatch_batches.weather_conditions,
            dispatch_batches.planned_attendance,
            dispatch_batches.email_subject,
            COUNT(call_outs.id) AS created_call_outs
        FROM dispatch_batches
        LEFT JOIN call_outs ON call_outs.dispatch_batch_id = dispatch_batches.id
        GROUP BY dispatch_batches.id
        ORDER BY dispatch_batches.called_out_at DESC
        LIMIT ?
        """,
        (limit,),
    ).fetchall()


def fetch_bulk_dispatch_candidates(called_out_at: str) -> tuple[list[sqlite3.Row], list[sqlite3.Row]]:
    all_sites = get_db().execute(
        """
        SELECT
            sites.id,
            sites.name AS site_name,
            sites.address,
            sites.route_notes,
            sites.service_days,
            sites.pricing_options,
            customers.name AS customer_name,
            customers.contact_name,
            customers.contact_value,
            customers.contact_method,
            NOTIFY_EMAILS_SQL,
            customers.emergency_phone,
            customers.emergency_contacts
        FROM sites
        JOIN customers ON customers.id = sites.customer_id
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """.replace("NOTIFY_EMAILS_SQL", NOTIFY_EMAILS_SQL)
    ).fetchall()

    eligible: list[sqlite3.Row] = []
    excluded: list[sqlite3.Row] = []
    required_day = dispatch_day_key(called_out_at)
    for site in all_sites:
        if required_day in load_service_days(site["service_days"]):
            eligible.append(site)
        else:
            excluded.append(site)
    return eligible, excluded


def smtp_configured() -> bool:
    smtp_settings = fetch_smtp_settings()
    return bool(smtp_settings["smtp_host"].strip() and smtp_settings["smtp_from_email"].strip())


def parse_recipient_list(raw_value: str | None) -> list[str]:
    if not raw_value:
        return []
    recipients: list[str] = []
    normalized = raw_value.replace(";", "\n").replace(",", "\n")
    for token in normalized.splitlines():
        email = normalize_email(token)
        if email and email not in recipients:
            recipients.append(email)
    return recipients


def send_email_message(
    recipients: list[str],
    subject: str,
    body: str,
    attachments: list[dict[str, Any]] | None = None,
) -> tuple[str, str]:
    if not recipients:
        return "skipped_no_email", "No email recipients are recorded."

    smtp_settings = fetch_smtp_settings()
    host = smtp_settings["smtp_host"].strip()
    from_email = smtp_settings["smtp_from_email"].strip()
    if not host or not from_email:
        return "logged_only", "SMTP is not configured, so the email was logged only."

    try:
        port = int((smtp_settings["smtp_port"] or "587").strip())
    except ValueError:
        return "failed", "SMTP port is not a valid number."
    username = smtp_settings["smtp_username"].strip()
    password = smtp_settings["smtp_password"]
    use_ssl = parse_bool(smtp_settings["smtp_use_ssl"], default=False)
    use_tls = parse_bool(smtp_settings["smtp_use_tls"], default=True)

    text_body, html_body = build_email_content(body)

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = from_email
    message["To"] = ", ".join(recipients)
    message.set_content(text_body)
    message.add_alternative(html_body, subtype="html")

    if EMAIL_LOGO_PATH.exists():
        html_part = message.get_payload()[-1]
        html_part.add_related(
            EMAIL_LOGO_PATH.read_bytes(),
            maintype="image",
            subtype="png",
            cid="<ngs-logo>",
            filename="ngs-logo.png",
            disposition="inline",
        )

    for attachment in attachments or []:
        message.add_attachment(
            attachment["data"],
            maintype=attachment["maintype"],
            subtype=attachment["subtype"],
            filename=attachment["filename"],
        )

    try:
        if use_ssl:
            with smtplib.SMTP_SSL(host, port, timeout=20) as server:
                if username:
                    server.login(username, password)
                server.send_message(message)
        else:
            with smtplib.SMTP(host, port, timeout=20) as server:
                if use_tls:
                    server.starttls()
                if username:
                    server.login(username, password)
                server.send_message(message)
        return "sent", "Email sent successfully."
    except Exception as exc:
        return "failed", str(exc)


def send_email_notification(recipient: str | None, subject: str, body: str) -> tuple[str, str]:
    recipients = split_emails(recipient)
    if not recipients:
        return "skipped_no_email", "No customer email address is recorded."
    return send_email_message(recipients, subject, body)


def log_notification(
    *,
    call_out_id: int,
    channel: str,
    recipient: str | None,
    subject: str,
    message: str,
    kind: str = "",
) -> tuple[str, str]:
    logged_message = message
    if channel.lower() == "email":
        recipient = ", ".join(split_emails(recipient)) or recipient
    delivery_status = "logged"
    delivery_details = ""
    if channel.lower() == "email":
        logged_message, _ = build_email_content(message)
        delivery_status, delivery_details = send_email_notification(recipient, subject, message)

    get_db().execute(
        """
        INSERT INTO notification_logs (
            call_out_id,
            channel,
            recipient,
            subject,
            message,
            sent_at,
            delivery_status,
            delivery_details,
            kind
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            call_out_id,
            channel,
            recipient or "",
            subject,
            logged_message,
            now_iso(),
            delivery_status,
            delivery_details,
            kind,
        ),
    )
    return delivery_status, delivery_details


UNSENT_STATUSES = ("logged", "logged_only", "failed", "skipped_no_email")


def strip_signature(message: str) -> str:
    """A logged email has the company signature on the end already; take it off before sending it again."""
    signature = fetch_email_signature()
    signature_text = render_email_template(signature["email_signature_text"], {"company_name": COMPANY_NAME}).strip()
    text = (message or "").rstrip()
    if signature_text and text.endswith(signature_text):
        text = text[: -len(signature_text)].rstrip()
    return text


def unsent_emails(db: sqlite3.Connection, call_out_ids: list[int], live_only: bool) -> list[sqlite3.Row]:
    """Logged emails that never went out (the mail server was not set up, or it failed) for these call-outs.

    Jobs that were cancelled are left alone. With live_only, so are finished jobs: their call-out is history."""
    if not call_out_ids:
        return []
    marks = ",".join("?" for _ in call_out_ids)
    statuses = ("called_out", "no_access") if live_only else ("called_out", "no_access", "completed")
    rows = db.execute(
        f"""
        SELECT notification_logs.*, call_outs.site_id, call_outs.status AS job_status
        FROM notification_logs
        JOIN call_outs ON call_outs.id = notification_logs.call_out_id
        WHERE notification_logs.call_out_id IN ({marks})
          AND LOWER(notification_logs.channel) = 'email'
          AND notification_logs.delivery_status IN ({",".join("?" for _ in UNSENT_STATUSES)})
          AND call_outs.status IN ({",".join("?" for _ in statuses)})
          AND (
            (call_outs.status = 'called_out' AND COALESCE(notification_logs.kind, '') IN ('', 'callout'))
            OR (call_outs.status = 'no_access' AND notification_logs.kind = 'no_access')
            OR (call_outs.status = 'completed' AND notification_logs.kind = 'completion')
          )
          AND NOT EXISTS (
            SELECT 1 FROM notification_logs sent
            WHERE sent.call_out_id = notification_logs.call_out_id AND sent.subject = notification_logs.subject AND sent.delivery_status = 'sent'
          )
        ORDER BY notification_logs.id
        """,
        (*call_out_ids, *UNSENT_STATUSES, *statuses),
    ).fetchall()
    latest: dict[tuple[int, str], sqlite3.Row] = {}
    for row in rows:
        latest[(row["call_out_id"], row["subject"] or "")] = row  # the latest attempt tells why it still has not gone
    return list(latest.values())


def resend_unsent_emails(db: sqlite3.Connection, call_out_ids: list[int], live_only: bool = True) -> dict[str, int]:
    """Send again the emails that were only logged. Goes to the addresses on the customer and site now."""
    counts = {"sent": 0, "failed": 0, "no_email": 0}
    for row in unsent_emails(db, call_out_ids, live_only):
        current = fetch_site_notification_row(row["site_id"])
        recipient = (current["email"] if current else "") or row["recipient"]
        status, _ = log_notification(
            call_out_id=row["call_out_id"], channel="Email", recipient=recipient, subject=row["subject"] or "",
            message=strip_signature(row["message"]), kind=row["kind"] or "",
        )
        key = "sent" if status == "sent" else "no_email" if status == "skipped_no_email" else "failed"
        counts[key] += 1
    return counts


def email_problems(db: sqlite3.Connection, call_out_ids: list[int]) -> dict[int, str]:
    """For each call-out whose email did not go, a short reason to show on its card, so it can be sorted and resent."""
    reasons: dict[int, str] = {}
    for row in unsent_emails(db, call_out_ids, live_only=False):
        status = row["delivery_status"]
        if status == "skipped_no_email":
            reason = "no email address on file"
        elif status == "failed":
            reason = "sending failed" + (f": {row['delivery_details']}" if row["delivery_details"] else "")
        else:
            reason = "the mail server was not set up, so it was only logged"
        reasons.setdefault(row["call_out_id"], f"Email not sent: {reason}")
    return reasons


def flash_resend_result(counts: dict[str, int]) -> None:
    total = sum(counts.values())
    if total == 0:
        flash("There are no unsent emails to resend.")
        return
    parts = [f"{counts['sent']} sent"]
    if counts["failed"]:
        parts.append(f"{counts['failed']} failed (see the job's notifications for the reason)")
    if counts["no_email"]:
        parts.append(f"{counts['no_email']} with no email address on file")
    flash("Resent: " + ", ".join(parts) + ".")


def extras_on(db: sqlite3.Connection, site_id: int, night: str) -> tuple[float, list[dict[str, Any]]]:
    """Extras charged on every visit to the site on the given night (YYYY-MM-DD): the total and each item.

    An extra applies from its first night to its last night inclusive, or from its first night onward if it has no end."""
    rows = db.execute(
        """
        SELECT description, amount FROM site_extras
        WHERE site_id = ? AND effective_from <= ? AND (effective_to IS NULL OR effective_to = '' OR effective_to >= ?)
        ORDER BY effective_from, id
        """,
        (site_id, night, night),
    ).fetchall()
    items = [{"description": row["description"], "amount": round(float(row["amount"]), 2)} for row in rows]
    return round(sum(item["amount"] for item in items), 2), items


def load_extras_items(raw: str | None) -> list[dict[str, Any]]:
    try:
        items = json.loads(raw or "[]")
    except ValueError:
        return []
    return [item for item in items if isinstance(item, dict) and "description" in item and "amount" in item] if isinstance(items, list) else []


def reprice_open_jobs(db: sqlite3.Connection, site_id: int) -> int:
    """Re-work the price of the site's jobs that are not done yet after its extras change. Completed jobs keep their price."""
    count = 0
    for row in db.execute(
        "SELECT id, called_out_at, pricing_amount, extras_amount, extras_json FROM call_outs WHERE site_id = ? AND status = 'called_out' AND pricing_amount IS NOT NULL",
        (site_id,),
    ).fetchall():
        base = round(float(row["pricing_amount"]) - float(row["extras_amount"] or 0), 2)
        extras_total, items = extras_on(db, site_id, (row["called_out_at"] or "")[:10])
        total = round(base + extras_total, 2)
        if total != round(float(row["pricing_amount"]), 2) or json.dumps(items) != row["extras_json"]:
            db.execute(
                "UPDATE call_outs SET pricing_amount = ?, extras_amount = ?, extras_json = ? WHERE id = ?",
                (total, extras_total, json.dumps(items), row["id"]),
            )
            count += 1
    return count


def extras_summary(raw: str | None) -> str:
    """The extras on a job in words, for example "Rear bay +£20.00"."""
    return ", ".join(
        f"{item['description']} {'+' if item['amount'] >= 0 else '-'}£{abs(item['amount']):.2f}" for item in load_extras_items(raw)
    )


TEMPLATE_FILTERS["extras_summary"] = extras_summary


def site_extras_redirect(site_id: int) -> Any:
    return redirect(url_for("gritting.customers") + f"#site-extras-{site_id}")


def parse_night(value: str, default: str = "") -> str | None:
    """A date typed as YYYY-MM-DD, or the default when blank; None if it is not a real date."""
    text = (value or "").strip() or default
    try:
        return datetime.strptime(text, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return None


@bp.post("/sites/<int:site_id>/extras")
def add_site_extra(site_id: int):
    """Log something the customer has added to the site: it adds to the price of every visit from the first night onward."""
    db = get_db()
    site = db.execute("SELECT id FROM sites WHERE id = ?", (site_id,)).fetchone()
    if site is None:
        flash("Site not found.")
        return redirect(url_for("gritting.customers"))
    description = request.form.get("description", "").strip()
    try:
        amount = round(float(request.form.get("amount", "").strip()), 2)
    except ValueError:
        amount = 0.0
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    first_night = parse_night(request.form.get("effective_from", ""), tomorrow)
    if not description or len(description) > 160:
        flash("Describe what was added, for example 'rear loading bay'.")
    elif amount == 0.0 or abs(amount) > 100000:
        flash("Enter the extra price per visit, ex VAT. Use a minus for a reduction.")
    elif first_night is None:
        flash("Choose a valid first night.")
    else:
        db.execute(
            "INSERT INTO site_extras (site_id, description, amount, effective_from, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (site_id, description, amount, first_night, request.form.get("note", "").strip()[:500], now_iso()),
        )
        changed = reprice_open_jobs(db, site_id)
        db.commit()
        flash(
            f"Added {description} ({'+' if amount >= 0 else '-'}£{abs(amount):.2f} per visit) from the night of {format_date_from_iso(first_night)}."
            + (f" {changed} open {'job was' if changed == 1 else 'jobs were'} repriced." if changed else "")
        )
    return site_extras_redirect(site_id)


@bp.post("/extras/<int:extra_id>/end")
def end_site_extra(extra_id: int):
    """Stop charging an extra after the last night it applies. The entry stays as the record."""
    db = get_db()
    extra = db.execute("SELECT * FROM site_extras WHERE id = ?", (extra_id,)).fetchone()
    if extra is None:
        flash("Extra not found.")
        return redirect(url_for("gritting.customers"))
    last_night = parse_night(request.form.get("last_night", ""), datetime.now().strftime("%Y-%m-%d"))
    if last_night is None:
        flash("Choose a valid last night.")
    elif last_night < extra["effective_from"]:
        flash("The last night cannot be before the first night. Remove the extra instead if it was entered by mistake.")
    else:
        db.execute("UPDATE site_extras SET effective_to = ? WHERE id = ?", (last_night, extra_id))
        changed = reprice_open_jobs(db, extra["site_id"])
        db.commit()
        flash(f"{extra['description']} now ends after the night of {format_date_from_iso(last_night)}." + (f" {changed} open {'job was' if changed == 1 else 'jobs were'} repriced." if changed else ""))
    return site_extras_redirect(extra["site_id"])


@bp.post("/extras/<int:extra_id>/delete")
def delete_site_extra(extra_id: int):
    """Remove an extra that was entered by mistake. Jobs already done keep their price."""
    db = get_db()
    extra = db.execute("SELECT * FROM site_extras WHERE id = ?", (extra_id,)).fetchone()
    if extra is None:
        flash("Extra not found.")
        return redirect(url_for("gritting.customers"))
    db.execute("DELETE FROM site_extras WHERE id = ?", (extra_id,))
    changed = reprice_open_jobs(db, extra["site_id"])
    db.commit()
    flash(f"Removed {extra['description']}." + (f" {changed} open {'job was' if changed == 1 else 'jobs were'} repriced." if changed else ""))
    return site_extras_redirect(extra["site_id"])


def assign_job_numbers(db: sqlite3.Connection, call_out_ids: list[int]) -> None:
    """Give each of these jobs the next job number, in the order given, unless it already has one.

    A job only gets a number once it is on the list to grit from. A job cancelled before that never has one."""
    for call_out_id in call_out_ids:
        db.execute(
            "UPDATE call_outs SET job_number = (SELECT COALESCE(MAX(job_number), 0) + 1 FROM call_outs) WHERE id = ? AND job_number IS NULL",
            (call_out_id,),
        )


def create_call_out(
    *,
    site_id: int,
    called_out_at: str,
    weather_conditions: str,
    planned_attendance: str,
    service_notes: str,
    pricing_label: str | None = None,
    pricing_amount: float | None = None,
    dispatch_batch_id: int | None = None,
) -> int:
    db = get_db()
    extras_total, extras_items = (0.0, [])
    if pricing_amount is not None:
        # the price for the night is the site price plus the extras in force on that night
        extras_total, extras_items = extras_on(db, site_id, (called_out_at or "")[:10])
        pricing_amount = round(pricing_amount + extras_total, 2)
    cursor = db.execute(
        """
        INSERT INTO call_outs (
            site_id,
            dispatch_batch_id,
            called_out_at,
            weather_conditions,
            planned_attendance,
            service_notes,
            pricing_label,
            pricing_amount,
            extras_amount,
            extras_json,
            status,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'called_out', ?)
        """,
        (
            site_id,
            dispatch_batch_id,
            called_out_at,
            weather_conditions,
            planned_attendance,
            service_notes,
            pricing_label,
            pricing_amount,
            extras_total,
            json.dumps(extras_items),
            now_iso(),
        ),
    )
    call_out_id = int(cursor.lastrowid)
    if dispatch_batch_id is None:
        assign_job_numbers(db, [call_out_id])  # a single call-out goes straight onto the live list; a batch waits for its final list
    return call_out_id


def build_bulk_email_subject(
    site: sqlite3.Row | dict[str, Any],
    called_out_at: str,
    weather_conditions: str = "",
    planned_attendance: str = "",
    service_notes: str = "",
    subject_template: str | None = None,
) -> str:
    context = build_callout_email_context(
        site,
        called_out_at=called_out_at,
        weather_conditions=weather_conditions,
        planned_attendance=planned_attendance,
        service_notes=service_notes,
    )
    if subject_template is None:
        subject_template = fetch_email_templates()["callout_subject"]
    return render_email_template(subject_template, context)


def build_bulk_email_message(
    site: sqlite3.Row | dict[str, Any],
    *,
    called_out_at: str,
    weather_conditions: str,
    planned_attendance: str,
    service_notes: str,
    message_template: str | None = None,
) -> str:
    context = build_callout_email_context(
        site,
        called_out_at=called_out_at,
        weather_conditions=weather_conditions,
        planned_attendance=planned_attendance,
        service_notes=service_notes,
    )
    if message_template is None:
        message_template = fetch_email_templates()["callout_body"]
    return render_email_template(message_template, context)


def build_no_access_email(call_out: sqlite3.Row, access_note: str) -> tuple[str, str]:
    templates = fetch_email_templates()
    context = {
        "company_name": COMPANY_NAME,
        "contact_name": call_out["contact_name"],
        "customer_name": call_out["customer_name"],
        "site_name": call_out["site_name"],
        "called_out_at": format_date_from_iso((call_out["called_out_at"] or "")[:10]),
        "access_note": access_note.strip() or "We could not get in to the site.",
        "service_notes": "",
    }
    return render_email_template(templates["no_access_subject"], context), render_email_template(templates["no_access_body"], context)


def build_completion_email_subject(call_out: sqlite3.Row, completed_at: str) -> str:
    templates = fetch_email_templates()
    context = {
        "company_name": COMPANY_NAME,
        "contact_name": call_out["contact_name"],
        "customer_name": call_out["customer_name"],
        "site_name": call_out["site_name"],
        "called_out_at": format_datetime(call_out["called_out_at"]),
        "service_notes": "",
    }
    return render_email_template(templates["completion_subject"], context)


def build_completion_email_message(
    call_out: sqlite3.Row,
    *,
    completed_at: str,
    service_notes: str,
    pricing_label: str | None,
    pricing_amount: float | None,
) -> str:
    templates = fetch_email_templates()
    context = {
        "company_name": COMPANY_NAME,
        "contact_name": call_out["contact_name"],
        "customer_name": call_out["customer_name"],
        "site_name": call_out["site_name"],
        "called_out_at": format_datetime(call_out["called_out_at"]),
        "service_notes": service_notes.strip(),
    }
    return render_email_template(templates["completion_body"], context)


def complete_call_out_record(
    *,
    call_out_id: int,
    completed_at: str,
    service_notes: str,
    pricing_label: str | None,
    pricing_amount: float | None,
    extras_amount: float = 0.0,
    extras_json: str = "[]",
) -> tuple[sqlite3.Row | None, str, str]:
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        return None, "", ""

    db.execute(
        """
        UPDATE call_outs
        SET status = 'completed',
            pricing_label = ?,
            pricing_amount = ?,
            extras_amount = ?,
            extras_json = ?,
            service_notes = ?,
            completed_at = ?
        WHERE id = ?
        """,
        (
            pricing_label,
            pricing_amount,
            extras_amount if pricing_amount is not None else 0.0,
            extras_json if pricing_amount is not None else "[]",
            service_notes,
            completed_at,
            call_out_id,
        ),
    )

    updated_call_out = fetch_call_out(call_out_id)
    subject = build_completion_email_subject(updated_call_out, completed_at)
    message = build_completion_email_message(
        updated_call_out,
        completed_at=completed_at,
        service_notes=service_notes,
        pricing_label=pricing_label,
        pricing_amount=pricing_amount,
    )
    delivery_status, delivery_details = log_notification(
        call_out_id=call_out_id,
        channel="Email",
        recipient=updated_call_out["email"],
        subject=subject,
        message=message,
        kind="completion",
    )
    return updated_call_out, delivery_status, delivery_details


def flash_completion_message(delivery_status: str) -> None:
    if delivery_status == "sent":
        flash("Job completed and confirmation email sent.")
    elif delivery_status == "logged_only":
        flash("Job completed. Confirmation email was logged, but SMTP is not configured yet.")
    elif delivery_status == "skipped_no_email":
        flash("Job completed, but no customer email address is recorded for the confirmation.")
    else:
        flash("Job completed and the confirmation was logged.")


def fetch_call_out(call_out_id: int) -> sqlite3.Row | None:
    return get_db().execute(
        """
        SELECT
            call_outs.*,
            dispatch_batches.batch_key,
            dispatch_batches.email_subject AS batch_email_subject,
            sites.name AS site_name,
            sites.address,
            sites.postcode,
            sites.route_notes,
            sites.service_days,
            sites.pricing_options,
            customers.name AS customer_name,
            customers.contact_name,
            customers.contact_method,
            customers.contact_value,
            NOTIFY_EMAILS_SQL,
            customers.emergency_phone,
            customers.emergency_contacts
        FROM call_outs
        LEFT JOIN dispatch_batches ON dispatch_batches.id = call_outs.dispatch_batch_id
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        WHERE call_outs.id = ?
        """.replace("NOTIFY_EMAILS_SQL", NOTIFY_EMAILS_SQL),
        (call_out_id,),
    ).fetchone()


def fetch_dashboard_data() -> dict[str, Any]:
    db = get_db()
    season = current_gritting_season()
    season_range = (season["start_date"], season["end_date"])
    stats = {
        "customer_count": db.execute("SELECT COUNT(*) FROM customers").fetchone()[0],
        "site_count": db.execute("SELECT COUNT(*) FROM sites").fetchone()[0],
        "season_call_outs": db.execute(
            """
            SELECT COUNT(*)
            FROM call_outs
            WHERE substr(called_out_at, 1, 10) BETWEEN ? AND ?
            """,
            season_range,
        ).fetchone()[0],
        "season_completed_jobs": db.execute(
            """
            SELECT COUNT(*)
            FROM call_outs
            WHERE substr(called_out_at, 1, 10) BETWEEN ? AND ?
              AND status = 'completed'
            """,
            season_range,
        ).fetchone()[0],
        "season_net_value": db.execute(
            """
            SELECT COALESCE(SUM(pricing_amount), 0)
            FROM call_outs
            WHERE substr(called_out_at, 1, 10) BETWEEN ? AND ?
            """,
            season_range,
        ).fetchone()[0],
        "season_total_value": db.execute(
            """
            SELECT COALESCE(SUM(pricing_amount + (pricing_amount * ?)), 0)
            FROM call_outs
            WHERE substr(called_out_at, 1, 10) BETWEEN ? AND ?
              AND pricing_amount IS NOT NULL
            """,
            (VAT_RATE, *season_range),
        ).fetchone()[0],
        "live_jobs": db.execute(
            """
            SELECT COUNT(*)
            FROM call_outs
            WHERE status NOT IN ('completed', 'cancelled', 'no_access')
            """
        ).fetchone()[0],
    }
    season_rows = db.execute(
        """
        SELECT
            call_outs.id,
            call_outs.job_number,
            call_outs.called_out_at,
            call_outs.completed_at,
            call_outs.status,
            call_outs.weather_conditions,
            call_outs.pricing_label,
            call_outs.pricing_amount,
            sites.name AS site_name,
            customers.name AS customer_name,
            COUNT(DISTINCT notification_logs.id) AS notifications_sent
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        LEFT JOIN notification_logs ON notification_logs.call_out_id = call_outs.id
        WHERE substr(call_outs.called_out_at, 1, 10) BETWEEN ? AND ?
        GROUP BY call_outs.id
        ORDER BY call_outs.called_out_at DESC
        """,
        season_range,
    ).fetchall()
    return {"stats": stats, "season_rows": season_rows, "season": season}


def fetch_season_archive_rows(start_date: str, end_date: str) -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT
            call_outs.id,
            call_outs.job_number,
            call_outs.called_out_at,
            call_outs.completed_at,
            call_outs.status,
            call_outs.weather_conditions,
            call_outs.pricing_label,
            call_outs.pricing_amount,
            call_outs.service_notes,
            sites.name AS site_name,
            customers.name AS customer_name,
            COUNT(DISTINCT notification_logs.id) AS notifications_sent
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        LEFT JOIN notification_logs ON notification_logs.call_out_id = call_outs.id
        WHERE substr(call_outs.called_out_at, 1, 10) BETWEEN ? AND ?
        GROUP BY call_outs.id
        ORDER BY call_outs.called_out_at DESC
        """,
        (start_date, end_date),
    ).fetchall()


def fetch_call_out_rows(*, live_only: bool = False) -> list[sqlite3.Row]:
    where_clause = "WHERE call_outs.status NOT IN ('completed', 'cancelled', 'no_access')" if live_only else ""
    return get_db().execute(
        f"""
        SELECT
            call_outs.id,
            call_outs.job_number,
            call_outs.dispatch_batch_id,
            call_outs.called_out_at,
            call_outs.status,
            call_outs.weather_conditions,
            call_outs.planned_attendance,
            sites.name AS site_name,
            customers.name AS customer_name
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        {where_clause}
        ORDER BY call_outs.called_out_at DESC
        """
    ).fetchall()


def build_new_callout_page_data() -> dict[str, Any]:
    preview_called_out_at = request.args.get("called_out_at") or now_iso()
    eligible_sites_raw, excluded_sites_raw = fetch_bulk_dispatch_candidates(preview_called_out_at)
    eligible_sites = [decorate_site_for_callout(site, preview_called_out_at) for site in eligible_sites_raw]
    excluded_sites = [decorate_site_for_callout(site, preview_called_out_at) for site in excluded_sites_raw]
    time_options = overnight_time_options()
    return {
        "sites": [decorate_site_for_callout(site, now_iso()) for site in fetch_sites()],
        "now_value": now_iso(),
        "smtp_ready": smtp_configured(),
        "email_templates": fetch_email_templates(),
        "preview_called_out_at": preview_called_out_at,
        "default_planned_attendance_start": "18:00",
        "default_planned_attendance_end": "00:00",
        "time_options": time_options,
        "eligible_sites": eligible_sites,
        "excluded_sites": excluded_sites,
        "special_dates_by_day": special_dates_by_day(),
    }


# --- Road surface temperature forecast -------------------------------------------------------------
# Hourly model surface (0 cm) temperature from the Met Office UK model via Open-Meteo, which needs no key.
# It is a model figure, not a road sensor reading: grit at the 0°C trigger or below, and keep an eye on it up to 2°C above.
FORECAST_PLACE = "Norwich"
FORECAST_LATITUDE = 52.63
FORECAST_LONGITUDE = 1.30
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
FORECAST_COMMERCIAL_URL = "https://customer-api.open-meteo.com/v1/forecast"
FORECAST_CACHE_SECONDS = 20 * 60
FORECAST_STALE_LIMIT_SECONDS = 6 * 60 * 60
_forecast_cache: dict[str, Any] = {"at": 0.0, "data": None}
_forecast_lock = threading.Lock()
RISK_ORDER = {"ok": 0, "watch": 1, "grit": 2}


def local_now() -> datetime:
    try:
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo("Europe/London")).replace(tzinfo=None)
    except Exception:
        return datetime.now()


def forecast_request_url() -> str:
    query = (
        f"latitude={FORECAST_LATITUDE}&longitude={FORECAST_LONGITUDE}"
        "&hourly=soil_temperature_0cm,temperature_2m,dew_point_2m,precipitation,snowfall,cloud_cover,wind_speed_10m"
        "&models=ukmo_seamless,ecmwf_ifs025,icon_seamless&past_days=1&forecast_days=8&timezone=Europe%2FLondon&wind_speed_unit=mph"
    )
    api_key = os.environ.get("OPEN_METEO_API_KEY", "").strip()
    if api_key:  # a paid Open-Meteo plan, for commercial use
        return f"{FORECAST_COMMERCIAL_URL}?{query}&apikey={api_key}"
    return f"{FORECAST_URL}?{query}"


def fetch_open_meteo_forecast(timeout: float = 10.0) -> tuple[dict[str, Any], bool]:
    """Return (raw forecast, is_stale). Cached for a few minutes; falls back to an older copy if the service is down."""
    with _forecast_lock:
        age = time.time() - _forecast_cache["at"]
        if _forecast_cache["data"] is not None and age < FORECAST_CACHE_SECONDS:
            return _forecast_cache["data"], False
        try:
            request_obj = urllib.request.Request(forecast_request_url(), headers={"User-Agent": "ngs-gritting-app"})
            with urllib.request.urlopen(request_obj, timeout=timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
            if not isinstance(data.get("hourly"), dict):
                raise ValueError("Forecast response had no hourly data.")
        except Exception:
            if _forecast_cache["data"] is not None and age < FORECAST_STALE_LIMIT_SECONDS:
                return _forecast_cache["data"], True
            raise
        _forecast_cache["at"] = time.time()
        _forecast_cache["data"] = data
        return data, False


GRIT_TRIGGER_C = 0.0  # our trigger: grit when the forecast road surface is at this temperature or lower
WATCH_SURFACE_C = GRIT_TRIGGER_C + 2.0  # "keep an eye on it" up to 2 degrees above the trigger


def classify_road_hour(row: dict[str, Any]) -> tuple[str, str]:
    """Risk for one hour from the forecast road surface temperature: grit at the trigger or below."""
    surface = row["surface"]
    if surface <= GRIT_TRIGGER_C:
        return "grit", f"Road surface at or below {GRIT_TRIGGER_C:g}°C"
    if surface <= WATCH_SURFACE_C:
        return "watch", "Road surface close to freezing"
    return "ok", ""


def forecast_rows(raw: dict[str, Any], bias: float = 0.0) -> list[dict[str, Any]]:
    hourly = raw.get("hourly") or {}
    times = hourly.get("time") or []

    def series(name: str) -> list[Any]:
        values = hourly.get(name) or []
        return values + [None] * (len(times) - len(values))

    surface, air, dew = series("soil_temperature_0cm"), series("temperature_2m"), series("dew_point_2m")
    rain, snow, cloud, wind = series("precipitation"), series("snowfall"), series("cloud_cover"), series("wind_speed_10m")
    rows = []
    for index, stamp in enumerate(times):
        if surface[index] is None or air[index] is None:
            continue
        try:
            moment = datetime.fromisoformat(stamp)
        except ValueError:
            continue
        row = {"time": moment, "surface": float(surface[index]) + bias, "air": float(air[index]),
               "dew": None if dew[index] is None else float(dew[index]), "rain": rain[index], "snow": snow[index],
               "cloud": cloud[index], "wind": wind[index]}
        row["level"], row["reason"] = classify_road_hour(row)
        rows.append(row)
    return rows


def night_label(start: datetime, today: datetime) -> str:
    days_ahead = (start.date() - today.date()).days
    if days_ahead <= 0:
        return "Tonight"
    if days_ahead == 1:
        return "Tomorrow night"
    return start.strftime("%A") + " night"


def is_wet_hour(row: dict[str, Any]) -> bool:
    return float(row["rain"] or 0) >= 0.1 or float(row["snow"] or 0) > 0


def rain_for_night(wet_hours: list[dict[str, Any]], risky_hours: list[dict[str, Any]]) -> dict[str, Any]:
    """Rain during a night, and when to grit if rain is forecast on a night that needs it.

    Salt put down before rain can be washed off, so on a night that needs grit the best time is once the rain has stopped."""
    info: dict[str, Any] = {"rain_from": "", "rain_until": "", "rain_mm": 0.0, "rain_advice": "", "grit_after_rain_at": ""}
    if not wet_hours:
        return info
    rain_end = wet_hours[-1]["time"] + timedelta(hours=1)
    info.update(rain_from=wet_hours[0]["time"].strftime("%H:%M"), rain_until=rain_end.strftime("%H:%M"),
                rain_mm=round(sum(float(row["rain"] or 0) + float(row["snow"] or 0) * 10 for row in wet_hours), 1))
    if risky_hours:
        risk_start, risk_end = risky_hours[0]["time"], risky_hours[-1]["time"] + timedelta(hours=1)
        if risk_end > rain_end:
            best = max(rain_end, risk_start)
            info.update(grit_after_rain_at=best.strftime("%H:%M"),
                        rain_advice=f"Rain until {rain_end.strftime('%H:%M')}: grit after it stops, from {best.strftime('%H:%M')}.")
        else:
            info["rain_advice"] = "Rain is forecast through the cold spell, so salt may wash off. Check again once the rain eases."
    return info


def night_summary(level: str, wet_hours: list[dict[str, Any]], rain_info: dict[str, Any]) -> str:
    """A plain sentence about the night for the call-out wording. No temperatures in it."""
    sky = {"grit": "Frost expected overnight, icy roads likely.",
           "watch": "Cold overnight, a chance of frost on the roads.",
           "ok": "No frost expected overnight."}[level]
    if not wet_hours:
        wet = "Dry."
    else:
        kind = "Snow" if any(float(row["snow"] or 0) > 0 for row in wet_hours) else "Rain"
        wet = f"{kind} from {rain_info['rain_from']} to {rain_info['rain_until']}."
    return f"{sky} {wet}"


def next_rain_spell(upcoming: list[dict[str, Any]]) -> dict[str, Any] | None:
    """The first run of wet hours from now: when it starts, when it stops and how much falls."""
    spell: list[dict[str, Any]] = []
    for row in upcoming:
        if is_wet_hour(row):
            if spell and row["time"] - spell[-1]["time"] > timedelta(hours=1):
                break
            spell.append(row)
        elif spell:
            break
    if not spell:
        return None
    return {"from": spell[0]["time"].strftime("%a %H:%M"), "until": (spell[-1]["time"] + timedelta(hours=1)).strftime("%a %H:%M"),
            "mm": round(sum(float(row["rain"] or 0) + float(row["snow"] or 0) * 10 for row in spell), 1)}


RISK_HEADLINES = {"grit": "GRITTING REQUIRED", "watch": "Keep an eye on it", "ok": "NO Gritting"}


# Three forecast models, each with its own idea of the weather. The page uses the middle value, so one model that is
# far colder or warmer than the others (common five or six days out) cannot swing the answer on its own.
MODELS = {"ukmo_seamless": "Met Office", "ecmwf_ifs025": "ECMWF", "icon_seamless": "ICON"}
OUTLOOK_AFTER_DAYS = 3  # nights further ahead than this are marked as an early outlook


def split_models(raw: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """A response for several models names every series after its model; return one plain response per model.

    A response without model names (older cached data, or the weather page's own forecast) is one model on its own."""
    hourly = raw.get("hourly") or {}
    if not any(key.endswith(f"_{model}") for key in hourly for model in MODELS):
        return {"ukmo_seamless": raw}
    found: dict[str, dict[str, Any]] = {}
    for model in MODELS:
        suffix = f"_{model}"
        data: dict[str, Any] = {"time": hourly.get("time") or []}
        for key, values in hourly.items():
            if key.endswith(suffix):
                data[key[: -len(suffix)]] = values
        if data.get("soil_temperature_0cm") and data.get("temperature_2m"):
            found[model] = {"hourly": data}
    return found


def _middle(values: list[float]) -> float:
    ordered = sorted(values)
    count = len(ordered)
    return ordered[count // 2] if count % 2 else (ordered[count // 2 - 1] + ordered[count // 2]) / 2


def consensus_rows(per_model: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    """One row per hour using the middle value of the models that have that hour; keeps each model's own reading."""
    by_time: dict[datetime, dict[str, dict[str, Any]]] = {}
    for model, rows in per_model.items():
        for row in rows:
            by_time.setdefault(row["time"], {})[model] = row
    merged = []
    for moment in sorted(by_time):
        readings = by_time[moment]
        group = list(readings.values())

        def middle(key: str) -> float | None:
            values = [float(row[key]) for row in group if row[key] is not None]
            return _middle(values) if values else None

        rain = [float(row["rain"] or 0) for row in group]
        snow = [float(row["snow"] or 0) for row in group]
        row = {"time": moment, "surface": middle("surface"), "air": middle("air"), "dew": middle("dew"),
               "rain": sum(rain) / len(rain), "snow": sum(snow) / len(snow), "cloud": middle("cloud"), "wind": middle("wind"),
               "models": {model: item["surface"] for model, item in readings.items()},
               "model_air": {model: item["air"] for model, item in readings.items()},
               "model_levels": [item["level"] for item in group]}
        row["level"], row["reason"] = classify_road_hour(row)
        row["lone_grit"] = len(group) > 1 and row["level"] != "grit" and "grit" in row["model_levels"]
        merged.append(row)
    return merged


def road_forecast_from_open_meteo(raw: dict[str, Any], now: datetime, night_count: int = 7, bias: float = 0.0) -> dict[str, Any]:
    """The next seven nights (18:00 to 06:00) and the next seven days hour by hour, from a raw Open-Meteo response."""
    per_model = {model: forecast_rows(part, bias) for model, part in split_models(raw).items()}
    per_model = {model: rows for model, rows in per_model.items() if rows}
    if not per_model:
        raise ValueError("The forecast had no surface temperatures.")
    rows = consensus_rows(per_model)
    near_cutoff = datetime.combine(now.date() + timedelta(days=OUTLOOK_AFTER_DAYS + 1), datetime.min.time())
    for row in rows:  # inside the first few days, one model showing a freeze is worth a look even when the others are mild
        if row["lone_grit"] and row["level"] == "ok" and row["time"] < near_cutoff:
            row["level"], row["reason"] = "watch", "One forecast model shows a freeze"
    first_start = (now if now.hour >= 6 else now - timedelta(days=1)).replace(hour=18, minute=0, second=0, microsecond=0)
    nights = []
    for offset in range(night_count):
        start = first_start + timedelta(days=offset)
        window = [row for row in rows if start <= row["time"] < start + timedelta(hours=12)]
        if not window:
            continue
        coldest = min(window, key=lambda row: row["surface"])
        model_lows = {MODELS[model]: round(min(row["models"][model] for row in window if model in row["models"]), 1)
                      for model in MODELS if any(model in row["models"] for row in window)}
        worst = max(window, key=lambda row: (RISK_ORDER[row["level"]], -row["surface"]))
        risky_hours = [row for row in window if row["level"] == "grit"]
        wet_hours = [row for row in window if is_wet_hour(row)]
        rain_info = rain_for_night(wet_hours, risky_hours)
        nights.append({
            "summary": night_summary(worst["level"], wet_hours, rain_info),
            "label": night_label(start, now),
            "date": start.strftime("%Y-%m-%d"),
            "level": worst["level"],
            "headline": RISK_HEADLINES[worst["level"]],
            "reason": worst["reason"],
            "min_surface": round(coldest["surface"], 1),
            "min_surface_time": coldest["time"].strftime("%H:%M"),
            "min_surface_day": coldest["time"].strftime("%a"),  # the low is often after midnight, so say which day
            "min_air": round(min(row["air"] for row in window), 1),
            "model_lows": model_lows,
            "models_disagree": any(row["lone_grit"] for row in window) or (len(model_lows) > 1 and max(model_lows.values()) - min(model_lows.values()) > 3),
            "outlook": (start.date() - now.date()).days > OUTLOOK_AFTER_DAYS,
            "grit_from": risky_hours[0]["time"].strftime("%H:%M") if risky_hours else "",
            "grit_until": (risky_hours[-1]["time"] + timedelta(hours=1)).strftime("%H:%M") if risky_hours else "",
            "hours_covered": len(window),  # fewer than 12 means the forecast stops part way through the night
            **rain_info,
        })
    this_hour = now.replace(minute=0, second=0, microsecond=0)
    upcoming = [row for row in rows if row["time"] >= this_hour][: 7 * 24]

    def hour_entry(row: dict[str, Any]) -> dict[str, Any]:
        return {"time": row["time"].strftime("%a %H:%M"), "hour": row["time"].strftime("%H:%M"), "day": row["time"].strftime("%a"),
                "date": row["time"].strftime("%Y-%m-%d"),
                "rain": round(float(row["rain"] or 0) + float(row["snow"] or 0) * 10, 1),  # snow as mm of water, roughly 10 cm per 10 mm
                "surface": round(row["surface"], 1), "air": round(row["air"], 1),
                "dew": None if row["dew"] is None else round(row["dew"], 1), "level": row["level"]}

    return {
        "place": FORECAST_PLACE,
        "nights": nights,
        "week": [hour_entry(row) for row in upcoming],
        # from 09:00 on the first night's date, so a night can be shown for its whole 24 hours, 09:00 to 09:00
        "hourly": [hour_entry(row) for row in rows if row["time"] >= first_start.replace(hour=9, minute=0)],
        "next_rain": next_rain_spell(upcoming),
    }


# ---- Checking the forecast against real road readings ------------------------------------------------------
BIAS_MIN_READINGS = 5      # no correction until there are this many readings to learn from
BIAS_LIMIT_C = 3.0         # never move the forecast by more than this
BIAS_RECENT_DAYS = 90


def forecast_surface_at(moment: datetime) -> tuple[float | None, float | None]:
    """What the forecast (before any correction) gave for the road surface and air at this hour, if it is still on hand."""
    try:
        raw, _ = fetch_open_meteo_forecast()
        per_model = {model: forecast_rows(part) for model, part in split_models(raw).items()}
        rows = consensus_rows({model: rows for model, rows in per_model.items() if rows})
    except Exception:
        return None, None
    hour = moment.replace(minute=0, second=0, microsecond=0)
    if moment.minute >= 30:
        hour += timedelta(hours=1)
    match = next((row for row in rows if row["time"] == hour), None)
    return (None, None) if match is None else (round(match["surface"], 1), round(match["air"], 1))


def road_accuracy(db: sqlite3.Connection) -> dict[str, Any]:
    """How the forecast compares with the readings taken, and the correction that follows from it."""
    since = (datetime.now() - timedelta(days=BIAS_RECENT_DAYS)).strftime("%Y-%m-%dT%H:%M")
    rows = db.execute(
        "SELECT reading, forecast_surface FROM road_readings WHERE forecast_surface IS NOT NULL AND read_at >= ? ORDER BY read_at",
        (since,),
    ).fetchall()
    diffs = sorted(row["reading"] - row["forecast_surface"] for row in rows)
    count = len(diffs)
    result: dict[str, Any] = {"count": count, "bias": 0.0, "median_miss": None, "typical_error": None, "needed": max(BIAS_MIN_READINGS - count, 0)}
    if count:
        middle = diffs[count // 2] if count % 2 else (diffs[count // 2 - 1] + diffs[count // 2]) / 2
        result["median_miss"] = round(middle, 1)
        result["typical_error"] = round(sum(abs(value) for value in diffs) / count, 1)
        if count >= BIAS_MIN_READINGS:
            shrunk = middle * count / (count + BIAS_MIN_READINGS)  # trust it more as readings build up
            result["bias"] = round(max(-BIAS_LIMIT_C, min(BIAS_LIMIT_C, shrunk)), 1)
    return result


def current_road_bias() -> float:
    return float(road_accuracy(get_db())["bias"])


@bp.post("/callouts/<int:call_out_id>/road-reading")
def save_road_reading(call_out_id: int):
    """A real road surface temperature taken on site, kept next to what the forecast said for that hour."""
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    back = redirect(url_for("gritting.live_jobs") if request.form.get("next", "").strip() == "live_jobs" else url_for("gritting.callout_detail", call_out_id=call_out_id))
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    try:
        reading = round(float(request.form.get("reading", "").strip()), 1)
    except ValueError:
        flash("Enter the road temperature as a number, for example -1.5.")
        return back
    if not -30 <= reading <= 45:
        flash("That does not look like a road temperature. Enter it in degrees C.")
        return back
    moment = datetime.now()
    forecast_surface, forecast_air = forecast_surface_at(moment)
    db.execute(
        "INSERT INTO road_readings (call_out_id, site_id, read_at, reading, forecast_surface, forecast_air, note) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (call_out_id, call_out["site_id"], moment.strftime("%Y-%m-%dT%H:%M"), reading, forecast_surface, forecast_air, request.form.get("note", "").strip()[:200]),
    )
    db.commit()
    where = f"{call_out['customer_name']} · {call_out['site_name']}"
    if forecast_surface is None:
        flash(f"Saved {reading}°C at {where}. The forecast for this hour was not available to compare with.")
    else:
        miss = round(reading - forecast_surface, 1)
        flash(f"Saved {reading}°C at {where}. The forecast said {forecast_surface}°C, so it was {abs(miss)}°C {'colder' if miss < 0 else 'warmer'} than forecast." if miss else f"Saved {reading}°C at {where}, exactly as forecast.")
    return back


@bp.route("/road-accuracy")
def road_accuracy_page() -> str:
    db = get_db()
    readings = db.execute(
        """
        SELECT road_readings.*, sites.name AS site_name, customers.name AS customer_name
        FROM road_readings
        LEFT JOIN sites ON sites.id = road_readings.site_id
        LEFT JOIN customers ON customers.id = sites.customer_id
        ORDER BY road_readings.read_at DESC
        LIMIT 60
        """
    ).fetchall()
    return render_template("gritting/road_accuracy.html", readings=readings, accuracy=road_accuracy(db), limit=BIAS_LIMIT_C, minimum=BIAS_MIN_READINGS)


@bp.get("/road-forecast.json")
def road_forecast_json() -> Any:
    try:
        raw, stale = fetch_open_meteo_forecast()
        accuracy = road_accuracy(get_db())
        payload = road_forecast_from_open_meteo(raw, local_now(), bias=accuracy["bias"])
        payload.update(ok=True, stale=stale, updated=datetime.now().strftime("%H:%M"), bias=accuracy["bias"], readings=accuracy["count"])
    except Exception:
        payload = {"ok": False, "message": "The forecast could not be loaded just now. Check the internet connection and try again."}
    response = jsonify(payload)
    response.headers["Cache-Control"] = "no-store"
    return response


@bp.route("/")
def dashboard() -> str:
    data = fetch_dashboard_data()
    return render_template("gritting/dashboard.html", forecast_place=FORECAST_PLACE, **data)


def fetch_night_lists() -> list[dict[str, Any]]:
    """The lists still being worked: one per call-out batch with a job still to do, plus single call-outs."""
    db = get_db()
    rows = db.execute(
        """
        SELECT
            call_outs.id,
            call_outs.job_number,
            call_outs.dispatch_batch_id,
            call_outs.called_out_at,
            call_outs.status,
            call_outs.planned_attendance,
            call_outs.weather_conditions,
            call_outs.pricing_label,
            call_outs.extras_json,
            call_outs.completed_at,
            call_outs.cancelled_at,
            call_outs.no_access_at,
            call_outs.access_note,
            sites.name AS site_name,
            sites.address,
            sites.postcode,
            sites.route_notes,
            customers.name AS customer_name,
            customers.emergency_phone,
            customers.emergency_contacts
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        WHERE call_outs.status = 'called_out'
           OR call_outs.dispatch_batch_id IN (
                SELECT dispatch_batch_id FROM call_outs WHERE status = 'called_out' AND dispatch_batch_id IS NOT NULL)
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """
    ).fetchall()
    batches = {row["id"]: row for row in db.execute("SELECT * FROM dispatch_batches").fetchall()}
    grouped: dict[Any, dict[str, Any]] = {}
    for row in rows:
        key = row["dispatch_batch_id"]
        if key is None and row["status"] != "called_out":
            continue
        group = grouped.setdefault(key, {"batch_id": key, "batch": batches.get(key), "rows": []})
        group["rows"].append(row)
    lists = []
    for group in grouped.values():
        batch = group["batch"]
        group["called_out_at"] = batch["called_out_at"] if batch else group["rows"][0]["called_out_at"]
        group["planned_attendance"] = batch["planned_attendance"] if batch else group["rows"][0]["planned_attendance"]
        group["weather"] = batch["weather_conditions"] if batch else ""
        group["finalised_at"] = batch["finalised_at"] if batch else None
        rank = {"called_out": 0, "no_access": 1, "cancelled": 2, "completed": 3}  # still to grit first, finished jobs last
        group["rows"].sort(key=lambda row: (rank.get(row["status"], 3), row["customer_name"].lower(), row["site_name"].lower()))
        group["counts"] = {
            "to_grit": sum(1 for row in group["rows"] if row["status"] == "called_out"),
            "cancelled": sum(1 for row in group["rows"] if row["status"] == "cancelled"),
            "no_access": sum(1 for row in group["rows"] if row["status"] == "no_access"),
            "done": sum(1 for row in group["rows"] if row["status"] == "completed"),
        }
        lists.append(group)
    lists.sort(key=lambda group: group["called_out_at"], reverse=True)
    lists.sort(key=lambda group: group["batch_id"] is None)  # batches first, single call-outs last
    return lists


@bp.route("/live-jobs")
def live_jobs() -> str:
    lists = fetch_night_lists()
    live_count = sum(group["counts"]["to_grit"] for group in lists)
    for group in lists:
        ids = [row["id"] for row in group["rows"]]
        group["email_problems"] = email_problems(get_db(), ids)
        group["unsent_emails"] = len(unsent_emails(get_db(), ids, live_only=True))
        group["smtp_ready"] = smtp_configured()
    return render_template("gritting/live_jobs.html", lists=lists, live_count=live_count)


@bp.post("/batches/<int:batch_id>/resend-emails")
def resend_batch_emails(batch_id: int):
    """Send the night's call-out emails again, for any that were only logged or failed."""
    db = get_db()
    if not smtp_configured():
        flash("Email is not set up yet. Add the mail server under Settings first.")
        return redirect(url_for("gritting.live_jobs"))
    ids = [row["id"] for row in db.execute("SELECT id FROM call_outs WHERE dispatch_batch_id = ?", (batch_id,)).fetchall()]
    counts = resend_unsent_emails(db, ids, live_only=True)
    db.commit()
    flash_resend_result(counts)
    return redirect(url_for("gritting.live_jobs") + f"#list-{batch_id}")


@bp.post("/callouts/<int:call_out_id>/resend-emails")
def resend_callout_emails(call_out_id: int):
    db = get_db()
    if not smtp_configured():
        flash("Email is not set up yet. Add the mail server under Settings first.")
    else:
        counts = resend_unsent_emails(db, [call_out_id], live_only=False)
        db.commit()
        flash_resend_result(counts)
    return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))


@bp.post("/batches/<int:batch_id>/finalise")
def finalise_batch(batch_id: int):
    """Lock in the night's list once the cancellations are done: this is the list to go gritting from."""
    db = get_db()
    batch = db.execute("SELECT * FROM dispatch_batches WHERE id = ?", (batch_id,)).fetchone()
    if batch is None:
        flash("List not found.")
        return redirect(url_for("gritting.live_jobs"))
    counts = db.execute(
        "SELECT SUM(status = 'called_out'), SUM(status = 'cancelled') FROM call_outs WHERE dispatch_batch_id = ?", (batch_id,)
    ).fetchone()
    to_grit, cancelled = int(counts[0] or 0), int(counts[1] or 0)
    if to_grit == 0:
        flash("There is nothing left to grit on this list, so it cannot be finalised.")
        return redirect(url_for("gritting.live_jobs"))
    db.execute("UPDATE dispatch_batches SET finalised_at = ? WHERE id = ?", (now_iso(), batch_id))
    ordered = db.execute(
        """
        SELECT call_outs.id FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        WHERE call_outs.dispatch_batch_id = ? AND call_outs.status IN ('called_out', 'no_access')
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """,
        (batch_id,),
    ).fetchall()
    assign_job_numbers(db, [row["id"] for row in ordered])  # job numbers are given now, to the sites we are going to
    db.commit()
    flash(f"List finalised: {to_grit} {'site' if to_grit == 1 else 'sites'} to grit" + (f", {cancelled} not needed." if cancelled else "."))
    return redirect(url_for("gritting.live_jobs") + f"#list-{batch_id}")


@bp.post("/batches/<int:batch_id>/reopen")
def reopen_batch(batch_id: int):
    db = get_db()
    started = db.execute("SELECT COUNT(*) FROM call_outs WHERE dispatch_batch_id = ? AND status = 'completed'", (batch_id,)).fetchone()[0]
    if started:
        # a draft list never holds finished work, so a list cannot go back to draft once a job is done
        flash("Some jobs on this list are already done, so it cannot be reopened.")
        return redirect(url_for("gritting.live_jobs") + f"#list-{batch_id}")
    db.execute("UPDATE dispatch_batches SET finalised_at = NULL WHERE id = ?", (batch_id,))
    db.execute("UPDATE call_outs SET job_number = NULL WHERE dispatch_batch_id = ?", (batch_id,))  # numbers are given again at the next finalise
    db.commit()
    flash("List reopened. Finalise it again when the changes are done.")
    return redirect(url_for("gritting.live_jobs") + f"#list-{batch_id}")


@bp.route("/batches/<int:batch_id>/run-sheet")
def run_sheet(batch_id: int) -> str:
    """The list to take out gritting: sites still to do first, then anything already done and anything not needed."""
    db = get_db()
    batch = db.execute("SELECT * FROM dispatch_batches WHERE id = ?", (batch_id,)).fetchone()
    if batch is None:
        flash("List not found.")
        return redirect(url_for("gritting.live_jobs"))
    rows = db.execute(
        """
        SELECT call_outs.id, call_outs.job_number, call_outs.status, call_outs.pricing_label, call_outs.extras_json, call_outs.planned_attendance, call_outs.access_note,
               sites.name AS site_name, sites.address, sites.postcode, sites.route_notes,
               customers.name AS customer_name, customers.contact_name, customers.emergency_phone, customers.emergency_contacts
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        WHERE call_outs.dispatch_batch_id = ?
        ORDER BY customers.name COLLATE NOCASE, sites.name COLLATE NOCASE
        """,
        (batch_id,),
    ).fetchall()
    return render_template(
        "gritting/run_sheet.html",
        batch=batch,
        to_grit=[row for row in rows if row["status"] == "called_out"],
        done=[row for row in rows if row["status"] == "completed"],
        cancelled=[row for row in rows if row["status"] == "cancelled"],
        no_access=[row for row in rows if row["status"] == "no_access"],
    )


@bp.route("/new-callout", methods=["GET", "POST"])
def new_callout() -> str:
    db = get_db()
    if request.method == "POST":
        called_out_at = request.form.get("called_out_at") or now_iso()
        planned_attendance = planned_attendance_from_form(request.form)
        site_id = int(request.form["site_id"])
        site_row = db.execute("SELECT pricing_options FROM sites WHERE id = ?", (request.form["site_id"],)).fetchone()
        site_email_row = fetch_site_notification_row(site_id)
        site_pricing_options = load_pricing_options(site_row["pricing_options"] if site_row else None)
        pricing_label, pricing_amount = selected_pricing_details(
            site_pricing_options,
            request.form.get("pricing_label"),
        )
        call_out_id = create_call_out(
            site_id=site_id,
            called_out_at=called_out_at,
            weather_conditions=request.form["weather_conditions"].strip(),
            planned_attendance=planned_attendance,
            service_notes=request.form["service_notes"].strip(),
            pricing_label=pricing_label,
            pricing_amount=pricing_amount,
        )
        rendered_subject = request.form.get("subject", "").strip()
        if not rendered_subject and site_email_row is not None:
            rendered_subject = build_bulk_email_subject(
                site_email_row,
                called_out_at,
                request.form["weather_conditions"].strip(),
                planned_attendance,
                request.form["service_notes"].strip(),
            )
        rendered_message = request.form.get("message", "").strip()
        if not rendered_message and site_email_row is not None:
            rendered_message = build_bulk_email_message(
                site_email_row,
                called_out_at=called_out_at,
                weather_conditions=request.form["weather_conditions"].strip(),
                planned_attendance=planned_attendance,
                service_notes=request.form["service_notes"].strip(),
            )
        delivery_status, _ = log_notification(
            kind="callout",
            call_out_id=call_out_id,
            channel="Email",
            recipient=request.form["recipient"].strip() or (site_email_row["email"] if site_email_row else ""),
            subject=rendered_subject,
            message=rendered_message,
        )
        db.commit()
        if delivery_status == "sent":
            flash("Call-out created and notification sent.")
        elif delivery_status == "logged_only":
            flash("Call-out created. Email logging worked, but SMTP is not configured yet.")
        else:
            flash("Call-out created and notification logged.")
        return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))

    return render_template("gritting/new_callout.html", **build_new_callout_page_data())


def fetch_extras_by_site() -> dict[int, list[dict[str, Any]]]:
    """Every extra for each site, newest first, with whether it is charged today, starts later or has ended."""
    db = get_db()
    today = datetime.now().strftime("%Y-%m-%d")
    grouped: dict[int, list[dict[str, Any]]] = {}
    for row in db.execute("SELECT * FROM site_extras ORDER BY effective_from DESC, id DESC").fetchall():
        item = dict(row)
        if item["effective_to"] and item["effective_to"] < today:
            item["state"] = "ended"
        elif item["effective_from"] > today:
            item["state"] = "later"
        else:
            item["state"] = "active"
        grouped.setdefault(int(row["site_id"]), []).append(item)
    return grouped


def read_contact_fields(*, need_customer: bool, need_site: bool = True) -> tuple[dict[str, Any] | None, str]:
    """Emails, emergency contacts and postcode from the posted form: (values, error). An error means nothing is saved."""
    form = request.form
    customer_emails, bad_customer = emails_from_form(form, "email")
    site_emails, bad_site = emails_from_form(form, "site_emails")
    bad = bad_customer + bad_site
    if bad:
        return None, "Check the email address: " + ", ".join(bad) + "."
    if need_customer and not customer_emails:
        return None, "Enter at least one notification email for the customer."
    contacts, contact_error = emergency_contacts_from_form(form)
    if contact_error:
        return None, contact_error
    postcode = normalize_postcode(form.get("postcode", "")) if need_site else ""
    if postcode is None:
        return None, "That does not look like a UK postcode. Use the form NR15 1BE."
    return {
        "emails": ", ".join(customer_emails),
        "site_emails": ", ".join(site_emails),
        "contacts_json": json.dumps(contacts),
        "first_phone": contacts[0]["phone"] if contacts else "",
        "postcode": postcode,
    }, ""


def add_customer_and_site() -> Any:
    """One form: a new customer with their first site, or another site for a customer already on the list."""
    db = get_db()
    form = request.form
    site_name, address = form.get("site_name", "").strip(), form.get("address", "").strip()
    choice = form.get("customer_choice", "new").strip()
    if not site_name or not address:
        flash("Enter the site name and address.")
        return redirect(url_for("gritting.customers"))
    if choice == "new":
        name, contact = form.get("customer_name", "").strip(), form.get("contact_name", "").strip()
        if not name or not contact:
            flash("Enter the customer name and contact name.")
            return redirect(url_for("gritting.customers"))
        values, error = read_contact_fields(need_customer=True)
        if values is None:
            flash(error)
            return redirect(url_for("gritting.customers"))
        cursor = db.execute(
            "INSERT INTO customers (name, contact_name, contact_method, contact_value, email, emergency_phone, emergency_contacts) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (name, contact, "Email", values["first_phone"], values["emails"], values["first_phone"], values["contacts_json"]),
        )
        customer_id = int(cursor.lastrowid)
        customer_label = name
    else:
        row = db.execute("SELECT id, name FROM customers WHERE id = ?", (choice,)).fetchone()
        if row is None:
            flash("That customer could not be found.")
            return redirect(url_for("gritting.customers"))
        customer_id, customer_label = int(row["id"]), row["name"]
        values, error = read_contact_fields(need_customer=False)  # only the site's own extra emails and postcode apply
        if values is None:
            flash(error)
            return redirect(url_for("gritting.customers"))
    selected_days = [day for day in DAY_KEYS if form.get(f"service_day_{day}") == "on"]
    db.execute(
        "INSERT INTO sites (customer_id, name, address, postcode, notify_emails, route_notes, service_days, pricing_options) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            customer_id,
            site_name,
            address,
            values["postcode"],
            values["site_emails"],
            form.get("route_notes", "").strip(),
            serialize_service_days(selected_days or DAY_KEYS),
            serialize_pricing_options(pricing_options_from_form(form)),
        ),
    )
    db.commit()
    flash(f"{customer_label}: site {site_name} added." if choice != "new" else f"Customer {customer_label} and site {site_name} added.")
    return redirect(url_for("gritting.customers"))


@bp.route("/customers", methods=["GET", "POST"])
def customers() -> str:
    db = get_db()
    if request.method == "POST":
        form_type = request.form.get("form_type", "customer").strip()
        if form_type == "customer_site":
            return add_customer_and_site()
        if form_type == "site":
            selected_days = [day for day in DAY_KEYS if request.form.get(f"service_day_{day}") == "on"]
            db.execute(
                """
                INSERT INTO sites (customer_id, name, address, route_notes, service_days, pricing_options)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    request.form["customer_id"],
                    request.form["name"].strip(),
                    request.form["address"].strip(),
                    request.form["route_notes"].strip(),
                    serialize_service_days(selected_days or DAY_KEYS),
                    serialize_pricing_options(pricing_options_from_form(request.form)),
                ),
            )
            db.commit()
            flash("Site added.")
        elif form_type == "edit_customer":
            customer_id = int(request.form["customer_id"])
            values, error = read_contact_fields(need_customer=True, need_site=False)
            if values is None:
                flash(error)
                return redirect(url_for("gritting.customers") + f"#customer-{customer_id}")
            db.execute(
                """
                UPDATE customers
                SET
                    name = ?,
                    contact_name = ?,
                    contact_method = 'Email',
                    contact_value = ?,
                    email = ?,
                    emergency_phone = ?,
                    emergency_contacts = ?
                WHERE id = ?
                """,
                (
                    request.form["name"].strip(),
                    request.form["contact_name"].strip(),
                    values["first_phone"],
                    values["emails"],
                    values["first_phone"],
                    values["contacts_json"],
                    customer_id,
                ),
            )
            db.commit()
            flash("Customer updated.")
            return redirect(url_for("gritting.customers") + f"#customer-{customer_id}")
        elif form_type == "edit_site":
            site_id = int(request.form["site_id"])
            values, error = read_contact_fields(need_customer=False)
            if values is None:
                flash(error)
                return redirect(url_for("gritting.customers") + f"#customer-site-{site_id}")
            selected_days = [day for day in DAY_KEYS if request.form.get(f"service_day_{day}") == "on"]
            db.execute(
                """
                UPDATE sites
                SET
                    name = ?,
                    address = ?,
                    postcode = ?,
                    notify_emails = ?,
                    route_notes = ?,
                    service_days = ?,
                    pricing_options = ?
                WHERE id = ?
                """,
                (
                    request.form["name"].strip(),
                    request.form["address"].strip(),
                    values["postcode"],
                    values["site_emails"],
                    request.form["route_notes"].strip(),
                    serialize_service_days(selected_days or DAY_KEYS),
                    serialize_pricing_options(pricing_options_from_form(request.form)),
                    site_id,
                ),
            )
            db.commit()
            flash("Site updated.")
            return redirect(url_for("gritting.customers") + f"#customer-site-{site_id}")
        else:
            old_contacts = emergency_contacts_from_form(request.form)[0]
            old_phone = old_contacts[0]["phone"] if old_contacts else ""
            db.execute(
                """
                INSERT INTO customers (
                    name,
                    contact_name,
                    contact_method,
                    contact_value,
                    email,
                    emergency_phone,
                    emergency_contacts
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    request.form["name"].strip(),
                    request.form["contact_name"].strip(),
                    "Email",
                    old_phone,
                    ", ".join(emails_from_form(request.form, "email")[0]),
                    old_phone,
                    json.dumps(old_contacts),
                ),
            )
            db.commit()
            flash("Customer added.")
        return redirect(url_for("gritting.customers"))

    return render_template(
        "gritting/customers.html",
        customers=fetch_customers(),
        sites=fetch_sites(),
        directory_rows=fetch_customer_directory_rows(),
        extras_by_site=fetch_extras_by_site(),
        event_names=[row["name"] for row in get_db().execute("SELECT DISTINCT name FROM special_dates ORDER BY name").fetchall()],
        today=datetime.now().strftime("%Y-%m-%d"),
        tomorrow=(datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d"),
    )


@bp.route("/settings", methods=["GET", "POST"])
def settings() -> str:
    db = get_db()
    if request.method == "POST":
        form_type = request.form.get("form_type", "smtp").strip()
        if form_type == "email_setup":
            save_email_templates(db, request.form)
            flash("Email setup updated.")
        elif form_type == "special_date_add":
            event_date = parse_night(request.form.get("event_date", ""))
            name = request.form.get("event_name", "").strip()
            if not event_date or not name:
                flash("Enter a date and a name, for example Norwich home match.")
            else:
                db.execute("INSERT OR IGNORE INTO special_dates (event_date, name) VALUES (?, ?)", (event_date, name))
                flash(f"{name} added on {event_date}.")
        elif form_type == "special_date_delete":
            db.execute("DELETE FROM special_dates WHERE id = ?", (request.form.get("special_date_id", "0"),))
            flash("Special date removed.")
        elif form_type == "season_rollover":
            archived_season, archive_path = start_new_gritting_season_rollover()
            flash(
                f"{archived_season['label']} archived to {archive_path.relative_to(DATA_ROOT)} and the app has been switched to the next season."
            )
        else:
            save_smtp_settings(db, request.form)
            save_report_settings(db, request.form)
            flash("Settings updated.")
        db.commit()
        return redirect(url_for("gritting.settings"))

    return render_template(
        "gritting/settings.html",
        email_settings=fetch_email_settings(),
        special_dates=db.execute("SELECT id, event_date, name FROM special_dates WHERE event_date >= ? ORDER BY event_date, name", (datetime.now().date().isoformat(),)).fetchall(),
        smtp_settings=fetch_smtp_settings(),
        report_settings=fetch_report_settings(),
        report_state=fetch_report_state(),
        season=current_gritting_season(),
        season_state=fetch_season_state(),
        monthly_day_options=monthly_day_options(),
        monthly_time_options=schedule_time_options(),
        smtp_ready=smtp_configured(),
    )


@bp.route("/email-signature-preview")
def email_signature_preview() -> str:
    signature = fetch_email_signature()
    signature_html = render_email_template(
        signature["email_signature_html"],
        {"company_name": COMPANY_NAME},
    ).replace("cid:ngs-logo", url_for("static", filename="gritting/brand/ngs-logo.png"))
    signature_text = render_email_template(
        signature["email_signature_text"],
        {"company_name": COMPANY_NAME},
    )
    return render_template(
        "gritting/email_signature_preview.html",
        signature_html=signature_html,
        signature_text=signature_text,
    )


@bp.route("/sites", methods=["GET", "POST"])
def sites() -> str:
    if request.method == "POST":
        return customers()
    return redirect(url_for("gritting.customers"))


@bp.route("/callouts")
def callouts() -> str:
    return render_template(
        "gritting/callouts.html",
        call_outs=(rows := fetch_call_out_rows()),
        email_problems=email_problems(get_db(), [row["id"] for row in rows]),
        dispatch_batches=fetch_dispatch_batches(),
    )


@bp.post("/callouts/bulk")
def bulk_callouts() -> Any:
    db = get_db()
    called_out_at = request.form.get("called_out_at") or now_iso()
    weather_conditions = request.form["weather_conditions"].strip()
    planned_attendance = planned_attendance_from_form(request.form)
    service_notes = request.form["service_notes"].strip()
    subject_template = ""
    if request.form.get("subject_override_enabled") == "1":
        subject_template = request.form.get("subject", "").strip()
    message_template = ""
    if request.form.get("message_override_enabled") == "1":
        message_template = request.form.get("message", "").strip()

    eligible_sites, excluded_sites = fetch_bulk_dispatch_candidates(called_out_at)
    selected_site_ids = {int(site_id) for site_id in request.form.getlist("selected_site_ids") if site_id.isdigit()}
    if not eligible_sites:
        flash("No eligible customers were found for that bulk call-out.")
        return redirect(url_for("gritting.new_callout", called_out_at=called_out_at))
    if not selected_site_ids:
        flash("Tick at least one site for the call-out.")
        return redirect(url_for("gritting.new_callout", called_out_at=called_out_at))

    selected_sites = [site for site in eligible_sites if site["id"] in selected_site_ids]
    if not selected_sites:
        flash("None of the ticked sites were eligible for that call-out.")
        return redirect(url_for("gritting.new_callout", called_out_at=called_out_at))

    batch_key = uuid.uuid4().hex[:10]
    batch_subject_for_log = build_bulk_email_subject(
        selected_sites[0],
        called_out_at,
        weather_conditions,
        planned_attendance,
        service_notes,
        subject_template=subject_template or None,
    )
    batch_cursor = db.execute(
        """
        INSERT INTO dispatch_batches (
            batch_key,
            called_out_at,
            weather_conditions,
            planned_attendance,
            service_notes,
            email_subject,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            batch_key,
            called_out_at,
            weather_conditions,
            planned_attendance,
            service_notes,
            batch_subject_for_log,
            now_iso(),
        ),
    )
    dispatch_batch_id = int(batch_cursor.lastrowid)

    status_counts: dict[str, int] = {}
    first_call_out_id: int | None = None
    for site in selected_sites:
        site_pricing_options = load_pricing_options(site["pricing_options"])
        pricing_label, pricing_amount = selected_pricing_details(
            site_pricing_options,
            request.form.get(f"pricing_label_{site['id']}"),
        )
        call_out_id = create_call_out(
            site_id=site["id"],
            dispatch_batch_id=dispatch_batch_id,
            called_out_at=called_out_at,
            weather_conditions=weather_conditions,
            planned_attendance=planned_attendance,
            service_notes=service_notes,
            pricing_label=pricing_label,
            pricing_amount=pricing_amount,
        )
        if first_call_out_id is None:
            first_call_out_id = call_out_id
        delivery_status, _ = log_notification(
            kind="callout",
            call_out_id=call_out_id,
            channel="Email",
            recipient=site["email"],
            subject=build_bulk_email_subject(
                site,
                called_out_at,
                weather_conditions,
                planned_attendance,
                service_notes,
                subject_template=subject_template or None,
            ),
            message=build_bulk_email_message(
                site,
                called_out_at=called_out_at,
                weather_conditions=weather_conditions,
                planned_attendance=planned_attendance,
                service_notes=service_notes,
                message_template=message_template or None,
            ),
        )
        status_counts[delivery_status] = status_counts.get(delivery_status, 0) + 1

    db.commit()

    flash(
        "Bulk call-out created. "
        f"{len(selected_sites)} customer sites included, "
        f"{max(len(eligible_sites) - len(selected_sites), 0)} unticked, "
        f"{len(excluded_sites)} excluded by site day settings, "
        f"{status_counts.get('sent', 0)} emails sent."
    )
    not_sent = len(selected_sites) - status_counts.get("sent", 0)
    if not_sent:
        flash(
            f"{not_sent} {'email was' if not_sent == 1 else 'emails were'} NOT sent: "
            f"{status_counts.get('logged_only', 0)} only logged (mail server not set up), "
            f"{status_counts.get('failed', 0)} failed, {status_counts.get('skipped_no_email', 0)} with no email address. "
            "They are marked on the list so you can sort them and resend."
        )
    if first_call_out_id is not None:
        # the new night's list: cancel anyone who doesn't need us, then finalise it
        return redirect(url_for("gritting.live_jobs"))
    return redirect(url_for("gritting.new_callout"))


def on_draft_list(call_out: Any) -> bool:
    """True while a job sits on a call-out list that has not been finalised: it can be opened or cancelled, not completed."""
    batch_id = call_out["dispatch_batch_id"]
    if batch_id is None:
        return False
    batch = get_db().execute("SELECT finalised_at FROM dispatch_batches WHERE id = ?", (batch_id,)).fetchone()
    return batch is not None and not batch["finalised_at"]


def on_finalised_list(call_out: Any) -> bool:
    """True once the list the job is on has been finalised: the list to grit from is fixed."""
    batch_id = call_out["dispatch_batch_id"]
    if batch_id is None:
        return False
    batch = get_db().execute("SELECT finalised_at FROM dispatch_batches WHERE id = ?", (batch_id,)).fetchone()
    return batch is not None and bool(batch["finalised_at"])


DRAFT_LIST_MESSAGE = "This job is on a draft list. Finalise the list first, then complete jobs from it."


@bp.route("/callouts/<int:call_out_id>")
def callout_detail(call_out_id: int) -> str:
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))

    notifications = db.execute(
        """
        SELECT *
        FROM notification_logs
        WHERE call_out_id = ?
        ORDER BY sent_at DESC
        """,
        (call_out_id,),
    ).fetchall()
    pricing_options = load_pricing_options(call_out["pricing_options"])
    suggested_label = call_out["pricing_label"] or suggested_pricing_label(call_out["called_out_at"], pricing_options)
    return render_template(
        "gritting/callout_detail.html",
        call_out=call_out,
        notifications=notifications,
        unsent_count=len(unsent_emails(db, [call_out_id], live_only=False)),
        email_problem=email_problems(db, [call_out_id]).get(call_out_id, ""),
        smtp_ready=smtp_configured(),
        now_value=now_iso(),
        pricing_options=pricing_options,
        suggested_pricing_label=suggested_label,
        draft_list=on_draft_list(call_out),
        final_list=on_finalised_list(call_out),
        extras_items=load_extras_items(call_out["extras_json"]),
    )


@bp.post("/callouts/<int:call_out_id>/complete")
def complete_callout(call_out_id: int):
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "completed":
        flash("This job is already completed.")
        return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))
    if call_out["status"] == "cancelled":
        flash("This job was cancelled. Reopen it first if we do need to attend.")
        return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))
    if call_out["status"] == "no_access":
        flash("This job was marked no access. Put it back on the list first if we did get in.")
        return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))
    if on_draft_list(call_out):
        flash(DRAFT_LIST_MESSAGE)
        return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))

    pricing_options = load_pricing_options(call_out["pricing_options"] if call_out else None)
    selected_label = request.form.get("pricing_label", "").strip()
    selected_amount = None
    extras_total, extras_items = float(call_out["extras_amount"] or 0), load_extras_items(call_out["extras_json"])
    if selected_label:
        matched = next((option for option in pricing_options if option["label"] == selected_label), None)
        if matched:
            # a price chosen now is the site price: add the extras in force on the night
            extras_total, extras_items = extras_on(db, call_out["site_id"], (call_out["called_out_at"] or "")[:10])
            selected_amount = round(matched["amount"] + extras_total, 2)
    elif call_out["pricing_label"] and call_out["pricing_amount"] is not None:
        selected_label = call_out["pricing_label"]
        selected_amount = call_out["pricing_amount"]

    _, delivery_status, _ = complete_call_out_record(
        call_out_id=call_out_id,
        completed_at=request.form.get("completed_at") or now_iso(),
        service_notes=request.form["service_notes"].strip(),
        pricing_label=selected_label or None,
        pricing_amount=selected_amount,
        extras_amount=extras_total,
        extras_json=json.dumps(extras_items),
    )
    db.commit()
    flash_completion_message(delivery_status)
    return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))


@bp.post("/callouts/<int:call_out_id>/cancel")
def cancel_callout(call_out_id: int):
    """The customer does not need us to attend: take the job off the live list and out of the totals."""
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "completed":
        flash("A completed job cannot be cancelled.")
    elif call_out["status"] == "cancelled":
        flash("This job is already cancelled.")
    elif call_out["status"] == "no_access":
        flash("This job was marked no access. Put it back on the list first if you want to cancel it.")
    else:
        db.execute(
            "UPDATE call_outs SET status = 'cancelled', cancelled_at = ?, pricing_label = NULL, pricing_amount = NULL, job_number = NULL WHERE id = ?",
            (now_iso(), call_out_id),
        )
        db.commit()
        flash(f"Cancelled {call_out['customer_name']} · {call_out['site_name']}. It is no longer a live job and is not counted in the totals."
              + (" This list was already finalised, so print the final list again." if call_out["dispatch_batch_id"] and db.execute("SELECT finalised_at FROM dispatch_batches WHERE id = ?", (call_out["dispatch_batch_id"],)).fetchone()["finalised_at"] else ""))
    if request.form.get("next", "").strip() == "live_jobs":
        return redirect(url_for("gritting.live_jobs"))
    return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))


@bp.post("/callouts/<int:call_out_id>/no-access")
def no_access_callout(call_out_id: int):
    """We attended but could not get in: email the customer, and take the job off the live list without completing it."""
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    back = redirect(url_for("gritting.live_jobs") if request.form.get("next", "").strip() == "live_jobs" else url_for("gritting.callout_detail", call_out_id=call_out_id))
    if call_out["status"] != "called_out":
        flash("Only a job still to do can be marked as no access.")
        return back
    if on_draft_list(call_out):
        flash("Finalise the list first. Until then a job can only be opened or marked not needed.")
        return back
    note = request.form.get("access_note", "").strip()[:500]
    db.execute("UPDATE call_outs SET status = 'no_access', no_access_at = ?, access_note = ? WHERE id = ?", (now_iso(), note, call_out_id))
    subject, message = build_no_access_email(call_out, note)
    delivery_status, _ = log_notification(call_out_id=call_out_id, channel="Email", recipient=call_out["email"], subject=subject, message=message, kind="no_access")
    db.commit()
    label = f"{call_out['customer_name']} · {call_out['site_name']}"
    if delivery_status == "sent":
        flash(f"{label} marked as no access and the customer has been emailed. It is off the live list and not counted as done.")
    else:
        flash(f"{label} marked as no access and off the live list, but the email was NOT sent ({delivery_status.replace('_', ' ')}). Open the job to resend it.")
    return back


@bp.post("/callouts/<int:call_out_id>/reopen")
def reopen_callout(call_out_id: int):
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "no_access":
        db.execute("UPDATE call_outs SET status = 'called_out', no_access_at = NULL, access_note = NULL WHERE id = ?", (call_out_id,))
        db.commit()
        flash("Job put back on the list.")
    elif call_out["status"] == "cancelled":
        db.execute("UPDATE call_outs SET status = 'called_out', cancelled_at = NULL WHERE id = ?", (call_out_id,))
        if not on_draft_list(call_out):
            assign_job_numbers(db, [call_out_id])  # back on a live list: it gets a number again
        db.commit()
        flash("Job reopened. Choose the pricing when you complete it.")
    if request.form.get("next", "").strip() == "live_jobs":
        return redirect(url_for("gritting.live_jobs"))
    return redirect(url_for("gritting.callout_detail", call_out_id=call_out_id))


@bp.post("/callouts/<int:call_out_id>/complete-quick")
def quick_complete_callout(call_out_id: int):
    db = get_db()
    call_out = fetch_call_out(call_out_id)
    if call_out is None:
        flash("Call-out not found.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "completed":
        flash("This job is already completed.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "cancelled":
        flash("This job was cancelled. Reopen it first if we do need to attend.")
        return redirect(url_for("gritting.callouts"))
    if call_out["status"] == "no_access":
        flash("This job was marked no access. Put it back on the list first if we did get in.")
        return redirect(url_for("gritting.callouts"))
    if on_draft_list(call_out):
        flash(DRAFT_LIST_MESSAGE)
        return redirect(url_for("gritting.live_jobs") if request.form.get("next", "").strip() == "live_jobs" else url_for("gritting.callouts"))

    pricing_options = load_pricing_options(call_out["pricing_options"])
    selected_label = call_out["pricing_label"] or suggested_pricing_label(call_out["called_out_at"], pricing_options)
    selected_amount = call_out["pricing_amount"]
    extras_total, extras_items = float(call_out["extras_amount"] or 0), load_extras_items(call_out["extras_json"])
    if selected_label and selected_amount is None:
        matched = next((option for option in pricing_options if option["label"] == selected_label), None)
        if matched:
            extras_total, extras_items = extras_on(db, call_out["site_id"], (call_out["called_out_at"] or "")[:10])
            selected_amount = round(matched["amount"] + extras_total, 2)

    _, delivery_status, _ = complete_call_out_record(
        call_out_id=call_out_id,
        completed_at=now_iso(),
        service_notes=(call_out["service_notes"] or "").strip(),
        pricing_label=selected_label or None,
        pricing_amount=selected_amount,
        extras_amount=extras_total,
        extras_json=json.dumps(extras_items),
    )
    db.commit()
    flash_completion_message(delivery_status)
    next_page = request.form.get("next", "").strip()
    if next_page == "live_jobs":
        return redirect(url_for("gritting.live_jobs"))
    return redirect(url_for("gritting.callouts"))


def fetch_monthly_rows(month_value: str) -> list[sqlite3.Row]:
    return get_db().execute(
        """
        SELECT
            call_outs.id,
            call_outs.job_number,
            call_outs.called_out_at,
            call_outs.completed_at,
            call_outs.status,
            call_outs.weather_conditions,
            call_outs.pricing_label,
            call_outs.pricing_amount,
            call_outs.extras_json,
            call_outs.service_notes,
            sites.name AS site_name,
            customers.name AS customer_name,
            COUNT(DISTINCT notification_logs.id) AS notifications_sent
        FROM call_outs
        JOIN sites ON sites.id = call_outs.site_id
        JOIN customers ON customers.id = sites.customer_id
        LEFT JOIN notification_logs ON notification_logs.call_out_id = call_outs.id
        WHERE substr(call_outs.called_out_at, 1, 7) = ?
        GROUP BY call_outs.id
        ORDER BY call_outs.called_out_at DESC
        """,
        (month_value,),
    ).fetchall()


def monthly_summary_export_header() -> list[str]:
    return [
        "Job number",
        "Customer",
        "Site",
        "Called out",
        "Status",
        "Pricing label",
        "Pricing amount ex VAT (£)",
        "VAT at 20% (£)",
        "Pricing total incl VAT (£)",
        "Notes",
    ]


def monthly_summary_export_rows(rows: list[sqlite3.Row]) -> list[list[Any]]:
    export_rows: list[list[Any]] = []
    sorted_rows = sorted(
        rows,
        key=lambda row: (
            str(row["customer_name"]).lower(),
            str(row["site_name"]).lower(),
            str(row["called_out_at"]),
        ),
    )
    previous_customer = None
    customer_net_total = 0.0

    def append_customer_total(customer_name: str | None, net_total: float) -> None:
        if not customer_name:
            return
        export_rows.append(
            [
                "",
                "",
                "",
                "",
                "",
                "Total",
                round(net_total, 2),
                vat_amount(net_total) or 0.0,
                gross_amount(net_total) or 0.0,
                "",
            ]
        )

    for row in sorted_rows:
        current_customer = row["customer_name"]
        if previous_customer is not None and current_customer != previous_customer:
            append_customer_total(previous_customer, customer_net_total)
            export_rows.append([""] * len(monthly_summary_export_header()))
            customer_net_total = 0.0

        pricing_amount = float(row["pricing_amount"]) if row["pricing_amount"] is not None else 0.0
        export_rows.append(
            [
                row["job_number"] or "",
                row["customer_name"],
                row["site_name"],
                format_date_from_iso(row["called_out_at"]),
                row["status"],
                ((row["pricing_label"] or "") + (f" + {extras_summary(row['extras_json'])}" if extras_summary(row["extras_json"]) else "")),
                row["pricing_amount"] if row["pricing_amount"] is not None else "",
                vat_amount(row["pricing_amount"]) if row["pricing_amount"] is not None else "",
                gross_amount(row["pricing_amount"]) if row["pricing_amount"] is not None else "",
                (row["service_notes"] or "").replace("\n", " "),
            ]
        )
        previous_customer = current_customer
        customer_net_total += pricing_amount

    append_customer_total(previous_customer, customer_net_total)
    return export_rows


def excel_column_name(index: int) -> str:
    result = ""
    current = index + 1
    while current:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def workbook_xml_escape(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=False)


def workbook_column_widths(source_rows: list[list[Any]], fallback_rows: list[list[Any]] | None = None) -> list[float]:
    rows_for_width = source_rows or (fallback_rows or [])
    if not rows_for_width:
        return []

    column_count = max(len(row) for row in rows_for_width)
    widths: list[float] = []
    for column_index in range(column_count):
        longest = 0
        for row in rows_for_width:
            value = row[column_index] if column_index < len(row) else ""
            lines = str("" if value is None else value).splitlines() or [""]
            longest = max(longest, max(len(line) for line in lines))
        widths.append(float(min(max(longest + 2, 6), 60)))
    return widths


def build_simple_xlsx(
    sheet_name: str,
    header: list[str],
    rows: list[list[Any]],
    top_rows: list[list[Any]] | None = None,
    width_source_rows: list[list[Any]] | None = None,
) -> bytes:
    top_row_values = top_rows or []
    all_rows = [*top_row_values, header, *rows]
    column_widths = workbook_column_widths(width_source_rows or rows, all_rows)
    worksheet_rows: list[str] = []
    header_row_index = len(top_row_values) + 1
    for row_index, row in enumerate(all_rows, start=1):
        cells: list[str] = []
        if row_index == 1 and top_row_values:
            style_index = 1
        elif row_index == 2 and len(top_row_values) > 1:
            style_index = 2
        elif row_index == header_row_index:
            style_index = 3
        else:
            style_index = 0
        style_attribute = f' s="{style_index}"' if style_index else ""
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{excel_column_name(column_index - 1)}{row_index}"
            if isinstance(value, (int, float)) and value != "":
                cells.append(f'<c r="{cell_ref}"{style_attribute}><v>{value}</v></c>')
            else:
                text = workbook_xml_escape(value)
                cells.append(
                    f'<c r="{cell_ref}" t="inlineStr"{style_attribute}><is><t>{text}</t></is></c>'
                )
        worksheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    column_xml = ""
    if column_widths:
        column_parts = []
        for index, width in enumerate(column_widths, start=1):
            column_parts.append(
                f'<col min="{index}" max="{index}" width="{width:.2f}" customWidth="1"/>'
            )
        column_xml = f'<cols>{"".join(column_parts)}</cols>'

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'{column_xml}<sheetData>{"".join(worksheet_rows)}</sheetData>'
        "</worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{workbook_xml_escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="4">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><name val="Calibri"/><color rgb="FF11275D"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FF3156B4"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>'
        '</fonts>'
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF4FF"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF3156B4"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf xfId="0" fontId="0" fillId="0" borderId="0"/>'
        '<xf xfId="0" fontId="1" fillId="0" borderId="0" applyFont="1"><alignment horizontal="left" vertical="center"/></xf>'
        '<xf xfId="0" fontId="2" fillId="2" borderId="0" applyFont="1" applyFill="1"><alignment horizontal="left" vertical="center"/></xf>'
        '<xf xfId="0" fontId="3" fillId="3" borderId="0" applyFont="1" applyFill="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:creator>{workbook_xml_escape(COMPANY_NAME)}</dc:creator>'
        f'<cp:lastModifiedBy>{workbook_xml_escape(COMPANY_NAME)}</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{datetime.utcnow().isoformat()}Z</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{datetime.utcnow().isoformat()}Z</dcterms:modified>'
        "</cp:coreProperties>"
    )
    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Norwich Gritting Services</Application>'
        "</Properties>"
    )

    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml)
        workbook.writestr("_rels/.rels", root_rels_xml)
        workbook.writestr("docProps/core.xml", core_xml)
        workbook.writestr("docProps/app.xml", app_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        workbook.writestr("xl/styles.xml", styles_xml)
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
    return buffer.getvalue()


def build_monthly_summary_xlsx(month_value: str, rows: list[sqlite3.Row]) -> bytes:
    export_rows = monthly_summary_export_rows(rows)
    return build_simple_xlsx(
        sheet_name=f"{month_value} Summary",
        top_rows=[
            [COMPANY_NAME],
            [f"Summary month: {format_month_label(month_value)}"],
            [""],
        ],
        header=monthly_summary_export_header(),
        rows=export_rows,
        width_source_rows=export_rows,
    )


def season_archive_filename(season: dict[str, Any], extension: str) -> str:
    label = season["label"].lower().replace("/", "-").replace(" ", "-")
    return f"norwich-gritting-services-{label}-archive.{extension}"


def build_season_archive_xlsx(season: dict[str, Any], rows: list[sqlite3.Row]) -> bytes:
    export_rows = monthly_summary_export_rows(rows)
    return build_simple_xlsx(
        sheet_name=f"{season['start_year']}-{season['end_year']} Season",
        top_rows=[
            [COMPANY_NAME],
            [f"Season archive: {season['label']}"],
            [f"Season period: {season['date_range_label']}"],
            [""],
        ],
        header=monthly_summary_export_header(),
        rows=export_rows,
        width_source_rows=export_rows,
    )


def monthly_summary_filename(month_value: str, extension: str) -> str:
    return f"norwich-gritting-services-summary-{month_value}.{extension}"


def send_monthly_summary_email(month_value: str, recipients: list[str]) -> tuple[str, str]:
    rows = fetch_monthly_rows(month_value)
    month_label = format_month_label(month_value)
    xlsx_bytes = build_monthly_summary_xlsx(month_value, rows)
    return send_email_message(
        recipients,
        f"{COMPANY_NAME} monthly summary - {month_label}",
        f"Please find attached the {COMPANY_NAME} monthly summary for {month_label}.",
        attachments=[
            {
                "filename": monthly_summary_filename(month_value, "xlsx"),
                "maintype": "application",
                "subtype": "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "data": xlsx_bytes,
            }
        ],
    )


def auto_monthly_summary_if_due() -> None:
    now = datetime.now()
    if _flask_app is None:
        return
    ensure_initialised()

    with _flask_app.app_context():
        report_settings = fetch_report_settings()
        if not parse_bool(report_settings.get("monthly_summary_auto_send", ""), default=False):
            return

        try:
            configured_day = int(report_settings.get("monthly_summary_auto_day", "1"))
        except ValueError:
            configured_day = 1
        configured_day = min(max(configured_day, 1), 31)
        last_day_of_month = monthrange(now.year, now.month)[1]
        scheduled_day = min(configured_day, last_day_of_month)

        scheduled_time = report_settings.get("monthly_summary_auto_time", "06:00")
        if scheduled_time not in schedule_time_options():
            scheduled_time = "06:00"
        scheduled_hour, scheduled_minute = [int(part) for part in scheduled_time.split(":", 1)]

        if now.day != scheduled_day:
            return
        if (now.hour, now.minute) < (scheduled_hour, scheduled_minute):
            return

        month_value = previous_month_value(now)
        report_state = fetch_report_state()
        if report_state.get("monthly_summary_last_attempt_month", "") == month_value:
            return

        db = get_db()
        recipients = parse_recipient_list(report_settings.get("monthly_summary_recipients", ""))
        month_label = format_month_label(month_value)
        status_message = ""
        delivery_status = "skipped_no_email"
        delivery_details = ""

        if not recipients:
            status_message = f"Skipped automatic send for {month_label}: no recipients are set."
        else:
            delivery_status, delivery_details = send_monthly_summary_email(month_value, recipients)
            if delivery_status == "sent":
                status_message = f"Sent {month_label} automatically to {len(recipients)} recipient(s)."
            elif delivery_status == "logged_only":
                status_message = f"Automatic send for {month_label} was logged only because SMTP is not configured."
            elif delivery_status == "skipped_no_email":
                status_message = f"Skipped automatic send for {month_label}: no recipients are set."
            else:
                status_message = f"Automatic send for {month_label} failed: {delivery_details}"

        state_updates = {
            "monthly_summary_last_attempt_month": month_value,
            "monthly_summary_last_attempt_at": now.replace(second=0, microsecond=0).isoformat(timespec="minutes"),
            "monthly_summary_last_status": status_message,
        }
        if delivery_status == "sent":
            state_updates["monthly_summary_last_sent_month"] = month_value
        save_report_state(db, state_updates)
        db.commit()


def auto_monthly_summary_loop() -> None:
    while True:
        try:
            auto_monthly_summary_if_due()
        except Exception:
            # Keep the scheduler alive even if a single attempt fails.
            pass
        time.sleep(AUTO_MONTHLY_SUMMARY_CHECK_INTERVAL_SECONDS)


def start_background_jobs() -> None:
    global _auto_monthly_summary_thread_started
    with _auto_monthly_summary_thread_lock:
        if _auto_monthly_summary_thread_started:
            return
        thread = threading.Thread(target=auto_monthly_summary_loop, name="monthly-summary-scheduler", daemon=True)
        thread.start()
        _auto_monthly_summary_thread_started = True


def start_new_gritting_season_rollover() -> tuple[dict[str, Any], Path]:
    db = get_db()
    season = current_gritting_season()
    season_rows = fetch_season_archive_rows(season["start_date"], season["end_date"])
    archive_bytes = build_season_archive_xlsx(season, season_rows)

    SEASON_ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    archive_path = SEASON_ARCHIVE_DIR / season_archive_filename(season, "xlsx")
    if archive_path.exists():
        archive_path = SEASON_ARCHIVE_DIR / (
            archive_path.stem + "-" + datetime.now().strftime("%Y%m%d-%H%M") + archive_path.suffix
        )
    archive_path.write_bytes(archive_bytes)

    next_start_year = season["start_year"] + 1
    save_season_settings(db, {"active_season_start_year": str(next_start_year)})
    save_season_state(
        db,
        {
            "last_archived_season_label": season["label"],
            "last_archived_season_path": str(archive_path.relative_to(DATA_ROOT)),
            "last_archived_at": now_iso(),
        },
    )
    db.commit()
    return season, archive_path


@bp.route("/monthly-summary")
def monthly_summary() -> str:
    month_value = request.args.get("month") or datetime.now().strftime("%Y-%m")
    rows = fetch_monthly_rows(month_value)
    total_revenue = sum(row["pricing_amount"] or 0 for row in rows)
    total_vat = round(total_revenue * VAT_RATE, 2)
    total_revenue_inc_vat = round(total_revenue + total_vat, 2)
    completed_jobs = sum(1 for row in rows if row["status"] == "completed")
    return render_template(
        "gritting/monthly_summary.html",
        month_value=month_value,
        rows=rows,
        total_revenue=total_revenue,
        total_vat=total_vat,
        total_revenue_inc_vat=total_revenue_inc_vat,
        completed_jobs=completed_jobs,
        report_settings=fetch_report_settings(),
        smtp_ready=smtp_configured(),
    )


@bp.post("/monthly-summary/email")
def email_monthly_summary() -> Any:
    month_value = request.form.get("month") or datetime.now().strftime("%Y-%m")
    raw_recipients = request.form.get("monthly_summary_recipients", "").strip()
    if not raw_recipients:
        raw_recipients = fetch_report_settings()["monthly_summary_recipients"]
    recipients = parse_recipient_list(raw_recipients)
    if not recipients:
        flash("Add at least one email address to the monthly summary list in Settings.")
        return redirect(url_for("gritting.monthly_summary", month=month_value))

    db = get_db()
    if request.form.get("monthly_summary_recipients") is not None:
        save_report_settings(db, request.form)
        db.commit()

    delivery_status, delivery_details = send_monthly_summary_email(month_value, recipients)
    if delivery_status == "sent":
        flash(f"Monthly .xlsx summary emailed to {len(recipients)} recipient(s).")
    elif delivery_status == "logged_only":
        flash("Monthly summary email is ready, but SMTP is not configured yet.")
    elif delivery_status == "skipped_no_email":
        flash("No monthly summary recipients are set.")
    else:
        flash(f"Monthly summary email failed: {delivery_details}")
    return redirect(url_for("gritting.monthly_summary", month=month_value))


@bp.route("/monthly-summary/export")
def export_monthly_summary() -> Response:
    month_value = request.args.get("month") or datetime.now().strftime("%Y-%m")
    rows = fetch_monthly_rows(month_value)
    export_rows = monthly_summary_export_rows(rows)
    header = monthly_summary_export_header()

    def generate() -> Any:
        header_buffer = StringIO()
        csv.writer(header_buffer).writerow(header)
        yield header_buffer.getvalue()
        for single_row in export_rows:
            row_buffer = StringIO()
            csv.writer(row_buffer).writerow(single_row)
            yield row_buffer.getvalue()

    filename = monthly_summary_filename(month_value, "csv")
    return Response(
        generate(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


def ensure_initialised() -> None:
    """Create the database (and read the bundled customer workbook) the first time it is needed."""
    global _initialised_database
    with _initialise_lock:
        if _initialised_database == DATABASE_PATH:
            return
        DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
        init_db()
        _initialised_database = DATABASE_PATH


@bp.before_request
def prepare_database() -> None:
    ensure_initialised()


def register(flask_app: Any, data_root: Any = None) -> None:
    """Add the gritting pages to the muck app under /gritting."""
    global _flask_app
    if data_root is not None:
        configure(data_root)
    _flask_app = flask_app
    flask_app.jinja_env.filters.update(TEMPLATE_FILTERS)
    flask_app.teardown_appcontext(close_db)
    flask_app.register_blueprint(bp)
