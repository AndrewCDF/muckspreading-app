from flask import Flask, Response, jsonify, redirect, render_template, render_template_string, request, session, url_for
import csv
import io
import hmac
import json
import mimetypes
import os
import re
import secrets
import shutil
import smtplib
import sqlite3
import struct
import subprocess
import tempfile
import threading
import time
import sys
import zipfile
import xml.etree.ElementTree as ET
import zlib
from datetime import datetime, timedelta
from email.message import EmailMessage
from xml.sax.saxutils import escape as xml_escape
from email.utils import parseaddr
from settings_workbook import Workbook, active as settings_row_active


app = Flask(__name__)
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
app.secret_key = os.environ.get("MUCKSPREADING_SESSION_SECRET") or secrets.token_hex(32)
SETTINGS_WORKBOOK_PATH = os.path.join(APP_ROOT, "settings.xlsx")
TIMESHEET_COMPANIES = ["Cherry Dene Farm Ltd.", "A. Farrell Contracting Ltd."]
DATA_DIR = os.path.join(APP_ROOT, "data")
JOBS_PATH = os.path.join(DATA_DIR, "jobs.ndjson")
CUSTOMERS_PATH = os.path.join(DATA_DIR, "customers.json")
FARMS_PATH = os.path.join(DATA_DIR, "farms.json")
MUCK_TYPES_PATH = os.path.join(DATA_DIR, "muck_types.json")
FIELD_MAP_PATH = os.path.join(DATA_DIR, "customer_fields.json")
EMAIL_CONFIG_PATH = os.path.join(DATA_DIR, "email_config.json")
EMAIL_STATE_PATH = os.path.join(DATA_DIR, "weekly_email_state.json")
APP_SETTINGS_PATH = os.path.join(DATA_DIR, "app_settings.json")
INVOICE_LEDGER_PATH = os.path.join(DATA_DIR, "invoice_ledger.json")
INVOICE_STATE_PATH = os.path.join(DATA_DIR, "invoice_state.json")
TIMESHEET_MONTH_STATE_PATH = os.path.join(DATA_DIR, "timesheet_month_state.json")
STRAW_STATE_PATH = os.path.join(DATA_DIR, "straw-records.json")
INVOICE_ARCHIVE_DIR = os.path.join(DATA_DIR, "invoices")
ISSUE_PHOTOS_DIR = os.path.join(DATA_DIR, "job_issue_photos")
CUSTOMER_MASTER_XLSX_PATH = os.path.join(APP_ROOT, "customer_master.xlsx")
EMAIL_SETTINGS_CSV_PATH = os.path.join(APP_ROOT, "email_settings.csv")
WEEKLY_SUMMARY_TEMPLATE_PATH = os.path.join(APP_ROOT, "weekly_summary_layout_template.xlsx")
INVOICE_TEMPLATE_CANDIDATES = [
    os.path.join(APP_ROOT, "Invoice_template.xlsx"),
    os.path.join(APP_ROOT, "invoice_layout_template.xlsx"),
]
EMAIL_CHECK_INTERVAL_SECONDS = 300
XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRICT_XLSX_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main"
STRICT_REL_NS = "http://purl.oclc.org/ooxml/officeDocument/relationships"
CUSTOMER_MASTER_HEADERS = [
    "customer_name",
    "farm_name",
    "email",
    "email_2",
    "address_line_1",
    "address_line_2",
    "town",
    "postcode",
    "rate_per_ton",
    "vat_rate",
    "active",
    "muck_type",
]
ALLOWED_ISSUE_PHOTO_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".heic", ".heif"}

DEFAULT_EMAIL_CONFIG = {
    "enabled": False,
    "smtp_host": "",
    "smtp_port": 587,
    "use_tls": True,
    "smtp_username": "",
    "smtp_password": "",
    "from_email": "",
    "to_emails": [],
    "send_weekday": 0,
    "send_hour": 5,
    "send_minute": 0,
    "monthly_enabled": True,
    "monthly_send_hour": 5,
    "monthly_send_minute": 0,
    "subject_prefix": "A. Farrell Contracting",
}
INVOICE_PAYMENT_TERMS_OPTIONS = ["7", "14", "21", "30"]

APP_SHORT_NAME = "Muck Jobs"
APP_THEME_COLOR = "#334d38"
DEFAULT_INVOICE_SUBJECT_TEMPLATE = "Invoice {invoice_number} - {customer}"
DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE = """{greeting},

Please find attached an invoice for recent spreading work.

Many Thanks

Andrew Farrell
A. Farrell Contracting Ltd.
07952683364"""
DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE = """{greeting},

Please find attached the invoice PDF and workbook for recent spreading work.

Many Thanks

Andrew Farrell
A. Farrell Contracting Ltd.
07952683364"""
DEFAULT_APP_SETTINGS = {
    "invoice_from_email": "andrew@afarrellcontracting.co.uk",
    "invoice_subject_template": DEFAULT_INVOICE_SUBJECT_TEMPLATE,
    "invoice_customer_message_template": DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE,
    "invoice_accounts_message_template": DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE,
    "invoice_default_payment_terms_days": "14",
}

DASHBOARD_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="#24442f">
  <title>A. Farrell Contracting</title>
  <link rel="icon" type="image/png" href="{{ url_for('app_icon_png', size=180) }}">
  <link rel="apple-touch-icon" sizes="180x180" href="{{ url_for('app_icon_png', size=180) }}">
    <link rel="stylesheet" href="{{ url_for('static', filename='af_brand.css') }}">
  <link rel="stylesheet" href="{{ url_for('static', filename='combined.css') }}">
</head>
<body>
  <div class="shell">
    <header class="topbar">
      <div class="brand-block">
        <p class="eyebrow">Operations</p>
        <h1>A. Farrell Contracting</h1>
      </div>
    </header>

    <main class="layout">
      <section class="branch-grid">
                <a class="primary" href="/muck">Muck Spreading</a>
                <a class="primary straw-button" href="/straw">Hay &amp; Straw</a>
                <a class="primary" href="{{ url_for('timesheet_home') }}">Timesheet</a>
                <a class="primary maintenance-button" href="{{ url_for('maintenance_home') }}">Maintenance</a>
      </section>
    </main>
  </div>
    <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
        <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
        <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
        <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
    </nav>
  <script src="{{ url_for('static', filename='combined.js') }}"></script>
</body>
</html>
"""


def discover_custom_app_icon_path():
    preferred_path = os.path.join(APP_ROOT, "static", "muckspreading-app-icon.png")
    if os.path.exists(preferred_path):
        return preferred_path
    try:
        names = sorted(os.listdir(APP_ROOT))
    except OSError:
        return ""
    for name in names:
        if name.lower().endswith(".png"):
            return os.path.join(APP_ROOT, name)
    return ""


CUSTOM_APP_ICON_PATH = discover_custom_app_icon_path()

ICON = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
<rect width="64" height="64" rx="12" fill="#334d38"/>
<rect x="5" y="5" width="54" height="54" rx="10" fill="none" stroke="#d7bf7a" stroke-width="2.5"/>
<path d="M14 42c4-9 9-14 15-17 5-2 10-3 14-2 2 0 5 1 7 2-2 3-4 6-6 9-3 3-7 5-12 6-5 2-11 2-18 2z" fill="#d7bf7a"/>
</svg>"""


def _png_chunk(tag, data):
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def _rounded_rect_contains(x, y, left, top, width, height, radius):
    right = left + width
    bottom = top + height
    if left + radius <= x <= right - radius or top + radius <= y <= bottom - radius:
        return True
    corner_centers = [
        (left + radius, top + radius),
        (right - radius, top + radius),
        (left + radius, bottom - radius),
        (right - radius, bottom - radius),
    ]
    radius_sq = radius * radius
    for cx, cy in corner_centers:
        dx = x - cx
        dy = y - cy
        if dx * dx + dy * dy <= radius_sq:
            return True
    return False


def _point_in_polygon(x, y, points):
    inside = False
    j = len(points) - 1
    i = 0
    while i < len(points):
        xi, yi = points[i]
        xj, yj = points[j]
        if ((yi > y) != (yj > y)) and (x < ((xj - xi) * (y - yi) / ((yj - yi) or 1e-9) + xi)):
            inside = not inside
        j = i
        i += 1
    return inside


def build_app_icon_png(size):
    size = max(120, min(int(size), 1024))
    bg = (0x33, 0x4D, 0x38, 255)
    gold = (0xD7, 0xBF, 0x7A, 255)
    transparent = (0, 0, 0, 0)

    outer_radius = size * 0.2
    border_inset = max(6.0, size * 0.08)
    inner_radius = max(outer_radius - (size * 0.04), 4.0)
    border_width = max(3.0, size * 0.025)

    leaf_points = [
        (size * 0.2, size * 0.66),
        (size * 0.28, size * 0.5),
        (size * 0.38, size * 0.39),
        (size * 0.5, size * 0.31),
        (size * 0.63, size * 0.28),
        (size * 0.78, size * 0.34),
        (size * 0.69, size * 0.51),
        (size * 0.58, size * 0.61),
        (size * 0.44, size * 0.67),
        (size * 0.3, size * 0.68),
    ]

    rows = []
    y = 0
    while y < size:
        row = bytearray([0])
        x = 0
        while x < size:
            px = x + 0.5
            py = y + 0.5
            color = transparent
            in_outer = _rounded_rect_contains(px, py, 0.0, 0.0, float(size), float(size), outer_radius)
            if in_outer:
                color = bg
                in_border_box = _rounded_rect_contains(
                    px,
                    py,
                    border_inset,
                    border_inset,
                    float(size) - (border_inset * 2),
                    float(size) - (border_inset * 2),
                    inner_radius,
                )
                in_inner_fill = _rounded_rect_contains(
                    px,
                    py,
                    border_inset + border_width,
                    border_inset + border_width,
                    float(size) - ((border_inset + border_width) * 2),
                    float(size) - ((border_inset + border_width) * 2),
                    max(inner_radius - border_width, 2.0),
                )
                if in_border_box and not in_inner_fill:
                    color = gold
                if _point_in_polygon(px, py, leaf_points):
                    color = gold
            row.extend(color)
            x += 1
        rows.append(bytes(row))
        y += 1

    raw = b"".join(rows)
    compressed = zlib.compress(raw, 9)
    return b"".join([
        b"\x89PNG\r\n\x1a\n",
        _png_chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)),
        _png_chunk(b"IDAT", compressed),
        _png_chunk(b"IEND", b""),
    ])


def app_icon_bytes(size):
    if CUSTOM_APP_ICON_PATH and os.path.exists(CUSTOM_APP_ICON_PATH):
        try:
            with open(CUSTOM_APP_ICON_PATH, "rb") as handle:
                return handle.read()
        except OSError:
            pass
    return build_app_icon_png(size)


def app_icon_dimensions():
    if CUSTOM_APP_ICON_PATH and os.path.exists(CUSTOM_APP_ICON_PATH):
        try:
            with open(CUSTOM_APP_ICON_PATH, "rb") as handle:
                header = handle.read(24)
            if header[:8] == b"\x89PNG\r\n\x1a\n" and header[12:16] == b"IHDR":
                width, height = struct.unpack(">II", header[16:24])
                if width > 0 and width == height:
                    return width
        except OSError:
            pass
    return 180


def resolve_invoice_template_path():
    for path in INVOICE_TEMPLATE_CANDIDATES:
        if os.path.exists(path):
            return path
    return ""



HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="#334d38">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <meta name="apple-mobile-web-app-title" content="Muck Jobs">
  <title>A. Farrell Contracting Muck Spreading Jobs</title>
  <link rel="icon" type="image/png" href="{{ url_for('app_icon_png', size=180) }}">
  <link rel="icon" type="image/svg+xml" href="{{ url_for('favicon') }}">
  <link rel="apple-touch-icon" sizes="180x180" href="{{ url_for('app_icon_png', size=180) }}">
  <link rel="manifest" href="{{ url_for('web_manifest') }}">
  <style>
    @import url('/static/af_brand.css');
    :root {
      --bg: #edf1ea;
      --panel: rgba(255, 255, 255, 0.94);
      --ink: #1d2b1f;
      --muted: #5b6b60;
      --line: rgba(36, 68, 47, 0.12);
      --green: #24442f;
      --gold: #c79e4f;
      --red: #8b4738;
      --shadow: 0 18px 36px rgba(29, 48, 35, 0.08);
      --font-main: Inter, "Segoe UI", system-ui, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font-main);
      color: var(--ink);
      background: linear-gradient(180deg, #edf1ea 0%, #e7ece3 100%);
    }
    input, select, button, table, th, td {
      font-family: var(--font-main);
    }
    .page {
      max-width: 1180px;
      margin: 0 auto;
      padding: max(18px, env(safe-area-inset-top)) max(14px, env(safe-area-inset-right)) max(28px, env(safe-area-inset-bottom)) max(14px, env(safe-area-inset-left));
    }
    .hero {
      display: grid;
      grid-template-columns: minmax(0, 1.3fr) minmax(280px, 0.9fr);
      gap: 18px;
      margin-bottom: 18px;
    }
    .card {
      background: var(--panel);
      border: 1px solid rgba(82, 69, 42, 0.12);
      border-radius: 24px;
      box-shadow: var(--shadow);
      padding: 24px;
    }
    .hero-title-card {
      text-align: left;
      display: grid;
      gap: 16px;
      align-content: start;
      background:
        linear-gradient(145deg, rgba(255,255,255,0.3), rgba(255,255,255,0) 38%),
        linear-gradient(180deg, rgba(60,95,70,0.06), rgba(186,148,80,0.08)),
        var(--panel);
    }
    .hero-copy {
      margin: 0;
      max-width: 40rem;
      color: #5f5d4f;
      font-size: 17px;
      line-height: 1.55;
    }
    .eyebrow {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      width: fit-content;
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(60,95,70,0.1);
      border: 1px solid rgba(60,95,70,0.12);
      color: var(--green);
      font-size: 12px;
      font-weight: bold;
      text-transform: uppercase;
      letter-spacing: 0.09em;
    }
    h1 {
      margin: 14px 0 10px;
      line-height: 1;
    }
    .title-line {
      display: block;
      white-space: nowrap;
    }
    .title-line-primary {
      font-size: clamp(32px, 5.6vw, 54px);
      letter-spacing: -0.035em;
    }
    .title-line-secondary {
      margin-top: 4px;
      font-size: clamp(20px, 3.2vw, 30px);
      letter-spacing: -0.02em;
    }
    .meta {
      margin-top: 6px;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .pill {
      padding: 10px 14px;
      border-radius: 999px;
      border: 1px solid rgba(255,255,255,0.12);
      background: linear-gradient(180deg, #3d6348, #2f4f39);
      font-size: 14px;
      color: #f5efe2;
      text-align: center;
      white-space: nowrap;
    }
    .stats-stack {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }
    .stats {
      display: grid;
      gap: 12px;
      color: #f5efe2;
      background: linear-gradient(180deg, #3d6348, #2f4f39);
    }
    .stats h2 {
      margin: 0;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      color: rgba(245,239,226,0.74);
      white-space: nowrap;
    }
    .stats-grid {
      display: grid;
      grid-template-columns: 1fr;
      gap: 8px;
    }
    .stat {
      padding: 12px;
      border-radius: 18px;
      background: rgba(255,255,255,0.08);
      border: 1px solid rgba(255,255,255,0.1);
    }
    .stat-label {
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: rgba(245,239,226,0.74);
    }
    .stat-value {
      margin-top: 8px;
      font-size: clamp(20px, 3vw, 28px);
      font-weight: bold;
    }
    .layout {
      display: grid;
      gap: 18px;
      align-items: start;
    }
    .dashboard-main-grid {
      display: grid;
      grid-template-columns: minmax(0, 0.95fr) minmax(360px, 1.05fr);
      gap: 18px;
      align-items: start;
    }
    .form-card,
    .recent-jobs-card {
      height: 100%;
    }
    .recent-jobs-card {
      position: sticky;
      top: 18px;
    }
    h2.panel-title {
      margin: 0 0 8px;
      font-size: 28px;
    }
    .copy {
      margin: 0 0 18px;
      color: var(--muted);
      line-height: 1.45;
    }
    .status {
      margin-bottom: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      font-size: 15px;
      border: 1px solid transparent;
      line-height: 1.45;
    }
    .status.ok {
      background: rgba(60,95,70,0.1);
      border-color: rgba(60,95,70,0.16);
      color: #284332;
    }
    .status.error {
      background: rgba(139,71,56,0.1);
      border-color: rgba(139,71,56,0.16);
      color: #6d3124;
    }
    .form-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }
    .field {
      display: grid;
      gap: 7px;
      position: relative;
    }
    .field-date {
      max-width: 220px;
    }
    .field-date input {
      padding-right: 46px;
    }
    .field-date input::-webkit-date-and-time-value {
      text-align: left;
    }
    .field-date input::-webkit-calendar-picker-indicator {
      opacity: 0.75;
      cursor: pointer;
      filter: sepia(0.35) saturate(0.9);
    }
    .invoice-date-field {
      max-width: 220px;
    }
    .field.full {
      grid-column: 1 / -1;
    }
    .field.notes-field {
      grid-column: 1 / 2;
    }
    .field.issue-photos-field {
      grid-column: 2 / 3;
    }
    label {
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-weight: bold;
      color: #5d5b48;
    }
    input, select, button, textarea {
      font: inherit;
    }
    input, select, textarea {
      width: 100%;
      min-height: 56px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: #fffdfa;
      color: var(--ink);
      padding: 12px 14px;
      outline: none;
      font-size: 16px;
    }
    textarea {
      min-height: 110px;
      resize: vertical;
      padding-top: 14px;
    }
    .file-input {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }
    .file-button {
      min-height: 56px;
      width: 100%;
    }
    .file-selection {
      margin-top: 2px;
      font-size: 13px;
      color: var(--muted);
    }
    .photo-help {
      margin-top: 2px;
    }
    .photo-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(84px, 1fr));
      gap: 10px;
      margin-top: 10px;
    }
    .photo-thumb {
      display: grid;
      gap: 6px;
      text-decoration: none;
      color: var(--muted);
      font-size: 12px;
    }
    .photo-thumb img {
      width: 100%;
      aspect-ratio: 1 / 1;
      object-fit: cover;
      border-radius: 14px;
      border: 1px solid rgba(82, 69, 42, 0.12);
      background: rgba(255,255,255,0.72);
      box-shadow: 0 8px 18px rgba(60, 49, 25, 0.08);
    }
    .photo-count {
      display: inline-flex;
      align-items: center;
      min-height: 30px;
      padding: 4px 10px;
      border-radius: 999px;
      background: rgba(60,95,70,0.08);
      border: 1px solid rgba(60,95,70,0.1);
      color: var(--green);
      font-size: 12px;
      font-weight: bold;
      white-space: nowrap;
    }
    .job-photo-grid {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .job-photo-grid .photo-thumb {
      width: 56px;
    }
    .job-photo-grid .photo-thumb span {
      display: none;
    }
    input:focus, select:focus, textarea:focus {
      border-color: var(--gold);
      box-shadow: 0 0 0 4px rgba(186,148,80,0.14);
    }
    .suggestions {
      display: none;
      position: absolute;
      left: 0;
      right: 0;
      top: calc(100% + 6px);
      z-index: 30;
      max-height: 220px;
      overflow-y: auto;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.16);
      background: rgba(250, 247, 240, 0.98);
      box-shadow: 0 16px 30px rgba(60, 49, 25, 0.16);
      -webkit-overflow-scrolling: touch;
      overscroll-behavior: contain;
      touch-action: pan-y;
    }
    .suggestions.is-open {
      display: block;
    }
    .suggestion-item {
      width: 100%;
      padding: 12px 14px;
      border: none;
      border-bottom: 1px solid rgba(82, 69, 42, 0.08);
      background: transparent;
      color: var(--ink);
      text-align: left;
      cursor: pointer;
      font: inherit;
      touch-action: manipulation;
    }
    .suggestion-item:last-child {
      border-bottom: none;
    }
    .suggestion-item:active,
    .suggestion-item:focus,
    .suggestion-item:hover {
      background: rgba(186,148,80,0.14);
      outline: none;
    }
    .suggestion-item.is-match {
      background: rgba(186,148,80,0.08);
      font-weight: bold;
    }
    .hint {
      font-size: 13px;
      color: var(--muted);
    }
    .hint-centered {
      text-align: center;
      margin-top: 10px;
    }
    .actions {
      margin-top: 18px;
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }
    .actions form {
      margin: 0;
      width: 100%;
    }
    .actions-inline {
      display: flex;
      flex-wrap: nowrap;
      gap: 8px;
      align-items: center;
    }
    .actions-inline form {
      margin: 0;
      flex: 0 0 auto;
    }
    .actions-inline .button {
      width: auto;
      white-space: nowrap;
    }
    .section-grid {
      margin-top: 18px;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
      align-items: start;
    }
    .section-grid > .card,
    .bottom-export .card {
      display: grid;
      align-content: start;
    }
    .bottom-export {
      margin-top: 18px;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }
    .bottom-export .card {
      grid-column: 1 / -1;
    }
    .button {
      min-height: 56px;
      padding: 0 20px;
      border-radius: 999px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
    }
    .button-small {
      min-height: 40px;
      padding: 0 14px;
      font-size: 14px;
    }
    .button-primary {
      background: linear-gradient(135deg, var(--gold), #cda961);
      color: #2b2212;
      box-shadow: 0 14px 28px rgba(186,148,80,0.24);
    }
    .button-secondary {
      background: rgba(60,95,70,0.1);
      color: var(--green);
      border: 1px solid rgba(60,95,70,0.12);
    }
    .button-danger {
      background: rgba(139,71,56,0.1);
      color: var(--red);
      border: 1px solid rgba(139,71,56,0.16);
    }
    .button-full {
      width: 100%;
    }
    .download-notice {
      position: fixed;
      left: 50%;
      bottom: calc(18px + env(safe-area-inset-bottom));
      transform: translateX(-50%);
      z-index: 1000;
      width: min(92vw, 420px);
      padding: 14px 16px;
      border-radius: 18px;
      background: rgba(39, 45, 33, 0.94);
      color: #f5efe2;
      text-align: center;
      box-shadow: 0 18px 44px rgba(60, 49, 25, 0.28);
    }
    .download-notice[hidden] {
      display: none;
    }
    .download-frame {
      display: none;
      width: 0;
      height: 0;
      border: 0;
    }
    .mini-form {
      display: grid;
      gap: 12px;
    }
    .mini-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }
    .checkbox-list {
      display: grid;
      gap: 10px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.1);
      background: rgba(255,255,255,0.68);
    }
    .checkbox-item {
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 15px;
      color: var(--ink);
    }
    .checkbox-item input {
      width: 18px;
      height: 18px;
      margin: 0;
      accent-color: var(--green);
    }
    .checkbox-item span {
      color: var(--muted);
      font-size: 13px;
    }
    .message-preview {
      margin-top: 14px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.1);
      background: rgba(255,255,255,0.68);
      font-family: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
      font-size: 13px;
      line-height: 1.5;
      white-space: pre-wrap;
      color: #2f3428;
    }
    .invoice-preview-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      margin-top: 14px;
    }
    .invoice-preview-card {
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.1);
      background: rgba(255,255,255,0.68);
      display: grid;
      gap: 8px;
    }
    .invoice-preview-card h3 {
      margin: 0;
      font-size: 17px;
    }
    .invoice-preview-card .hint {
      margin: 0;
    }
    .metric-block {
      margin-top: 14px;
      padding: 14px 16px;
      border-radius: 16px;
      background: rgba(60,95,70,0.08);
      border: 1px solid rgba(60,95,70,0.1);
      display: grid;
      gap: 8px;
    }
    .metric-row {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      font-size: 15px;
    }
    .table-actions {
      min-width: 170px;
    }
    .table-wrap {
      overflow-x: auto;
      border: 1px solid rgba(82, 69, 42, 0.1);
      border-radius: 18px;
    }
    .mobile-jobs {
      display: none;
      gap: 12px;
    }
    .job-row {
      padding: 12px 14px;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.1);
      background: rgba(255,255,255,0.68);
    }
    .job-row-top {
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: flex-start;
    }
    .job-row-main {
      font-size: 16px;
      line-height: 1.2;
      min-width: 0;
      display: grid;
      gap: 10px;
    }
    .job-row-customer {
      font-size: 16px;
      font-weight: bold;
      line-height: 1.2;
    }
    .job-row-farm {
      font-size: 13px;
      line-height: 1.25;
      color: #6a6657;
      font-weight: normal;
    }
    .job-row-date {
      color: var(--muted);
      font-size: 13px;
      white-space: nowrap;
    }
    .job-row-meta {
      display: grid;
      gap: 10px;
      justify-items: end;
      flex: 0 0 auto;
    }
    .job-row-summary {
      font-size: 14px;
      line-height: 1.3;
      color: #6a6657;
      min-width: 0;
      word-break: break-word;
      display: grid;
      gap: 10px;
    }
    .job-row-tons {
      font-size: 13px;
      line-height: 1.25;
      color: #6a6657;
      white-space: nowrap;
    }
    .job-row-status {
      font-size: 13px;
      line-height: 1.25;
      color: #6a6657;
    }
    .status-chip {
      display: inline-flex;
      align-items: center;
      min-height: 30px;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: bold;
      border: 1px solid transparent;
      white-space: nowrap;
    }
    .status-chip.open {
      background: rgba(186,148,80,0.12);
      border-color: rgba(186,148,80,0.18);
      color: #7a5d22;
    }
    .status-chip.invoiced {
      background: rgba(60,95,70,0.1);
      border-color: rgba(60,95,70,0.16);
      color: #284332;
    }
    .status-chip.manual {
      background: rgba(82, 69, 42, 0.08);
      border-color: rgba(82, 69, 42, 0.12);
      color: #5d5b48;
    }
    .job-row .actions-inline {
      flex-direction: column;
      align-items: stretch;
    }
    table {
      width: 100%;
      min-width: 860px;
      border-collapse: collapse;
      background: rgba(255,255,255,0.68);
    }
    th, td {
      padding: 14px 16px;
      text-align: left;
      border-bottom: 1px solid rgba(82, 69, 42, 0.1);
    }
    th {
      background: rgba(240,232,216,0.85);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #5d5b48;
    }
    tr:last-child td {
      border-bottom: none;
    }
    .empty {
      padding: 26px;
      border-radius: 18px;
      border: 1px dashed var(--line);
      background: rgba(255,255,255,0.48);
      color: var(--muted);
      text-align: center;
    }
    .mono {
      font-family: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
    }
    @media (max-width: 980px) {
      .hero, .layout {
        grid-template-columns: 1fr;
      }
      .dashboard-main-grid {
        grid-template-columns: 1fr;
      }
      .recent-jobs-card {
        position: static;
      }
      .section-grid {
        grid-template-columns: 1fr;
      }
    }
    @media (max-width: 760px) {
      .page {
        padding: max(14px, env(safe-area-inset-top)) max(12px, env(safe-area-inset-right)) max(24px, env(safe-area-inset-bottom)) max(12px, env(safe-area-inset-left));
      }
      .card {
        padding: 16px;
        border-radius: 20px;
      }
      .hero-title-card {
        text-align: center;
        justify-items: center;
      }
      .hero-copy {
        text-align: center;
        font-size: 15px;
      }
      .hero {
        gap: 14px;
        margin-bottom: 14px;
      }
      .layout {
        gap: 14px;
      }
      .bottom-export {
        margin-top: 14px;
        grid-template-columns: 1fr;
        gap: 14px;
      }
      h1 {
        margin: 12px 0 8px;
        line-height: 1.02;
      }
      .title-line-primary {
        font-size: clamp(24px, 8vw, 34px);
      }
      .title-line-secondary {
        margin-top: 2px;
        font-size: clamp(17px, 5.4vw, 21px);
      }
      .stats h2 {
        font-size: 11px;
        letter-spacing: 0.02em;
      }
      .lead {
        font-size: 16px;
        line-height: 1.45;
      }
      .meta {
        margin-top: 14px;
      }
      .stats-stack {
        gap: 10px;
      }
      .pill {
        min-width: 0;
        padding: 10px 12px;
        font-size: 13px;
      }
      .form-grid {
        grid-template-columns: 1fr;
        gap: 12px;
      }
      .mini-grid {
        grid-template-columns: 1fr;
      }
      .invoice-preview-grid {
        grid-template-columns: 1fr;
      }
      .field {
        gap: 8px;
      }
      label {
        font-size: 14px;
        text-transform: none;
        letter-spacing: 0.02em;
      }
      input, select, textarea {
        min-height: 60px;
        padding: 14px 15px;
        font-size: 17px;
      }
      textarea {
        min-height: 128px;
      }
      .field-date,
      .invoice-date-field {
        max-width: none;
      }
      .field.notes-field,
      .field.issue-photos-field,
      .field.full {
        grid-column: 1 / -1;
      }
      .suggestion-item {
        padding: 14px 15px;
        font-size: 16px;
      }
      .file-button {
        min-height: 60px;
      }
      .actions {
        margin-top: 16px;
      }
      .button {
        width: 100%;
        min-height: 58px;
      }
      .actions-inline .button {
        width: auto;
      }
      .job-row {
        padding: 14px 15px;
      }
      .job-row-main {
        gap: 8px;
      }
      .job-row-customer {
        font-size: 17px;
      }
      .job-row-farm,
      .job-row-date,
      .job-row-tons,
      .job-row-status {
        font-size: 14px;
      }
      .job-row-summary {
        font-size: 15px;
        line-height: 1.38;
        gap: 8px;
      }
      .photo-grid {
        grid-template-columns: repeat(auto-fill, minmax(92px, 1fr));
      }
      .table-wrap {
        display: none;
      }
      .mobile-jobs {
        display: grid;
      }
    }
    @media (max-width: 420px) {
      .job-row-top,
      .job-row-bottom {
        display: block;
      }
      .job-row-date {
        display: block;
        margin-top: 4px;
        white-space: normal;
      }
      .job-row .actions-inline {
        margin-top: 8px;
      }
    }
  </style>
</head>
<body>
  <div class="page">
    {% if invoice_page %}
    <section class="layout">
      <div class="card hero-title-card">
        <h1>
          <span class="title-line title-line-primary">Create Invoice</span>
          <span class="title-line title-line-secondary">Invoice Page</span>
        </h1>
        <div class="actions">
          <a class="button button-secondary" href="{{ url_for('home') }}">Back to Dashboard</a>
          <a class="button button-secondary" href="{{ url_for('invoice_history') }}">Invoice History</a>
        </div>
      </div>

      {% if status_msg %}
      <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
      {% endif %}

      <div class="card">
        <h2 class="panel-title">Invoices</h2>
        <p class="copy">Create an invoice from uninvoiced jobs in the selected customer or farm scope. The customer receives the PDF only, and your accounts copies receive both the PDF and Excel workbook.</p>
        <form method="post" action="{{ url_for('invoice_create_and_send') }}">
          <input type="hidden" name="history_ledger_index" value="{{ invoice_form.history_ledger_index }}">
          {% if invoice_form.edit_reference_label %}
          <div class="status ok">Editing {{ invoice_form.edit_reference_label }}. Preview or send here will update and re-send this invoice.</div>
          {% endif %}
          <div class="form-grid">
            <div class="field">
              <label for="invoice_customer">Customer</label>
              <select id="invoice_customer" name="customer" required>
                <option value="">Select customer</option>
                {% for customer in customers %}
                <option value="{{ customer }}" {% if invoice_form.customer == customer %}selected{% endif %}>{{ customer }}</option>
                {% endfor %}
              </select>
            </div>
            <div class="field">
              <label for="invoice_farm_name">Farm</label>
              <select id="invoice_farm_name" name="farm_name" data-selected="{{ invoice_form.farm_name }}">
                <option value="">All Farms For Customer</option>
              </select>
            </div>
                        <div class="field">
                            <label for="invoice_customer_name_override">Invoice Customer Name</label>
                            <input id="invoice_customer_name_override" name="customer_name_override" type="text" value="{{ invoice_form.customer_name_override }}" placeholder="Uses saved customer name">
                        </div>
            <div class="field">
              <label for="invoice_number">Invoice Number</label>
              <input id="invoice_number" name="invoice_number" type="number" inputmode="numeric" min="1" step="1" value="{{ invoice_form.invoice_number }}" required>
              <div class="hint">You can enter any unused invoice number, including a missing number from earlier in the sequence.</div>
            </div>
            <div class="field invoice-date-field">
              <label for="invoice_date">Invoice Date</label>
              <input id="invoice_date" name="invoice_date" type="date" value="{{ invoice_form.invoice_date }}" required>
            </div>
            <div class="field invoice-date-field">
              <label for="invoice_job_date_from">Job Date From</label>
              <input id="invoice_job_date_from" name="job_date_from" type="date" value="{{ invoice_form.job_date_from }}">
              <div id="invoice_job_date_from_hint" class="hint">{% if invoice_form.job_date_from_hint %}Suggested from last invoice: {{ invoice_form.job_date_from_hint }}{% else %}Leave blank to include all uninvoiced jobs in scope.{% endif %}</div>
            </div>
            <div class="field">
              <label for="invoice_payment_terms">Payment Terms</label>
              <select id="invoice_payment_terms" name="payment_terms_days">
                {% for option in invoice_payment_terms_options %}
                <option value="{{ option }}" {% if invoice_form.payment_terms_days == option %}selected{% endif %}>{{ option }} days</option>
                {% endfor %}
              </select>
            </div>
            <div class="field notes-field">
              <label for="invoice_rate_override">Rate Per Ton</label>
              <input id="invoice_rate_override" name="rate_override" type="number" inputmode="decimal" min="0" step="0.01" placeholder="Uses customer master rate" value="{{ invoice_form.rate_override }}">
            </div>
                        <div class="field notes-field">
                            <label for="invoice_address_line_1">Billing Address</label>
                            <input id="invoice_address_line_1" name="address_line_1_override" type="text" value="{{ invoice_form.address_line_1_override }}" placeholder="Address line 1">
                            <input name="address_line_2_override" type="text" value="{{ invoice_form.address_line_2_override }}" placeholder="Address line 2">
                            <input name="town_override" type="text" value="{{ invoice_form.town_override }}" placeholder="Town">
                            <input name="postcode_override" type="text" value="{{ invoice_form.postcode_override }}" placeholder="Postcode">
                        </div>
            <div class="field notes-field">
              <label for="invoice_subject">Subject</label>
              <input id="invoice_subject" name="subject" type="text" value="{{ invoice_form.subject }}" placeholder="Invoice email subject">
            </div>
            <div class="field notes-field">
              <label for="invoice_customer_message">Customer Message</label>
              <textarea id="invoice_customer_message" name="customer_message" placeholder="Customer email message">{{ invoice_form.customer_message }}</textarea>
            </div>
            <div class="field notes-field">
              <label>Additional Fees</label>
              <div id="invoice_fee_rows" class="mini-form">
                {% for fee_row in invoice_form.additional_fee_rows %}
                <div class="mini-grid">
                  <input type="text" name="additional_fee_description" placeholder="Description" value="{{ fee_row.description }}">
                  <input type="number" name="additional_fee_amount" inputmode="decimal" min="0" step="0.01" placeholder="Amount" value="{{ fee_row.amount }}">
                </div>
                {% endfor %}
              </div>
              <div class="actions-inline">
                <button id="add_fee_row" class="button button-secondary button-small" type="button">Add Another Fee</button>
              </div>
            </div>
          </div>
          <div class="hint">Use Job Date From if you want to stop the invoice pulling in older uninvoiced jobs. The rate field defaults from the customer master but can be changed for this invoice. Additional fees use separate description and amount boxes, with VAT fixed at 20%.</div>
          <div class="hint">The customer email comes from the saved job snapshot or the current customer master. Office and Andrew receive accounts copies automatically. Owen is excluded from invoice copies.</div>
          <div class="actions">
            <button class="button button-secondary" type="submit" formaction="{{ url_for('invoice_preview_pdf') }}" formmethod="post" formtarget="_blank">Preview PDF</button>
            <button class="button button-secondary" type="submit" formaction="{{ url_for('invoice_preview_xlsx') }}" formmethod="post">Download Preview .xlsx</button>
            <button class="button button-primary" type="submit">Create and Email Invoice</button>
          </div>
        </form>
      </div>

      <div class="card">
        <h2 class="panel-title">Mark Already Invoiced</h2>
        <p class="copy">Use this for jobs that were invoiced outside the app so they are not picked up again.</p>
        <form method="post" action="{{ url_for('invoice_mark_existing') }}">
          <div class="form-grid">
            <div class="field">
              <label for="mark_customer">Customer</label>
              <select id="mark_customer" name="customer" required>
                <option value="">Select customer</option>
                {% for customer in customers %}
                <option value="{{ customer }}">{{ customer }}</option>
                {% endfor %}
              </select>
            </div>
            <div class="field">
              <label for="mark_farm_name">Farm</label>
              <select id="mark_farm_name" name="farm_name">
                <option value="">All Farms For Customer</option>
              </select>
            </div>
            <div class="field invoice-date-field">
              <label for="mark_through_date">Job Date Through</label>
              <input id="mark_through_date" name="through_date" type="date" value="{{ today_iso }}" required>
            </div>
            <div class="field">
              <label for="mark_note">Note</label>
              <input id="mark_note" name="note" type="text" placeholder="Optional note for history">
            </div>
          </div>
          <div class="actions">
            <button class="button button-secondary" type="submit">Mark As Already Invoiced</button>
          </div>
        </form>
      </div>

            <div class="card">
                <h2 class="panel-title">Unmark Invoiced Jobs</h2>
                <p class="copy">Choose a customer and invoice date. That invoice and any later invoice records for the customer will be removed, making those jobs available again.</p>
                <form method="post" action="{{ url_for('invoice_unmark_from_date') }}" id="invoice_unmark_form">
                    <div class="form-grid">
                        <div class="field">
                            <label for="unmark_customer">Customer</label>
                            <select id="unmark_customer" name="customer" required>
                                <option value="">Select customer</option>
                                {% for customer_name in invoice_unmark_options.keys()|sort %}
                                <option value="{{ customer_name }}">{{ customer_name }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="field">
                            <label for="unmark_invoice_date">Invoice Date</label>
                            <select id="unmark_invoice_date" name="invoice_date" required disabled>
                                <option value="">Select customer first</option>
                            </select>
                        </div>
                    </div>
                    <div class="hint">This removes invoice history from the selected date onward. Use with care for invoices already emailed.</div>
                    <div class="actions">
                        <button class="button button-danger" type="submit">Unmark From Selected Date</button>
                    </div>
                </form>
            </div>
            <script>
                (function () {
                    const choices = {{ invoice_unmark_options_json|safe }};
                    const customer = document.getElementById("unmark_customer");
                    const date = document.getElementById("unmark_invoice_date");
                    if (!customer || !date) return;
                    customer.addEventListener("change", function () {
                        const rows = choices[customer.value] || [];
                        date.innerHTML = "";
                        rows.forEach(function (row) {
                            const option = document.createElement("option");
                            option.value = row.value;
                            option.textContent = row.label;
                            date.appendChild(option);
                        });
                        date.disabled = rows.length === 0;
                    });
                }());
            </script>
    </section>
    {% else %}
    <section class="hero">
      <div class="card hero-title-card">
        <div class="eyebrow">Live Dashboard</div>
        <div>
          <h1>
            <span class="title-line title-line-primary">A. Farrell Contracting</span>
            <span class="title-line title-line-secondary">Muck Spreading Records</span>
          </h1>
          <p class="hero-copy">Record jobs quickly, keep customer data tidy, and export clean summaries and invoices without leaving the dashboard.</p>
        </div>
        <div class="meta">
          <div class="pill">Today: {{ today_human }}</div>
          <div class="pill">Saved jobs: {{ total_jobs }}</div>
        </div>
      </div>
      <div class="stats-stack">
        <div class="card stats">
          <h2>Today's Totals</h2>
          <div class="stats-grid">
            <div class="stat">
              <div class="stat-label">Jobs</div>
              <div class="stat-value">{{ today_job_count }}</div>
            </div>
            <div class="stat">
              <div class="stat-label">Spreader Tons</div>
              <div class="stat-value">{{ today_spreader_tons }}</div>
            </div>
            <div class="stat">
              <div class="stat-label">Ops Center Tons</div>
              <div class="stat-value">{{ today_john_deere_tons }}</div>
            </div>
          </div>
        </div>
        <div class="card stats">
          <h2>Year to Date Totals</h2>
          <div class="stats-grid">
            <div class="stat">
              <div class="stat-label">Jobs</div>
              <div class="stat-value">{{ year_job_count }}</div>
            </div>
            <div class="stat">
              <div class="stat-label">Spreader Tons</div>
              <div class="stat-value">{{ year_spreader_tons }}</div>
            </div>
            <div class="stat">
              <div class="stat-label">Ops Center Tons</div>
              <div class="stat-value">{{ year_john_deere_tons }}</div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <section class="layout">
      <div class="dashboard-main-grid">
      <div class="card form-card">
        <h2 class="panel-title">{{ form_title }}</h2>

        {% if status_msg %}
        <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
        {% endif %}

        <form method="post" action="{{ url_for('save_job') }}" enctype="multipart/form-data">
          <input type="hidden" name="edit_job_id" value="{{ form_job.id }}">
          <div class="form-grid">
            <div class="field field-date">
              <label for="job_date">Date</label>
                            <input id="job_date" name="job_date" type="date" value="{{ form_job.job_date or today_iso }}" required>
            </div>
            <div class="field">
              <label for="customer">Customer</label>
              <input id="customer" name="customer" type="text" placeholder="Start typing a customer name" autocomplete="off" value="{{ form_job.customer }}" required>
              <div id="customer_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="farm_name">Farm Name</label>
              <input id="farm_name" name="farm_name" type="text" placeholder="Start typing a farm name (optional)" autocomplete="off" value="{{ form_job.farm_name }}">
              <div id="farm_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="field_name">Field Name</label>
              <input id="field_name" name="field_name" type="text" placeholder="Start typing a field name" autocomplete="off" value="{{ form_job.field_name }}" required>
              <div id="field_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="muck_type">Muck Type</label>
              <input id="muck_type" name="muck_type" type="text" placeholder="Start typing a muck type" autocomplete="off" value="{{ form_job.muck_type }}" required>
              <div id="muck_type_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="spreader_tons">Total Spreader Tons</label>
              <input id="spreader_tons" name="spreader_tons" type="number" inputmode="decimal" min="0" step="0.01" placeholder="0.00" value="{{ form_job.total_spreader_tons }}" required>
            </div>
            <div class="field">
              <label for="john_deere_tons">Total Ops Center Tons</label>
              <input id="john_deere_tons" name="john_deere_tons" type="number" inputmode="decimal" min="0" step="0.01" placeholder="0.00" value="{{ form_job.total_john_deere_tons }}" required>
            </div>
            <div class="field notes-field">
              <label for="job_notes">Job Notes</label>
              <textarea id="job_notes" name="job_notes" placeholder="Optional notes for this job">{{ form_job.job_notes }}</textarea>
            </div>
            <div class="field issue-photos-field">
              <label for="issue_photos">Issue Photos</label>
              <input id="issue_photos" class="file-input" name="issue_photos" type="file" accept="image/*" multiple>
              <label class="button button-secondary file-button" for="issue_photos">Add Photos</label>
              <div id="issue_photo_selection" class="file-selection">No new photos selected</div>
              <div class="hint photo-help">Add photos of any issue found during the job. You can add more when editing later.</div>
              {% if form_job.issue_photo_items %}
              <div class="photo-grid">
                {% for photo in form_job.issue_photo_items %}
                <a class="photo-thumb" href="{{ photo.url }}" target="_blank" rel="noopener">
                  <img src="{{ photo.url }}" alt="Issue photo">
                  <span>{{ loop.index }}</span>
                </a>
                {% endfor %}
              </div>
              {% endif %}
            </div>
          </div>
          <div class="actions">
            <button class="button button-primary" type="submit">{{ form_submit_label }}</button>
            {% if is_editing %}
            <a class="button button-secondary" href="{{ url_for('home') }}">Cancel Edit</a>
            {% endif %}
          </div>
        </form>
      </div>

      <div class="card recent-jobs-card">
        <h2 class="panel-title">Recent Jobs</h2>
        <p class="copy">Newest entries are shown first.</p>
        {% if recent_jobs %}
        <div class="table-wrap">
          <table>
            <thead>
                <tr>
                  <th>Date</th>
                  <th>Customer</th>
                  <th>Farm</th>
                  <th>Field</th>
                  <th>Muck Type</th>
                  <th>Notes</th>
                  <th>Issue Photos</th>
                  <th>Spreader Tons</th>
                  <th>Ops Center Tons</th>
                  <th>Invoice Status</th>
                  <th>Saved</th>
                  <th class="table-actions">Actions</th>
                </tr>
            </thead>
            <tbody>
              {% for job in recent_jobs %}
                <tr>
                  <td>{{ job.job_date_label }}</td>
                  <td>{{ job.customer }}</td>
                  <td>{{ job.farm_name }}</td>
                  <td>{{ job.field_name }}</td>
                  <td>{{ job.muck_type }}</td>
                  <td>{{ job.job_notes or '--' }}</td>
                  <td>
                    {% if job.issue_photo_items %}
                    <div class="job-photo-grid">
                      {% for photo in job.issue_photo_items %}
                      <a class="photo-thumb" href="{{ photo.url }}" target="_blank" rel="noopener">
                        <img src="{{ photo.url }}" alt="Issue photo">
                        <span>{{ loop.index }}</span>
                      </a>
                      {% endfor %}
                    </div>
                    {% else %}
                    --
                    {% endif %}
                  </td>
                  <td>{{ job.spreader_tons_label }}</td>
                  <td>{{ job.john_deere_tons_label }}</td>
                  <td><span class="status-chip {{ job.invoice_status_key }}">{{ job.invoice_status_label }}</span></td>
                  <td>{{ job.saved_label }}</td>
                  <td>
                    <div class="actions-inline">
                      <a class="button button-secondary button-small" href="{{ url_for('muck_home', edit_id=job.id) }}">Edit</a>
                      <form method="post" action="{{ url_for('delete_job', job_id=job.id) }}" class="delete-job-form">
                        <button class="button button-danger button-small" type="submit">Delete</button>
                      </form>
                    </div>
                  </td>
                </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
        <div class="mobile-jobs">
          {% for job in recent_jobs %}
          <div class="job-row">
            <div class="job-row-top">
              <div class="job-row-main">
                <div class="job-row-customer">{{ job.customer }}</div>
                {% if job.farm_name %}
                <div class="job-row-farm">{{ job.farm_name }}</div>
                {% endif %}
                <div class="job-row-summary">
                  <div>{{ job.field_name }} | {{ job.muck_type }}</div>
                  <div class="job-row-tons">Spreader {{ job.spreader_tons_label }} | Ops Center {{ job.john_deere_tons_label }}</div>
                  <div class="job-row-status">Invoice: {{ job.invoice_status_label }}</div>
                  {% if job.issue_photo_count %}
                  <div class="photo-count">{{ job.issue_photo_count }} issue photo{% if job.issue_photo_count != 1 %}s{% endif %}</div>
                  {% endif %}
                  {% if job.issue_photo_items %}
                  <div class="job-photo-grid">
                    {% for photo in job.issue_photo_items %}
                    <a class="photo-thumb" href="{{ photo.url }}" target="_blank" rel="noopener">
                      <img src="{{ photo.url }}" alt="Issue photo">
                      <span>{{ loop.index }}</span>
                    </a>
                    {% endfor %}
                  </div>
                  {% endif %}
                </div>
              </div>
              <div class="job-row-meta">
                <div class="job-row-date">{{ job.job_date_label }}</div>
                <div class="actions-inline">
                  <a class="button button-secondary button-small" href="{{ url_for('muck_home', edit_id=job.id) }}">Edit</a>
                  <form method="post" action="{{ url_for('delete_job', job_id=job.id) }}" class="delete-job-form">
                    <button class="button button-danger button-small" type="submit">Delete</button>
                  </form>
                </div>

              </div>
            </div>
          </div>
          {% endfor %}
        </div>
        {% if has_more_recent_jobs %}
        <div class="actions">
          <a class="button button-secondary button-full" href="{{ url_for('muck_home', jobs_page=jobs_page + 1) }}">Load 20 More Jobs</a>
        </div>
        {% endif %}
        {% else %}
        <div class="empty">No jobs saved yet.</div>
        {% endif %}
      </div>
      </div>
    </section>

    <section class="section-grid">
      <div class="card">
        <h2 class="panel-title">Invoices</h2>
        <div class="actions">
          <a class="button button-secondary button-full" href="{{ url_for('invoice_home') }}">Create Invoice</a>
          <a class="button button-secondary button-full" href="{{ url_for('invoice_history') }}">Previous Invoices</a>
        </div>
      </div>

      <div class="card">
        <h2 class="panel-title">Tools</h2>
        <p class="copy">Use these for admin and backups.</p>
        <div class="actions">
          <a class="button button-secondary button-full" href="{{ url_for('admin_home') }}">Open Data Admin</a>
          <a class="button button-secondary button-full" href="{{ url_for('settings_home') }}">Open Settings</a>
          <a class="button button-secondary button-full" href="{{ url_for('backup_export_zip') }}">Download Backup ZIP</a>
        </div>
      </div>
    </section>

    <section class="bottom-export">
      <div class="card">
        <h2 class="panel-title">Summary Exports</h2>
        <p class="copy">Download period summaries or the full saved job list as Excel files.</p>
        <div class="actions">
          <a class="button button-secondary button-full" data-download-link="1" href="{{ url_for('export_period_summary_xlsx', period_key='current-week') }}">Download Current Week Summary .xlsx</a>
          <a class="button button-secondary button-full" data-download-link="1" href="{{ url_for('export_period_summary_xlsx', period_key='current-month') }}">Download Current Month Summary .xlsx</a>
          <a class="button button-secondary button-full" data-download-link="1" href="{{ url_for('export_period_summary_xlsx', period_key='last-month') }}">Download Last Month Summary .xlsx</a>
          <a class="button button-secondary button-full" data-download-link="1" href="{{ url_for('export_jobs_xlsx') }}">Download Full Job List .xlsx</a>
        </div>
      </div>
    </section>
    {% endif %}
  </div>
  <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
    <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
  </nav>
  <div id="download_notice" class="download-notice" hidden>Preparing download...</div>

  <script>
    const fieldMap = {{ field_map_json|safe }};
    const customerFarmMap = {{ customer_farm_map_json|safe }};
    const customerRateMap = {{ customer_rate_map_json|safe }};
    const customerInvoiceFromMap = {{ customer_invoice_from_map_json|safe }};
    const allCustomers = {{ customers_json|safe }};
    const customerInput = document.getElementById("customer");
    const customerSuggestions = document.getElementById("customer_suggestions");
    const farmInput = document.getElementById("farm_name");
    const farmSuggestions = document.getElementById("farm_suggestions");
    const fieldInput = document.getElementById("field_name");
    const fieldSuggestions = document.getElementById("field_suggestions");
    const muckTypeInput = document.getElementById("muck_type");
    const muckTypeSuggestions = document.getElementById("muck_type_suggestions");
    const invoiceCustomerSelect = document.getElementById("invoice_customer");
    const invoiceFarmSelect = document.getElementById("invoice_farm_name");
    const invoiceRateInput = document.getElementById("invoice_rate_override");
    const invoiceJobDateFromInput = document.getElementById("invoice_job_date_from");
    const invoiceJobDateFromHint = document.getElementById("invoice_job_date_from_hint");
    const markCustomerSelect = document.getElementById("mark_customer");
    const markFarmSelect = document.getElementById("mark_farm_name");
    const invoiceFeeRows = document.getElementById("invoice_fee_rows");
    const addFeeRowButton = document.getElementById("add_fee_row");
    const issuePhotosInput = document.getElementById("issue_photos");
    const issuePhotoSelection = document.getElementById("issue_photo_selection");
    const downloadNotice = document.getElementById("download_notice");
    const updateAppForm = document.getElementById("update_app_form");
    const updateAppButton = updateAppForm ? updateAppForm.querySelector('button[type="submit"]') : null;
    const allFields = {{ all_fields_json|safe }};
    const allFarms = {{ all_farms_json|safe }};
    const allMuckTypes = {{ muck_types_json|safe }};
    const suggestionBoxes = [customerSuggestions, farmSuggestions, fieldSuggestions, muckTypeSuggestions];
    function closeSuggestions(exceptBox) {
      for (const box of suggestionBoxes) {
        if (box !== exceptBox) {
          box.classList.remove("is-open");
          box.innerHTML = "";
        }
      }
    }

    function findCustomerKey(rawValue) {
      const typed = String(rawValue || "").trim().toLowerCase();
      if (!typed) return "";
      for (const key of allCustomers) {
        if (key.toLowerCase() === typed) {
          return key;
        }
      }
      return "";
    }

    function filterOptions(options, typedValue) {
      const typed = String(typedValue || "").trim().toLowerCase();
      const filtered = [];
      for (const option of options) {
        const normalized = option.toLowerCase();
        if (!typed || normalized.includes(typed)) {
          filtered.push({
            value: option,
            starts: typed ? normalized.startsWith(typed) : false,
            containsIndex: typed ? normalized.indexOf(typed) : 9999,
          });
        }
      }
      filtered.sort(function (a, b) {
        if (a.starts !== b.starts) return a.starts ? -1 : 1;
        if (a.containsIndex !== b.containsIndex) return a.containsIndex - b.containsIndex;
        return a.value.localeCompare(b.value);
      });
      return filtered;
    }

    let downloadNoticeTimer = null;
    function showDownloadNotice(message) {
      if (!downloadNotice) {
        return;
      }
      downloadNotice.textContent = message;
      downloadNotice.hidden = false;
      if (downloadNoticeTimer) {
        window.clearTimeout(downloadNoticeTimer);
      }
      downloadNoticeTimer = window.setTimeout(function () {
        downloadNotice.hidden = true;
      }, 3200);
    }

    function parseDownloadFilename(response, fallbackName) {
      const disposition = response.headers.get("Content-Disposition") || "";
      const utf8Match = disposition.match(/filename\\*=UTF-8''([^;]+)/i);
      if (utf8Match && utf8Match[1]) {
        try {
          return decodeURIComponent(utf8Match[1]).replace(/^["']|["']$/g, "");
        } catch (error) {
        }
      }
      const plainMatch = disposition.match(/filename=([^;]+)/i);
      if (plainMatch && plainMatch[1]) {
        return plainMatch[1].trim().replace(/^["']|["']$/g, "");
      }
      return fallbackName;
    }

    function formatDisplayDate(isoValue) {
      const raw = String(isoValue || "").trim();
      if (!raw) {
        return "";
      }
      const parts = raw.split("-");
      if (parts.length !== 3) {
        return raw;
      }
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const year = parts[0];
      const monthIndex = Number(parts[1]) - 1;
      const day = Number(parts[2]);
      if (!Number.isFinite(monthIndex) || monthIndex < 0 || monthIndex > 11 || !Number.isFinite(day)) {
        return raw;
      }
      return day + " " + months[monthIndex] + " " + year;
    }

    async function triggerExportDownload(link) {
      const targetUrl = new URL(link.getAttribute("href"), window.location.href);
      targetUrl.searchParams.set("_dl", Date.now().toString());
      showDownloadNotice("Preparing download...");

      const response = await fetch(targetUrl.toString(), {
        credentials: "same-origin",
      });
      if (!response.ok) {
        throw new Error("Download failed");
      }

      const blob = await response.blob();
      const fallbackName = link.textContent.trim().replace(/^Download\\s+/i, "") || "download.xlsx";
      const filename = parseDownloadFilename(response, fallbackName);

      if (navigator.share && window.File) {
        try {
          const sharedFile = new File([blob], filename, {type: blob.type || "application/octet-stream"});
          if (!navigator.canShare || navigator.canShare({files: [sharedFile]})) {
            showDownloadNotice("Opening share options...");
            await navigator.share({
              files: [sharedFile],
              title: filename,
            });
            return;
          }
        } catch (error) {
          if (error && error.name === "AbortError") {
            showDownloadNotice("Download cancelled");
            return;
          }
        }
      }

      const objectUrl = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = filename;
      anchor.rel = "noopener";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(function () {
        URL.revokeObjectURL(objectUrl);
      }, 1000);
      showDownloadNotice("Download started");
    }

    async function waitForUpdatedApp(targetUrl) {
      showDownloadNotice("App updated. Restarting now...");
      const startedAt = Date.now();

      async function checkReady() {
        try {
          const probeUrl = new URL(targetUrl, window.location.href);
          probeUrl.searchParams.set("_ping", Date.now().toString());
          const response = await fetch(probeUrl.toString(), {
            credentials: "same-origin",
            cache: "no-store",
          });
          if (response.ok) {
            window.location.replace(targetUrl);
            return;
          }
        } catch (error) {
        }

        if (Date.now() - startedAt >= 60000) {
          showDownloadNotice("App updated. Refresh this page if it does not reconnect.");
          if (updateAppButton) {
            updateAppButton.disabled = false;
            updateAppButton.textContent = "Update App";
          }
          return;
        }

        window.setTimeout(checkReady, 1500);
      }

      window.setTimeout(checkReady, 4500);
    }

    function openSuggestionBox(box, values, onSelect) {
      closeSuggestions(box);
      box.innerHTML = "";
      if (!values.length) {
        box.classList.remove("is-open");
        return;
      }
      for (const item of values) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "suggestion-item";
        if (item.starts) {
          button.classList.add("is-match");
        }
        button.textContent = item.value;
        button.addEventListener("click", function () {
          onSelect(item.value);
          closeSuggestions();
        });
        box.appendChild(button);
      }
      box.classList.add("is-open");
    }

    function currentFarmOptions() {
      const customerKey = findCustomerKey(customerInput.value);
      return customerKey ? (customerFarmMap[customerKey] || []) : allFarms;
    }

    function valueInOptions(rawValue, options) {
      const typed = String(rawValue || "").trim().toLowerCase();
      if (!typed) return false;
      for (const option of options) {
        if (String(option || "").trim().toLowerCase() === typed) {
          return true;
        }
      }
      return false;
    }

    function currentFieldOptions() {
      const customerKey = findCustomerKey(customerInput.value);
      if (!customerKey) return [];

      const customerFields = fieldMap[customerKey] || {};
      const farmKey = String(farmInput.value || "").trim().toLowerCase();
      const combined = [];

      function addFields(values) {
        if (!Array.isArray(values)) return;
        for (const value of values) {
          if (value && !combined.includes(value)) {
            combined.push(value);
          }
        }
      }

      addFields(customerFields[""]);
      addFields(customerFields["__all__"]);

      if (farmKey) {
        for (const key of Object.keys(customerFields)) {
          if (key.toLowerCase() === farmKey) {
            addFields(customerFields[key]);
          }
        }
      } else {
        for (const key of Object.keys(customerFields)) {
          addFields(customerFields[key]);
        }
      }

      return combined;
    }

    function syncDependentInputs() {
      // Keep existing edit values in place when the customer changes.
      // Suggestions still update to the new customer/farm scope, but we do not
      // silently clear the farm or field while the user is correcting a job.
    }

    function showCustomerSuggestions() {
      openSuggestionBox(customerSuggestions, filterOptions(allCustomers, customerInput.value), function (value) {
        customerInput.value = value;
        syncDependentInputs();
        if (!String(farmInput.value || "").trim()) {
          const farms = currentFarmOptions();
          if (farms.length === 1) {
            farmInput.value = farms[0];
          }
        }
        showFarmSuggestions();
        showFieldSuggestions();
      });
    }

    function showFarmSuggestions() {
      openSuggestionBox(farmSuggestions, filterOptions(currentFarmOptions(), farmInput.value), function (value) {
        farmInput.value = value;
        showFieldSuggestions();
      });
    }

    function showFieldSuggestions() {
      openSuggestionBox(fieldSuggestions, filterOptions(currentFieldOptions(), fieldInput.value), function (value) {
        fieldInput.value = value;
      });
    }

    function showMuckTypeSuggestions() {
      openSuggestionBox(muckTypeSuggestions, filterOptions(allMuckTypes, muckTypeInput.value), function (value) {
        muckTypeInput.value = value;
      });
    }

    if (customerInput && customerSuggestions && farmInput && farmSuggestions && fieldInput && fieldSuggestions) {
      customerInput.addEventListener("input", function () {
        syncDependentInputs();
        showCustomerSuggestions();
        if (!String(farmInput.value || "").trim()) {
          showFarmSuggestions();
        }
        showFieldSuggestions();
      });
      customerInput.addEventListener("keyup", showCustomerSuggestions);
      customerInput.addEventListener("focus", function () {
        showCustomerSuggestions();
      });
      farmInput.addEventListener("input", function () {
        showFarmSuggestions();
        showFieldSuggestions();
      });
      farmInput.addEventListener("keyup", function () {
        showFarmSuggestions();
        showFieldSuggestions();
      });
      farmInput.addEventListener("focus", function () {
        showFarmSuggestions();
      });
      fieldInput.addEventListener("input", showFieldSuggestions);
      fieldInput.addEventListener("keyup", showFieldSuggestions);
      fieldInput.addEventListener("focus", function () {
        showFieldSuggestions();
      });
    }
    if (muckTypeInput && muckTypeSuggestions) {
      muckTypeInput.addEventListener("input", showMuckTypeSuggestions);
      muckTypeInput.addEventListener("keyup", showMuckTypeSuggestions);
      muckTypeInput.addEventListener("focus", function () {
        showMuckTypeSuggestions();
      });
    }

    document.addEventListener("click", function (event) {
      if (!event.target.closest(".field")) {
        closeSuggestions();
      }
    });

    for (const form of document.querySelectorAll(".delete-job-form")) {
      form.addEventListener("submit", function (event) {
        if (!window.confirm("Delete this saved job?")) {
          event.preventDefault();
          return;
        }
      });
    }

    for (const link of document.querySelectorAll("[data-download-link]")) {
      link.addEventListener("click", async function (event) {
        event.preventDefault();
        try {
          await triggerExportDownload(link);
        } catch (error) {
          showDownloadNotice("Download failed");
        }
      });
    }

    if (updateAppForm && updateAppButton) {
      updateAppForm.addEventListener("submit", async function (event) {
        event.preventDefault();
        updateAppButton.disabled = true;
        updateAppButton.textContent = "Updating...";
        showDownloadNotice("Checking for updates...");

        try {
          const response = await fetch(updateAppForm.action, {
            method: "POST",
            credentials: "same-origin",
            headers: {
              "Accept": "application/json",
              "X-Requested-With": "fetch",
            },
          });
          const payload = await response.json();
          if (!payload || !payload.ok) {
            throw new Error((payload && payload.msg) || "Update failed");
          }

          if (payload.status === "up_to_date") {
            showDownloadNotice(payload.msg || "App is already up to date");
            updateAppButton.disabled = false;
            updateAppButton.textContent = "Update App";
            return;
          }

          await waitForUpdatedApp(payload.redirect_url || {{ url_for('home', ok=1, msg='App updated')|tojson }});
        } catch (error) {
          showDownloadNotice((error && error.message) || "Update failed");
          updateAppButton.disabled = false;
          updateAppButton.textContent = "Update App";
        }
      });
    }

    function syncInvoiceFarmOptions() {
      if (!invoiceCustomerSelect || !invoiceFarmSelect) {
        return;
      }
      const customerName = String(invoiceCustomerSelect.value || "").trim();
      const farms = customerName ? (customerFarmMap[customerName] || []) : [];
      const previous = String(invoiceFarmSelect.value || invoiceFarmSelect.dataset.selected || "");
      invoiceFarmSelect.innerHTML = "";
      const allOption = document.createElement("option");
      allOption.value = "";
      allOption.textContent = "All Farms For Customer";
      invoiceFarmSelect.appendChild(allOption);
      for (const farm of farms) {
        const option = document.createElement("option");
        option.value = farm;
        option.textContent = farm;
        if (farm === previous) {
          option.selected = true;
        }
        invoiceFarmSelect.appendChild(option);
      }
      if (previous && !farms.includes(previous)) {
        invoiceFarmSelect.value = "";
      }
      invoiceFarmSelect.dataset.selected = invoiceFarmSelect.value;
    }

    function syncMarkFarmOptions() {
      if (!markCustomerSelect || !markFarmSelect) {
        return;
      }
      const customerName = String(markCustomerSelect.value || "").trim();
      const farms = customerName ? (customerFarmMap[customerName] || []) : [];
      const previous = String(markFarmSelect.value || markFarmSelect.dataset.selected || "");
      markFarmSelect.innerHTML = "";
      const allOption = document.createElement("option");
      allOption.value = "";
      allOption.textContent = "All Farms For Customer";
      markFarmSelect.appendChild(allOption);
      for (const farm of farms) {
        const option = document.createElement("option");
        option.value = farm;
        option.textContent = farm;
        if (farm === previous) {
          option.selected = true;
        }
        markFarmSelect.appendChild(option);
      }
      if (previous && !farms.includes(previous)) {
        markFarmSelect.value = "";
      }
      markFarmSelect.dataset.selected = markFarmSelect.value;
    }

    function invoiceDefaultRate() {
      if (!invoiceCustomerSelect) {
        return "";
      }
      const customerName = String(invoiceCustomerSelect.value || "").trim();
      const farmName = String((invoiceFarmSelect && invoiceFarmSelect.value) || "").trim();
      const customerRates = customerRateMap[customerName] || {};
      if (farmName && customerRates[farmName]) {
        return customerRates[farmName];
      }
      if (customerRates[""]) {
        return customerRates[""];
      }
      const rateKeys = Object.keys(customerRates);
      return rateKeys.length ? customerRates[rateKeys[0]] : "";
    }

    function invoiceDefaultJobDateFrom() {
      if (!invoiceCustomerSelect) {
        return "";
      }
      const customerName = String(invoiceCustomerSelect.value || "").trim();
      const farmName = String((invoiceFarmSelect && invoiceFarmSelect.value) || "").trim();
      const customerDates = customerInvoiceFromMap[customerName] || {};
      if (farmName && customerDates[farmName]) {
        return customerDates[farmName];
      }
      if (customerDates[""]) {
        return customerDates[""];
      }
      return "";
    }

    function syncInvoiceRate(force) {
      if (!invoiceRateInput) {
        return;
      }
      const defaultRate = invoiceDefaultRate();
      if (force || invoiceRateInput.dataset.userEdited !== "1" || !String(invoiceRateInput.value || "").trim()) {
        invoiceRateInput.value = defaultRate;
      }
    }

    function syncInvoiceJobDateFrom(force) {
      if (!invoiceJobDateFromInput) {
        return;
      }
      const defaultDate = invoiceDefaultJobDateFrom();
      if (force || invoiceJobDateFromInput.dataset.userEdited !== "1" || !String(invoiceJobDateFromInput.value || "").trim()) {
        invoiceJobDateFromInput.value = defaultDate;
      }
      if (invoiceJobDateFromHint) {
        invoiceJobDateFromHint.textContent = defaultDate
          ? ("Suggested from last invoice: " + formatDisplayDate(defaultDate))
          : "Leave blank to include all uninvoiced jobs in scope.";
      }
    }

    if (invoiceCustomerSelect && invoiceFarmSelect) {
      invoiceCustomerSelect.addEventListener("change", function () {
        syncInvoiceFarmOptions();
        syncInvoiceRate(false);
        syncInvoiceJobDateFrom(false);
      });
      invoiceFarmSelect.addEventListener("change", function () {
        syncInvoiceRate(false);
        syncInvoiceJobDateFrom(false);
      });
      syncInvoiceFarmOptions();
      syncInvoiceRate(false);
      syncInvoiceJobDateFrom(false);
    }
    if (markCustomerSelect && markFarmSelect) {
      markCustomerSelect.addEventListener("change", function () {
        syncMarkFarmOptions();
      });
      syncMarkFarmOptions();
    }
    if (invoiceRateInput) {
      invoiceRateInput.addEventListener("input", function () {
        invoiceRateInput.dataset.userEdited = "1";
      });
    }
    if (invoiceJobDateFromInput) {
      invoiceJobDateFromInput.addEventListener("input", function () {
        invoiceJobDateFromInput.dataset.userEdited = "1";
      });
    }
    if (invoiceFeeRows && addFeeRowButton) {
      addFeeRowButton.addEventListener("click", function () {
        const row = document.createElement("div");
        row.className = "mini-grid";
        row.innerHTML = '<input type="text" name="additional_fee_description" placeholder="Description"><input type="number" name="additional_fee_amount" inputmode="decimal" min="0" step="0.01" placeholder="Amount">';
        invoiceFeeRows.appendChild(row);
      });
    }
    if (issuePhotosInput && issuePhotoSelection) {
      issuePhotosInput.addEventListener("change", function () {
        const count = issuePhotosInput.files ? issuePhotosInput.files.length : 0;
        issuePhotoSelection.textContent = count ? (count + " new photo" + (count === 1 ? "" : "s") + " selected") : "No new photos selected";
      });
    }
  </script>
</body>
</html>
"""

ADMIN_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="#334d38">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <meta name="apple-mobile-web-app-title" content="Muck Jobs">
  <title>A. Farrell Contracting Data Admin</title>
  <link rel="icon" type="image/png" href="{{ url_for('app_icon_png', size=180) }}">
  <link rel="icon" type="image/svg+xml" href="{{ url_for('favicon') }}">
  <link rel="apple-touch-icon" sizes="180x180" href="{{ url_for('app_icon_png', size=180) }}">
  <link rel="manifest" href="{{ url_for('web_manifest') }}">
  <style>
    @import url('/static/af_brand.css');
    :root {
      --bg: #edf1ea;
      --panel: rgba(255, 255, 255, 0.94);
      --ink: #1d2b1f;
      --muted: #5b6b60;
      --line: rgba(36, 68, 47, 0.12);
      --green: #24442f;
      --gold: #c79e4f;
      --red: #8b4738;
      --shadow: 0 18px 36px rgba(29, 48, 35, 0.08);
      --font-main: Inter, "Segoe UI", system-ui, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font-main);
      color: var(--ink);
      background: linear-gradient(180deg, #edf1ea 0%, #e7ece3 100%);
    }
    .page {
      max-width: 1180px;
      margin: 0 auto;
      padding: 18px 14px 28px;
    }
    .grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
      align-items: start;
    }
    .card {
      background: var(--panel);
      border: 1px solid rgba(82, 69, 42, 0.12);
      border-radius: 24px;
      box-shadow: var(--shadow);
      padding: 20px;
    }
    h1, h2, h3 { margin-top: 0; }
    .copy { color: var(--muted); line-height: 1.45; }
    .status {
      margin-bottom: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      font-size: 15px;
      border: 1px solid transparent;
    }
    .status.ok {
      background: rgba(60,95,70,0.1);
      border-color: rgba(60,95,70,0.16);
      color: #284332;
    }
    .status.error {
      background: rgba(139,71,56,0.1);
      border-color: rgba(139,71,56,0.16);
      color: #6d3124;
    }
    .mini-form, .stack { display: grid; gap: 12px; }
    .mini-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }
    label {
      display: block;
      margin-bottom: 6px;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-weight: bold;
      color: #5d5b48;
    }
    input, select, button {
      font: inherit;
    }
    input, select {
      width: 100%;
      min-height: 48px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: #fffdfa;
      color: var(--ink);
      padding: 12px 14px;
      outline: none;
      font-size: 16px;
    }
    .suggestion-field {
      position: relative;
    }
    .suggestions {
      display: none;
      position: absolute;
      left: 0;
      right: 0;
      top: calc(100% + 6px);
      z-index: 30;
      max-height: 220px;
      overflow-y: auto;
      border-radius: 16px;
      border: 1px solid rgba(82, 69, 42, 0.16);
      background: rgba(250, 247, 240, 0.98);
      box-shadow: 0 16px 30px rgba(60, 49, 25, 0.16);
      -webkit-overflow-scrolling: touch;
      overscroll-behavior: contain;
      touch-action: pan-y;
    }
    .suggestions.is-open {
      display: block;
    }
    .suggestion-item {
      width: 100%;
      padding: 12px 14px;
      border: none;
      border-bottom: 1px solid rgba(82, 69, 42, 0.08);
      background: transparent;
      color: var(--ink);
      text-align: left;
      cursor: pointer;
      font: inherit;
      touch-action: manipulation;
    }
    .suggestion-item:last-child {
      border-bottom: none;
    }
    .suggestion-item:active,
    .suggestion-item:focus,
    .suggestion-item:hover {
      background: rgba(186,148,80,0.14);
      outline: none;
    }
    .suggestion-item.is-match {
      background: rgba(186,148,80,0.08);
      font-weight: bold;
    }
    .button {
      min-height: 46px;
      padding: 0 18px;
      border-radius: 14px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-weight: 700;
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .button:hover { transform: translateY(-1px); }
    .button-secondary {
      background: #edf5ee;
      color: var(--green);
      border: 1px solid rgba(36, 68, 47, 0.08);
    }
    .button-danger {
      background: rgba(139,71,56,0.1);
      color: var(--red);
      border: 1px solid rgba(139,71,56,0.16);
    }
    .button-full { width: 100%; }
    .button-small {
      min-height: 38px;
      padding: 0 12px;
      font-size: 14px;
    }
    .actions-inline {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }
    .map-list { display: grid; gap: 12px; }
    .map-customer, .map-farm {
      border: 1px solid rgba(82, 69, 42, 0.12);
      border-radius: 16px;
      background: rgba(255,255,255,0.55);
    }
    .map-customer {
      padding: 0;
      overflow: hidden;
    }
    .map-customer summary {
      list-style: none;
      cursor: pointer;
      padding: 14px;
    }
    .map-customer summary::-webkit-details-marker {
      display: none;
    }
    .map-customer summary::marker {
      content: "";
    }
    .map-customer-title {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
    }
    .map-customer-name {
      font-weight: bold;
      font-size: 18px;
    }
    .map-customer-meta {
      font-size: 12px;
      color: var(--muted);
      white-space: nowrap;
    }
    .map-customer-body {
      padding: 0 14px 14px;
      border-top: 1px solid rgba(82, 69, 42, 0.08);
      background: rgba(255,255,255,0.42);
    }
    .farm-list {
      display: grid;
      gap: 10px;
      margin-top: 12px;
    }
    .map-farm {
      padding: 0;
      overflow: hidden;
    }
    .map-farm summary {
      list-style: none;
      cursor: pointer;
      padding: 14px;
    }
    .map-farm summary::-webkit-details-marker {
      display: none;
    }
    .map-farm summary::marker {
      content: "";
    }
    .map-farm-title {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
    }
    .map-farm-name {
      font-weight: bold;
    }
    .map-farm-meta {
      font-size: 12px;
      color: var(--muted);
      white-space: nowrap;
    }
    .map-farm-body {
      padding: 0 14px 14px;
      border-top: 1px solid rgba(82, 69, 42, 0.08);
      background: rgba(255,255,255,0.42);
    }
    .field-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }
    .field-tag {
      display: inline-flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 6px;
      padding: 8px 10px;
      border-radius: 999px;
      background: rgba(60,95,70,0.08);
      border: 1px solid rgba(60,95,70,0.1);
    }
    .field-move-form {
      display: inline-flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 6px;
    }
        .field-move-details {
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .field-move-details > summary {
            cursor: pointer;
            color: var(--green);
            font-size: 13px;
            font-weight: bold;
        }
    .field-move-input {
      min-width: 160px;
      max-width: 220px;
      padding: 8px 10px;
      border-radius: 999px;
      border: 1px solid rgba(82, 69, 42, 0.16);
      background: rgba(255,255,255,0.92);
      font: inherit;
      color: var(--ink);
    }
    .top-links {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
    }
    .customer-meta-form {
      margin: 12px 0 14px;
      padding: 12px;
      border-radius: 16px;
      background: rgba(255,255,255,0.44);
      border: 1px solid rgba(82, 69, 42, 0.08);
    }
    .customer-meta-summary {
      display: grid;
      gap: 6px;
      margin: 8px 0 0;
      color: var(--muted);
      font-size: 14px;
    }
        .admin-action-group {
            border-top: 1px solid rgba(82, 69, 42, 0.12);
            padding-top: 12px;
        }
        .admin-action-group > summary {
            cursor: pointer;
            color: var(--green);
            font-weight: bold;
            padding: 6px 0 12px;
        }
        .admin-action-group > summary::marker {
            color: var(--gold);
        }
    @media (max-width: 860px) {
      .grid, .mini-grid {
        grid-template-columns: 1fr;
      }
    }
  </style>
</head>
<body>
  <div class="page">
    <div class="top-links">
      <a class="button button-secondary" href="{{ url_for('home') }}">Back To Jobs</a>
      <a class="button button-secondary" href="{{ url_for('backup_export_zip') }}">Download Backup ZIP</a>
    </div>

    {% if status_msg %}
    <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
    {% endif %}

    <datalist id="admin_customer_move_options">
      {% for customer in customers %}
      <option value="{{ customer }}"></option>
      {% endfor %}
    </datalist>

    <div class="grid">
      <div class="card stack">
        <div>
          <h1>Data Admin</h1>
          <p class="copy">Manage customer, farm, and field links used by saved jobs.</p>
        </div>

        <div>
          <h2>Add Field Link</h2>
          <form class="mini-form" method="post" action="{{ url_for('admin_add_field') }}">
            <div class="mini-grid">
              <div class="suggestion-field">
                <label for="admin_customer_select">Existing Customer</label>
                <input id="admin_customer_select" name="customer" type="text" placeholder="Select customer" autocomplete="off">
                <div id="admin_customer_suggestions" class="suggestions"></div>
              </div>
              <div>
                <label for="admin_customer_new">New Customer Name</label>
                <input id="admin_customer_new" name="new_customer" type="text" placeholder="Optional new customer name">
              </div>
            </div>
            <div>
              <label for="admin_farm_name">Farm Name</label>
              <input id="admin_farm_name" name="farm_name" type="text" placeholder="Optional farm name">
            </div>
            <div>
              <label for="admin_field_name">Field Name</label>
              <input id="admin_field_name" name="field_name" type="text" required>
            </div>
            <button class="button button-secondary button-full" type="submit">Add Field</button>
          </form>
        </div>

                <details class="admin-action-group">
                    <summary>Move jobs</summary>
          <form class="mini-form" method="post" action="{{ url_for('admin_move_jobs') }}">
            <div class="mini-grid">
              <div>
                <label for="admin_move_jobs_from_customer">From Customer</label>
                <input id="admin_move_jobs_from_customer" name="from_customer" type="text" list="admin_customer_move_options" placeholder="Current customer" required>
              </div>
              <div>
                <label for="admin_move_jobs_to_customer">To Customer</label>
                <input id="admin_move_jobs_to_customer" name="to_customer" type="text" list="admin_customer_move_options" placeholder="New customer" required>
              </div>
            </div>
            <div class="mini-grid">
              <div>
                <label for="admin_move_jobs_from_farm">From Farm (optional)</label>
                <input id="admin_move_jobs_from_farm" name="from_farm_name" type="text" placeholder="Only this farm">
              </div>
              <div>
                <label for="admin_move_jobs_to_farm">To Farm (optional)</label>
                <input id="admin_move_jobs_to_farm" name="to_farm_name" type="text" placeholder="Keep same if blank">
              </div>
            </div>
            <div>
              <label for="admin_move_jobs_field">Field (optional)</label>
              <input id="admin_move_jobs_field" name="field_name" type="text" placeholder="Only this field">
            </div>
            <button class="button button-secondary button-full" type="submit">Move Matching Jobs</button>
          </form>
                </details>

                <details class="admin-action-group">
                    <summary>Merge customers</summary>
          <form class="mini-form" method="post" action="{{ url_for('admin_merge_customers') }}">
            <div class="mini-grid">
              <div>
                <label for="admin_merge_from_customer">Merge From</label>
                <input id="admin_merge_from_customer" name="from_customer" type="text" list="admin_customer_move_options" placeholder="Old customer" required>
              </div>
              <div>
                <label for="admin_merge_to_customer">Into Customer</label>
                <input id="admin_merge_to_customer" name="to_customer" type="text" list="admin_customer_move_options" placeholder="Target customer" required>
              </div>
            </div>
            <div class="hint">This moves all jobs, field links, and invoice history for the source customer into the target customer.</div>
            <button class="button button-secondary button-full" type="submit">Merge Customer</button>
          </form>
                </details>
      </div>

      <div class="card">
        <h2>Customer / Farm Links</h2>
        <div class="map-list">
          {% for customer in admin_tree %}
          <details class="map-customer">
            <summary>
              <div class="map-customer-title">
                <div class="map-customer-name">{{ customer.customer_name }}</div>
                <div class="map-customer-meta">{{ customer.farms|length }} farms</div>
              </div>
            </summary>
            <div class="map-customer-body">
              <form class="mini-form customer-meta-form" method="post" action="{{ url_for('admin_save_customer_details') }}">
                <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                <div class="mini-grid">
                  <div>
                    <label for="customer_email_{{ loop.index }}">Billing Email</label>
                    <input id="customer_email_{{ loop.index }}" name="customer_email" type="text" value="{{ customer.customer_email }}" placeholder="billing@example.com, accounts@example.com">
                    <div class="hint">Use commas to add more than one invoice email address.</div>
                  </div>
                  <div>
                    <label for="customer_rate_{{ loop.index }}">Rate Per Ton</label>
                    <input id="customer_rate_{{ loop.index }}" name="rate_per_ton" type="number" inputmode="decimal" min="0" step="0.01" value="{{ customer.rate_per_ton_text }}" placeholder="0.00">
                  </div>
                </div>
                <div class="actions-inline">
                  <button class="button button-secondary button-small" type="submit">Save Customer Info</button>
                </div>
                <div class="customer-meta-summary">
                  <div>Email: {{ customer.customer_email or 'Not set' }}</div>
                  <div>Rate: {{ customer.rate_per_ton_label or 'Not set' }}</div>
                </div>
              </form>
                            <details class="customer-meta-form admin-action-group">
                                <summary>Move or merge farms</summary>
                            <form class="mini-form" method="post" action="{{ url_for('admin_merge_farms') }}">
                                <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                                <div class="mini-grid">
                                    <div>
                                        <label>Merge Farm From</label>
                                        <input name="from_farm_name" type="text" list="farm_move_options_{{ loop.index }}" placeholder="Duplicate or old farm" required>
                                    </div>
                                    <div>
                                        <label>Merge Farm Into</label>
                                        <input name="to_farm_name" type="text" list="farm_move_options_{{ loop.index }}" placeholder="Keep this farm" required>
                                    </div>
                                </div>
                                <button class="button button-secondary button-small" type="submit">Merge Farms</button>
                            </form>
                            <form class="mini-form" method="post" action="{{ url_for('admin_consolidate_farm_name') }}">
                                <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                                <div class="mini-grid">
                                    <div>
                                        <label>Farm Name To Consolidate</label>
                                        <input name="farm_name" type="text" list="farm_move_options_{{ loop.index }}" placeholder="Same farm, different spelling" required>
                                    </div>
                                    <div>
                                        <label>Keep This Spelling</label>
                                        <input name="canonical_farm_name" type="text" placeholder="One final farm name" required>
                                    </div>
                                </div>
                                <button class="button button-secondary button-small" type="submit">Consolidate Same-Name Farm</button>
                            </form>
                            </details>
                            {% if customer.master_records %}
                            <details class="customer-meta-form">
                                <summary>Edit Customer Spreadsheet Rows ({{ customer.master_records|length }})</summary>
                                <div class="stack">
                                {% for record in customer.master_records %}
                                <form class="mini-form customer-meta-form" method="post" action="{{ url_for('admin_save_customer_master_row') }}">
                                    <input type="hidden" name="row_number" value="{{ record._row_number }}">
                                    <div class="mini-grid">
                                        <div><label>Customer Name</label><input name="customer_name" type="text" value="{{ record.customer_name }}" required></div>
                                        <div><label>Farm Name</label><input name="farm_name" type="text" value="{{ record.farm_name }}"></div>
                                        <div><label>Email</label><input name="email" type="email" value="{{ record.email }}"></div>
                                        <div><label>Second Email</label><input name="email_2" type="email" value="{{ record.email_2 }}" placeholder="second@example.com"></div>
                                        <div><label>Address Line 1</label><input name="address_line_1" type="text" value="{{ record.address_line_1 }}"></div>
                                        <div><label>Address Line 2</label><input name="address_line_2" type="text" value="{{ record.address_line_2 }}"></div>
                                        <div><label>Town</label><input name="town" type="text" value="{{ record.town }}"></div>
                                        <div><label>Postcode</label><input name="postcode" type="text" value="{{ record.postcode }}"></div>
                                        <div><label>Rate Per Ton</label><input name="rate_per_ton" type="number" min="0" step="0.01" value="{{ record.rate_per_ton }}"></div>
                                        <div><label>VAT Rate</label><input name="vat_rate" type="number" min="0" step="0.01" value="{{ record.vat_rate }}"></div>
                                        <div><label>Active</label><select name="active"><option value="1" {% if record.active|string in ['1', 'True', 'true'] %}selected{% endif %}>Yes</option><option value="0" {% if record.active|string in ['0', 'False', 'false'] %}selected{% endif %}>No</option></select></div>
                                        <div><label>Muck Type</label><input name="muck_type" type="text" value="{{ record.muck_type }}"></div>
                                    </div>
                                    <button class="button button-secondary button-small" type="submit">Save Spreadsheet Row</button>
                                </form>
                                {% endfor %}
                                </div>
                            </details>
                            {% endif %}
              <div class="farm-list">
                {% set customer_index = loop.index %}
                <datalist id="farm_move_options_{{ customer_index }}">
                  {% for option_farm in customer.farms %}
                  {% if option_farm.farm_name %}
                  <option value="{{ option_farm.farm_name }}"></option>
                  {% endif %}
                  {% endfor %}
                </datalist>
                {% for farm in customer.farms %}
                <details class="map-farm">
                  <summary>
                    <div class="map-farm-title">
                      <div class="map-farm-name">{{ farm.farm_label }}</div>
                      <div class="map-farm-meta">{{ farm.fields|length }} fields</div>
                    </div>
                  </summary>
                  <div class="map-farm-body">
                    <div class="actions-inline">
                      <form method="post" action="{{ url_for('admin_remove_farm') }}">
                        <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                        <input type="hidden" name="farm_name" value="{{ farm.farm_name }}">
                        <button class="button button-danger button-small" type="submit">Remove Farm</button>
                      </form>
                    </div>
                    {% if farm.fields %}
                    <div class="field-tags">
                      {% for field_name in farm.fields %}
                      <div class="field-tag">
                        <span>{{ field_name }}</span>
                                                <details class="field-move-details">
                                                    <summary>Move</summary>
                                                    <form class="field-move-form" method="post" action="{{ url_for('admin_move_field') }}">
                                                        <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                                                        <input type="hidden" name="farm_name" value="{{ farm.farm_name }}">
                                                        <input type="hidden" name="field_name" value="{{ field_name }}">
                                                        <input class="field-move-input" name="to_customer" type="text" list="admin_customer_move_options" placeholder="New customer" value="{{ customer.customer_name }}">
                                                        <input class="field-move-input" name="to_farm_name" type="text" list="farm_move_options_{{ customer_index }}" placeholder="Move to farm" required>
                                                        <button class="button button-secondary button-small" type="submit">Move</button>
                                                    </form>
                                                </details>
                        <form method="post" action="{{ url_for('admin_delete_field') }}">
                          <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                          <input type="hidden" name="farm_name" value="{{ farm.farm_name }}">
                          <input type="hidden" name="field_name" value="{{ field_name }}">
                          <button class="button button-danger button-small" type="submit">Remove</button>
                        </form>
                      </div>
                      {% endfor %}
                    </div>
                    {% else %}
                    <p class="copy">No saved fields yet.</p>
                    {% endif %}
                  </div>
                </details>
                {% endfor %}
              </div>
            </div>
          </details>
          {% endfor %}
        </div>
      </div>

      <div class="card stack">
        <div>
          <h2>Muck Types</h2>
          <p class="copy">Manage the saved muck type list used in job entry.</p>
        </div>

        <div>
          <form class="mini-form" method="post" action="{{ url_for('admin_add_muck_type') }}">
            <div>
              <label for="admin_muck_type_name">New Muck Type</label>
              <input id="admin_muck_type_name" name="muck_type" type="text" placeholder="Enter muck type name" required>
            </div>
            <button class="button button-secondary button-full" type="submit">Add Muck Type</button>
          </form>
        </div>

        <div class="field-tags">
          {% for muck_type in muck_types %}
          <div class="field-tag">
            <span>{{ muck_type }}</span>
            <form method="post" action="{{ url_for('admin_delete_muck_type') }}">
              <input type="hidden" name="muck_type" value="{{ muck_type }}">
              <button class="button button-danger button-small" type="submit">Remove</button>
            </form>
          </div>
          {% else %}
          <p class="copy">No saved muck types yet.</p>
          {% endfor %}
        </div>
      </div>
    </div>
  </div>
  <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
    <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
  </nav>
  <script>
    const adminCustomers = {{ customers_json|safe }};
    const adminCustomerInput = document.getElementById("admin_customer_select");
    const adminCustomerSuggestions = document.getElementById("admin_customer_suggestions");

    function closeAdminSuggestions() {
      adminCustomerSuggestions.classList.remove("is-open");
      adminCustomerSuggestions.innerHTML = "";
    }

    function filterAdminOptions(options, typedValue) {
      const typed = String(typedValue || "").trim().toLowerCase();
      if (!typed) {
        return options.slice();
      }
      const starts = [];
      const contains = [];
      for (const option of options) {
        const lower = option.toLowerCase();
        if (lower.startsWith(typed)) {
          starts.push(option);
        } else if (lower.includes(typed)) {
          contains.push(option);
        }
      }
      return starts.concat(contains);
    }

    function openAdminSuggestions(options) {
      adminCustomerSuggestions.innerHTML = "";
      if (!options.length) {
        closeAdminSuggestions();
        return;
      }
      for (const option of options) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "suggestion-item";
        if (String(adminCustomerInput.value || "").trim() && option.toLowerCase().startsWith(String(adminCustomerInput.value || "").trim().toLowerCase())) {
          button.classList.add("is-match");
        }
        button.textContent = option;
        button.addEventListener("click", function () {
          adminCustomerInput.value = option;
          closeAdminSuggestions();
        });
        adminCustomerSuggestions.appendChild(button);
      }
      adminCustomerSuggestions.classList.add("is-open");
    }

    function showAdminCustomerSuggestions() {
      openAdminSuggestions(filterAdminOptions(adminCustomers, adminCustomerInput.value));
    }

    adminCustomerInput.addEventListener("input", showAdminCustomerSuggestions);
    adminCustomerInput.addEventListener("keyup", showAdminCustomerSuggestions);
    adminCustomerInput.addEventListener("focus", showAdminCustomerSuggestions);

    document.addEventListener("click", function (event) {
      if (!event.target.closest(".suggestion-field")) {
        closeAdminSuggestions();
      }
    });
  </script>
</body>
</html>
"""

INVOICE_HISTORY_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="#334d38">
<title>Invoice History</title>
  <style>
    @import url('/static/af_brand.css');
    .page {
      max-width: 1180px;
      margin: 0 auto;
      padding: 18px 14px 28px;
    }
    .top-links {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
    }
    .card {
      background: var(--panel);
      border: 1px solid rgba(82, 69, 42, 0.12);
      border-radius: 24px;
      box-shadow: var(--shadow);
      padding: 24px;
    }
    .button {
      min-height: 48px;
      padding: 0 18px;
      border-radius: 14px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-weight: 700;
      background: #edf5ee;
      color: var(--green);
      border: 1px solid rgba(36, 68, 47, 0.08);
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .button:hover { transform: translateY(-1px); }
    .button-small {
      min-height: 40px;
      padding: 0 14px;
      font-size: 14px;
    }
    .copy {
      margin: 0 0 18px;
      color: var(--muted);
      line-height: 1.45;
    }
    .status {
      margin-bottom: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      font-size: 15px;
      border: 1px solid transparent;
      line-height: 1.45;
    }
    .status.ok {
      background: rgba(60,95,70,0.1);
      border-color: rgba(60,95,70,0.16);
      color: #284332;
    }
    .status.error {
      background: rgba(139,71,56,0.1);
      border-color: rgba(139,71,56,0.16);
      color: #6d3124;
    }
    .table-wrap {
      overflow-x: auto;
      border: 1px solid rgba(82, 69, 42, 0.1);
      border-radius: 18px;
    }
    table {
      width: 100%;
      min-width: 860px;
      border-collapse: collapse;
      background: rgba(255,255,255,0.68);
      font-family: var(--font-main);
    }
    th, td {
      padding: 14px 16px;
      text-align: left;
      border-bottom: 1px solid rgba(82, 69, 42, 0.1);
    }
    th {
      background: rgba(240,232,216,0.85);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #5d5b48;
    }
    tr:last-child td { border-bottom: none; }
    .actions-inline {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .empty {
      padding: 26px;
      border-radius: 18px;
      border: 1px dashed var(--line);
      background: rgba(255,255,255,0.48);
      color: var(--muted);
      text-align: center;
    }
    .download-notice {
      position: fixed;
      left: 50%;
      bottom: calc(82px + env(safe-area-inset-bottom));
      transform: translateX(-50%);
      z-index: 1000;
      width: min(92vw, 420px);
      padding: 14px 16px;
      border-radius: 18px;
      background: rgba(39, 45, 33, 0.94);
      color: #f5efe2;
      text-align: center;
      box-shadow: 0 18px 44px rgba(60, 49, 25, 0.28);
    }
    .download-notice[hidden] { display: none; }
  </style>
</head>
<body>
  <div class="page">
    <div class="top-links">
      <a class="button" href="{{ url_for('invoice_home') }}">Back To Invoices</a>
      <a class="button" href="{{ url_for('home') }}">Back To Jobs</a>
    </div>

    {% if status_msg %}
    <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
    {% endif %}

    <div class="card">
      <h1>Invoice History</h1>
      <p class="copy">View, edit or save any previous invoice. Invoices are listed newest first.</p>
      {% if history_rows %}
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Reference</th>
              <th>Type</th>
              <th>Customer</th>
              <th>Farm</th>
              <th>Period</th>
              <th>Jobs</th>
              <th>Total</th>
              <th>Created</th>
              <th>Note</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {% for row in history_rows %}
            <tr>
              <td>{{ row.reference_label }}</td>
              <td>{{ row.type_label }}</td>
              <td>{{ row.customer }}</td>
              <td>{{ row.farm_name }}</td>
              <td>{{ row.period_label }}</td>
              <td>{{ row.job_count }}</td>
              <td>{{ row.grand_total_label }}</td>
              <td>{{ row.created_label }}</td>
              <td>{{ row.note or '--' }}</td>
              <td>
                {% if row.can_edit %}
                <div class="actions-inline">
                  <a class="button button-small" href="{{ url_for('invoice_history_edit', ledger_index=row.ledger_index) }}">Edit / Re-send</a>
                  <a class="button button-small" href="{{ url_for('invoice_history_show_pdf', ledger_index=row.ledger_index) }}" target="_blank" rel="noopener">View PDF</a>
                  <a class="button button-small" data-invoice-download="1" data-fallback-name="invoice.pdf" href="{{ url_for('invoice_history_download_pdf', ledger_index=row.ledger_index) }}">Save PDF</a>
                  <a class="button button-small" data-invoice-download="1" data-fallback-name="invoice.xlsx" href="{{ url_for('invoice_history_download_xlsx', ledger_index=row.ledger_index) }}">Save Excel</a>
                </div>
                {% else %}
                --
                {% endif %}
              </td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      {% else %}
      <div class="empty">No invoice history yet.</div>
      {% endif %}
    </div>
  </div>

  <div id="download_notice" class="download-notice" hidden>Preparing file...</div>

  <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
    <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
  </nav>
  <script>
    const downloadNotice = document.getElementById("download_notice");
    let downloadNoticeTimer = null;

    function showDownloadNotice(message) {
      downloadNotice.textContent = message;
      downloadNotice.hidden = false;
      if (downloadNoticeTimer) window.clearTimeout(downloadNoticeTimer);
      downloadNoticeTimer = window.setTimeout(function () {
        downloadNotice.hidden = true;
      }, 3200);
    }

    function downloadFilename(response, fallbackName) {
      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename="?([^";]+)"?/i);
      return match && match[1] ? match[1].trim() : fallbackName;
    }

    async function openInvoiceDownload(link) {
      const targetUrl = new URL(link.href, window.location.href);
      targetUrl.searchParams.set("_dl", Date.now().toString());
      showDownloadNotice("Preparing file...");
      const response = await fetch(targetUrl.toString(), {credentials: "same-origin"});
      const contentType = response.headers.get("Content-Type") || "";
      if (!response.ok) throw new Error("Download failed");
      if (response.redirected && contentType.includes("text/html")) {
        window.location.href = response.url;
        return;
      }

      const blob = await response.blob();
      const filename = downloadFilename(response, link.dataset.fallbackName || "invoice");
      if (navigator.share && window.File) {
        try {
          const file = new File([blob], filename, {type: blob.type || contentType || "application/octet-stream"});
          if (!navigator.canShare || navigator.canShare({files: [file]})) {
            showDownloadNotice("Choose Excel, Files or another app");
            await navigator.share({files: [file], title: filename});
            return;
          }
        } catch (error) {
          if (error && error.name === "AbortError") {
            showDownloadNotice("File closed");
            return;
          }
        }
      }

      const objectUrl = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = filename;
      anchor.rel = "noopener";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(function () { URL.revokeObjectURL(objectUrl); }, 1000);
      showDownloadNotice("File downloaded");
    }

    for (const link of document.querySelectorAll("[data-invoice-download]")) {
      link.addEventListener("click", async function (event) {
        event.preventDefault();
        try {
          await openInvoiceDownload(link);
        } catch (error) {
          showDownloadNotice("File could not be opened");
        }
      });
    }
  </script>
</body>
</html>
"""

SETTINGS_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <meta name="theme-color" content="#334d38">
<title>Settings</title>
  <style>
    @import url('/static/af_brand.css');
    :root {
      --bg: #edf1ea;
      --panel: rgba(255, 255, 255, 0.94);
      --ink: #1d2b1f;
      --muted: #5b6b60;
      --line: rgba(36, 68, 47, 0.12);
      --green: #24442f;
      --gold: #c79e4f;
      --red: #8b4738;
      --shadow: 0 18px 36px rgba(29, 48, 35, 0.08);
      --font-main: Inter, "Segoe UI", system-ui, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font-main);
      color: var(--ink);
      background: linear-gradient(180deg, #edf1ea 0%, #e7ece3 100%);
    }
    .page {
      max-width: 980px;
      margin: 0 auto;
      padding: 18px 14px 28px;
    }
    .top-links {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
    }
    .card {
      background: var(--panel);
      border: 1px solid rgba(82, 69, 42, 0.12);
      border-radius: 24px;
      box-shadow: var(--shadow);
      padding: 24px;
    }
    .card + .card {
      margin-top: 18px;
    }
    .button {
      min-height: 48px;
      padding: 0 18px;
      border-radius: 999px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
      background: rgba(60,95,70,0.1);
      color: var(--green);
      border: 1px solid rgba(60,95,70,0.12);
      font-family: var(--font-main);
    }
    .button-primary {
      background: linear-gradient(135deg, var(--green), #2c5641);
      color: #fff;
      box-shadow: 0 12px 22px rgba(36, 68, 47, 0.16);
      border: none;
    }
    .status {
      margin-bottom: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      font-size: 15px;
      border: 1px solid transparent;
      line-height: 1.45;
    }
    .status.ok {
      background: rgba(60,95,70,0.1);
      border-color: rgba(60,95,70,0.16);
      color: #284332;
    }
    .status.error {
      background: rgba(139,71,56,0.1);
      border-color: rgba(139,71,56,0.16);
      color: #6d3124;
    }
    .copy {
      margin: 0 0 18px;
      color: var(--muted);
      line-height: 1.45;
    }
    .hint {
      font-size: 13px;
      color: var(--muted);
      margin-top: 8px;
    }
    .form-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
    }
    .field {
      display: grid;
      gap: 7px;
    }
    .field.full {
      grid-column: 1 / -1;
    }
    label {
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-weight: bold;
      color: #5d5b48;
    }
    input, select, textarea {
      width: 100%;
      min-height: 56px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: #fffdfa;
      color: var(--ink);
      padding: 12px 14px;
      outline: none;
      font-size: 16px;
      font-family: var(--font-main);
    }
    textarea {
      min-height: 140px;
      resize: vertical;
      padding-top: 14px;
    }
    input:focus, select:focus, textarea:focus {
      border-color: var(--gold);
      box-shadow: 0 0 0 4px rgba(186,148,80,0.14);
    }
    .actions {
      margin-top: 18px;
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }
    @media (max-width: 760px) {
      .form-grid {
        grid-template-columns: 1fr;
      }
      .button {
        width: 100%;
      }
    }
  </style>
</head>
<body>
  <div class="page">
    <div class="top-links">
      <a class="button" href="{{ url_for('home') }}">Back To Jobs</a>
    </div>

    {% if status_msg %}
    <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
    {% endif %}

    <div class="card">
      <h1>Settings</h1>
      <p class="copy">All settings are stored in <strong>settings.xlsx</strong> in the app folder. Use its tabs to edit staff names, companies, customers, email settings and recipients, invoice settings, farms, fields and muck types. Save and close Excel, then refresh the app. Changes saved here update the same workbook.</p>
      <p class="copy">Change the default invoice wording and payment terms used when you open the invoice page.</p>
      <form method="post" action="{{ url_for('settings_save') }}">
        <div class="form-grid">
                    <div class="field full">
                        <h2>Email Sending Account</h2>
                        <div class="hint">These settings control the SMTP account used for invoice and summary emails. Leave the password blank to keep the current password.</div>
                    </div>
                    <div class="field">
                        <label for="settings_smtp_host">SMTP Host</label>
                        <input id="settings_smtp_host" name="smtp_host" type="text" value="{{ email_config.smtp_host }}">
                    </div>
                    <div class="field">
                        <label for="settings_smtp_port">SMTP Port</label>
                        <input id="settings_smtp_port" name="smtp_port" type="number" min="1" max="65535" value="{{ email_config.smtp_port }}">
                    </div>
                    <div class="field">
                        <label for="settings_smtp_username">SMTP Username</label>
                        <input id="settings_smtp_username" name="smtp_username" type="email" value="{{ email_config.smtp_username }}">
                    </div>
                    <div class="field">
                        <label for="settings_smtp_password">SMTP Password</label>
                        <input id="settings_smtp_password" name="smtp_password" type="password" placeholder="Leave blank to keep current password" autocomplete="new-password">
                    </div>
                    <div class="field">
                        <label for="settings_smtp_from_email">SMTP Account From</label>
                        <input id="settings_smtp_from_email" name="smtp_from_email" type="email" value="{{ email_config.from_email }}">
                    </div>
                    <div class="field">
                        <label for="settings_smtp_use_tls">Use TLS</label>
                        <select id="settings_smtp_use_tls" name="smtp_use_tls">
                            <option value="1" {% if email_config.use_tls %}selected{% endif %}>Yes</option>
                            <option value="0" {% if not email_config.use_tls %}selected{% endif %}>No</option>
                        </select>
                    </div>
          <div class="field">
            <label for="settings_invoice_from_email">Invoice From Email</label>
            <input id="settings_invoice_from_email" name="invoice_from_email" type="email" value="{{ settings.invoice_from_email }}">
          </div>
          <div class="field">
            <label for="settings_invoice_default_payment_terms_days">Default Payment Terms</label>
            <select id="settings_invoice_default_payment_terms_days" name="invoice_default_payment_terms_days">
              {% for option in invoice_payment_terms_options %}
              <option value="{{ option }}" {% if settings.invoice_default_payment_terms_days == option %}selected{% endif %}>{{ option }} days</option>
              {% endfor %}
            </select>
          </div>
          <div class="field">
            <label for="settings_invoice_subject_template">Invoice Subject Template</label>
            <input id="settings_invoice_subject_template" name="invoice_subject_template" type="text" value="{{ settings.invoice_subject_template }}">
            <div class="hint">Placeholders: `{invoice_number}` `{customer}` `{farm}` `{scope}` `{greeting}`</div>
          </div>
          <div class="field full">
            <label for="settings_invoice_customer_message_template">Customer Message Template</label>
            <textarea id="settings_invoice_customer_message_template" name="invoice_customer_message_template">{{ settings.invoice_customer_message_template }}</textarea>
          </div>
          <div class="field full">
            <label for="settings_invoice_accounts_message_template">Accounts Message Template</label>
            <textarea id="settings_invoice_accounts_message_template" name="invoice_accounts_message_template">{{ settings.invoice_accounts_message_template }}</textarea>
          </div>
        </div>
        <div class="hint">The greeting changes automatically to `Good Morning` or `Good Afternoon` when the email is sent.</div>
        <div class="actions">
          <button class="button button-primary" type="submit">Save Settings</button>
        </div>
      </form>
    </div>

    <div class="card">
      <h2>App Update</h2>
      <p class="copy">Install the latest version of the app from GitHub.</p>
      <form method="post" action="{{ url_for('update_app') }}" id="update_app_form">
        <button class="button button-primary" type="submit">Update App</button>
      </form>
      <div class="hint">Current version: {{ app_version }}</div>
    </div>
  </div>
  <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
    <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
  </nav>
</body>
</html>
"""


def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(ISSUE_PHOTOS_DIR, exist_ok=True)
    normalize_customer_case_storage()


def read_json_file(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return default


def write_json_atomic(path, payload):
    parent = os.path.dirname(path) or "."
    os.makedirs(parent, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix=os.path.basename(path) + ".", suffix=".tmp", dir=parent)
    with os.fdopen(fd, "w") as handle:
        json.dump(payload, handle, indent=2)
    os.replace(temp_path, path)


def append_json_line(path, payload):
    with open(path, "a") as handle:
        handle.write(json.dumps(payload))
        handle.write("\n")


def write_json_lines_atomic(path, rows):
    parent = os.path.dirname(path) or "."
    os.makedirs(parent, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix=os.path.basename(path) + ".", suffix=".tmp", dir=parent)
    with os.fdopen(fd, "w") as handle:
        for row in rows:
            if isinstance(row, dict):
                handle.write(json.dumps(row))
                handle.write("\n")
    os.replace(temp_path, path)


def sanitize_issue_photo_filename(name):
    base = os.path.basename(str(name or "").strip())
    return base if base else ""


def normalize_issue_photo_names(values):
    names = []
    seen = set()
    if not isinstance(values, list):
        return names
    for value in values:
        filename = sanitize_issue_photo_filename(value)
        if not filename:
            continue
        lowered = filename.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        names.append(filename)
    return names


def save_issue_photo_uploads(job_id, uploads):
    ensure_data_dir()
    saved = []
    uploads = uploads if isinstance(uploads, list) else []
    timestamp = int(time.time() * 1000)
    index = 0
    for upload in uploads:
        filename = sanitize_issue_photo_filename(getattr(upload, "filename", ""))
        if not filename:
            continue
        extension = os.path.splitext(filename)[1].lower()
        mimetype = str(getattr(upload, "mimetype", "") or "").strip().lower()
        if extension not in ALLOWED_ISSUE_PHOTO_EXTENSIONS and not mimetype.startswith("image/"):
            continue
        if extension not in ALLOWED_ISSUE_PHOTO_EXTENSIONS:
            extension = ".jpg"
        safe_name = "%s_%s_%s%s" % (int(job_id or 0), timestamp, index, extension)
        output_path = os.path.join(ISSUE_PHOTOS_DIR, safe_name)
        try:
            upload.save(output_path)
        except Exception:
            index += 1
            continue
        saved.append(safe_name)
        index += 1
    return saved


def delete_issue_photo_files(values):
    for filename in normalize_issue_photo_names(values):
        file_path = os.path.join(ISSUE_PHOTOS_DIR, filename)
        if os.path.isfile(file_path):
            try:
                os.remove(file_path)
            except OSError:
                pass


def issue_photo_items(values):
    items = []
    for filename in normalize_issue_photo_names(values):
        file_path = os.path.join(ISSUE_PHOTOS_DIR, filename)
        if not os.path.isfile(file_path):
            continue
        items.append({
            "name": filename,
            "url": url_for("job_issue_photo", filename=filename),
        })
    return items


def settings_workbook_exists():
    return os.path.isfile(SETTINGS_WORKBOOK_PATH)


def settings_key_values(sheet):
    return {row["setting"]: row.get("value", "") for row in Workbook(SETTINGS_WORKBOOK_PATH).records(sheet) if row.get("setting")}


def update_settings_key_values(sheet, values):
    def update(rows):
        result = [list(row) for row in rows]
        seen = set()
        for row in result[1:]:
            if row and row[0] in values:
                while len(row) < 2:
                    row.append("")
                row[1] = str(values[row[0]])
                seen.add(row[0])
        result.extend([[key, str(value)] for key, value in values.items() if key not in seen])
        return result
    Workbook(SETTINGS_WORKBOOK_PATH).update(sheet, update)


def settings_names(sheet, default):
    if not settings_workbook_exists():
        return list(default)
    return list(dict.fromkeys(row.get("name", "").strip() for row in Workbook(SETTINGS_WORKBOOK_PATH).records(sheet) if settings_row_active(row) and row.get("name", "").strip()))


def timesheet_staff_records():
    if not settings_workbook_exists():
        return [{"name": "Owen Curl", "pin": "", "active": "1"}]
    records = []
    seen = set()
    for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Staff"):
        name = clean_name(row.get("name"))
        key = name.casefold()
        if not name or key in seen or not settings_row_active(row):
            continue
        seen.add(key)
        records.append({"name": name, "pin": str(row.get("pin", "") or "").strip(), "active": "1"})
    return records


def save_timesheet_staff_pin(staff_name, pin):
    ensure_settings_workbook()
    matched = {"value": False}

    def update(rows):
        result = [list(row) for row in rows]
        if not result:
            raise ValueError("The Staff tab is missing from settings.xlsx.")
        headers = [str(value).strip().lower() for value in result[0]]
        if "name" not in headers:
            raise ValueError("The Staff tab must contain a name column.")
        if "pin" not in headers:
            result[0].append("pin")
            headers.append("pin")
        name_index = headers.index("name")
        pin_index = headers.index("pin")
        for row in result[1:]:
            row.extend([""] * max(0, len(headers) - len(row)))
            if clean_name(row[name_index]).casefold() == staff_name.casefold():
                row[pin_index] = pin
                matched["value"] = True
                break
        if not matched["value"]:
            raise ValueError("That staff member is no longer available.")
        return result

    Workbook(SETTINGS_WORKBOOK_PATH).update("Staff", update)


def current_timesheet_staff():
    selected = clean_name(session.get("timesheet_staff"))
    for row in timesheet_staff_records():
        if row["name"].casefold() == selected.casefold():
            return row["name"]
    session.pop("timesheet_staff", None)
    return ""


def save_settings_names(sheet, names):
    def update(rows):
        inactive = [row for row in rows[1:] if len(row) > 1 and not settings_row_active({"active": row[1]}) and row[0] not in names]
        return [["name", "active"]] + [[name, "1"] for name in names] + inactive
    Workbook(SETTINGS_WORKBOOK_PATH).update(sheet, update)


def sync_workbook_customers(names):
    def update(rows):
        headers = rows[0]
        name_index = headers.index("customer_name")
        active_index = headers.index("active") if "active" in headers else None
        requested = {name.casefold(): name for name in names}
        result = [list(headers)]
        seen = set()
        for previous in rows[1:]:
            row = list(previous) + [""] * max(0, len(headers) - len(previous))
            key = row[name_index].casefold()
            if key in requested:
                row[name_index] = requested[key]
                result.append(row)
                seen.add(key)
            elif active_index is not None and not settings_row_active({"active": row[active_index]}):
                result.append(row)
        for key, name in requested.items():
            if key not in seen:
                row = [""] * len(headers)
                row[name_index] = name
                if active_index is not None:
                    row[active_index] = "1"
                result.append(row)
        return result
    Workbook(SETTINGS_WORKBOOK_PATH).update("Customers", update)


def save_workbook_customer(name, email, rate, row_number, field_values):
    rate = str(rate or "").strip()
    if rate:
        try:
            rate = ("%.2f" % float(rate)).rstrip("0").rstrip(".")
        except ValueError:
            return False, "Rate per ton must be a number"
    def update(rows):
        result = [list(row) for row in rows]
        headers = result[0]
        name_index = headers.index("customer_name")
        matches = [index for index, row in enumerate(result[1:], 1) if (index + 1 == int(row_number) if row_number is not None else len(row) > name_index and clean_name(row[name_index]).casefold() == name.casefold())]
        if not matches:
            if row_number is not None:
                raise ValueError("Customer row no longer exists. Reload the page.")
            result.append([""] * len(headers))
            matches = [len(result) - 1]
            for key, value in {"customer_name": name, "active": "1", "vat_rate": "20"}.items():
                if key in headers:
                    result[-1][headers.index(key)] = value
        values = {"email": str(email or "").strip(), "rate_per_ton": rate}
        values.update(field_values or {})
        for index in matches:
            result[index].extend([""] * max(0, len(headers) - len(result[index])))
            for key, value in values.items():
                if key in headers:
                    result[index][headers.index(key)] = str(value or "")
        return result
    try:
        Workbook(SETTINGS_WORKBOOK_PATH).update("Customers", update)
        return True, "%s updated" % name
    except (OSError, ValueError, zipfile.BadZipFile):
        return False, "Could not update Customers in settings.xlsx. Close Excel and try again."


def ensure_settings_workbook():
    if settings_workbook_exists():
        ensure_customer_second_email_column()
        ensure_timesheet_recipient_column()
        ensure_staff_pin_column()
        ensure_settings_name_sheet("Machinery")
        return
    # Read the legacy files only for the initial migration, before creating the workbook.
    config = load_email_config()
    invoice_settings = load_app_settings()
    raw_customers = load_customer_master_raw_rows()
    header_index = next((index for index, row in enumerate(raw_customers) if "customer_name" in [clean_name(cell).lower() for cell in row]), None)
    if header_index is None:
        customers = [list(CUSTOMER_MASTER_HEADERS)]
    else:
        customers = [[clean_name(cell).lower() for cell in raw_customers[header_index]]] + raw_customers[header_index + 1:]
    if "email_2" not in customers[0]:
        customers[0].append("email_2")
        customers = [customers[0]] + [list(row) + [""] for row in customers[1:]]
    name_index = customers[0].index("customer_name")
    existing_names = {row[name_index].casefold() for row in customers[1:] if len(row) > name_index}
    for name in load_customers():
        if name.casefold() not in existing_names:
            row = [""] * len(customers[0])
            row[name_index] = name
            if "active" in customers[0]:
                row[customers[0].index("active")] = "1"
            customers.append(row)
            existing_names.add(name.casefold())
    recipient_options = load_email_recipient_options()
    recipients = [["name", "email", "active", "summary", "invoice_option", "timesheet"]]
    known_emails = set()
    if os.path.exists(EMAIL_SETTINGS_CSV_PATH):
        with open(EMAIL_SETTINGS_CSV_PATH, newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                if row.get("record_type", "").strip().lower() == "recipient" and row.get("email", "").strip():
                    email = row["email"].strip()
                    summary = "1" if email in config.get("to_emails", []) or not settings_row_active(row) else "0"
                    recipients.append([row.get("name", ""), email, row.get("active", "1"), summary, "1", summary])
                    known_emails.add(email.lower())
    for email in config.get("to_emails", []):
        if email.lower() not in known_emails:
            option = next((row for row in recipient_options if row["email"].lower() == email.lower()), None)
            recipients.append([option["name"] if option else "", email, "1", "1", "1" if option else "0", "1"])
            known_emails.add(email.lower())
    summary_order = {email.lower(): index for index, email in enumerate(config.get("to_emails", []))}
    recipients[1:] = sorted(recipients[1:], key=lambda row: summary_order.get(row[1].lower(), len(summary_order)))
    staff = ["Owen Curl"]
    timesheets_path = os.path.join(DATA_DIR, "timesheets.sqlite3")
    if os.path.exists(timesheets_path):
        connection = sqlite3.connect(timesheets_path)
        try:
            for row in connection.execute("SELECT DISTINCT name FROM timesheets"):
                if row[0] and row[0] not in staff:
                    staff.append(row[0])
        finally:
            connection.close()
    field_map = load_field_map()
    Workbook(SETTINGS_WORKBOOK_PATH).create({
        "Customers": customers,
        "Staff": [["name", "active", "pin"]] + [[name, "1", ""] for name in staff],
        "Companies": [["name", "active"]] + [[name, "1"] for name in TIMESHEET_COMPANIES],
        "Email Settings": [["setting", "value"]] + [[key, value] for key, value in config.items() if key != "to_emails"],
        "Email Recipients": recipients,
        "Invoice Settings": [["setting", "value"]] + [[key, value] for key, value in invoice_settings.items()],
        "Farms": [["name", "active"]] + [[name, "1"] for name in load_farms()],
        "Fields": [["customer_name", "farm_name", "field_name", "active"]] + [[customer, farm, field, "1"] for customer, farms in field_map.items() for farm, fields in farms.items() for field in fields],
        "Muck Types": [["name", "active"]] + [[name, "1"] for name in load_muck_types()],
        "Straw Crops": [["name", "active"]] + [[name, "1"] for name in ["Wheat", "Barley", "Spring Barley", "Oats", "Hay"]],
        "Machinery": [["name", "active"]],
        "Read Me": [["topic", "instructions"],
                    ["Editing", "Edit values under the existing headings. Save and close Excel, then refresh the app. Keep tab names and column headings unchanged."],
                    ["Staff", "Each active staff name appears as a Timesheet button. A person with a blank PIN sets one the first time they choose their name. active 1 = available, 0 = hidden."],
                    ["Companies", "Company names used by the timesheet dropdown."],
                    ["Machinery", "Machinery names offered when adding maintenance records. New names entered in the app are added here."],
                    ["Email Recipients", "summary 1 = receives scheduled summaries; invoice_option 1 = available for invoice copies; timesheet 1 = receives completed monthly timesheets. active 0 disables the row."],
                    ["Email Settings", "enabled controls weekly email; monthly_enabled controls month-end email. 1 = enabled, 0 = disabled. send_weekday: Monday 0 through Sunday 6. Hours use 0–23."],
                    ["Customers", "Customer and farm addresses, email, price per ton and VAT. One row per customer/farm. active 0 hides the row."],
                    ["Saved work", "Timesheets, jobs, invoices and straw records remain in their existing data stores."],
                    ["Configuration", "settings.xlsx is the live configuration source. Legacy files are only read if an older installation still has them during first-time migration."]],
    })


def ensure_timesheet_recipient_column():
    try:
        rows = Workbook(SETTINGS_WORKBOOK_PATH).rows("Email Recipients")
    except Exception:
        return
    if not rows or "timesheet" in [str(value).strip().lower() for value in rows[0]]:
        return
    updated = [list(rows[0]) + ["timesheet"]]
    headers = [str(value).strip().lower() for value in rows[0]]
    summary_index = headers.index("summary") if "summary" in headers else None
    for row in rows[1:]:
        value = row[summary_index] if summary_index is not None and summary_index < len(row) else "0"
        updated.append(list(row) + [value])
    Workbook(SETTINGS_WORKBOOK_PATH).set_rows("Email Recipients", updated)


def ensure_customer_second_email_column():
    try:
        rows = Workbook(SETTINGS_WORKBOOK_PATH).rows("Customers")
    except Exception:
        return
    headers = [str(value).strip().lower() for value in rows[0]] if rows else []
    if "email_2" in headers:
        return
    updated = [list(rows[0]) + ["email_2"]]
    for row in rows[1:]:
        updated.append(list(row) + [""])
    Workbook(SETTINGS_WORKBOOK_PATH).set_rows("Customers", updated)


def ensure_staff_pin_column():
    try:
        rows = Workbook(SETTINGS_WORKBOOK_PATH).rows("Staff")
    except Exception:
        return
    headers = [str(value).strip().lower() for value in rows[0]] if rows else []
    if "pin" in headers:
        return
    updated = [list(rows[0]) + ["pin"]]
    for row in rows[1:]:
        updated.append(list(row) + [""])
    Workbook(SETTINGS_WORKBOOK_PATH).set_rows("Staff", updated)


def ensure_settings_name_sheet(sheet):
    try:
        Workbook(SETTINGS_WORKBOOK_PATH).rows(sheet)
        return
    except Exception:
        pass
    Workbook(SETTINGS_WORKBOOK_PATH).add_sheet(sheet, [["name", "active"]])


def load_app_settings():
    data = settings_key_values("Invoice Settings") if settings_workbook_exists() else read_json_file(APP_SETTINGS_PATH, {})
    merged = dict(DEFAULT_APP_SETTINGS)
    if isinstance(data, dict):
        merged.update(data)
    merged["invoice_subject_template"] = str(merged.get("invoice_subject_template", DEFAULT_INVOICE_SUBJECT_TEMPLATE) or DEFAULT_INVOICE_SUBJECT_TEMPLATE).strip()
    merged["invoice_customer_message_template"] = str(merged.get("invoice_customer_message_template", DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE) or DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE).strip()
    merged["invoice_accounts_message_template"] = str(merged.get("invoice_accounts_message_template", DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE) or DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE).strip()
    merged["invoice_from_email"] = str(merged.get("invoice_from_email", "andrew@afarrellcontracting.co.uk") or "andrew@afarrellcontracting.co.uk").strip()
    payment_terms = str(merged.get("invoice_default_payment_terms_days", "14") or "14").strip()
    if payment_terms not in INVOICE_PAYMENT_TERMS_OPTIONS:
        payment_terms = "14"
    merged["invoice_default_payment_terms_days"] = payment_terms
    return merged


def save_app_settings(settings):
    current = load_app_settings()
    current.update(settings if isinstance(settings, dict) else {})
    if settings_workbook_exists():
        update_settings_key_values("Invoice Settings", current)
    else:
        write_json_atomic(APP_SETTINGS_PATH, current)


def invoice_from_email(config):
    settings = load_app_settings()
    return str(settings.get("invoice_from_email", "") or config.get("from_email", "") or "").strip()


def parse_csv_decimal(value):
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return ("%.2f" % float(text)).rstrip("0").rstrip(".")
    except Exception:
        return text


def worksheet_ref_col_index(ref):
    letters = []
    for char in str(ref or ""):
        if char.isalpha():
            letters.append(char.upper())
        else:
            break
    value = 0
    for char in letters:
        value = (value * 26) + (ord(char) - 64)
    return value


def first_child_by_local_name(parent, local_name):
    if parent is None:
        return None
    for child in list(parent):
        if str(child.tag).rsplit("}", 1)[-1] == local_name:
            return child
    return None


def children_by_local_name(parent, local_name):
    if parent is None:
        return []
    matches = []
    for child in list(parent):
        if str(child.tag).rsplit("}", 1)[-1] == local_name:
            matches.append(child)
    return matches


def xlsx_shared_strings(archive):
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except Exception:
        return []

    values = []
    for item in children_by_local_name(root, "si"):
        text_parts = []
        text_node = first_child_by_local_name(item, "t")
        if text_node is not None and text_node.text is not None:
            text_parts.append(text_node.text)
        for run in children_by_local_name(item, "r"):
            run_text = first_child_by_local_name(run, "t")
            if run_text is not None and run_text.text is not None:
                text_parts.append(run_text.text)
        values.append("".join(text_parts))
    return values


def xlsx_first_sheet_rows(path):
    try:
        with zipfile.ZipFile(path, "r") as archive:
            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relationships = {}
            for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                relationships[rel.attrib.get("Id")] = rel.attrib.get("Target", "")

            sheets_node = first_child_by_local_name(workbook_root, "sheets")
            first_sheet = first_child_by_local_name(sheets_node, "sheet")
            if first_sheet is None:
                return []
            rel_id = (
                first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                or first_sheet.attrib.get("{%s}id" % STRICT_REL_NS, "")
            )
            target = relationships.get(rel_id, "worksheets/sheet1.xml")
            if not target.startswith("xl/"):
                target = "xl/%s" % target.lstrip("/")

            sheet_root = ET.fromstring(archive.read(target))
            shared_strings = xlsx_shared_strings(archive)
    except Exception:
        return []

    rows = []
    sheet_data = first_child_by_local_name(sheet_root, "sheetData")
    if sheet_data is None:
        return rows

    for row_node in children_by_local_name(sheet_data, "row"):
        row_values = []
        for cell in children_by_local_name(row_node, "c"):
            ref = cell.attrib.get("r", "")
            col_index = worksheet_ref_col_index(ref)
            while len(row_values) < max(col_index - 1, 0):
                row_values.append("")

            value = ""
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                inline_node = first_child_by_local_name(first_child_by_local_name(cell, "is"), "t")
                if inline_node is not None and inline_node.text is not None:
                    value = inline_node.text
            else:
                value_node = first_child_by_local_name(cell, "v")
                if value_node is not None and value_node.text is not None:
                    value = value_node.text
                    if cell_type == "s":
                        try:
                            value = shared_strings[int(value)]
                        except Exception:
                            pass
            row_values.append(value)
        rows.append(row_values)
    return rows


def load_customer_master_raw_rows():
    if settings_workbook_exists():
        return Workbook(SETTINGS_WORKBOOK_PATH).rows("Customers")
    if os.path.exists(CUSTOMER_MASTER_XLSX_PATH):
        rows = xlsx_first_sheet_rows(CUSTOMER_MASTER_XLSX_PATH)
        if rows:
            return rows
    return []


def load_customer_master_rows():
    raw_rows = load_customer_master_raw_rows()
    if not raw_rows:
        return []

    rows = []
    header_index = None
    headers = []
    i = 0
    while i < len(raw_rows):
        candidate = [clean_name(cell).lower() for cell in raw_rows[i]]
        if "customer_name" in candidate:
            header_index = i
            headers = candidate
            break
        i += 1

    if header_index is None:
        return []

    j = header_index + 1
    while j < len(raw_rows):
        raw_row = raw_rows[j]
        if not any(clean_name(cell) for cell in raw_row):
            j += 1
            continue

        row_dict = {}
        k = 0
        while k < len(headers):
            header = headers[k]
            if header:
                row_dict[header] = raw_row[k] if k < len(raw_row) else ""
            k += 1

        customer_name = clean_name(row_dict.get("customer_name"))
        farm_name = clean_name(row_dict.get("farm_name"))
        if not customer_name:
            j += 1
            continue
        active_text = str(row_dict.get("active", "1") or "1").strip().lower()
        if active_text in ["0", "false", "no", "n", "off"]:
            j += 1
            continue
        rows.append({
            "_row_number": j + 1,
            "customer_name": customer_name,
            "farm_name": farm_name,
            "muck_type": clean_name(row_dict.get("muck_type")),
            "email": ", ".join(normalize_email_list([row_dict.get("email", "")])),
            "email_2": normalize_email_address(row_dict.get("email_2", "")),
            "address_line_1": clean_name(row_dict.get("address_line_1")),
            "address_line_2": clean_name(row_dict.get("address_line_2")),
            "town": clean_name(row_dict.get("town")),
            "postcode": clean_name(row_dict.get("postcode")),
            "rate_per_ton": parse_csv_decimal(row_dict.get("rate_per_ton")),
            "vat_rate": parse_csv_decimal(row_dict.get("vat_rate")),
        })
        j += 1

    rows.sort(key=lambda row: (row["customer_name"].lower(), row["farm_name"].lower()))
    return rows


def namespace_from_tag(tag, fallback=XLSX_NS):
    text = str(tag or "")
    if text.startswith("{") and "}" in text:
        return text[1:].split("}", 1)[0]
    return fallback


def worksheet_cell_string_value(cell, shared_strings):
    if cell is None:
        return ""
    cell_type = str(cell.attrib.get("t", "") or "").strip()
    if cell_type == "inlineStr":
        text_parts = []
        inline_node = first_child_by_local_name(cell, "is")
        if inline_node is not None:
            text_node = first_child_by_local_name(inline_node, "t")
            if text_node is not None and text_node.text is not None:
                text_parts.append(text_node.text)
            for run in children_by_local_name(inline_node, "r"):
                run_text = first_child_by_local_name(run, "t")
                if run_text is not None and run_text.text is not None:
                    text_parts.append(run_text.text)
        return "".join(text_parts)
    value_node = first_child_by_local_name(cell, "v")
    raw_value = value_node.text if value_node is not None and value_node.text is not None else ""
    if cell_type == "s":
        try:
            return shared_strings[int(str(raw_value or "0").strip())]
        except Exception:
            return ""
    return str(raw_value or "")


def worksheet_row_values(row_node, shared_strings):
    row_values = []
    for cell in children_by_local_name(row_node, "c"):
        ref = cell.attrib.get("r", "")
        col_index = worksheet_ref_col_index(ref)
        while len(row_values) < max(col_index - 1, 0):
            row_values.append("")
        row_values.append(worksheet_cell_string_value(cell, shared_strings))
    return row_values


def worksheet_find_or_create_cell(row_node, row_number, col_index, namespace):
    target_ref = "%s%s" % (xlsx_col_name(col_index), row_number)
    for cell in children_by_local_name(row_node, "c"):
        if str(cell.attrib.get("r", "") or "").upper() == target_ref.upper():
            return cell

    new_cell = ET.Element("{%s}c" % namespace, {"r": target_ref})
    inserted = False
    existing_children = list(row_node)
    insert_at = len(existing_children)
    for index, child in enumerate(existing_children):
        if str(child.tag).rsplit("}", 1)[-1] != "c":
            continue
        child_ref = child.attrib.get("r", "")
        if xlsx_col_index(child_ref) > col_index:
            insert_at = index
            break
    row_node.insert(insert_at, new_cell)
    return new_cell


def worksheet_set_cell_inline_text(cell, value, namespace):
    style_value = cell.attrib.get("s")
    ref_value = cell.attrib.get("r")
    cell.clear()
    if ref_value:
        cell.attrib["r"] = ref_value
    if style_value not in [None, ""]:
        cell.attrib["s"] = style_value
    text = str(value or "")
    if not text:
        return
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, "{%s}is" % namespace)
    t_node = ET.SubElement(is_node, "{%s}t" % namespace)
    t_node.text = text


def worksheet_set_cell_number(cell, value_text, namespace):
    style_value = cell.attrib.get("s")
    ref_value = cell.attrib.get("r")
    cell.clear()
    if ref_value:
        cell.attrib["r"] = ref_value
    if style_value not in [None, ""]:
        cell.attrib["s"] = style_value
    text = str(value_text or "").strip()
    if not text:
        return
    value_node = ET.SubElement(cell, "{%s}v" % namespace)
    value_node.text = text


def save_customer_master_customer_details(customer_name, customer_email, rate_per_ton_text, row_number=None, field_values=None):
    customer_name = clean_name(customer_name)
    if not customer_name:
        return False, "Customer is required"
    raw_email = str(customer_email or "").strip()
    normalized_emails = normalize_email_list([raw_email]) if raw_email else []
    normalized_email = ", ".join(normalized_emails)
    if raw_email and not normalized_emails:
        return False, "Customer email is not valid"
    normalized_field_values = dict(field_values) if isinstance(field_values, dict) else field_values
    if isinstance(normalized_field_values, dict):
        if "email" in normalized_field_values:
            normalized_field_values["email"] = normalized_email
        if "email_2" in normalized_field_values:
            second_email = normalize_email_address(normalized_field_values.get("email_2", ""))
            if str(normalized_field_values.get("email_2", "") or "").strip() and not second_email:
                return False, "Second customer email is not valid"
            normalized_field_values["email_2"] = second_email
    if settings_workbook_exists():
        return save_workbook_customer(customer_name, normalized_email, rate_per_ton_text, row_number, normalized_field_values)
    if not os.path.exists(CUSTOMER_MASTER_XLSX_PATH):
        return False, "customer_master.xlsx could not be found"

    normalized_rate = str(rate_per_ton_text or "").strip()
    if normalized_rate:
        try:
            normalized_rate = ("%.2f" % float(normalized_rate)).rstrip("0").rstrip(".")
        except Exception:
            return False, "Rate per ton must be a number"

    try:
        with zipfile.ZipFile(CUSTOMER_MASTER_XLSX_PATH, "r") as source_archive:
            archive_names = source_archive.namelist()
            entries = {name: source_archive.read(name) for name in archive_names}
            workbook_root = ET.fromstring(entries["xl/workbook.xml"])
            rels_root = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
            relationships = {}
            for rel in rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                relationships[rel.attrib.get("Id")] = rel.attrib.get("Target", "")
            sheets_node = first_child_by_local_name(workbook_root, "sheets")
            first_sheet = first_child_by_local_name(sheets_node, "sheet")
            if first_sheet is None:
                return False, "No worksheet found in customer master"
            rel_id = (
                first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                or first_sheet.attrib.get("{%s}id" % STRICT_REL_NS, "")
            )
            target = relationships.get(rel_id, "worksheets/sheet1.xml")
            if not target.startswith("xl/"):
                target = "xl/%s" % target.lstrip("/")
            sheet_root = ET.fromstring(entries[target])
            shared_strings = xlsx_shared_strings(source_archive)
    except Exception:
        return False, "Could not open customer_master.xlsx"

    sheet_namespace = namespace_from_tag(sheet_root.tag)
    ET.register_namespace("", sheet_namespace)
    sheet_data = first_child_by_local_name(sheet_root, "sheetData")
    if sheet_data is None:
        return False, "Customer master sheet is missing row data"

    header_row = None
    headers = []
    all_rows = children_by_local_name(sheet_data, "row")
    row_number_filter = row_number
    for row_node in all_rows:
        candidate = [clean_name(cell).lower() for cell in worksheet_row_values(row_node, shared_strings)]
        if "customer_name" in candidate:
            header_row = row_node
            headers = candidate
            break
    if header_row is None:
        return False, "Customer master headers could not be found"

    email_col = headers.index("email") + 1 if "email" in headers else 0
    rate_col = headers.index("rate_per_ton") + 1 if "rate_per_ton" in headers else 0
    if not email_col or not rate_col:
        return False, "Customer master must include email and rate_per_ton columns"

    matched_rows = 0
    header_found = False
    for row_node in all_rows:
        if row_node is header_row:
            header_found = True
            continue
        if not header_found:
            continue
        row_values = worksheet_row_values(row_node, shared_strings)
        row_dict = customer_master_row_to_dict(headers, row_values)
        if row_number_filter is not None:
            try:
                if int(row_node.attrib.get("r", "0") or "0") != int(row_number_filter):
                    continue
            except (TypeError, ValueError):
                continue
        elif clean_name(row_dict.get("customer_name")).lower() != customer_name.lower():
            continue
        row_number_text = str(row_node.attrib.get("r", "") or "").strip()
        try:
            row_number_value = int(row_number_text)
        except Exception:
            continue
        email_cell = worksheet_find_or_create_cell(row_node, row_number_value, email_col, sheet_namespace)
        rate_cell = worksheet_find_or_create_cell(row_node, row_number_value, rate_col, sheet_namespace)
        worksheet_set_cell_inline_text(email_cell, normalized_email, sheet_namespace)
        worksheet_set_cell_number(rate_cell, normalized_rate, sheet_namespace)
        if isinstance(normalized_field_values, dict):
            for header, value in normalized_field_values.items():
                col_index = headers.index(header) + 1 if header in headers else 0
                if not col_index:
                    continue
                cell = worksheet_find_or_create_cell(row_node, row_number_value, col_index, sheet_namespace)
                if header in ["rate_per_ton", "vat_rate", "active"]:
                    worksheet_set_cell_number(cell, value, sheet_namespace)
                else:
                    worksheet_set_cell_inline_text(cell, value, sheet_namespace)
        matched_rows += 1

    if not matched_rows:
        data_rows = [row_node for row_node in all_rows if row_node is not header_row]
        if not data_rows:
            return False, "Customer master has no data row template"
        template_row = data_rows[-1]
        row_number_text = str(template_row.attrib.get("r", "0") or "0").strip()
        try:
            new_row_number = int(row_number_text) + 1
        except ValueError:
            return False, "Customer master row numbering is invalid"
        new_row = ET.fromstring(ET.tostring(template_row, encoding="utf-8"))
        new_row.attrib["r"] = str(new_row_number)
        for cell in children_by_local_name(new_row, "c"):
            cell_ref = cell.attrib.get("r", "")
            col_index = worksheet_ref_col_index(cell_ref)
            if not col_index:
                continue
            cell.attrib["r"] = "%s%s" % (xlsx_col_name(col_index), new_row_number)
            worksheet_set_cell_inline_text(cell, "", sheet_namespace)

        header_positions = {header: index + 1 for index, header in enumerate(headers) if header}
        customer_cell = worksheet_find_or_create_cell(new_row, new_row_number, header_positions.get("customer_name", 0), sheet_namespace)
        email_cell = worksheet_find_or_create_cell(new_row, new_row_number, header_positions.get("email", 0), sheet_namespace)
        rate_cell = worksheet_find_or_create_cell(new_row, new_row_number, header_positions.get("rate_per_ton", 0), sheet_namespace)
        worksheet_set_cell_inline_text(customer_cell, customer_name, sheet_namespace)
        worksheet_set_cell_inline_text(email_cell, normalized_email, sheet_namespace)
        worksheet_set_cell_number(rate_cell, normalized_rate, sheet_namespace)
        if header_positions.get("active"):
            active_cell = worksheet_find_or_create_cell(new_row, new_row_number, header_positions["active"], sheet_namespace)
            worksheet_set_cell_number(active_cell, "1", sheet_namespace)
        if header_positions.get("vat_rate"):
            vat_cell = worksheet_find_or_create_cell(new_row, new_row_number, header_positions["vat_rate"], sheet_namespace)
            worksheet_set_cell_number(vat_cell, "20", sheet_namespace)
        sheet_data.append(new_row)
        matched_rows = 1

        dimension = first_child_by_local_name(sheet_root, "dimension")
        if dimension is not None and ":" in dimension.attrib.get("ref", ""):
            start_ref, end_ref = dimension.attrib["ref"].split(":", 1)
            end_match = re.match(r"^([A-Z]+)(\d+)$", end_ref)
            if end_match and int(end_match.group(2)) < new_row_number:
                dimension.attrib["ref"] = "%s:%s%s" % (start_ref, end_match.group(1), new_row_number)

    entries[target] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    temp_path = ""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=APP_ROOT) as temp_handle:
            temp_path = temp_handle.name
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output_archive:
            for name in archive_names:
                output_archive.writestr(name, entries[name])
        os.replace(temp_path, CUSTOMER_MASTER_XLSX_PATH)
    except Exception:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        return False, "Could not save customer_master.xlsx"

    return True, "%s updated" % customer_name


def customer_master_row_to_dict(header_row, raw_row):
    row_dict = {}
    i = 0
    while i < len(header_row):
        header = clean_name(header_row[i]).lower()
        if header:
            row_dict[header] = raw_row[i] if i < len(raw_row) else ""
        i += 1
    return row_dict


def build_customer_farm_map(master_rows, jobs=None, field_map=None):
    out = {}

    def add_farm(customer_name, farm_name):
        customer_name = clean_name(customer_name)
        farm_name = clean_name(farm_name)
        if not customer_name or not farm_name:
            return
        bucket = out.setdefault(customer_name, [])
        if farm_name not in bucket:
            bucket.append(farm_name)

    for row in master_rows:
        add_farm(row.get("customer_name", ""), row.get("farm_name"))

    if isinstance(jobs, list):
        for row in jobs:
            if isinstance(row, dict):
                add_farm(row.get("customer", ""), row.get("farm_name"))

    if isinstance(field_map, dict):
        for customer_name, farms_by_field in field_map.items():
            if not isinstance(farms_by_field, dict):
                continue
            for farm_name in farms_by_field.keys():
                add_farm(customer_name, farm_name)

    for customer_name in out:
        out[customer_name].sort(key=lambda item: item.lower())
    return out


def build_customer_rate_map(master_rows):
    out = {}
    for row in master_rows:
        customer_name = row.get("customer_name", "")
        rate_text = str(row.get("rate_per_ton", "") or "").strip()
        if not customer_name or not rate_text:
            continue
        customer_bucket = out.setdefault(customer_name, {})
        farm_name = clean_name(row.get("farm_name"))
        if farm_name:
            customer_bucket[farm_name] = rate_text
        elif "" not in customer_bucket:
            customer_bucket[""] = rate_text
    return out


def iso_day_after(value):
    try:
        return (datetime.strptime(str(value or "").strip(), "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        return ""


def build_customer_invoice_from_map():
    out = {}
    for row in load_invoice_ledger():
        customer_name = clean_name(row.get("customer"))
        if not customer_name:
            continue
        end_date = str(row.get("end_date", "") or "").strip()
        next_date = iso_day_after(end_date)
        if not next_date:
            continue
        customer_bucket = out.setdefault(customer_name, {})
        farm_name = clean_name(row.get("farm_name"))
        previous = customer_bucket.get(farm_name, "")
        if not previous or next_date > previous:
            customer_bucket[farm_name] = next_date
    return out


def build_customer_field_admin_map(master_rows, jobs, field_map):
    tree = {}
    customer_meta = {}

    def ensure_bucket(customer_name, farm_name):
        customer_name = clean_name(customer_name)
        farm_name = clean_name(farm_name)
        if not customer_name:
            return None
        customer_bucket = tree.setdefault(customer_name, {})
        farm_bucket = customer_bucket.setdefault(farm_name, [])
        return farm_bucket

    for row in master_rows:
        customer_name = row.get("customer_name")
        ensure_bucket(customer_name, row.get("farm_name"))
        if customer_name:
            meta = customer_meta.setdefault(customer_name, {"customer_email": "", "rate_per_ton": ""})
            email_values = normalize_email_list([row.get("email", ""), row.get("email_2", "")])
            existing_emails = normalize_email_list([meta["customer_email"]])
            for email in email_values:
                if email.lower() not in {item.lower() for item in existing_emails}:
                    existing_emails.append(email)
            meta["customer_email"] = ", ".join(existing_emails)
            rate_value = str(row.get("rate_per_ton", "") or "").strip()
            if rate_value and not meta["rate_per_ton"]:
                meta["rate_per_ton"] = rate_value

    for job in jobs:
        ensure_bucket(job.get("customer"), job.get("farm_name"))

    if isinstance(field_map, dict):
        for customer_name, fields_by_farm in field_map.items():
            if not isinstance(fields_by_farm, dict):
                continue
            for farm_name, field_names in fields_by_farm.items():
                bucket = ensure_bucket(customer_name, farm_name)
                if bucket is None or not isinstance(field_names, list):
                    continue
                for field_name in field_names:
                    cleaned = clean_name(field_name)
                    if cleaned and cleaned not in bucket:
                        bucket.append(cleaned)

    output = []
    for customer_name in sorted(tree.keys(), key=lambda item: item.lower()):
        master_records = [dict(row) for row in master_rows if clean_name(row.get("customer_name")).lower() == customer_name.lower()]
        farms = []
        for farm_name in sorted(tree[customer_name].keys(), key=lambda item: item.lower()):
            fields = sorted(tree[customer_name][farm_name], key=lambda item: item.lower())
            farms.append({
                "farm_name": farm_name or "",
                "farm_label": farm_name or "No Farm Name",
                "fields": fields,
            })
        output.append({
            "customer_name": customer_name,
            "customer_email": customer_meta.get(customer_name, {}).get("customer_email", ""),
            "rate_per_ton_text": customer_meta.get(customer_name, {}).get("rate_per_ton", ""),
            "rate_per_ton_label": format_money(parse_decimal_or_zero(customer_meta.get(customer_name, {}).get("rate_per_ton", ""))) if str(customer_meta.get(customer_name, {}).get("rate_per_ton", "") or "").strip() else "",
            "master_records": master_records,
            "farms": farms,
        })
    return output


def find_customer_master_record(master_rows, customer_name, farm_name=""):
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    if not customer_name:
        return None

    matches = []
    for row in master_rows:
        if str(row.get("customer_name", "")).lower() == customer_name.lower():
            matches.append(row)

    if not matches:
        return None
    if farm_name:
        for row in matches:
            if str(row.get("farm_name", "")).lower() == farm_name.lower():
                return row
    if len(matches) == 1:
        return matches[0]
    for row in matches:
        if not str(row.get("farm_name", "")).strip():
            return row
    return None


def apply_customer_master_snapshot(record, master_rows, customer_name, farm_name):
    if not isinstance(record, dict):
        return
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    master_record = find_customer_master_record(master_rows, customer_name, farm_name)
    if isinstance(master_record, dict):
        record.update({
            "customer_email": master_record.get("email", ""),
            "customer_address_line_1": master_record.get("address_line_1", ""),
            "customer_address_line_2": master_record.get("address_line_2", ""),
            "customer_town": master_record.get("town", ""),
            "customer_postcode": master_record.get("postcode", ""),
            "rate_per_ton": master_record.get("rate_per_ton", ""),
            "vat_rate": master_record.get("vat_rate", ""),
            "customer_master_match": True,
        })
    else:
        record.update({
            "customer_email": "",
            "customer_address_line_1": "",
            "customer_address_line_2": "",
            "customer_town": "",
            "customer_postcode": "",
            "rate_per_ton": "",
            "vat_rate": "",
            "customer_master_match": False,
        })


def merge_customer_into_customer(from_customer, to_customer, master_rows=None, jobs=None, field_map=None, invoice_ledger=None):
    from_customer = canonical_customer_name(
        from_customer,
        master_rows=master_rows if isinstance(master_rows, list) else load_customer_master_rows(),
        customers=load_customers(),
        jobs=jobs if isinstance(jobs, list) else load_jobs(),
        field_map=field_map if isinstance(field_map, dict) else load_field_map(),
        invoice_ledger=invoice_ledger if isinstance(invoice_ledger, list) else load_invoice_ledger(),
    )
    to_customer = canonical_customer_name(
        to_customer,
        master_rows=master_rows if isinstance(master_rows, list) else load_customer_master_rows(),
        customers=load_customers(),
        jobs=jobs if isinstance(jobs, list) else load_jobs(),
        field_map=field_map if isinstance(field_map, dict) else load_field_map(),
        invoice_ledger=invoice_ledger if isinstance(invoice_ledger, list) else load_invoice_ledger(),
    )
    if not from_customer or not to_customer:
        raise ValueError("Both customer names are required")
    if normalized_name_key(from_customer) == normalized_name_key(to_customer):
        raise ValueError("You cannot merge a customer into itself")

    jobs_rows = jobs if isinstance(jobs, list) else load_jobs()
    invoice_rows = invoice_ledger if isinstance(invoice_ledger, list) else load_invoice_ledger()
    field_map_data = field_map if isinstance(field_map, dict) else load_field_map()

    updated_jobs = 0
    for row in jobs_rows:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) != normalized_name_key(from_customer):
            continue
        row["customer"] = to_customer
        updated_jobs += 1

    updated_invoices = 0
    for row in invoice_rows:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) != normalized_name_key(from_customer):
            continue
        row["customer"] = to_customer
        updated_invoices += 1

    source_bucket = field_map_data.get(from_customer, {}) if isinstance(field_map_data.get(from_customer, {}), dict) else {}
    target_bucket = field_map_data.get(to_customer, {}) if isinstance(field_map_data.get(to_customer, {}), dict) else {}
    for farm_name, fields in source_bucket.items():
        if not isinstance(fields, list):
            continue
        merged_fields = target_bucket.get(farm_name, [])
        for field_name in fields:
            cleaned_field = clean_name(field_name)
            if cleaned_field and normalized_name_key(cleaned_field) not in [normalized_name_key(item) for item in merged_fields]:
                merged_fields.append(cleaned_field)
        merged_fields.sort(key=lambda item: item.lower())
        target_bucket[farm_name] = merged_fields
    if source_bucket:
        field_map_data[to_customer] = target_bucket
    if from_customer in field_map_data:
        del field_map_data[from_customer]

    customer_names = load_customers()
    customer_names = [name for name in customer_names if normalized_name_key(name) != normalized_name_key(from_customer)]
    if to_customer not in customer_names:
        customer_names.append(to_customer)
    customer_names.sort(key=lambda item: item.lower())
    save_customers(customer_names)

    if updated_jobs:
        save_jobs(jobs_rows)
    if updated_invoices:
        save_invoice_ledger(invoice_rows)
    save_field_map(field_map_data)
    sync_farms_store(master_rows if isinstance(master_rows, list) else load_customer_master_rows(), jobs_rows, field_map_data)

    return {
        "updated_jobs": updated_jobs,
        "updated_invoices": updated_invoices,
        "to_customer": to_customer,
    }


def parse_csv_bool(value, default=False):
    text = str(value or "").strip().lower()
    if not text:
        return bool(default)
    if text in ["1", "true", "yes", "y", "on"]:
        return True
    if text in ["0", "false", "no", "n", "off"]:
        return False
    return bool(default)


def parse_csv_int(value, default):
    text = str(value if value is not None else "").strip()
    if not text:
        return int(default)
    try:
        return int(text)
    except Exception:
        return int(default)


def load_email_settings_csv():
    if settings_workbook_exists():
        values = settings_key_values("Email Settings")
        parsed = dict(DEFAULT_EMAIL_CONFIG)
        for key, default in DEFAULT_EMAIL_CONFIG.items():
            value = values.get(key, default)
            if isinstance(default, bool):
                parsed[key] = parse_csv_bool(value, default)
            elif isinstance(default, int):
                parsed[key] = parse_csv_int(value, default)
            elif key != "to_emails":
                parsed[key] = str(value or "")
        parsed["to_emails"] = [row["email"].strip() for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Email Recipients") if settings_row_active(row) and parse_csv_bool(row.get("summary", "1"), True) and row.get("email", "").strip()]
        return parsed
    if not os.path.exists(EMAIL_SETTINGS_CSV_PATH):
        return {}

    try:
        with open(EMAIL_SETTINGS_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            rows = [row for row in csv.DictReader(handle) if isinstance(row, dict)]
    except Exception:
        return {}

    selected_row = None
    recipient_emails = []
    has_record_type_column = bool(rows and "record_type" in rows[0])
    for row in rows:
        active_text = str(row.get("active", "1") or "1").strip().lower()
        if active_text in ["0", "false", "no", "n", "off"]:
            continue
        if has_record_type_column:
            record_type = str(row.get("record_type", "") or "").strip().lower()
            if record_type == "recipient":
                email = str(row.get("email", "") or "").strip()
                if email and email not in recipient_emails:
                    recipient_emails.append(email)
                continue
            if record_type and record_type != "settings":
                continue
        if selected_row is None:
            selected_row = row

    if not isinstance(selected_row, dict):
        return {}

    parsed = {
        "enabled": parse_csv_bool(selected_row.get("enabled"), False),
        "smtp_host": str(selected_row.get("smtp_host", "") or "").strip(),
        "smtp_port": parse_csv_int(selected_row.get("smtp_port"), 587),
        "use_tls": parse_csv_bool(selected_row.get("use_tls"), True),
        "smtp_username": str(selected_row.get("smtp_username", "") or "").strip(),
        "smtp_password": str(selected_row.get("smtp_password", "") or ""),
        "from_email": str(selected_row.get("from_email", "") or "").strip(),
        "send_weekday": parse_csv_int(selected_row.get("send_weekday"), 0),
        "send_hour": parse_csv_int(selected_row.get("send_hour"), 7),
        "send_minute": parse_csv_int(selected_row.get("send_minute"), 0),
        "monthly_enabled": parse_csv_bool(selected_row.get("monthly_enabled"), parse_csv_bool(selected_row.get("enabled"), False)),
        "monthly_send_hour": parse_csv_int(selected_row.get("monthly_send_hour"), parse_csv_int(selected_row.get("send_hour"), 7)),
        "monthly_send_minute": parse_csv_int(selected_row.get("monthly_send_minute"), parse_csv_int(selected_row.get("send_minute"), 0)),
        "subject_prefix": str(selected_row.get("subject_prefix", "") or "").strip(),
    }

    raw_to_emails = str(selected_row.get("to_emails", "") or "").strip()
    if raw_to_emails:
        parsed["to_emails"] = [email.strip() for email in raw_to_emails.split(",") if email.strip()]
    for email in recipient_emails:
        if email not in parsed.get("to_emails", []):
            parsed.setdefault("to_emails", []).append(email)

    return parsed


def update_email_settings_csv(values):
    if settings_workbook_exists():
        updates = {key: values.get(key, "") for key in ["smtp_host", "smtp_port", "smtp_username", "from_email", "use_tls"]}
        if values.get("smtp_password"):
            updates["smtp_password"] = values["smtp_password"]
        update_settings_key_values("Email Settings", updates)
        return True
    if not os.path.exists(EMAIL_SETTINGS_CSV_PATH):
        return False
    try:
        with open(EMAIL_SETTINGS_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            fieldnames = list(reader.fieldnames or [])
            rows = [dict(row) for row in reader]
    except Exception:
        return False
    if not fieldnames:
        return False

    settings_row = None
    for row in rows:
        record_type = str(row.get("record_type", "") or "").strip().lower()
        if record_type != "recipient":
            settings_row = row
            break
    if settings_row is None:
        return False

    for key in ["smtp_host", "smtp_port", "smtp_username", "from_email", "use_tls"]:
        settings_row[key] = str(values.get(key, "") or "")
    new_password = str(values.get("smtp_password", "") or "")
    if new_password:
        settings_row["smtp_password"] = new_password

    temp_path = ""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", dir=APP_ROOT, mode="w", newline="", encoding="utf-8") as handle:
            temp_path = handle.name
            writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temp_path, EMAIL_SETTINGS_CSV_PATH)
        return True
    except Exception:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        return False


def load_email_recipient_options():
    if settings_workbook_exists():
        options = [{"email": row["email"].strip(), "name": row.get("name", "").strip() or row["email"].strip()} for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Email Recipients") if settings_row_active(row) and parse_csv_bool(row.get("invoice_option", "1"), True) and row.get("email", "").strip()]
        return sorted(options, key=lambda row: (row["name"].lower(), row["email"].lower()))
    if not os.path.exists(EMAIL_SETTINGS_CSV_PATH):
        return []

    try:
        with open(EMAIL_SETTINGS_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            rows = [row for row in csv.DictReader(handle) if isinstance(row, dict)]
    except Exception:
        return []

    options = []
    seen = set()
    has_record_type_column = bool(rows and "record_type" in rows[0])
    for row in rows:
        active_text = str(row.get("active", "1") or "1").strip().lower()
        if active_text in ["0", "false", "no", "n", "off"]:
            continue
        if has_record_type_column and str(row.get("record_type", "") or "").strip().lower() != "recipient":
            continue
        email = str(row.get("email", "") or "").strip()
        if not email or email.lower() in seen:
            continue
        seen.add(email.lower())
        options.append({
            "email": email,
            "name": clean_name(row.get("name")) or email,
        })
    options.sort(key=lambda item: (item["name"].lower(), item["email"].lower()))
    return options


def load_timesheet_email_recipients():
    if settings_workbook_exists():
        recipients = []
        for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Email Recipients"):
            enabled = row.get("timesheet", row.get("summary", "0"))
            email = str(row.get("email", "") or "").strip()
            if settings_row_active(row) and parse_csv_bool(enabled, False) and email:
                recipients.append(email)
        return normalize_email_list(recipients)
    return normalize_email_list(load_email_config().get("to_emails", []))


def invoice_accounts_copy_emails(invoice_recipient_options):
    preferred = []
    fallback = []
    for item in invoice_recipient_options:
        email = str(item.get("email", "") or "").strip()
        name = str(item.get("name", "") or "").strip().lower()
        combined = "%s %s" % (name, email.lower())
        if not email:
            continue
        if "owen" in combined:
            continue
        fallback.append(email)
        if "office" in combined or "andrew" in combined:
            preferred.append(email)
    return normalize_email_list(preferred or fallback)


def render_invoice_template(template_text, invoice):
    template = str(template_text or "").strip()
    if not template:
        return ""
    farm_name = clean_name(invoice.get("farm_name"))
    greeting = "Good Morning" if datetime.now().hour < 12 else "Good Afternoon"
    return (
        template
        .replace("{greeting}", greeting)
        .replace("{invoice_number}", str(invoice.get("invoice_number_label", "")))
        .replace("{customer}", invoice_customer_label(invoice))
        .replace("{farm}", farm_name)
        .replace("{scope}", ("%s / %s" % (invoice_customer_label(invoice), farm_name)) if farm_name else invoice_customer_label(invoice))
    )


def load_email_config():
    data = {} if settings_workbook_exists() else read_json_file(EMAIL_CONFIG_PATH, {})
    merged = dict(DEFAULT_EMAIL_CONFIG)
    if isinstance(data, dict):
        merged.update(data)
    csv_data = load_email_settings_csv()
    if isinstance(csv_data, dict) and csv_data:
        merged.update(csv_data)
    merged["enabled"] = bool(merged.get("enabled", False))
    merged["smtp_host"] = str(merged.get("smtp_host", "") or "").strip()
    merged["smtp_port"] = int(merged.get("smtp_port", 587) or 587)
    merged["use_tls"] = bool(merged.get("use_tls", True))
    merged["smtp_username"] = str(merged.get("smtp_username", "") or "").strip()
    merged["smtp_password"] = str(merged.get("smtp_password", "") or "")
    merged["from_email"] = str(merged.get("from_email", "") or "").strip()
    raw_to_emails = merged.get("to_emails", [])
    if not isinstance(raw_to_emails, list):
        raw_to_emails = []
    to_emails = []
    for item in raw_to_emails:
        email = str(item or "").strip()
        if email and email not in to_emails:
            to_emails.append(email)
    merged["to_emails"] = to_emails
    merged["send_weekday"] = max(0, min(6, int(merged.get("send_weekday", 0) or 0)))
    merged["send_hour"] = max(0, min(23, parse_csv_int(merged.get("send_hour"), 7)))
    merged["send_minute"] = max(0, min(59, int(merged.get("send_minute", 0) or 0)))
    merged["monthly_enabled"] = bool(merged.get("monthly_enabled", merged.get("enabled", False)))
    merged["monthly_send_hour"] = max(0, min(23, parse_csv_int(merged.get("monthly_send_hour"), merged["send_hour"])))
    merged["monthly_send_minute"] = max(0, min(59, parse_csv_int(merged.get("monthly_send_minute"), merged["send_minute"])))
    merged["subject_prefix"] = str(merged.get("subject_prefix", "A. Farrell Contracting") or "A. Farrell Contracting").strip()
    return merged


def load_email_state():
    data = read_json_file(EMAIL_STATE_PATH, {})
    return data if isinstance(data, dict) else {}


def save_email_state(state):
    write_json_atomic(EMAIL_STATE_PATH, state if isinstance(state, dict) else {})


SUMMARY_SEND_LOCK_SECONDS = 4 * 60 * 60


def summary_send_locked(state, lock_key, period_key, now_ts=None):
    state = state if isinstance(state, dict) else {}
    current_period = str(state.get(lock_key, "") or "")
    if current_period != str(period_key or ""):
        return False
    try:
        locked_at = int(state.get(lock_key + "_at", 0) or 0)
    except Exception:
        locked_at = 0
    now_ts = int(now_ts or time.time())
    return bool(locked_at and (now_ts - locked_at) < SUMMARY_SEND_LOCK_SECONDS)


def claim_summary_send_lock(lock_key, period_key):
    now_ts = int(time.time())
    state = load_email_state()
    if summary_send_locked(state, lock_key, period_key, now_ts):
        return False
    state[lock_key] = str(period_key or "")
    state[lock_key + "_at"] = now_ts
    save_email_state(state)
    refreshed = load_email_state()
    return str(refreshed.get(lock_key, "") or "") == str(period_key or "")


def clear_summary_send_lock(lock_key, period_key):
    state = load_email_state()
    if str(state.get(lock_key, "") or "") == str(period_key or ""):
        state.pop(lock_key, None)
        state.pop(lock_key + "_at", None)
        save_email_state(state)


def load_invoice_ledger():
    data = read_json_file(INVOICE_LEDGER_PATH, [])
    if not isinstance(data, list):
        return []
    return [row for row in data if isinstance(row, dict)]


def save_invoice_ledger(rows):
    write_json_atomic(INVOICE_LEDGER_PATH, [row for row in rows if isinstance(row, dict)])


def load_invoice_state():
    data = read_json_file(INVOICE_STATE_PATH, {})
    return data if isinstance(data, dict) else {}


def save_invoice_state(state):
    write_json_atomic(INVOICE_STATE_PATH, state if isinstance(state, dict) else {})


def email_sender_ready(config):
    if not isinstance(config, dict):
        return False
    required = [
        str(config.get("smtp_host", "") or "").strip(),
        str(config.get("from_email", "") or "").strip(),
    ]
    if not all(required):
        return False
    return True


def email_config_ready(config):
    if not email_sender_ready(config):
        return False
    to_emails = config.get("to_emails", [])
    if not isinstance(to_emails, list) or not to_emails:
        return False
    return True


def clean_name(value):
    return " ".join(str(value or "").strip().split())


def normalized_name_key(value):
    return clean_name(value).lower()


_customer_case_normalized = False


def canonical_customer_name(value, master_rows=None, customers=None, jobs=None, field_map=None, invoice_ledger=None):
    name = clean_name(value)
    if not name:
        return ""
    wanted_key = normalized_name_key(name)
    sources = []
    if isinstance(master_rows, list):
        sources.append([clean_name(row.get("customer_name")) for row in master_rows if isinstance(row, dict)])
    if isinstance(customers, list):
        sources.append([clean_name(item) for item in customers])
    if isinstance(field_map, dict):
        sources.append([clean_name(item) for item in field_map.keys()])
    if isinstance(jobs, list):
        sources.append([clean_name(row.get("customer")) for row in jobs if isinstance(row, dict)])
    if isinstance(invoice_ledger, list):
        sources.append([clean_name(row.get("customer")) for row in invoice_ledger if isinstance(row, dict)])

    for source in sources:
        for candidate in source:
            if candidate and normalized_name_key(candidate) == wanted_key:
                return candidate
    return name


def normalize_customer_case_storage():
    global _customer_case_normalized
    if _customer_case_normalized:
        return

    master_rows = load_customer_master_rows()
    customers_data = load_customers()
    field_map_data = load_field_map()
    invoice_ledger_data = read_json_file(INVOICE_LEDGER_PATH, [])

    jobs_rows = []
    if os.path.exists(JOBS_PATH):
        try:
            with open(JOBS_PATH, "r") as handle:
                for raw_line in handle:
                    line = raw_line.strip()
                    if not line:
                        continue
                    try:
                        row = json.loads(line)
                    except Exception:
                        continue
                    if isinstance(row, dict):
                        jobs_rows.append(row)
        except Exception:
            jobs_rows = []

    changed = False

    normalized_customers = []
    seen_customer_keys = set()
    if isinstance(customers_data, list):
        for item in customers_data:
            canonical = canonical_customer_name(
                item,
                master_rows=master_rows,
                customers=customers_data,
                jobs=jobs_rows,
                field_map=field_map_data,
                invoice_ledger=invoice_ledger_data,
            )
            key = normalized_name_key(canonical)
            if not canonical or key in seen_customer_keys:
                if clean_name(item):
                    changed = True
                continue
            seen_customer_keys.add(key)
            normalized_customers.append(canonical)
            if canonical != clean_name(item):
                changed = True
    normalized_customers.sort(key=lambda item: item.lower())

    normalized_field_map = {}
    if isinstance(field_map_data, dict):
        for customer_name, farms_by_field in field_map_data.items():
            canonical_customer = canonical_customer_name(
                customer_name,
                master_rows=master_rows,
                customers=normalized_customers,
                jobs=jobs_rows,
                field_map=field_map_data,
                invoice_ledger=invoice_ledger_data,
            )
            if not canonical_customer or not isinstance(farms_by_field, dict):
                continue
            existing_customer_bucket = normalized_field_map.setdefault(canonical_customer, {})
            if canonical_customer != clean_name(customer_name):
                changed = True
            for farm_name, fields in farms_by_field.items():
                farm_key = clean_name(farm_name)
                field_bucket = existing_customer_bucket.setdefault(farm_key, [])
                if isinstance(fields, list):
                    for field_name in fields:
                        cleaned_field = clean_name(field_name)
                        if cleaned_field and cleaned_field not in field_bucket:
                            field_bucket.append(cleaned_field)
        for customer_name in normalized_field_map:
            for farm_name in normalized_field_map[customer_name]:
                normalized_field_map[customer_name][farm_name].sort(key=lambda item: item.lower())

    normalized_jobs = []
    for row in jobs_rows:
        updated = dict(row)
        canonical_customer = canonical_customer_name(
            row.get("customer"),
            master_rows=master_rows,
            customers=normalized_customers,
            jobs=jobs_rows,
            field_map=field_map_data,
            invoice_ledger=invoice_ledger_data,
        )
        if canonical_customer and canonical_customer != clean_name(row.get("customer")):
            updated["customer"] = canonical_customer
            changed = True
        normalized_jobs.append(updated)

    normalized_invoice_ledger = []
    if isinstance(invoice_ledger_data, list):
        for row in invoice_ledger_data:
            if not isinstance(row, dict):
                continue
            updated = dict(row)
            canonical_customer = canonical_customer_name(
                row.get("customer"),
                master_rows=master_rows,
                customers=normalized_customers,
                jobs=normalized_jobs,
                field_map=normalized_field_map,
                invoice_ledger=invoice_ledger_data,
            )
            if canonical_customer and canonical_customer != clean_name(row.get("customer")):
                updated["customer"] = canonical_customer
                changed = True
            normalized_invoice_ledger.append(updated)

    if changed:
        save_customers(normalized_customers)
        save_field_map(normalized_field_map)
        write_json_atomic(INVOICE_LEDGER_PATH, normalized_invoice_ledger)
        write_json_lines_atomic(JOBS_PATH, normalized_jobs)

    _customer_case_normalized = True


def load_customers():
    data = [row.get("customer_name", "") for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Customers") if settings_row_active(row)] if settings_workbook_exists() else read_json_file(CUSTOMERS_PATH, [])
    if not isinstance(data, list):
        return []
    names = []
    seen = set()
    for item in data:
        name = clean_name(item)
        key = normalized_name_key(name)
        if name and key not in seen:
            seen.add(key)
            names.append(name)
    names.sort(key=lambda item: item.lower())
    return names


def save_customers(customers):
    cleaned = []
    seen = set()
    for item in customers:
        name = clean_name(item)
        key = normalized_name_key(name)
        if name and key not in seen:
            seen.add(key)
            cleaned.append(name)
    cleaned.sort(key=lambda item: item.lower())
    if settings_workbook_exists():
        sync_workbook_customers(cleaned)
    else:
        write_json_atomic(CUSTOMERS_PATH, cleaned)


def load_farms():
    data = settings_names("Farms", []) if settings_workbook_exists() else read_json_file(FARMS_PATH, [])
    if not isinstance(data, list):
        return []
    names = []
    for item in data:
        name = clean_name(item)
        if name and name not in names:
            names.append(name)
    names.sort(key=lambda item: item.lower())
    return names


def save_farms(farms):
    cleaned = []
    for item in farms:
        name = clean_name(item)
        if name and name not in cleaned:
            cleaned.append(name)
    cleaned.sort(key=lambda item: item.lower())
    if settings_workbook_exists():
        save_settings_names("Farms", cleaned)
    else:
        write_json_atomic(FARMS_PATH, cleaned)


def sync_farms_store(master_rows=None, jobs=None, field_map=None):
    master_rows = master_rows if isinstance(master_rows, list) else load_customer_master_rows()
    jobs = jobs if isinstance(jobs, list) else load_jobs()
    field_map = field_map if isinstance(field_map, dict) else load_field_map()
    farms = []
    seen = set()

    def add_farm(value):
        name = clean_name(value)
        key = normalized_name_key(name)
        if not name or key in seen:
            return
        seen.add(key)
        farms.append(name)

    if settings_workbook_exists():
        for name in load_farms():
            add_farm(name)

    for row in master_rows:
        if isinstance(row, dict):
            add_farm(row.get("farm_name"))

    for row in jobs:
        if isinstance(row, dict):
            add_farm(row.get("farm_name"))

    for fields_by_farm in field_map.values():
        if not isinstance(fields_by_farm, dict):
            continue
        for farm_name in fields_by_farm.keys():
            add_farm(farm_name)

    farms.sort(key=lambda item: item.lower())
    save_farms(farms)
    return farms


def load_muck_types(master_rows=None):
    names = []

    master_rows = master_rows if isinstance(master_rows, list) else load_customer_master_rows()
    for raw_row in master_rows:
        if not isinstance(raw_row, dict):
            continue
        name = clean_name(raw_row.get("muck_type"))
        if name and name not in names:
            names.append(name)

    data = settings_names("Muck Types", []) if settings_workbook_exists() else read_json_file(MUCK_TYPES_PATH, [])
    if isinstance(data, list):
        for item in data:
            name = clean_name(item)
            if name and name not in names:
                names.append(name)

    names.sort(key=lambda item: item.lower())
    return names


def save_muck_types(muck_types):
    cleaned = []
    for item in muck_types:
        name = clean_name(item)
        if name and name not in cleaned:
            cleaned.append(name)
    cleaned.sort(key=lambda item: item.lower())
    if settings_workbook_exists():
        save_settings_names("Muck Types", cleaned)
    else:
        write_json_atomic(MUCK_TYPES_PATH, cleaned)


def load_field_map():
    if settings_workbook_exists():
        data = {}
        for row in Workbook(SETTINGS_WORKBOOK_PATH).records("Fields"):
            if settings_row_active(row) and row.get("customer_name") and row.get("field_name"):
                data.setdefault(row["customer_name"], {}).setdefault(row.get("farm_name", ""), []).append(row["field_name"])
    else:
        data = read_json_file(FIELD_MAP_PATH, {})
    if not isinstance(data, dict):
        return {}
    cleaned = {}
    for customer, fields_by_farm in data.items():
        customer_name = clean_name(customer)
        if not customer_name:
            continue
        customer_bucket = {}
        if isinstance(fields_by_farm, list):
            bucket = []
            for field_name in fields_by_farm:
                cleaned_field = clean_name(field_name)
                if cleaned_field and cleaned_field not in bucket:
                    bucket.append(cleaned_field)
            bucket.sort(key=lambda item: item.lower())
            if bucket:
                customer_bucket[""] = bucket
        elif isinstance(fields_by_farm, dict):
            for farm_name, fields in fields_by_farm.items():
                farm_key = clean_name(farm_name)
                bucket = []
                if isinstance(fields, list):
                    for field_name in fields:
                        cleaned_field = clean_name(field_name)
                        if cleaned_field and cleaned_field not in bucket:
                            bucket.append(cleaned_field)
                bucket.sort(key=lambda item: item.lower())
                if bucket:
                    customer_bucket[farm_key] = bucket
        if customer_bucket:
            cleaned[customer_name] = customer_bucket
    return cleaned


def save_field_map(field_map):
    out = {}
    for customer, fields_by_farm in field_map.items():
        customer_name = clean_name(customer)
        if not customer_name:
            continue
        customer_bucket = {}
        if isinstance(fields_by_farm, dict):
            for farm_name, fields in fields_by_farm.items():
                farm_key = clean_name(farm_name)
                bucket = []
                if isinstance(fields, list):
                    for field_name in fields:
                        cleaned_field = clean_name(field_name)
                        if cleaned_field and cleaned_field not in bucket:
                            bucket.append(cleaned_field)
                bucket.sort(key=lambda item: item.lower())
                if bucket:
                    customer_bucket[farm_key] = bucket
        if customer_bucket:
            out[customer_name] = customer_bucket
    if settings_workbook_exists():
        Workbook(SETTINGS_WORKBOOK_PATH).set_rows("Fields", [["customer_name", "farm_name", "field_name", "active"]] + [[customer, farm, field, "1"] for customer, farms in out.items() for farm, fields in farms.items() for field in fields])
    else:
        write_json_atomic(FIELD_MAP_PATH, out)


def load_jobs():
    ensure_data_dir()
    if not os.path.exists(JOBS_PATH):
        return []
    rows = []
    try:
        with open(JOBS_PATH, "r") as handle:
            for raw_line in handle:
                line = raw_line.strip()
                if not line:
                    continue
                try:
                    row = json.loads(line)
                except Exception:
                    continue
                if isinstance(row, dict):
                    row["issue_photos"] = normalize_issue_photo_names(row.get("issue_photos"))
                    rows.append(row)
    except Exception:
        return []
    rows.sort(key=lambda row: int(row.get("created_ts", 0)), reverse=True)
    return rows


def save_jobs(rows):
    ordered = []
    if isinstance(rows, list):
        ordered = [row for row in rows if isinstance(row, dict)]
    ordered.sort(key=lambda row: (int(row.get("created_ts", 0) or 0), int(row.get("id", 0) or 0)))
    write_json_lines_atomic(JOBS_PATH, ordered)


def find_job_by_id(job_id):
    try:
        wanted = int(job_id)
    except Exception:
        return None
    for row in load_jobs():
        try:
            if int(row.get("id", 0) or 0) == wanted:
                return row
        except Exception:
            continue
    return None


def upsert_job(job_record):
    rows = load_jobs()
    updated = False
    new_rows = []
    for row in rows:
        try:
            matches = int(row.get("id", 0) or 0) == int(job_record.get("id", 0) or 0)
        except Exception:
            matches = False
        if matches:
            new_rows.append(job_record)
            updated = True
        else:
            new_rows.append(row)
    if not updated:
        new_rows.append(job_record)
    save_jobs(new_rows)


def delete_job_by_id(job_id):
    try:
        wanted = int(job_id)
    except Exception:
        return False
    rows = load_jobs()
    kept = []
    removed = False
    for row in rows:
        try:
            matches = int(row.get("id", 0) or 0) == wanted
        except Exception:
            matches = False
        if matches:
            delete_issue_photo_files(row.get("issue_photos"))
            removed = True
            continue
        kept.append(row)
    if removed:
        save_jobs(kept)
    return removed


def parse_job_date(value):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except Exception:
        return None


def parse_optional_iso_date(value):
    text = str(value or "").strip()
    if not text:
        return None
    return parse_job_date(text)


def filter_jobs_by_date_range(rows, date_from=None, date_to=None):
    filtered = []
    for row in rows:
        job_date = parse_job_date(row.get("job_date"))
        if job_date is None:
            continue
        if date_from and job_date < date_from:
            continue
        if date_to and job_date > date_to:
            continue
        filtered.append(row)
    return filtered


def summarize_jobs(rows):
    total_spreader = 0.0
    total_john_deere = 0.0
    for row in rows:
        try:
            total_spreader += float(row.get("total_spreader_tons", 0) or 0)
        except Exception:
            pass
        try:
            total_john_deere += float(row.get("total_john_deere_tons", 0) or 0)
        except Exception:
            pass
    return {
        "job_count": len(rows),
        "total_spreader_tons": round(total_spreader, 2),
        "total_john_deere_tons": round(total_john_deere, 2),
    }


def previous_full_week_range(now=None):
    now = now or datetime.now()
    today = now.date()
    current_week_start = today - timedelta(days=today.weekday())
    start_date = current_week_start - timedelta(days=7)
    end_date = current_week_start - timedelta(days=1)
    return start_date, end_date


def current_week_range(now=None):
    now = now or datetime.now()
    today = now.date()
    start_date = today - timedelta(days=today.weekday())
    return start_date, today


def current_month_range(now=None):
    now = now or datetime.now()
    today = now.date()
    start_date = today.replace(day=1)
    return start_date, today


def last_month_range(now=None):
    now = now or datetime.now()
    today = now.date()
    first_of_current_month = today.replace(day=1)
    end_date = first_of_current_month - timedelta(days=1)
    start_date = end_date.replace(day=1)
    return start_date, end_date


def current_full_month_range(now=None):
    now = now or datetime.now()
    today = now.date()
    start_date = today.replace(day=1)
    if today.month == 12:
        next_month_start = today.replace(year=today.year + 1, month=1, day=1)
    else:
        next_month_start = today.replace(month=today.month + 1, day=1)
    end_date = next_month_start - timedelta(days=1)
    return start_date, end_date


def is_last_day_of_month(date_value):
    if date_value is None:
        return False
    return date_value == current_full_month_range(datetime.combine(date_value, datetime.min.time()))[1]


def jobs_summary_for_range(start_date, end_date, title=None):
    rows = []
    total_spreader = 0.0
    total_john_deere = 0.0
    for job in reversed(load_jobs()):
        try:
            job_date = datetime.strptime(str(job.get("job_date", "")), "%Y-%m-%d").date()
        except Exception:
            continue
        if not (start_date <= job_date <= end_date):
            continue
        row = dict(job)
        row["job_date"] = job_date.isoformat()
        rows.append(row)
        try:
            total_spreader += float(job.get("total_spreader_tons", 0) or 0)
        except Exception:
            pass
        try:
            total_john_deere += float(job.get("total_john_deere_tons", 0) or 0)
        except Exception:
            pass

    return {
        "title": title or "Jobs Summary",
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "job_count": len(rows),
        "total_spreader_tons": round(total_spreader, 2),
        "total_john_deere_tons": round(total_john_deere, 2),
        "rows": rows,
    }


def weekly_jobs_summary(start_date=None, end_date=None):
    if start_date is None or end_date is None:
        start_date, end_date = previous_full_week_range()
    return jobs_summary_for_range(start_date, end_date, "Weekly Jobs Summary")


def monthly_jobs_summary(now=None):
    start_date, end_date = current_month_range(now or datetime.now())
    return jobs_summary_for_range(start_date, end_date, "Monthly Jobs Summary")


def summary_email_subject(summary, config, label):
    prefix = str(config.get("subject_prefix", "A. Farrell Contracting") or "A. Farrell Contracting").strip()
    start_label = format_job_date(summary.get("start_date"))
    end_label = format_job_date(summary.get("end_date"))
    return "%s %s: %s - %s" % (prefix, label, start_label, end_label)


def weekly_email_subject(summary, config):
    return summary_email_subject(summary, config, "Weekly Jobs Summary")


def monthly_email_subject(summary, config):
    return summary_email_subject(summary, config, "Monthly Jobs Summary")


def summary_email_body(summary, label):
    return "\n".join([
        label,
        "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date"))),
        "",
        "Jobs: %s" % summary.get("job_count", 0),
        "Spreader Tons: %s" % format_tons(summary.get("total_spreader_tons", 0)),
        "Ops Center Tons: %s" % format_tons(summary.get("total_john_deere_tons", 0)),
        "",
        "The full summary is attached as XLSX and PDF files.",
    ])


def weekly_email_body(summary):
    return summary_email_body(summary, "Weekly Jobs Summary")


def monthly_email_body(summary):
    return summary_email_body(summary, "Monthly Jobs Summary")


def summary_attachment_filename(summary, prefix):
    return "%s_%s_to_%s.xlsx" % (
        prefix,
        str(summary.get("start_date", "")).replace("-", ""),
        str(summary.get("end_date", "")).replace("-", ""),
    )


def weekly_summary_attachment_filename(summary):
    return summary_attachment_filename(summary, "weekly_jobs_summary")


def monthly_summary_attachment_filename(summary):
    return summary_attachment_filename(summary, "monthly_jobs_summary")


def summary_pdf_attachment_filename(summary, prefix):
    return "%s_%s_to_%s.pdf" % (
        prefix,
        str(summary.get("start_date", "")).replace("-", ""),
        str(summary.get("end_date", "")).replace("-", ""),
    )


def weekly_summary_pdf_attachment_filename(summary):
    return summary_pdf_attachment_filename(summary, "weekly_jobs_summary")


def monthly_summary_pdf_attachment_filename(summary):
    return summary_pdf_attachment_filename(summary, "monthly_jobs_summary")


def summary_export_config(period_key):
    now = datetime.now()
    options = {
        "current-week": {
            "title": "Current Week Jobs Summary",
            "filename_prefix": "current_week_jobs_summary",
            "range": current_week_range(now),
        },
        "current-month": {
            "title": "Current Month Jobs Summary",
            "filename_prefix": "current_month_jobs_summary",
            "range": current_month_range(now),
        },
        "last-month": {
            "title": "Last Month Jobs Summary",
            "filename_prefix": "last_month_jobs_summary",
            "range": last_month_range(now),
        },
    }
    return options.get(str(period_key or "").strip().lower())


def xlsx_col_name(index):
    name = ""
    value = int(index)
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xlsx_col_index(ref):
    letters = "".join(char for char in str(ref or "") if char.isalpha()).upper()
    value = 0
    for char in letters:
        value = (value * 26) + (ord(char) - 64)
    return value


def xlsx_cell_xml(row_number, col_number, value, style_id=None):
    ref = "%s%s" % (xlsx_col_name(col_number), row_number)
    style_attr = ' s="%s"' % style_id if style_id not in [None, ""] else ""
    if value is None or value == "":
        return '<c r="%s"%s/>' % (ref, style_attr)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, style_attr, value)
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, style_attr, xml_escape(str(value)))


def template_string_cell_value(cell, shared_strings):
    if cell is None:
        return ""
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s":
        try:
            index = int((cell.find("{%s}v" % XLSX_NS).text or "0").strip())
            return shared_strings[index] if 0 <= index < len(shared_strings) else ""
        except Exception:
            return ""
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//{%s}t" % XLSX_NS))
    value_node = cell.find("{%s}v" % XLSX_NS)
    return value_node.text if value_node is not None and value_node.text else ""


def load_weekly_summary_template():
    if not os.path.exists(WEEKLY_SUMMARY_TEMPLATE_PATH):
        return None

    try:
        with zipfile.ZipFile(WEEKLY_SUMMARY_TEMPLATE_PATH, "r") as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
    except Exception:
        return None

    sheet_bytes = entries.get("xl/worksheets/sheet1.xml")
    if not sheet_bytes:
        return None

    try:
        root = ET.fromstring(sheet_bytes)
    except Exception:
        return None

    shared_strings = []
    shared_bytes = entries.get("xl/sharedStrings.xml")
    if shared_bytes:
        try:
            shared_root = ET.fromstring(shared_bytes)
            for item in shared_root.findall("{%s}si" % XLSX_NS):
                shared_strings.append("".join(node.text or "" for node in item.findall(".//{%s}t" % XLSX_NS)))
        except Exception:
            shared_strings = []

    rows_by_number = {}
    cells_by_ref = {}
    sheet_data = root.find("{%s}sheetData" % XLSX_NS)
    if sheet_data is not None:
        for row in sheet_data.findall("{%s}row" % XLSX_NS):
            try:
                row_number = int(row.attrib.get("r", "0") or "0")
            except Exception:
                continue
            rows_by_number[row_number] = row
            for cell in row.findall("{%s}c" % XLSX_NS):
                ref = cell.attrib.get("r", "")
                if ref:
                    cells_by_ref[ref] = cell

    def cell_text(ref):
        return template_string_cell_value(cells_by_ref.get(ref), shared_strings)

    def cell_style(ref):
        cell = cells_by_ref.get(ref)
        return cell.attrib.get("s") if cell is not None else None

    def row_attrs(row_number):
        row = rows_by_number.get(row_number)
        if row is None:
            return {}
        cleaned = {}
        for key, value in row.attrib.items():
            if key in ["r", "spans"] or key.startswith("{"):
                continue
            cleaned[key] = value
        return cleaned

    def row_style_map(row_number):
        out = {}
        col_index = 1
        while col_index <= 6:
            style_id = cell_style("%s%s" % (xlsx_col_name(col_index), row_number))
            if style_id not in [None, ""]:
                out[col_index] = style_id
            col_index += 1
        return out

    def row_layout(row_number):
        row = rows_by_number.get(row_number)
        if row is None:
            return {"columns": [], "styles": {}, "attrs": {}}

        columns = []
        styles = {}
        for cell in row.findall("{%s}c" % XLSX_NS):
            ref = cell.attrib.get("r", "")
            col_index = xlsx_col_index(ref)
            if not col_index:
                continue
            columns.append(col_index)
            style_id = cell.attrib.get("s")
            if style_id not in [None, ""]:
                styles[col_index] = style_id

        return {
            "columns": columns,
            "styles": styles,
            "attrs": row_attrs(row_number),
        }

    customer_style_map = row_style_map(10)
    farm_style_map = {}
    detail_style_map = {}
    farm_total_style_map = {}
    detail_row_layouts = []
    customer_row_number = 10
    farm_row_number = None
    detail_row_number = None
    farm_total_row_number = None
    in_first_farm_block = False

    for row_number in sorted(rows_by_number.keys()):
        if row_number <= 9:
            continue
        a_text = cell_text("A%s" % row_number).strip()
        b_text = cell_text("B%s" % row_number).strip()
        c_text = cell_text("C%s" % row_number).strip()
        d_text = cell_text("D%s" % row_number).strip()
        e_text = cell_text("E%s" % row_number).strip()
        f_text = cell_text("F%s" % row_number).strip()

        if a_text == "Farm Totals" and not farm_total_style_map:
            farm_total_style_map = row_style_map(row_number)
            farm_total_row_number = row_number
            continue

        if not a_text and any([b_text, c_text, d_text, e_text, f_text]):
            if not detail_style_map:
                detail_style_map = row_style_map(row_number)
                detail_row_number = row_number
            if in_first_farm_block:
                detail_row_layouts.append(row_layout(row_number))
            continue

        if a_text and not any([b_text, c_text, d_text, e_text, f_text]) and not farm_style_map and row_number > 10:
            farm_style_map = row_style_map(row_number)
            farm_row_number = row_number
            in_first_farm_block = True
            continue

        if in_first_farm_block and a_text == "Farm Totals":
            in_first_farm_block = False
            continue

        if in_first_farm_block and a_text and not any([b_text, c_text, d_text, e_text, f_text]) and row_number > (farm_row_number or 10):
            in_first_farm_block = False

    cols = []
    cols_node = root.find("{%s}cols" % XLSX_NS)
    if cols_node is not None:
        for col in cols_node.findall("{%s}col" % XLSX_NS):
            cols.append(dict(col.attrib))

    page_margins = {}
    page_margins_node = root.find("{%s}pageMargins" % XLSX_NS)
    if page_margins_node is not None:
        page_margins = dict(page_margins_node.attrib)

    return {
        "entries": entries,
        "cols": cols,
        "page_margins": page_margins,
        "sheet_format_attrs": dict((k, v) for k, v in (root.find("{%s}sheetFormatPr" % XLSX_NS) or ET.Element("x")).attrib.items() if not k.startswith("{")),
        "row_attrs": {
            "ops_total": row_attrs(6),
            "blank_after_summary": row_attrs(7),
            "headers": row_attrs(8),
            "blank_after_headers": row_attrs(9),
        },
        "row_styles": {
            "title": row_style_map(1),
            "period": row_style_map(2),
            "jobs": row_style_map(4),
            "spreader_total": row_style_map(5),
            "ops_total": row_style_map(6),
            "headers": row_style_map(8),
            "customer": customer_style_map,
            "farm": {},
            "detail": detail_style_map,
            "farm_total": farm_total_style_map,
        },
        "row_templates": {
            "title": row_layout(1),
            "period": row_layout(2),
            "blank": {"columns": [], "styles": {}, "attrs": {}},
            "jobs": row_layout(4),
            "spreader_total": row_layout(5),
            "ops_total": row_layout(6),
            "blank_after_summary": {"columns": [], "styles": {}, "attrs": row_attrs(7)},
            "headers": row_layout(8),
            "blank_after_headers": {"columns": [], "styles": {}, "attrs": row_attrs(9)},
            "customer": row_layout(customer_row_number),
            "farm": {
                "columns": (row_layout(farm_row_number or customer_row_number).get("columns") or [1]),
                "styles": {},
                "attrs": row_attrs(farm_row_number or customer_row_number),
            },
            "detail": row_layout(detail_row_number or 11),
            "farm_total": row_layout(farm_total_row_number or 12),
        },
        "detail_row_layouts": detail_row_layouts or [row_layout(detail_row_number or 11)],
        "labels": {
            "title": cell_text("A1") or "Weekly Jobs Summary",
            "period": cell_text("A2") or "Period",
            "jobs": cell_text("A4") or "Jobs",
            "spreader_total": cell_text("A5") or "Spreader Tons",
            "ops_total": cell_text("A6") or "Ops Center Tons",
            "headers": [
                cell_text("A8") or "Customer / Farm",
                cell_text("B8") or "Date",
                cell_text("C8") or "Field",
                cell_text("D8") or "Product",
                cell_text("E8") or "Spreader Weight",
                cell_text("F8") or "Ops Center Weight",
            ],
            "farm_totals": "Farm Totals",
        },
    }


def weekly_summary_template_rows(summary, labels):
    summary_title = clean_name(summary.get("title")) or labels["title"]
    rows = [
        ("title", [summary_title]),
        ("period", [labels["period"], "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date")))]),
        ("blank", []),
        ("jobs", [labels["jobs"], int(summary.get("job_count", 0) or 0)]),
        ("spreader_total", [labels["spreader_total"], float(summary.get("total_spreader_tons", 0) or 0)]),
        ("ops_total", [labels["ops_total"], float(summary.get("total_john_deere_tons", 0) or 0)]),
        ("blank_after_summary", []),
        ("headers", labels["headers"]),
        ("blank_after_headers", []),
    ]

    detail_rows = summary.get("rows", [])
    if not detail_rows:
        rows.append(("customer", ["No jobs were recorded in this period."]))
        return rows

    grouped_rows = {}
    for row in detail_rows:
        customer_name = clean_name(row.get("customer")) or "Unknown Customer"
        grouped_rows.setdefault(customer_name, []).append(row)

    for customer_name in sorted(grouped_rows.keys(), key=lambda item: item.lower()):
        rows.append(("customer", [customer_name]))

        farm_groups = {}
        for row in grouped_rows[customer_name]:
            farm_name = clean_name(row.get("farm_name"))
            farm_groups.setdefault(farm_name, []).append(row)

        named_farms = [farm_name for farm_name in farm_groups.keys() if clean_name(farm_name)]
        show_all_farm_labels = len(named_farms) > 1 or (len(named_farms) >= 1 and "" in farm_groups)

        for farm_name in sorted(farm_groups.keys(), key=lambda item: (item == "", item.lower())):
            if farm_name and (show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower()):
                rows.append(("farm", [farm_name]))

            farm_rows = sorted(
                farm_groups[farm_name],
                key=lambda row: (
                    str(row.get("job_date", "")),
                    clean_name(row.get("muck_type")).lower(),
                    clean_name(row.get("field_name")).lower(),
                ),
            )
            farm_spread_total = 0.0
            farm_ops_total = 0.0
            for row in farm_rows:
                try:
                    spread_value = float(row.get("total_spreader_tons", 0) or 0)
                except Exception:
                    spread_value = 0.0
                try:
                    ops_value = float(row.get("total_john_deere_tons", 0) or 0)
                except Exception:
                    ops_value = 0.0

                farm_spread_total += spread_value
                farm_ops_total += ops_value
                rows.append((
                    "detail",
                    [format_job_date(row.get("job_date")), row.get("field_name", ""), row.get("muck_type", ""), spread_value, ops_value],
                ))

            rows.append(("farm_total", [labels["farm_totals"], "", "", "", farm_spread_total, farm_ops_total]))
            rows.append(("blank", []))

    return rows


def worksheet_row_xml(row_number, values, columns=None, style_map=None, row_attrs=None):
    columns = list(columns or [])
    style_map = style_map or {}
    row_attrs = row_attrs or {}
    attrs = ['r="%s"' % row_number]
    for key, value in row_attrs.items():
        attrs.append('%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})))
    cells = []
    for value_index, value in enumerate(values):
        if value_index < len(columns):
            col_index = columns[value_index]
        else:
            col_index = value_index + 1
        cells.append(xlsx_cell_xml(row_number, col_index, value, style_map.get(col_index)))
    return "<row %s>%s</row>" % (" ".join(attrs), "".join(cells))


def build_template_based_xlsx(summary, template):
    labels = template.get("labels", {})
    row_templates = template.get("row_templates", {})
    row_specs = weekly_summary_template_rows(summary, labels)
    detail_row_layouts = template.get("detail_row_layouts") or [row_templates.get("detail", {})]
    detail_row_cycle_index = 0

    xml_rows = []
    row_number = 1
    for kind, values in row_specs:
        if kind == "detail" and detail_row_layouts:
            template_row = detail_row_layouts[detail_row_cycle_index % len(detail_row_layouts)]
            detail_row_cycle_index += 1
        else:
            template_row = row_templates.get(kind, {})
        xml_rows.append(
            worksheet_row_xml(
                row_number,
                values,
                template_row.get("columns"),
                template_row.get("styles"),
                template_row.get("attrs"),
            )
        )
        row_number += 1

    cols_xml = ""
    if template.get("cols"):
        col_parts = []
        for col in template["cols"]:
            attrs = []
            for key, value in col.items():
                attrs.append('%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})))
            col_parts.append("<col %s/>" % " ".join(attrs))
        cols_xml = "<cols>%s</cols>" % "".join(col_parts)

    sheet_format_attrs = template.get("sheet_format_attrs", {}) or {"defaultRowHeight": "15"}
    sheet_format_xml = "<sheetFormatPr %s/>" % " ".join(
        '%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})) for key, value in sheet_format_attrs.items()
    )

    page_margins = template.get("page_margins", {})
    page_margins_xml = ""
    if page_margins:
        page_margins_xml = "<pageMargins %s/>" % " ".join(
            '%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})) for key, value in page_margins.items()
        )

    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="%s">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  %s
  %s
  <sheetData>%s</sheetData>
  %s
</worksheet>
""" % (
        XLSX_NS,
        sheet_format_xml,
        cols_xml,
        "".join(xml_rows),
        page_margins_xml,
    )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in template.get("entries", {}).items():
            if name == "xl/worksheets/sheet1.xml":
                archive.writestr(name, worksheet_xml)
            else:
                archive.writestr(name, data)
    return output.getvalue()


def build_weekly_summary_sheet_rows(summary):
    rows = [
        [clean_name(summary.get("title")) or "Weekly Jobs Summary"],
        ["Period", "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date")))],
        [],
        ["Jobs", int(summary.get("job_count", 0) or 0)],
        ["Spreader Tons", float(summary.get("total_spreader_tons", 0) or 0)],
        ["Ops Center Tons", float(summary.get("total_john_deere_tons", 0) or 0)],
        [],
        ["Customer / Farm", "Date", "Field", "Product", "Spreader Weight", "Ops Center Weight"],
        [],
        [],
    ]

    detail_rows = summary.get("rows", [])
    if not detail_rows:
        rows.append(["No jobs were recorded in this period."])
        return rows

    grouped_rows = {}
    for row in detail_rows:
        customer_name = clean_name(row.get("customer")) or "Unknown Customer"
        grouped_rows.setdefault(customer_name, []).append(row)

    for customer_name in sorted(grouped_rows.keys(), key=lambda item: item.lower()):
        rows.append([customer_name])

        farm_groups = {}
        for row in grouped_rows[customer_name]:
            farm_name = clean_name(row.get("farm_name"))
            farm_groups.setdefault(farm_name, []).append(row)

        named_farms = [farm_name for farm_name in farm_groups.keys() if clean_name(farm_name)]
        show_all_farm_labels = len(named_farms) > 1 or (len(named_farms) >= 1 and "" in farm_groups)

        for farm_name in sorted(farm_groups.keys(), key=lambda item: (item == "", item.lower())):
            if farm_name and (show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower()):
                rows.append(["  %s" % farm_name])

            farm_rows = sorted(
                farm_groups[farm_name],
                key=lambda row: (
                    str(row.get("job_date", "")),
                    clean_name(row.get("muck_type")).lower(),
                    clean_name(row.get("field_name")).lower(),
                ),
            )
            farm_spread_total = 0.0
            farm_ops_total = 0.0
            for row in farm_rows:
                try:
                    spread_value = float(row.get("total_spreader_tons", 0) or 0)
                except Exception:
                    spread_value = 0.0
                try:
                    ops_value = float(row.get("total_john_deere_tons", 0) or 0)
                except Exception:
                    ops_value = 0.0

                farm_spread_total += spread_value
                farm_ops_total += ops_value
                rows.append([
                    "",
                    format_job_date(row.get("job_date")),
                    row.get("field_name", ""),
                    row.get("muck_type", ""),
                    spread_value,
                    ops_value,
                ])

            rows.append(["  Farm Totals", "", "", "", farm_spread_total, farm_ops_total])
            rows.append([])

    return rows


def pdf_escape(text):
    return str(text or "").replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def build_weekly_summary_pdf_rows(summary):
    rows = []
    detail_rows = summary.get("rows", [])
    if not detail_rows:
        rows.append(("detail", ["No jobs were recorded in this period.", "", "", "", "", ""]))
        return rows

    grouped_rows = {}
    for row in detail_rows:
        customer_name = clean_name(row.get("customer")) or "Unknown Customer"
        grouped_rows.setdefault(customer_name, []).append(row)

    for customer_name in sorted(grouped_rows.keys(), key=lambda item: item.lower()):
        rows.append(("customer", [customer_name, "", "", "", "", ""]))

        farm_groups = {}
        for row in grouped_rows[customer_name]:
            farm_name = clean_name(row.get("farm_name"))
            farm_groups.setdefault(farm_name, []).append(row)

        named_farms = [farm_name for farm_name in farm_groups.keys() if clean_name(farm_name)]
        show_all_farm_labels = len(named_farms) > 1 or (len(named_farms) >= 1 and "" in farm_groups)

        for farm_name in sorted(farm_groups.keys(), key=lambda item: (item == "", item.lower())):
            if farm_name and (show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower()):
                rows.append(("farm", [farm_name, "", "", "", "", ""]))

            farm_rows = sorted(
                farm_groups[farm_name],
                key=lambda row: (
                    str(row.get("job_date", "")),
                    clean_name(row.get("muck_type")).lower(),
                    clean_name(row.get("field_name")).lower(),
                ),
            )
            farm_spread_total = 0.0
            farm_ops_total = 0.0
            for row in farm_rows:
                try:
                    spread_value = float(row.get("total_spreader_tons", 0) or 0)
                except Exception:
                    spread_value = 0.0
                try:
                    ops_value = float(row.get("total_john_deere_tons", 0) or 0)
                except Exception:
                    ops_value = 0.0

                farm_spread_total += spread_value
                farm_ops_total += ops_value
                rows.append((
                    "detail",
                    [
                        "",
                        format_job_date(row.get("job_date")),
                        clean_name(row.get("field_name")),
                        clean_name(row.get("muck_type")),
                        format_tons(spread_value),
                        format_tons(ops_value),
                    ],
                ))

            rows.append((
                "farm_total",
                ["Farm Totals", "", "", "", format_tons(farm_spread_total), format_tons(farm_ops_total)],
            ))
            rows.append(("blank", ["", "", "", "", "", ""]))

    return rows


def truncate_pdf_text(text, max_chars):
    value = str(text or "")
    if len(value) <= max_chars:
        return value
    if max_chars <= 1:
        return value[:max_chars]
    return value[:max_chars - 1] + "…"


def pdf_text_command(x, y, text, font_name, font_size):
    return "0 g BT /%s %s Tf 1 0 0 1 %.2f %.2f Tm (%s) Tj ET" % (
        font_name,
        font_size,
        x,
        y,
        pdf_escape(text),
    )


def build_pdf_attachment_bytes(summary):
    page_width = 842
    page_height = 595
    left_margin = 24
    right_margin = 24
    top_margin = 22
    bottom_margin = 24
    table_top = 500
    header_height = 22
    row_heights = {
        "customer": 22,
        "farm": 20,
        "detail": 18,
        "farm_total": 20,
        "blank": 10,
    }
    column_widths = [182, 72, 170, 130, 115, 125]
    column_headers = ["Customer / Farm", "Date", "Field", "Product", "Spreader Weight", "Ops Center Weight"]
    column_x = [left_margin]
    for width in column_widths[:-1]:
        column_x.append(column_x[-1] + width)
    table_width = sum(column_widths)

    rows = build_weekly_summary_pdf_rows(summary)

    available_height = table_top - bottom_margin - header_height
    pages = []
    current_rows = []
    used_height = 0
    for row_kind, row_values in rows:
        row_height = row_heights.get(row_kind, 18)
        if current_rows and used_height + row_height > available_height:
            pages.append(current_rows)
            current_rows = []
            used_height = 0
        current_rows.append((row_kind, row_values))
        used_height += row_height
    if current_rows:
        pages.append(current_rows)

    if not pages:
        pages = [[("detail", ["No jobs were recorded in this period.", "", "", "", "", ""])]]

    def row_fill(row_kind):
        if row_kind == "header":
            return "0.88 0.88 0.88"
        if row_kind == "customer":
            return "0.93 0.93 0.93"
        if row_kind == "farm_total":
            return "0.95 0.95 0.95"
        return None

    def row_font(row_kind):
        if row_kind in ["header", "customer", "farm_total"]:
            if row_kind == "header":
                return ("F2", 8.6)
            return ("F2", 9.5)
        return ("F1", 9.2)

    def row_text_values(row_kind, values):
        if row_kind == "blank":
            return ["", "", "", "", "", ""]
        return list(values) + ([""] * (6 - len(values)))

    def draw_row(commands, y_top, row_kind, values):
        row_height = row_heights.get(row_kind, 18)
        y_bottom = y_top - row_height
        fill = row_fill(row_kind)
        if fill:
            commands.append("%s rg" % fill)
            commands.append("%.2f %.2f %.2f %.2f re f" % (left_margin, y_bottom, table_width, row_height))
        commands.append("0.65 G")
        commands.append("0.5 w")
        commands.append("%.2f %.2f %.2f %.2f re S" % (left_margin, y_bottom, table_width, row_height))
        for x_value in column_x[1:]:
            commands.append("%.2f %.2f m %.2f %.2f l S" % (x_value, y_bottom, x_value, y_top))

        font_name, font_size = row_font(row_kind)
        text_values = row_text_values(row_kind, values)
        max_chars = [32, 12, 30, 24, 12, 12]
        for index, text_value in enumerate(text_values):
            if not text_value:
                continue
            x_value = column_x[index] + 4
            y_value = y_bottom + ((row_height - font_size) / 2.0) + 2
            if index >= 4:
                text_value = str(text_value)
                approx_width = len(text_value) * (font_size * 0.5)
                x_value = column_x[index] + column_widths[index] - approx_width - 4
            commands.append(
                pdf_text_command(
                    x_value,
                    y_value,
                    truncate_pdf_text(text_value, max_chars[index]),
                    font_name,
                    font_size,
                )
            )

    page_streams = []
    for page_index, page_rows in enumerate(pages):
        commands = []
        if page_index == 0:
            commands.append(pdf_text_command(left_margin, page_height - top_margin - 10, "Weekly Jobs Summary", "F2", 18))
            commands.append(
                pdf_text_command(
                    left_margin,
                    page_height - top_margin - 30,
                    "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date"))),
                    "F1",
                    10,
                )
            )
            commands.append(
                pdf_text_command(
                    left_margin,
                    page_height - top_margin - 47,
                    "Jobs: %s" % summary.get("job_count", 0),
                    "F2",
                    9.5,
                )
            )
            commands.append(
                pdf_text_command(
                    left_margin + 110,
                    page_height - top_margin - 47,
                    "Spreader Tons: %s" % format_tons(summary.get("total_spreader_tons", 0)),
                    "F2",
                    9.5,
                )
            )
            commands.append(
                pdf_text_command(
                    left_margin + 290,
                    page_height - top_margin - 47,
                    "Ops Center Tons: %s" % format_tons(summary.get("total_john_deere_tons", 0)),
                    "F2",
                    9.5,
                )
            )

        y_cursor = table_top
        draw_row(commands, y_cursor, "header", column_headers)
        y_cursor -= header_height
        for row_kind, row_values in page_rows:
            draw_row(commands, y_cursor, row_kind, row_values)
            y_cursor -= row_heights.get(row_kind, 18)
        page_streams.append("\n".join(commands).encode("latin-1", "replace"))

    objects = []

    def add_object(payload):
        if isinstance(payload, str):
            payload = payload.encode("latin-1")
        objects.append(payload)
        return len(objects)

    font_regular_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    font_bold_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
    page_ids = []
    pages_id_placeholder = add_object(b"<<>>")

    for page_stream in page_streams:
        content_id = add_object(
            b"<< /Length %d >>\nstream\n%s\nendstream" % (len(page_stream), page_stream)
        )
        page_id = add_object(
            (
                "<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] "
                "/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> /Contents %d 0 R >>"
            ) % (pages_id_placeholder, page_width, page_height, font_regular_id, font_bold_id, content_id)
        )
        page_ids.append(page_id)

    pages_payload = "<< /Type /Pages /Count %d /Kids [%s] >>" % (
        len(page_ids),
        " ".join("%d 0 R" % page_id for page_id in page_ids),
    )
    objects[pages_id_placeholder - 1] = pages_payload.encode("latin-1")

    catalog_id = add_object(("<< /Type /Catalog /Pages %d 0 R >>" % pages_id_placeholder).encode("latin-1"))

    pdf = io.BytesIO()
    pdf.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, payload in enumerate(objects, start=1):
        offsets.append(pdf.tell())
        pdf.write(("%d 0 obj\n" % index).encode("latin-1"))
        pdf.write(payload)
        pdf.write(b"\nendobj\n")

    xref_start = pdf.tell()
    pdf.write(("xref\n0 %d\n" % (len(objects) + 1)).encode("latin-1"))
    pdf.write(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.write(("%010d 00000 n \n" % offset).encode("latin-1"))
    pdf.write(
        (
            "trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF" % (
                len(objects) + 1,
                catalog_id,
                xref_start,
            )
        ).encode("latin-1")
    )
    return pdf.getvalue()


def build_xlsx_attachment_bytes(summary):
    template = load_weekly_summary_template()
    if template:
        try:
            return build_template_based_xlsx(summary, template)
        except Exception:
            pass

    sheet_rows = build_weekly_summary_sheet_rows(summary)
    sheet_xml_rows = []
    row_index = 1
    for row in sheet_rows:
        cell_xml = []
        col_index = 1
        for value in row:
            cell_xml.append(xlsx_cell_xml(row_index, col_index, value))
            col_index += 1
        sheet_xml_rows.append('<row r="%s">%s</row>' % (row_index, "".join(cell_xml)))
        row_index += 1

    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="24" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="3" width="24" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
    <col min="5" max="6" width="14" customWidth="1"/>
  </cols>
  <sheetData>%s</sheetData>
</worksheet>
""" % "".join(sheet_xml_rows)

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Weekly Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

    timestamp = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Weekly Jobs Summary</dc:title>
  <dc:creator>A. Farrell Contracting</dc:creator>
  <cp:lastModifiedBy>A. Farrell Contracting</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>
""" % (timestamp, timestamp)

    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
</Properties>
"""

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
    return output.getvalue()


def build_basic_xlsx_bytes(sheet_name, title, sheet_rows, column_widths=None):
    sheet_xml_rows = []
    row_index = 1
    for row in sheet_rows:
        cell_xml = []
        col_index = 1
        for value in row:
            cell_xml.append(xlsx_cell_xml(row_index, col_index, value))
            col_index += 1
        sheet_xml_rows.append('<row r="%s">%s</row>' % (row_index, "".join(cell_xml)))
        row_index += 1

    cols_xml = ""
    if isinstance(column_widths, list) and column_widths:
        col_parts = []
        for index, width in enumerate(column_widths, start=1):
            col_parts.append('<col min="%s" max="%s" width="%s" customWidth="1"/>' % (index, index, width))
        cols_xml = "<cols>%s</cols>" % "".join(col_parts)

    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  %s
  <sheetData>%s</sheetData>
</worksheet>
""" % (cols_xml, "".join(sheet_xml_rows))

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="%s" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""" % xml_escape(sheet_name)

    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

    timestamp = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>%s</dc:title>
  <dc:creator>A. Farrell Contracting</dc:creator>
  <cp:lastModifiedBy>A. Farrell Contracting</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>
""" % (xml_escape(title), timestamp, timestamp)

    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
</Properties>
"""

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
    return output.getvalue()


def strip_invoice_office_extension_metadata(root):
    allowed_attribute_namespaces = {
        XLSX_NS,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://www.w3.org/XML/1998/namespace",
    }
    removable_children = {"AlternateContent", "revisionPtr", "extLst"}
    for element in root.iter():
        for attribute_name in list(element.attrib):
            if not str(attribute_name).startswith("{"):
                continue
            namespace = str(attribute_name)[1:].split("}", 1)[0]
            if namespace not in allowed_attribute_namespaces:
                element.attrib.pop(attribute_name, None)
        for child in list(element):
            local_name = str(child.tag).rsplit("}", 1)[-1]
            child_namespace = str(child.tag)[1:].split("}", 1)[0] if str(child.tag).startswith("{") else ""
            if local_name in removable_children or (child_namespace and child_namespace != XLSX_NS):
                element.remove(child)
    return root


def enforce_invoice_single_page_print_settings(xlsx_bytes):
    if not xlsx_bytes:
        return xlsx_bytes

    try:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
    except Exception:
        return xlsx_bytes

    # The invoice sheet is rebuilt from scratch, so the template's cached
    # calculation chain and links to the old invoice workbook are not needed.
    entries.pop("xl/calcChain.xml", None)
    for entry_name in list(entries):
        if entry_name.startswith("xl/externalLinks/"):
            entries.pop(entry_name, None)
    workbook_rels_path = "xl/_rels/workbook.xml.rels"
    workbook_rels_bytes = entries.get(workbook_rels_path)
    if workbook_rels_bytes:
        try:
            rels_root = ET.fromstring(workbook_rels_bytes)
            for relationship in list(rels_root):
                relationship_type = str(relationship.attrib.get("Type", ""))
                if relationship_type.endswith("/calcChain") or relationship_type.endswith("/externalLink"):
                    rels_root.remove(relationship)
            entries[workbook_rels_path] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
        except Exception:
            pass
    content_types_path = "[Content_Types].xml"
    content_types_bytes = entries.get(content_types_path)
    if content_types_bytes:
        try:
            content_types_root = ET.fromstring(content_types_bytes)
            for override in list(content_types_root):
                part_name = str(override.attrib.get("PartName", ""))
                if part_name == "/xl/calcChain.xml" or part_name.startswith("/xl/externalLinks/"):
                    content_types_root.remove(override)
            entries[content_types_path] = ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True)
        except Exception:
            pass

    sheet_path = "xl/worksheets/sheet1.xml"
    sheet_bytes = entries.get(sheet_path)
    if not sheet_bytes:
        return xlsx_bytes

    try:
        root = ET.fromstring(sheet_bytes)
    except Exception:
        return xlsx_bytes

    namespace = namespace_from_tag(root.tag, XLSX_NS)

    def ensure_root_child(local_name, insert_before=None):
        existing = first_child_by_local_name(root, local_name)
        if existing is not None:
            return existing
        node = ET.Element("{%s}%s" % (namespace, local_name))
        children = list(root)
        insert_index = len(children)
        if isinstance(insert_before, list):
            for index, child in enumerate(children):
                if str(child.tag).rsplit("}", 1)[-1] in insert_before:
                    insert_index = index
                    break
        root.insert(insert_index, node)
        return node

    sheet_pr = ensure_root_child("sheetPr", insert_before=["dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData"])
    page_setup_pr = first_child_by_local_name(sheet_pr, "pageSetUpPr")
    if page_setup_pr is None:
        page_setup_pr = ET.SubElement(sheet_pr, "{%s}pageSetUpPr" % namespace)
    page_setup_pr.attrib["fitToPage"] = "1"
    page_setup_pr.attrib["autoPageBreaks"] = "0"

    print_options = ensure_root_child("printOptions", insert_before=["pageMargins", "pageSetup", "headerFooter", "drawing"])
    print_options.attrib["horizontalCentered"] = "1"
    print_options.attrib["verticalCentered"] = "0"

    page_margins = ensure_root_child("pageMargins", insert_before=["pageSetup", "headerFooter", "drawing"])
    page_margins.attrib["left"] = page_margins.attrib.get("left", "0.3")
    page_margins.attrib["right"] = page_margins.attrib.get("right", "0.3")
    page_margins.attrib["top"] = page_margins.attrib.get("top", "0.35")
    page_margins.attrib["bottom"] = page_margins.attrib.get("bottom", "0.35")
    page_margins.attrib["header"] = page_margins.attrib.get("header", "0.2")
    page_margins.attrib["footer"] = page_margins.attrib.get("footer", "0.2")

    page_setup = ensure_root_child("pageSetup", insert_before=["headerFooter", "drawing"])
    page_setup.attrib["paperSize"] = "9"
    page_setup.attrib["orientation"] = "portrait"
    page_setup.attrib["fitToWidth"] = "1"
    page_setup.attrib["fitToHeight"] = "0"
    page_setup.attrib.pop("scale", None)
    page_setup.attrib["usePrinterDefaults"] = "0"

    strip_invoice_office_extension_metadata(root)
    ET.register_namespace("", namespace)
    ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    entries[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    workbook_path = "xl/workbook.xml"
    workbook_bytes = entries.get(workbook_path)
    if workbook_bytes:
        try:
            workbook_root = ET.fromstring(workbook_bytes)
            sheets_node = first_child_by_local_name(workbook_root, "sheets")
            sheet_nodes = children_by_local_name(sheets_node, "sheet") if sheets_node is not None else []
            first_sheet = sheet_nodes[0] if sheet_nodes else None
            sheet_name = first_sheet.attrib.get("name", "Invoice") if first_sheet is not None else "Invoice"
            max_row = 1
            dimension = first_child_by_local_name(root, "dimension")
            if dimension is not None:
                dimension_ref = dimension.attrib.get("ref", "")
                max_ref = dimension_ref.split(":", 1)[-1]
                row_match = re.search(r"(\d+)$", max_ref)
                if row_match:
                    max_row = max(1, int(row_match.group(1)))
            defined_names = first_child_by_local_name(workbook_root, "definedNames")
            if defined_names is not None:
                print_area = None
                for defined_name in children_by_local_name(defined_names, "definedName"):
                    if "[" in str(defined_name.text or ""):
                        defined_names.remove(defined_name)
                        continue
                    if defined_name.attrib.get("name") == "_xlnm.Print_Area":
                        print_area = defined_name
                if print_area is not None:
                    print_area.text = "%s!$A$1:$G$%s" % (sheet_name, max_row)
                if not list(defined_names):
                    workbook_root.remove(defined_names)
            external_references = first_child_by_local_name(workbook_root, "externalReferences")
            if external_references is not None:
                workbook_root.remove(external_references)
            strip_invoice_office_extension_metadata(workbook_root)
            ET.register_namespace("", XLSX_NS)
            ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            entries[workbook_path] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
        except Exception:
            pass

    styles_path = "xl/styles.xml"
    styles_bytes = entries.get(styles_path)
    if styles_bytes:
        try:
            styles_root = ET.fromstring(styles_bytes)
            strip_invoice_office_extension_metadata(styles_root)
            ET.register_namespace("", XLSX_NS)
            entries[styles_path] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
        except Exception:
            pass

    output = io.BytesIO()
    try:
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
    except Exception:
        return xlsx_bytes
    return output.getvalue()


def build_invoice_sheet_rows(invoice):
    scope_label = invoice_customer_label(invoice)
    if invoice.get("farm_name"):
        scope_label = "%s / %s" % (scope_label, invoice.get("farm_name"))

    sheet_rows = [
        ["Invoice"],
        ["Invoice Number", invoice.get("invoice_number_label", "")],
        ["Invoice Date", invoice.get("invoice_date_label", "")],
        ["Customer / Farm", scope_label],
        ["Billing Email", invoice.get("customer_email", "")],
        ["Period", "%s to %s" % (invoice.get("start_date_label", ""), invoice.get("end_date_label", ""))],
        ["Last Invoiced Through", invoice.get("last_invoiced_end_date_label", "")],
        [],
    ]

    for line in invoice_address_lines(invoice):
        sheet_rows.append(["Billing Address", line])

    sheet_rows.extend([
        [],
        ["Date", "Farm", "Field", "Product", "Notes", "Tons", "Rate Per Ton", "VAT %", "Line Total"],
    ])

    for row in invoice.get("line_rows", []):
        sheet_rows.append([
            row.get("job_date_label", ""),
            row.get("farm_name", ""),
            invoice_line_field_label(row),
            row.get("muck_type", ""),
            row.get("job_notes", ""),
            row.get("tons", ""),
            row.get("rate_per_ton", ""),
            row.get("vat_rate", 0),
            row.get("line_total", 0),
        ])

    sheet_rows.extend([
        [],
        ["Jobs", invoice.get("job_count", 0)],
        ["Total Tons", invoice.get("total_tons", 0)],
        ["Subtotal", invoice.get("subtotal", 0)],
        ["VAT", invoice.get("vat_total", 0)],
        ["Grand Total", invoice.get("grand_total", 0)],
    ])

    return sheet_rows


def build_plain_invoice_xlsx_bytes(invoice):
    sheet_rows = build_invoice_sheet_rows(invoice)

    return build_basic_xlsx_bytes(
        "Invoice",
        "Invoice %s" % invoice.get("invoice_number_label", ""),
        sheet_rows,
        column_widths=[16, 20, 24, 20, 28, 12, 14, 10, 14],
    )


def load_invoice_template():
    invoice_template_path = resolve_invoice_template_path()
    if not invoice_template_path:
        return None

    try:
        with zipfile.ZipFile(invoice_template_path, "r") as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
    except Exception:
        return None

    sheet_bytes = entries.get("xl/worksheets/sheet1.xml")
    if not sheet_bytes:
        return None

    try:
        root = ET.fromstring(sheet_bytes)
    except Exception:
        return None

    rows_by_number = {}
    sheet_data = root.find("{%s}sheetData" % XLSX_NS)
    if sheet_data is not None:
        for row in sheet_data.findall("{%s}row" % XLSX_NS):
            try:
                row_number = int(row.attrib.get("r", "0") or "0")
            except Exception:
                continue
            rows_by_number[row_number] = row

    def row_attrs(row_number):
        row = rows_by_number.get(row_number)
        if row is None:
            return {}
        cleaned = {}
        for key, value in row.attrib.items():
            if key in ["r", "spans"] or key.startswith("{"):
                continue
            cleaned[key] = value
        return cleaned

    def row_layout(row_number):
        row = rows_by_number.get(row_number)
        if row is None:
            return {"columns": [], "styles": {}, "attrs": {}}

        columns = []
        styles = {}
        for cell in row.findall("{%s}c" % XLSX_NS):
            ref = cell.attrib.get("r", "")
            col_index = xlsx_col_index(ref)
            if not col_index:
                continue
            columns.append(col_index)
            style_id = cell.attrib.get("s")
            if style_id not in [None, ""]:
                styles[col_index] = style_id

        return {
            "columns": columns,
            "styles": styles,
            "attrs": row_attrs(row_number),
        }

    cols = []
    cols_node = root.find("{%s}cols" % XLSX_NS)
    if cols_node is not None:
        for col in cols_node.findall("{%s}col" % XLSX_NS):
            cols.append(dict(col.attrib))

    page_margins = {}
    page_margins_node = root.find("{%s}pageMargins" % XLSX_NS)
    if page_margins_node is not None:
        page_margins = dict(page_margins_node.attrib)

    return {
        "entries": entries,
        "cols": cols,
        "page_margins": page_margins,
        "sheet_format_attrs": dict((k, v) for k, v in (root.find("{%s}sheetFormatPr" % XLSX_NS) or ET.Element("x")).attrib.items() if not k.startswith("{")),
        "row_templates": {
            "title": row_layout(1),
            "invoice_number": row_layout(2),
            "invoice_date": row_layout(3),
            "customer_scope": row_layout(4),
            "billing_email": row_layout(5),
            "period": row_layout(6),
            "last_invoiced": row_layout(7),
            "blank_meta": row_layout(8),
            "address_rows": [row_layout(9), row_layout(10), row_layout(11), row_layout(12)],
            "blank_before_lines": row_layout(13),
            "headers": row_layout(14),
            "detail_rows": [row_layout(15), row_layout(16), row_layout(17), row_layout(18), row_layout(19), row_layout(20)],
            "blank_before_totals": row_layout(21),
            "jobs": row_layout(22),
            "total_tons": row_layout(23),
            "subtotal": row_layout(24),
            "vat": row_layout(25),
            "grand_total": row_layout(26),
        },
    }


def excel_date_serial(date_text):
    try:
        date_value = datetime.strptime(str(date_text or ""), "%Y-%m-%d")
    except Exception:
        return None
    excel_epoch = datetime(1899, 12, 30)
    return (date_value - excel_epoch).days


def invoice_template_is_layout_workbook(template):
    entries = template.get("entries", {}) if isinstance(template, dict) else {}
    sheet_bytes = entries.get("xl/worksheets/sheet1.xml")
    if not sheet_bytes:
        return False
    try:
        root = ET.fromstring(sheet_bytes)
    except Exception:
        return False
    sheet_data = root.find("{%s}sheetData" % XLSX_NS)
    if sheet_data is None:
        return False
    cell_text = {}
    shared_strings = []
    shared_bytes = entries.get("xl/sharedStrings.xml")
    if shared_bytes:
        try:
            shared_root = ET.fromstring(shared_bytes)
            for item in shared_root.findall("{%s}si" % XLSX_NS):
                shared_strings.append("".join(node.text or "" for node in item.findall(".//{%s}t" % XLSX_NS)))
        except Exception:
            shared_strings = []
    for row in sheet_data.findall("{%s}row" % XLSX_NS):
        for cell in row.findall("{%s}c" % XLSX_NS):
            ref = cell.attrib.get("r", "")
            if ref:
                cell_text[ref] = template_string_cell_value(cell, shared_strings)
    return clean_name(cell_text.get("A8")).lower() == "to:" and clean_name(cell_text.get("E11")).lower() == "invoice no."


def template_row_cell_map(row):
    return {
        cell.attrib.get("r", ""): cell
        for cell in row.findall("{%s}c" % XLSX_NS)
        if cell.attrib.get("r")
    }


def set_template_cell_value(row, ref, value):
    cells = template_row_cell_map(row)
    cell = cells.get(ref)
    if cell is None:
        cell = ET.Element("{%s}c" % XLSX_NS, {"r": ref})
        row.append(cell)
    style_id = cell.attrib.get("s")
    cell.attrib.clear()
    cell.attrib["r"] = ref
    if style_id not in [None, ""]:
        cell.attrib["s"] = style_id
    for child in list(cell):
        cell.remove(child)
    if value is None or value == "":
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        value_node = ET.SubElement(cell, "{%s}v" % XLSX_NS)
        value_node.text = str(value)
        return
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, "{%s}is" % XLSX_NS)
    text_node = ET.SubElement(is_node, "{%s}t" % XLSX_NS)
    text_node.text = str(value)


def set_template_cell_formula(row, ref, formula, cached_value=None):
    """Write an Excel formula while retaining a cached value for previewers."""
    cells = template_row_cell_map(row)
    cell = cells.get(ref)
    if cell is None:
        cell = ET.Element("{%s}c" % XLSX_NS, {"r": ref})
        row.append(cell)
    style_id = cell.attrib.get("s")
    cell.attrib.clear()
    cell.attrib["r"] = ref
    if style_id not in [None, ""]:
        cell.attrib["s"] = style_id
    for child in list(cell):
        cell.remove(child)
    ET.SubElement(cell, "{%s}f" % XLSX_NS).text = str(formula)
    value_node = ET.SubElement(cell, "{%s}v" % XLSX_NS)
    if cached_value is not None and cached_value != "":
        value_node.text = str(cached_value)


def patch_currency_number_formats(styles_bytes):
    try:
        root = ET.fromstring(styles_bytes)
    except Exception:
        return styles_bytes

    custom_formats = {}
    numfmts_node = first_child_by_local_name(root, "numFmts")
    if numfmts_node is None:
        numfmts_node = ET.Element("{%s}numFmts" % XLSX_NS, {"count": "0"})
        inserted = False
        for index, child in enumerate(list(root)):
            local_name = str(child.tag).rsplit("}", 1)[-1]
            if local_name in ["fonts", "fills", "borders", "cellStyleXfs", "cellXfs"]:
                root.insert(index, numfmts_node)
                inserted = True
                break
        if not inserted:
            root.insert(0, numfmts_node)

    for numfmt in children_by_local_name(numfmts_node, "numFmt"):
        numfmt_id = str(numfmt.attrib.get("numFmtId", "") or "").strip()
        format_code = str(numfmt.attrib.get("formatCode", "") or "")
        if numfmt_id:
            custom_formats[numfmt_id] = format_code

    currency_format_code = '£#,##0.00'
    currency_numfmt_id = None
    for numfmt_id, format_code in custom_formats.items():
        if format_code == currency_format_code:
            currency_numfmt_id = numfmt_id
            break
    if currency_numfmt_id is None:
        highest_id = 163
        for numfmt_id in custom_formats.keys():
            try:
                highest_id = max(highest_id, int(numfmt_id))
            except Exception:
                continue
        currency_numfmt_id = str(highest_id + 1)
        ET.SubElement(
            numfmts_node,
            "{%s}numFmt" % XLSX_NS,
            {"numFmtId": currency_numfmt_id, "formatCode": currency_format_code},
        )
        custom_formats[currency_numfmt_id] = currency_format_code
    numfmts_node.attrib["count"] = str(len(children_by_local_name(numfmts_node, "numFmt")))

    cellxfs_node = first_child_by_local_name(root, "cellXfs")
    if cellxfs_node is None:
        return styles_bytes

    for xf in children_by_local_name(cellxfs_node, "xf"):
        numfmt_id = str(xf.attrib.get("numFmtId", "") or "").strip()
        if not numfmt_id:
            continue
        format_code = custom_formats.get(numfmt_id, "")
        is_currency = numfmt_id in ["5", "6", "7", "8", "44"] or ("£" in format_code)
        if is_currency:
            xf.attrib["numFmtId"] = currency_numfmt_id
            xf.attrib["applyNumberFormat"] = "1"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sort_template_row_cells(row):
    cells = row.findall("{%s}c" % XLSX_NS)
    cells.sort(key=lambda cell: xlsx_col_index(cell.attrib.get("r", "")))
    for cell in list(row):
        row.remove(cell)
    for cell in cells:
        row.append(cell)


def fill_layout_invoice_template(invoice, template):
    entries = template.get("entries", {})
    sheet_bytes = entries.get("xl/worksheets/sheet1.xml")
    if not sheet_bytes:
        raise ValueError("Invoice template worksheet is missing.")

    root = ET.fromstring(sheet_bytes)
    sheet_data = root.find("{%s}sheetData" % XLSX_NS)
    if sheet_data is None:
        raise ValueError("Invoice template sheet data is missing.")

    rows_by_number = {}
    for row in sheet_data.findall("{%s}row" % XLSX_NS):
        try:
            rows_by_number[int(row.attrib.get("r", "0") or "0")] = row
        except Exception:
            continue

    def row(number):
        found = rows_by_number.get(number)
        if found is None:
            raise ValueError("Invoice template row %s is missing." % number)
        return found

    address_lines = invoice_address_lines(invoice)[:5]
    while len(address_lines) < 5:
        address_lines.append("")

    set_template_cell_value(row(9), "A9", address_lines[0])
    set_template_cell_value(row(10), "A10", address_lines[1])
    set_template_cell_value(row(11), "A11", address_lines[2])
    set_template_cell_value(row(12), "A12", address_lines[3])
    set_template_cell_value(row(13), "A13", address_lines[4])

    invoice_date_serial = excel_date_serial(invoice.get("invoice_date"))
    set_template_cell_value(row(10), "G10", invoice.get("invoice_date_label", ""))
    set_template_cell_value(row(11), "G11", invoice.get("invoice_number_label", ""))
    set_template_cell_value(row(12), "G12", invoice.get("payment_terms_days", "14"))

    rate_text = invoice.get("rate_override_label", "") or ""
    set_template_cell_value(row(17), "A17", "Muck Spreading @  %s/ton, Description/Field Name" % rate_text if rate_text else "Muck Spreading, Description/Field Name")
    set_template_cell_value(row(17), "E17", "Product")
    set_template_cell_value(row(17), "F17", "Tons")
    set_template_cell_value(row(17), "G17", "Total")

    detail_rows = invoice.get("line_rows", []) or []
    detail_start_row = 18
    detail_end_row = 41

    def shift_row_reference(value, offset):
        match = re.match(r"^([A-Z]+)(\d+)$", str(value or ""))
        if not match:
            return value
        return "%s%s" % (match.group(1), int(match.group(2)) + offset)

    def shift_row_element(row_element, offset):
        row_element.attrib["r"] = str(int(row_element.attrib.get("r", "0")) + offset)
        for cell in row_element.findall("{%s}c" % XLSX_NS):
            cell_ref = cell.attrib.get("r")
            if cell_ref:
                cell.attrib["r"] = shift_row_reference(cell_ref, offset)

    available_detail_rows = max(0, (detail_end_row - detail_start_row) + 1)
    extra_detail_rows = max(0, len(detail_rows) - available_detail_rows)
    if extra_detail_rows:
        for existing_row in list(sheet_data.findall("{%s}row" % XLSX_NS)):
            try:
                existing_row_number = int(existing_row.attrib.get("r", "0") or "0")
            except Exception:
                continue
            if existing_row_number >= detail_end_row + 1:
                shift_row_element(existing_row, extra_detail_rows)

        template_row = rows_by_number[detail_end_row]
        for offset in range(1, extra_detail_rows + 1):
            new_row = ET.fromstring(ET.tostring(template_row, encoding="utf-8"))
            new_row.attrib["r"] = str(detail_end_row + offset)
            for cell in new_row.findall("{%s}c" % XLSX_NS):
                cell_ref = cell.attrib.get("r")
                if cell_ref:
                    cell.attrib["r"] = shift_row_reference(cell_ref, offset)
            sheet_data.insert(list(sheet_data).index(template_row) + offset, new_row)

        merge_cells = first_child_by_local_name(root, "mergeCells")
        if merge_cells is not None:
            for merge_cell in children_by_local_name(merge_cells, "mergeCell"):
                ref = merge_cell.attrib.get("ref", "")
                parts = ref.split(":")
                try:
                    first_row = int(re.search(r"\d+$", parts[0]).group(0))
                except (AttributeError, ValueError):
                    continue
                if first_row >= detail_end_row + 1:
                    merge_cell.attrib["ref"] = ":".join(
                        shift_row_reference(part, extra_detail_rows) for part in parts
                    )
            for row_number in range(detail_end_row + 1, detail_end_row + extra_detail_rows + 1):
                ET.SubElement(merge_cells, "{%s}mergeCell" % XLSX_NS, {"ref": "A%s:D%s" % (row_number, row_number)})

        dimension = first_child_by_local_name(root, "dimension")
        if dimension is not None and ":" in dimension.attrib.get("ref", ""):
            start_ref, end_ref = dimension.attrib["ref"].split(":", 1)
            dimension.attrib["ref"] = "%s:%s" % (start_ref, shift_row_reference(end_ref, extra_detail_rows))

        rows_by_number = {}
        for existing_row in sheet_data.findall("{%s}row" % XLSX_NS):
            try:
                rows_by_number[int(existing_row.attrib.get("r", "0") or "0")] = existing_row
            except Exception:
                continue
        detail_end_row += extra_detail_rows

    current_row_number = detail_start_row
    for detail in detail_rows:
        detail_row = row(current_row_number)
        description = invoice_line_field_label(detail)
        if not description:
            description = clean_name(detail.get("job_notes"))
        if detail.get("is_extra_line"):
            product = ""
        else:
            product = clean_name(detail.get("muck_type"))
        set_template_cell_value(detail_row, "A%s" % current_row_number, description)
        set_template_cell_value(detail_row, "E%s" % current_row_number, product)
        set_template_cell_value(detail_row, "F%s" % current_row_number, detail.get("tons", ""))
        set_template_cell_value(detail_row, "G%s" % current_row_number, detail.get("line_total", "") if str(detail.get("line_total", "")).strip() != "" else "")
        sort_template_row_cells(detail_row)
        current_row_number += 1

    while current_row_number <= detail_end_row:
        detail_row = row(current_row_number)
        set_template_cell_value(detail_row, "A%s" % current_row_number, "")
        set_template_cell_value(detail_row, "E%s" % current_row_number, "")
        set_template_cell_value(detail_row, "F%s" % current_row_number, "")
        set_template_cell_value(detail_row, "G%s" % current_row_number, "")
        sort_template_row_cells(detail_row)
        current_row_number += 1

    totals_start_row = detail_end_row + 1
    set_template_cell_formula(
        row(totals_start_row),
        "F%s" % totals_start_row,
        "SUM(F%s:F%s)" % (detail_start_row, detail_end_row),
        invoice.get("total_tons", ""),
    )
    set_template_cell_formula(
        row(totals_start_row + 1),
        "G%s" % (totals_start_row + 1),
        "SUM(G%s:G%s)" % (detail_start_row, detail_end_row),
        invoice.get("subtotal", ""),
    )
    set_template_cell_value(row(totals_start_row + 2), "F%s" % (totals_start_row + 2), "VAT 20%")
    set_template_cell_formula(
        row(totals_start_row + 2),
        "G%s" % (totals_start_row + 2),
        "G%s*20%%" % (totals_start_row + 1),
        invoice.get("vat_total", ""),
    )
    set_template_cell_formula(
        row(totals_start_row + 3),
        "G%s" % (totals_start_row + 3),
        "SUM(G%s:G%s)" % (totals_start_row + 1, totals_start_row + 2),
        invoice.get("grand_total", ""),
    )

    for target_row in [9, 10, 11, 12, 13, 17, totals_start_row, totals_start_row + 1, totals_start_row + 2, totals_start_row + 3]:
        sort_template_row_cells(row(target_row))

    ET.register_namespace("", XLSX_NS)
    ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
    ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
    worksheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    try:
        original_xml_text = sheet_bytes.decode("utf-8")
        generated_xml_text = worksheet_xml.decode("utf-8")
        original_match = re.search(r"<worksheet\b[^>]*>", original_xml_text)
        generated_match = re.search(r"<worksheet\b[^>]*>", generated_xml_text)
        if original_match and generated_match:
            generated_xml_text = (
                generated_xml_text[:generated_match.start()]
                + original_match.group(0)
                + generated_xml_text[generated_match.end():]
            )
            worksheet_xml = generated_xml_text.encode("utf-8")
    except Exception:
        pass

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            if name == "xl/worksheets/sheet1.xml":
                archive.writestr(name, worksheet_xml)
            elif name == "xl/styles.xml":
                archive.writestr(name, patch_currency_number_formats(data))
            else:
                archive.writestr(name, data)
    return output.getvalue()


def invoice_template_rows(invoice):
    scope_label = invoice_customer_label(invoice)
    if invoice.get("farm_name"):
        scope_label = "%s / %s" % (scope_label, invoice.get("farm_name"))

    rows = [
        ("title", ["Invoice"]),
        ("invoice_number", ["Invoice Number", invoice.get("invoice_number_label", "")]),
        ("invoice_date", ["Invoice Date", invoice.get("invoice_date_label", "")]),
        ("customer_scope", ["Customer / Farm", scope_label]),
        ("billing_email", ["Billing Email", invoice.get("customer_email", "")]),
        ("period", ["Period", "%s to %s" % (invoice.get("start_date_label", ""), invoice.get("end_date_label", ""))]),
        ("last_invoiced", ["Last Invoiced Through", invoice.get("last_invoiced_end_date_label", "")]),
        ("blank_meta", []),
    ]

    address_lines = invoice_address_lines(invoice)[:4]
    while len(address_lines) < 4:
        address_lines.append("")
    for line in address_lines:
        rows.append(("address", ["Billing Address", line] if line else []))

    rows.extend([
        ("blank_before_lines", []),
        ("headers", ["Date", "Farm", "Field", "Product", "Notes", "Tons", "Rate Per Ton", "VAT %", "Line Total"]),
    ])

    for row in invoice.get("line_rows", []):
        rows.append((
            "detail",
            [
                row.get("job_date_label", ""),
                row.get("farm_name", ""),
                invoice_line_field_label(row),
                row.get("muck_type", ""),
                row.get("job_notes", ""),
                row.get("tons", ""),
                row.get("rate_per_ton", ""),
                row.get("vat_rate", 0),
                row.get("line_total", 0),
            ],
        ))

    rows.extend([
        ("blank_before_totals", []),
        ("jobs", ["Jobs", invoice.get("job_count", 0)]),
        ("total_tons", ["Total Tons", invoice.get("total_tons", 0)]),
        ("subtotal", ["Subtotal", invoice.get("subtotal", 0)]),
        ("vat", ["VAT", invoice.get("vat_total", 0)]),
        ("grand_total", ["Grand Total", invoice.get("grand_total", 0)]),
    ])
    return rows


def build_template_based_invoice_xlsx(invoice, template):
    if invoice_template_is_layout_workbook(template):
        return fill_layout_invoice_template(invoice, template)

    row_templates = template.get("row_templates", {})
    row_specs = invoice_template_rows(invoice)
    detail_templates = row_templates.get("detail_rows") or [row_templates.get("headers", {})]
    address_templates = row_templates.get("address_rows") or [row_templates.get("billing_email", {})]
    detail_index = 0
    address_index = 0
    xml_rows = []
    row_number = 1

    for kind, values in row_specs:
        if kind == "detail":
            template_row = detail_templates[detail_index % len(detail_templates)]
            detail_index += 1
        elif kind == "address":
            template_row = address_templates[address_index % len(address_templates)]
            address_index += 1
        else:
            template_row = row_templates.get(kind, {})
        xml_rows.append(
            worksheet_row_xml(
                row_number,
                values,
                template_row.get("columns"),
                template_row.get("styles"),
                template_row.get("attrs"),
            )
        )
        row_number += 1

    cols_xml = ""
    if template.get("cols"):
        col_parts = []
        for col in template["cols"]:
            attrs = []
            for key, value in col.items():
                attrs.append('%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})))
            col_parts.append("<col %s/>" % " ".join(attrs))
        cols_xml = "<cols>%s</cols>" % "".join(col_parts)

    sheet_format_attrs = template.get("sheet_format_attrs", {}) or {"defaultRowHeight": "15"}
    sheet_format_xml = "<sheetFormatPr %s/>" % " ".join(
        '%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})) for key, value in sheet_format_attrs.items()
    )

    page_margins = template.get("page_margins", {})
    page_margins_xml = ""
    if page_margins:
        page_margins_xml = "<pageMargins %s/>" % " ".join(
            '%s="%s"' % (key, xml_escape(str(value), {'"': '&quot;'})) for key, value in page_margins.items()
        )

    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="%s">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  %s
  %s
  <sheetData>%s</sheetData>
  %s
</worksheet>
""" % (
        XLSX_NS,
        sheet_format_xml,
        cols_xml,
        "".join(xml_rows),
        page_margins_xml,
    )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in template.get("entries", {}).items():
            if name == "xl/worksheets/sheet1.xml":
                archive.writestr(name, worksheet_xml)
            else:
                archive.writestr(name, data)
    return output.getvalue()


def build_invoice_layout_xlsx_bytes(invoice):
    template = load_invoice_template()
    if template:
        try:
            return enforce_invoice_single_page_print_settings(build_template_based_invoice_xlsx(invoice, template))
        except Exception:
            pass
    return enforce_invoice_single_page_print_settings(build_plain_invoice_xlsx_bytes(invoice))


def build_invoice_xlsx_bytes(invoice):
    return build_invoice_layout_xlsx_bytes(invoice)


def invoice_pdf_filename(invoice):
    xlsx_name = str(invoice.get("filename", "invoice.xlsx") or "invoice.xlsx")
    if xlsx_name.lower().endswith(".xlsx"):
        return xlsx_name[:-5] + ".pdf"
    return xlsx_name + ".pdf"


def office_pdf_converter_path():
    for candidate in [
        os.environ.get("MUCKSPREADING_PDF_CONVERTER", ""),
        shutil.which("soffice") or "",
        shutil.which("libreoffice") or "",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/snap/bin/libreoffice",
    ]:
        candidate = str(candidate or "").strip()
        if candidate and os.path.exists(candidate):
            return candidate
    return ""


def convert_xlsx_bytes_to_pdf_bytes(xlsx_bytes, xlsx_name="invoice.xlsx"):
    converter = office_pdf_converter_path()
    if not converter or not xlsx_bytes:
        return b""

    safe_xlsx_name = sanitize_filename_part(xlsx_name) or "invoice.xlsx"
    if not safe_xlsx_name.lower().endswith(".xlsx"):
        safe_xlsx_name += ".xlsx"
    pdf_name = safe_xlsx_name[:-5] + ".pdf"

    with tempfile.TemporaryDirectory(prefix="muckspreading-pdf-") as temp_dir:
        xlsx_path = os.path.join(temp_dir, safe_xlsx_name)
        pdf_path = os.path.join(temp_dir, pdf_name)
        with open(xlsx_path, "wb") as handle:
            handle.write(xlsx_bytes)

        command = [
            converter,
            "--headless",
            "--convert-to",
            "pdf:calc_pdf_Export",
            "--outdir",
            temp_dir,
            xlsx_path,
        ]
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=60,
            )
        except Exception:
            return b""

        if completed.returncode != 0 or not os.path.exists(pdf_path):
            return b""

        try:
            with open(pdf_path, "rb") as handle:
                return handle.read()
        except OSError:
            return b""


def build_invoice_pdf_bytes(invoice, xlsx_bytes=None):
    layout_xlsx_bytes = build_invoice_layout_xlsx_bytes(invoice)
    converted_pdf = convert_xlsx_bytes_to_pdf_bytes(layout_xlsx_bytes, invoice.get("filename", "invoice.xlsx"))
    if converted_pdf:
        return converted_pdf

    page_width = 842
    page_height = 595
    left = 34
    top = 560
    row_height = 18
    usable_width = 774
    columns = [82, 88, 138, 102, 165, 52, 58, 89]
    headers = ["Date", "Farm", "Field", "Product", "Notes", "Tons", "Rate", "Line Total"]
    x_positions = [left]
    for width in columns[:-1]:
        x_positions.append(x_positions[-1] + width)

    def draw_text(x, y, text, font="F1", size=9.2, max_chars=48):
        return pdf_text_command(x, y, truncate_pdf_text(text, max_chars), font, size)

    customer_scope = invoice_customer_label(invoice)
    if invoice.get("farm_name"):
        customer_scope = "%s / %s" % (customer_scope, invoice.get("farm_name"))

    rows = []
    for line in invoice.get("line_rows", []):
        rows.append([
            line.get("job_date_label", ""),
            line.get("farm_name", ""),
            invoice_line_field_label(line),
            line.get("muck_type", ""),
            line.get("job_notes", ""),
            format_tons(line.get("tons", 0)) if str(line.get("tons", "")).strip() != "" else "",
            format_tons(line.get("rate_per_ton", 0)) if str(line.get("rate_per_ton", "")).strip() != "" else "",
            format_money(line.get("line_total", 0)),
        ])

    page_rows = []
    current = []
    for row in rows:
        if len(current) >= 18:
            page_rows.append(current)
            current = []
        current.append(row)
    if current:
        page_rows.append(current)
    if not page_rows:
        page_rows = [[["", "", "No uninvoiced jobs", "", "", "", "", ""]]]

    objects = []

    def add_object(payload):
        if isinstance(payload, str):
            payload = payload.encode("latin-1", "replace")
        objects.append(payload)
        return len(objects)

    font_regular_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    font_bold_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
    pages_id_placeholder = add_object(b"<<>>")
    page_ids = []

    for page_index, page_data in enumerate(page_rows):
        commands = []
        commands.append(draw_text(left, top, "Invoice %s" % invoice.get("invoice_number_label", ""), "F2", 18, 64))
        commands.append(draw_text(left + 180, top, customer_scope, "F2", 14, 64))
        commands.append(draw_text(left, top - 20, "Invoice Date: %s" % invoice.get("invoice_date_label", ""), "F1", 10, 64))
        commands.append(draw_text(left + 170, top - 20, "Period: %s to %s" % (invoice.get("start_date_label", ""), invoice.get("end_date_label", "")), "F1", 10, 64))

        address_y = top - 42
        for line in invoice_address_lines(invoice):
            commands.append(draw_text(left, address_y, line, "F1", 9.5, 72))
            address_y -= 13

        table_top = top - 112
        commands.append("0 g")
        commands.append("0.82 0.82 0.82 rg")
        commands.append("%.2f %.2f %.2f %.2f re f" % (left, table_top - row_height, usable_width, row_height))
        commands.append("0.65 G")
        commands.append("0.5 w")
        commands.append("%.2f %.2f %.2f %.2f re S" % (left, table_top - row_height, usable_width, row_height))
        for x_pos in x_positions[1:]:
            commands.append("%.2f %.2f m %.2f %.2f l S" % (x_pos, table_top - row_height, x_pos, table_top))
        for idx, header in enumerate(headers):
            commands.append(draw_text(x_positions[idx] + 4, table_top - 12, header, "F2", 8.4, 20))

        y = table_top - row_height
        for row in page_data:
            y_next = y - row_height
            commands.append("%.2f %.2f %.2f %.2f re S" % (left, y_next, usable_width, row_height))
            for x_pos in x_positions[1:]:
                commands.append("%.2f %.2f m %.2f %.2f l S" % (x_pos, y_next, x_pos, y))
            for idx, value in enumerate(row):
                text = str(value or "")
                x = x_positions[idx] + 4
                if idx >= 5:
                    approx = len(text) * 4.8
                    x = x_positions[idx] + columns[idx] - approx - 4
                    commands.append(draw_text(x, y_next + 6, text, "F1", 8.8, 16))
                else:
                    commands.append(draw_text(x, y_next + 6, text, "F1", 8.8, 34 if idx == 4 else 22))
            y = y_next

        if page_index == len(page_rows) - 1:
            totals_y = y - 18
            commands.append(draw_text(left + 470, totals_y, "Subtotal:", "F2", 10, 20))
            commands.append(draw_text(left + 610, totals_y, format_tons(invoice.get("subtotal", 0)), "F2", 10, 16))
            commands.append(draw_text(left + 470, totals_y - 16, "VAT:", "F2", 10, 20))
            commands.append(draw_text(left + 610, totals_y - 16, format_tons(invoice.get("vat_total", 0)), "F2", 10, 16))
            commands.append(draw_text(left + 470, totals_y - 34, "Grand Total:", "F2", 10.5, 20))
            commands.append(draw_text(left + 610, totals_y - 34, format_tons(invoice.get("grand_total", 0)), "F2", 10.5, 16))

        stream = "\n".join(commands).encode("latin-1", "replace")
        content_id = add_object(b"<< /Length %d >>\nstream\n%s\nendstream" % (len(stream), stream))
        page_id = add_object(
            (
                "<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] "
                "/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> /Contents %d 0 R >>"
            ) % (pages_id_placeholder, page_width, page_height, font_regular_id, font_bold_id, content_id)
        )
        page_ids.append(page_id)

    pages_payload = "<< /Type /Pages /Count %d /Kids [%s] >>" % (
        len(page_ids),
        " ".join("%d 0 R" % page_id for page_id in page_ids),
    )
    objects[pages_id_placeholder - 1] = pages_payload.encode("latin-1")
    catalog_id = add_object(("<< /Type /Catalog /Pages %d 0 R >>" % pages_id_placeholder).encode("latin-1"))

    pdf = io.BytesIO()
    pdf.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for index, payload in enumerate(objects, start=1):
        offsets.append(pdf.tell())
        pdf.write(("%d 0 obj\n" % index).encode("latin-1"))
        pdf.write(payload)
        pdf.write(b"\nendobj\n")

    xref_start = pdf.tell()
    pdf.write(("xref\n0 %d\n" % (len(objects) + 1)).encode("latin-1"))
    pdf.write(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.write(("%010d 00000 n \n" % offset).encode("latin-1"))
    pdf.write(
        (
            "trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF" % (
                len(objects) + 1,
                catalog_id,
                xref_start,
            )
        ).encode("latin-1")
    )
    return pdf.getvalue()


def invoice_email_subject(invoice, config):
    return render_invoice_template(load_app_settings().get("invoice_subject_template", DEFAULT_INVOICE_SUBJECT_TEMPLATE), invoice)


def invoice_email_body(invoice, audience):
    settings = load_app_settings()
    if audience == "accounts":
        return str(settings.get("invoice_accounts_message_template", DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE) or DEFAULT_INVOICE_ACCOUNTS_MESSAGE_TEMPLATE).strip()
    return str(settings.get("invoice_customer_message_template", DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE) or DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE).strip()


def save_invoice_archive(attachment_name, attachment_bytes):
    os.makedirs(INVOICE_ARCHIVE_DIR, exist_ok=True)
    file_path = os.path.join(INVOICE_ARCHIVE_DIR, attachment_name)
    with open(file_path, "wb") as handle:
        handle.write(attachment_bytes)
    return file_path


def invoice_archive_path(filename):
    name = str(filename or "").strip()
    if not name or os.path.basename(name) != name:
        return ""
    return os.path.join(INVOICE_ARCHIVE_DIR, name)


def invoice_history_row_at(ledger_index):
    ledger = load_invoice_ledger()
    try:
        wanted = int(ledger_index)
    except Exception:
        return None, -1, ledger
    if wanted < 0 or wanted >= len(ledger):
        return None, -1, ledger
    row = ledger[wanted]
    if not isinstance(row, dict):
        return None, -1, ledger
    return row, wanted, ledger


def invoice_jobs_for_history_row(history_row):
    if not isinstance(history_row, dict):
        return []
    wanted_ids = []
    for value in history_row.get("job_ids", []):
        try:
            wanted_ids.append(int(value))
        except Exception:
            continue
    if not wanted_ids:
        return []
    jobs_by_id = {}
    for row in load_jobs():
        try:
            job_id = int(row.get("id", 0) or 0)
        except Exception:
            continue
        if job_id:
            jobs_by_id[job_id] = dict(row)
    rows = [jobs_by_id[job_id] for job_id in wanted_ids if job_id in jobs_by_id]
    rows.sort(
        key=lambda row: (
            str(row.get("job_date", "")),
            clean_name(row.get("farm_name")).lower(),
            clean_name(row.get("field_name")).lower(),
            int(row.get("created_ts", 0) or 0),
        )
    )
    return rows


def invoice_history_form_values(history_row, ledger_index):
    if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
        return None
    stored_fee_rows = history_row.get("additional_fee_rows", [])
    if not isinstance(stored_fee_rows, list):
        stored_fee_rows = []
    invoice_date_value = str(history_row.get("invoice_date", "") or "").strip()
    if not invoice_date_value:
        try:
            invoice_date_value = datetime.fromtimestamp(int(history_row.get("created_ts", 0) or 0)).strftime("%Y-%m-%d")
        except Exception:
            invoice_date_value = datetime.now().strftime("%Y-%m-%d")
    return {
        "history_ledger_index": str(ledger_index),
        "customer": clean_name(history_row.get("customer")),
        "farm_name": clean_name(history_row.get("farm_name")),
        "customer_name_override": clean_name(history_row.get("customer_name_override")),
        "address_line_1_override": clean_name(history_row.get("address_line_1_override")),
        "address_line_2_override": clean_name(history_row.get("address_line_2_override")),
        "town_override": clean_name(history_row.get("town_override")),
        "postcode_override": clean_name(history_row.get("postcode_override")),
        "invoice_number": str(history_row.get("invoice_number", "") or "").strip(),
        "invoice_date": invoice_date_value,
        "job_date_from": str(history_row.get("job_date_from", "") or "").strip(),
        "payment_terms_days": str(history_row.get("payment_terms_days", "") or DEFAULT_APP_SETTINGS.get("invoice_default_payment_terms_days", "14")).strip(),
        "rate_override": str(history_row.get("rate_override", "") or "").strip(),
        "additional_fee_rows": [
            {
                "description": clean_name(row.get("description")),
                "amount": str(row.get("amount", "") or "").strip(),
            }
            for row in stored_fee_rows if isinstance(row, dict)
        ],
        "subject": str(history_row.get("subject_text", "") or "").strip(),
        "customer_message": str(history_row.get("customer_message", "") or "").strip(),
        "edit_reference_label": ("Invoice %s" % format_invoice_number(history_row.get("invoice_number", ""))).strip(),
    }


def record_invoice(invoice, accounts_emails, subject_text="", customer_message="", history_ledger_index=""):
    ledger = load_invoice_ledger()
    entry = {
        "invoice_number": int(invoice.get("invoice_number", 0) or 0),
        "customer": invoice.get("customer", ""),
        "customer_name_override": invoice.get("customer_name_override", ""),
        "farm_name": invoice.get("farm_name", ""),
        "address_line_1_override": invoice.get("address_line_1_override", ""),
        "address_line_2_override": invoice.get("address_line_2_override", ""),
        "town_override": invoice.get("town_override", ""),
        "postcode_override": invoice.get("postcode_override", ""),
        "customer_email": invoice.get("customer_email", ""),
        "customer_emails": list(invoice.get("customer_emails", [invoice.get("customer_email", "")]) or []),
        "accounts_emails": list(accounts_emails or []),
        "invoice_date": invoice.get("invoice_date", ""),
        "job_date_from": invoice.get("job_date_from", ""),
        "payment_terms_days": str(invoice.get("payment_terms_days", "") or ""),
        "rate_override": str(invoice.get("rate_override_input", "") or "").strip(),
        "additional_fee_rows": [
            {
                "description": clean_name(row.get("description")),
                "amount": str(row.get("amount", "") or "").strip(),
            }
            for row in invoice.get("additional_fee_input_rows", []) if isinstance(row, dict)
        ],
        "subject_text": str(subject_text or "").strip(),
        "customer_message": str(customer_message or "").strip(),
        "start_date": invoice.get("start_date", ""),
        "end_date": invoice.get("end_date", ""),
        "job_ids": list(invoice.get("job_ids", [])),
        "job_count": int(invoice.get("job_count", 0) or 0),
        "grand_total": invoice.get("grand_total", 0),
        "xlsx_filename": invoice.get("filename", ""),
        "pdf_filename": invoice_pdf_filename(invoice),
        "created_ts": int(time.time()),
    }
    try:
        wanted_index = int(str(history_ledger_index or "").strip())
    except Exception:
        wanted_index = -1
    if 0 <= wanted_index < len(ledger) and isinstance(ledger[wanted_index], dict) and not bool(ledger[wanted_index].get("manual_only", False)):
        existing_row = ledger[wanted_index]
        entry["created_ts"] = int(existing_row.get("created_ts", 0) or entry["created_ts"])
        entry["updated_ts"] = int(time.time())
        ledger[wanted_index] = entry
    else:
        ledger.append(entry)
    save_invoice_ledger(ledger)
    reserve_next_invoice_number(int(invoice.get("invoice_number", 0) or 0) + 1)


def invoice_status_map():
    status_by_job_id = {}
    for row in load_invoice_ledger():
        if not isinstance(row, dict):
            continue
        invoice_number = int(row.get("invoice_number", 0) or 0)
        manual_only = bool(row.get("manual_only", False))
        if invoice_number > 0:
            label = "Invoice %s" % format_invoice_number(invoice_number)
            status_key = "invoiced"
        elif manual_only:
            label = "Marked Invoiced"
            status_key = "manual"
        else:
            label = "Invoiced"
            status_key = "invoiced"
        for value in row.get("job_ids", []):
            try:
                job_id = int(value)
            except Exception:
                continue
            status_by_job_id[job_id] = {
                "label": label,
                "status_key": status_key,
                "manual_only": manual_only,
                "invoice_number": invoice_number,
            }
    return status_by_job_id


def invoice_history_rows():
    rows = []
    ledger = load_invoice_ledger()
    index = 0
    while index < len(ledger):
        row = ledger[index]
        if not isinstance(row, dict):
            index += 1
            continue
        invoice_number = int(row.get("invoice_number", 0) or 0)
        manual_only = bool(row.get("manual_only", False))
        pdf_filename = str(row.get("pdf_filename", "") or "").strip()
        xlsx_filename = str(row.get("xlsx_filename", "") or "").strip()
        rows.append({
            "ledger_index": index,
            "invoice_number": invoice_number,
            "reference_label": ("Invoice %s" % format_invoice_number(invoice_number)) if invoice_number > 0 else "Marked Invoiced",
            "type_label": "Manual" if manual_only else "Invoice",
            "customer": clean_name(row.get("customer")),
            "farm_name": clean_name(row.get("farm_name")) or "--",
            "period_label": "%s to %s" % (format_job_date(row.get("start_date")), format_job_date(row.get("end_date"))),
            "job_count": int(row.get("job_count", 0) or 0),
            "grand_total_label": format_money(row.get("grand_total", 0)) if not manual_only and str(row.get("grand_total", "")).strip() != "" else "--",
            "created_label": format_saved_time(row.get("created_ts")),
            "created_ts": int(row.get("created_ts", 0) or 0),
            "note": clean_name(row.get("note")),
            "can_edit": not manual_only and invoice_number > 0,
            "pdf_available": bool(pdf_filename),
            "xlsx_available": bool(xlsx_filename),
        })
        index += 1
    rows.sort(
        key=lambda item: (
            -item["created_ts"],
            -(item["invoice_number"] if item["invoice_number"] > 0 else 0),
        )
    )
    return rows


def invoice_unmark_options(customers=None):
    options = {}
    for customer in customers if isinstance(customers, list) else []:
        customer_name = clean_name(customer)
        if customer_name:
            options.setdefault(customer_name, {})
    for row in load_invoice_ledger():
        if not isinstance(row, dict):
            continue
        customer = clean_name(row.get("customer"))
        invoice_date = str(row.get("invoice_date", "") or row.get("end_date", "") or "").strip()
        if not customer or not invoice_date:
            continue
        options.setdefault(customer, {})[invoice_date] = {
            "value": invoice_date,
                "label": "%s - %s%s" % (
                format_invoice_date(invoice_date),
                ("Invoice " + format_invoice_number(row.get("invoice_number", ""))) if int(row.get("invoice_number", 0) or 0) else "Marked through",
                " (manual mark)" if bool(row.get("manual_only", False)) else "",
            ),
        }
    return {
        customer: sorted(dates.values(), key=lambda item: item["value"])
        for customer, dates in options.items()
    }


def manual_mark_jobs_invoiced(customer_name, farm_name="", through_date="", note=""):
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    through_date = str(through_date or "").strip()
    note = clean_name(note)
    if not customer_name:
        return 0, "Customer is required"
    try:
        datetime.strptime(through_date, "%Y-%m-%d")
    except Exception:
        return 0, "Job date through must be a valid date"

    eligible_rows = []
    for row in invoice_scope_jobs(customer_name, farm_name):
        row_job_date = str(row.get("job_date", "") or "").strip()
        if not row_job_date or row_job_date > through_date:
            continue
        eligible_rows.append(row)

    if not eligible_rows:
        return 0, "No uninvoiced jobs were found in that scope up to the selected date"

    ledger = load_invoice_ledger()
    ledger.append({
        "invoice_number": 0,
        "manual_only": True,
        "customer": customer_name,
        "farm_name": farm_name,
        "customer_email": "",
        "accounts_emails": [],
        "start_date": min(str(row.get("job_date", "") or "") for row in eligible_rows),
        "end_date": max(str(row.get("job_date", "") or "") for row in eligible_rows),
        "job_ids": [int(row.get("id", 0) or 0) for row in eligible_rows if int(row.get("id", 0) or 0)],
        "job_count": len(eligible_rows),
        "grand_total": "",
        "xlsx_filename": "",
        "pdf_filename": "",
        "created_ts": int(time.time()),
        "note": note or "Marked as already invoiced",
    })
    save_invoice_ledger(ledger)
    return len(eligible_rows), ""


def send_invoice_email(invoice, config, accounts_emails, subject_text="", customer_message="", history_ledger_index=""):
    customer_emails = normalize_email_list(invoice.get("customer_emails", [invoice.get("customer_email", "")]))
    if not customer_emails:
        raise RuntimeError("Customer email is missing for this invoice.")
    customer_email = customer_emails[0]

    xlsx_bytes = build_invoice_xlsx_bytes(invoice)
    pdf_bytes = build_invoice_pdf_bytes(invoice, xlsx_bytes)
    pdf_name = invoice_pdf_filename(invoice)
    subject_value = render_invoice_template(subject_text or invoice_email_subject(invoice, config), invoice)
    customer_message_value = render_invoice_template(customer_message or invoice_email_body(invoice, "customer"), invoice)
    accounts_message_value = render_invoice_template(invoice_email_body(invoice, "accounts"), invoice)
    from_email_value = invoice_from_email(config)
    envelope_from_value = str(config.get("smtp_username", "") or from_email_value).strip()

    customer_attachments = [
        (pdf_bytes, pdf_name, "application", "pdf"),
    ]
    accounts_attachments = [
        (pdf_bytes, pdf_name, "application", "pdf"),
        (xlsx_bytes, invoice.get("filename", "invoice.xlsx"), "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ]

    with smtplib.SMTP(config["smtp_host"], int(config["smtp_port"]), timeout=30) as server:
        server.ehlo()
        if config.get("use_tls", True):
            server.starttls()
            server.ehlo()
        if config.get("smtp_username"):
            server.login(config.get("smtp_username", ""), config.get("smtp_password", ""))

        send_invoice_smtp_batch(
            server,
            customer_emails,
            from_email_value,
            subject_value,
            customer_message_value,
            customer_attachments,
            require_all=True,
            envelope_from=envelope_from_value,
        )

        if accounts_emails:
            failed_accounts = send_invoice_smtp_batch(
                server,
                accounts_emails,
                from_email_value,
                subject_value,
                accounts_message_value,
                accounts_attachments,
                require_all=False,
                envelope_from=envelope_from_value,
            )
            if failed_accounts:
                app.logger.warning("Invoice accounts copy rejected by SMTP for: %s", ", ".join(failed_accounts))

    save_invoice_archive(invoice.get("filename", "invoice.xlsx"), xlsx_bytes)
    save_invoice_archive(pdf_name, pdf_bytes)
    record_invoice(invoice, accounts_emails, subject_text, customer_message, history_ledger_index)


def send_summary_email(summary, config, subject_line, body_text, xlsx_filename, pdf_filename):
    if not email_config_ready(config):
        raise RuntimeError("Email configuration is incomplete")

    msg = EmailMessage()
    msg["Subject"] = subject_line
    msg["From"] = config["from_email"]
    msg["To"] = ", ".join(config["to_emails"])
    msg.set_content(body_text)
    msg.add_attachment(
        build_xlsx_attachment_bytes(summary),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=xlsx_filename,
    )
    msg.add_attachment(
        build_pdf_attachment_bytes(summary),
        maintype="application",
        subtype="pdf",
        filename=pdf_filename,
    )

    with smtplib.SMTP(config["smtp_host"], int(config["smtp_port"]), timeout=30) as server:
        server.ehlo()
        if config.get("use_tls", True):
            server.starttls()
            server.ehlo()
        if config.get("smtp_username"):
            server.login(config.get("smtp_username", ""), config.get("smtp_password", ""))
        server.send_message(msg)


def send_weekly_summary_email(summary, config):
    send_summary_email(
        summary,
        config,
        weekly_email_subject(summary, config),
        weekly_email_body(summary),
        weekly_summary_attachment_filename(summary),
        weekly_summary_pdf_attachment_filename(summary),
    )


def send_monthly_summary_email(summary, config):
    send_summary_email(
        summary,
        config,
        monthly_email_subject(summary, config),
        monthly_email_body(summary),
        monthly_summary_attachment_filename(summary),
        monthly_summary_pdf_attachment_filename(summary),
    )


def maybe_send_weekly_summary(now=None):
    config = load_email_config()
    if not config.get("enabled", False):
        return {"ok": False, "reason": "disabled"}
    if not email_config_ready(config):
        return {"ok": False, "reason": "config_incomplete"}

    now = now or datetime.now()
    scheduled_at = now.replace(
        hour=int(config.get("send_hour", 7)),
        minute=int(config.get("send_minute", 0)),
        second=0,
        microsecond=0,
    )
    if now.weekday() != int(config.get("send_weekday", 0)):
        return {"ok": False, "reason": "not_scheduled_day"}
    if now < scheduled_at:
        return {"ok": False, "reason": "before_scheduled_time"}

    summary = weekly_jobs_summary()
    period_key = "%s_%s" % (summary["start_date"], summary["end_date"])
    state = load_email_state()
    if str(state.get("last_weekly_sent_period_key", "") or state.get("last_sent_period_key", "")) == period_key:
        return {"ok": False, "reason": "already_sent"}
    if summary_send_locked(state, "weekly_sending_period_key", period_key):
        return {"ok": False, "reason": "send_in_progress"}
    if not claim_summary_send_lock("weekly_sending_period_key", period_key):
        return {"ok": False, "reason": "send_in_progress"}

    try:
        send_weekly_summary_email(summary, config)
        existing_state = load_email_state()
        existing_state.update({
            "last_sent_period_key": period_key,
            "last_weekly_sent_period_key": period_key,
            "last_weekly_sent_at": int(time.time()),
            "last_weekly_summary_job_count": summary.get("job_count", 0),
        })
        existing_state.pop("weekly_sending_period_key", None)
        existing_state.pop("weekly_sending_period_key_at", None)
        save_email_state(existing_state)
        return {"ok": True, "summary": summary}
    except Exception:
        clear_summary_send_lock("weekly_sending_period_key", period_key)
        raise


def maybe_send_monthly_summary(now=None):
    config = load_email_config()
    if not config.get("enabled", False):
        return {"ok": False, "reason": "disabled"}
    if not config.get("monthly_enabled", config.get("enabled", False)):
        return {"ok": False, "reason": "monthly_disabled"}
    if not email_config_ready(config):
        return {"ok": False, "reason": "config_incomplete"}

    now = now or datetime.now()
    scheduled_at = now.replace(
        hour=int(config.get("monthly_send_hour", config.get("send_hour", 7))),
        minute=int(config.get("monthly_send_minute", config.get("send_minute", 0))),
        second=0,
        microsecond=0,
    )
    if not is_last_day_of_month(now.date()):
        return {"ok": False, "reason": "not_month_end"}
    if now < scheduled_at:
        return {"ok": False, "reason": "before_scheduled_time"}

    summary = monthly_jobs_summary(now)
    period_key = "%s_%s" % (summary["start_date"], summary["end_date"])
    state = load_email_state()
    if str(state.get("last_monthly_sent_period_key", "")) == period_key:
        return {"ok": False, "reason": "already_sent"}
    if summary_send_locked(state, "monthly_sending_period_key", period_key):
        return {"ok": False, "reason": "send_in_progress"}
    if not claim_summary_send_lock("monthly_sending_period_key", period_key):
        return {"ok": False, "reason": "send_in_progress"}

    try:
        send_monthly_summary_email(summary, config)
        save_email_state({
            **load_email_state(),
            "last_monthly_sent_period_key": period_key,
            "last_monthly_sent_at": int(time.time()),
            "last_monthly_summary_job_count": summary.get("job_count", 0),
            "monthly_sending_period_key": "",
            "monthly_sending_period_key_at": 0,
        })
        return {"ok": True, "summary": summary}
    except Exception:
        clear_summary_send_lock("monthly_sending_period_key", period_key)
        raise


def weekly_email_worker():
    while True:
        try:
            maybe_send_weekly_summary()
        except Exception:
            pass
        try:
            maybe_send_monthly_summary()
        except Exception:
            pass
        time.sleep(EMAIL_CHECK_INTERVAL_SECONDS)


def start_background_workers(debug_mode=False):
    if debug_mode and os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        return
    worker = threading.Thread(target=weekly_email_worker, daemon=True)
    worker.start()


def run_git_command(args, timeout_seconds=120):
    return subprocess.run(
        args,
        cwd=APP_ROOT,
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
    )


def git_update_message(result, fallback):
    text = clean_name(result.stdout) or clean_name(result.stderr)
    if not text:
        return fallback
    return text.splitlines()[0]


def app_version_label():
    git_dir = os.path.join(APP_ROOT, ".git")
    if not os.path.isdir(git_dir):
        return "no-git"
    result = run_git_command(["git", "rev-parse", "--short", "HEAD"], timeout_seconds=30)
    if result.returncode != 0:
        return "unknown"
    return clean_name(result.stdout) or "unknown"


def restart_app_process(delay_seconds=1.5):
    app_script = os.path.join(APP_ROOT, "app.py")

    def _restart():
        time.sleep(delay_seconds)
        try:
            subprocess.Popen(
                [sys.executable, app_script],
                cwd=APP_ROOT,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
        finally:
            os._exit(0)

    threading.Thread(target=_restart, daemon=True).start()


def parse_tons(raw_value, label):
    text = str(raw_value or "").strip()
    if not text:
        raise ValueError("%s is required" % label)
    try:
        value = float(text)
    except Exception:
        raise ValueError("%s must be a number" % label)
    if value < 0:
        raise ValueError("%s cannot be negative" % label)
    return round(value, 2)


def format_tons(value):
    try:
        number = float(value)
    except Exception:
        return "--"
    if abs(number - round(number)) < 0.000001:
        return str(int(round(number)))
    return ("%.2f" % number).rstrip("0").rstrip(".")


def parse_decimal_or_zero(value):
    try:
        return float(value or 0)
    except Exception:
        return 0.0


def parse_optional_decimal(value, label):
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return round(float(text), 2)
    except Exception:
        raise ValueError("%s must be a number" % label)


def format_money(value):
    try:
        return "£%0.2f" % float(value or 0)
    except Exception:
        return str(value or "")


def normalize_email_address(value):
    text = str(value or "").strip().replace("\\@", "@").replace("＠", "@")
    if text.lower().startswith("mailto:"):
        text = text[7:].strip()
    _, parsed_address = parseaddr(text)
    candidate = str(parsed_address or text).strip()
    if candidate.count("@") != 1 or "\\" in candidate or any(char.isspace() for char in candidate):
        return ""
    local_part, domain = candidate.rsplit("@", 1)
    if not local_part or not domain or domain.startswith(".") or domain.endswith("."):
        return ""
    return candidate


def normalize_email_list(values):
    unique = []
    seen = set()
    for value in values:
        for part in str(value or "").replace(";", ",").split(","):
            email = normalize_email_address(part)
            if not email:
                continue
            lowered = email.lower()
            if lowered in seen:
                continue
            seen.add(lowered)
            unique.append(email)
    return unique


def first_valid_email(values):
    for email in normalize_email_list(values):
        return email
    return ""


def build_invoice_email_message(to_email, from_email, subject_value, body_value, attachments):
    msg = EmailMessage()
    msg["Subject"] = subject_value
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(body_value)
    for attachment_bytes, attachment_name, maintype, subtype in attachments:
        msg.add_attachment(
            attachment_bytes,
            maintype=maintype,
            subtype=subtype,
            filename=attachment_name,
        )
    return msg


def smtp_recipient_refusal_is_transient(exception, recipient):
    refused = getattr(exception, "recipients", {}) or {}
    for refused_recipient, response in refused.items():
        if str(refused_recipient).lower() != str(recipient).lower():
            continue
        try:
            code = int(response[0] if isinstance(response, tuple) else response)
        except (TypeError, ValueError):
            return False
        return 400 <= code < 500
    return False


def send_invoice_smtp_batch(server, recipients, from_email, subject_value, body_value, attachments, require_all=True, envelope_from=""):
    recipients = normalize_email_list(recipients)
    if not recipients:
        return []

    if len(recipients) == 1:
        recipient = recipients[0]
        for attempt in range(3):
            msg = build_invoice_email_message(recipient, from_email, subject_value, body_value, attachments)
            try:
                server.send_message(msg, from_addr=envelope_from or from_email)
                return []
            except smtplib.SMTPRecipientsRefused as exc:
                if not smtp_recipient_refusal_is_transient(exc, recipient) or attempt == 2:
                    if require_all:
                        raise
                    failed = getattr(exc, "recipients", {}) or {}
                    return list(failed.keys())
                time.sleep(2)

    failed = []
    for recipient in recipients:
        sent = False
        for attempt in range(3):
            msg = build_invoice_email_message(recipient, from_email, subject_value, body_value, attachments)
            try:
                server.send_message(msg, from_addr=envelope_from or from_email)
                sent = True
                break
            except smtplib.SMTPRecipientsRefused as exc:
                if not smtp_recipient_refusal_is_transient(exc, recipient) or attempt == 2:
                    refused = getattr(exc, "recipients", {}) or {}
                    if recipient.lower() in {key.lower() for key in refused}:
                        failed.append(recipient)
                        break
                    raise
                time.sleep(2)
            except Exception:
                if require_all:
                    raise
                failed.append(recipient)
                break
        if not sent and recipient not in failed and require_all:
            raise RuntimeError("Invoice email was not accepted for %s." % recipient)
    if require_all and not failed and len(failed) == 0:
        return []
    return failed


def default_invoice_form(invoice_recipient_options, values=None):
    values = values if isinstance(values, dict) else {}
    settings = load_app_settings()
    default_invoice_date = datetime.now().strftime("%Y-%m-%d")
    descriptions = values.get("additional_fee_descriptions", [])
    amounts = values.get("additional_fee_amounts", [])
    if (not isinstance(descriptions, list) or not isinstance(amounts, list) or (not descriptions and not amounts)) and isinstance(values.get("additional_fee_rows"), list):
        fee_rows_input = [row for row in values.get("additional_fee_rows", []) if isinstance(row, dict)]
        descriptions = [str(row.get("description", "") or "").strip() for row in fee_rows_input]
        amounts = [str(row.get("amount", "") or "").strip() for row in fee_rows_input]
    if not isinstance(descriptions, list):
        descriptions = []
    if not isinstance(amounts, list):
        amounts = []
    fee_rows = []
    max_len = max(len(descriptions), len(amounts), 1)
    for index in range(max_len):
        fee_rows.append({
            "description": str(descriptions[index] if index < len(descriptions) else "" or "").strip(),
            "amount": str(amounts[index] if index < len(amounts) else "" or "").strip(),
        })
    return {
        "customer": clean_name(values.get("customer", "")),
        "farm_name": clean_name(values.get("farm_name", "")),
        "customer_name_override": clean_name(values.get("customer_name_override", "")),
        "address_line_1_override": clean_name(values.get("address_line_1_override", "")),
        "address_line_2_override": clean_name(values.get("address_line_2_override", "")),
        "town_override": clean_name(values.get("town_override", "")),
        "postcode_override": clean_name(values.get("postcode_override", "")),
        "invoice_number": str(values.get("invoice_number", "") or next_invoice_number()).strip(),
        "invoice_date": str(values.get("invoice_date", "") or default_invoice_date).strip(),
        "job_date_from": str(values.get("job_date_from", "") or "").strip(),
        "job_date_from_hint": str(values.get("job_date_from_hint", "") or "").strip(),
        "payment_terms_days": str(values.get("payment_terms_days", "") or settings.get("invoice_default_payment_terms_days", "14")).strip(),
        "rate_override": str(values.get("rate_override", "") or "").strip(),
        "additional_fee_rows": fee_rows,
        "subject": str(values.get("subject", "") or settings.get("invoice_subject_template", DEFAULT_INVOICE_SUBJECT_TEMPLATE)).strip(),
        "customer_message": str(values.get("customer_message", "") or settings.get("invoice_customer_message_template", DEFAULT_INVOICE_CUSTOMER_MESSAGE_TEMPLATE)).strip(),
        "history_ledger_index": str(values.get("history_ledger_index", "") or "").strip(),
        "edit_reference_label": str(values.get("edit_reference_label", "") or "").strip(),
    }


def invoice_form_from_request(req):
    return {
        "customer": clean_name(req.form.get("customer")),
        "farm_name": clean_name(req.form.get("farm_name")),
        "customer_name_override": clean_name(req.form.get("customer_name_override")),
        "address_line_1_override": clean_name(req.form.get("address_line_1_override")),
        "address_line_2_override": clean_name(req.form.get("address_line_2_override")),
        "town_override": clean_name(req.form.get("town_override")),
        "postcode_override": clean_name(req.form.get("postcode_override")),
        "invoice_number": str(req.form.get("invoice_number", "") or "").strip(),
        "invoice_date": str(req.form.get("invoice_date", "") or "").strip(),
        "job_date_from": str(req.form.get("job_date_from", "") or "").strip(),
        "payment_terms_days": str(req.form.get("payment_terms_days", "") or "").strip(),
        "rate_override": str(req.form.get("rate_override", "") or "").strip(),
        "additional_fee_descriptions": list(req.form.getlist("additional_fee_description")),
        "additional_fee_amounts": list(req.form.getlist("additional_fee_amount")),
        "subject": str(req.form.get("subject", "") or "").strip(),
        "customer_message": str(req.form.get("customer_message", "") or "").strip(),
        "history_ledger_index": str(req.form.get("history_ledger_index", "") or "").strip(),
    }


def build_invoice_from_form(invoice_form):
    if not invoice_form.get("customer"):
        return None, "Invoice customer is required"
    history_row = None
    history_index_text = str(invoice_form.get("history_ledger_index", "") or "").strip()
    jobs_override = None
    existing_invoice_number = None
    if history_index_text:
        history_row, history_index, _ = invoice_history_row_at(history_index_text)
        if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
            return None, "That invoice history entry could not be loaded."
        jobs_override = invoice_jobs_for_history_row(history_row)
        if not jobs_override:
            return None, "The original jobs for that invoice could not be found."
        existing_invoice_number = int(history_row.get("invoice_number", 0) or 0)
    invoice = build_invoice_payload(
        clean_name(history_row.get("customer")) if isinstance(history_row, dict) else invoice_form.get("customer", ""),
        clean_name(history_row.get("farm_name")) if isinstance(history_row, dict) else invoice_form.get("farm_name", ""),
        invoice_form.get("rate_override", ""),
        invoice_form.get("additional_fee_descriptions", []),
        invoice_form.get("additional_fee_amounts", []),
        invoice_form.get("invoice_number", ""),
        invoice_form.get("invoice_date", ""),
        invoice_form.get("job_date_from", ""),
        invoice_form.get("payment_terms_days", ""),
        jobs_override=jobs_override,
        existing_invoice_number=existing_invoice_number,
        customer_name_override=invoice_form.get("customer_name_override", ""),
        address_line_1_override=invoice_form.get("address_line_1_override", ""),
        address_line_2_override=invoice_form.get("address_line_2_override", ""),
        town_override=invoice_form.get("town_override", ""),
        postcode_override=invoice_form.get("postcode_override", ""),
    )
    if not invoice:
        return None, "No uninvoiced jobs were found for that customer/farm"
    if isinstance(invoice, dict) and invoice.get("error"):
        return None, str(invoice.get("error"))
    if history_index_text:
        invoice["history_ledger_index"] = history_index_text
    return invoice, ""


def build_invoice_preview(invoice, config, accounts_emails, subject_text="", customer_message=""):
    customer_emails = normalize_email_list(invoice.get("customer_emails", [invoice.get("customer_email", "")]))
    customer_email = customer_emails[0] if customer_emails else ""
    normalized_accounts = normalize_email_list(accounts_emails)
    customer_email_keys = {email.lower() for email in customer_emails}
    deduped_accounts = [email for email in normalized_accounts if email.lower() not in customer_email_keys]
    removed_accounts = [email for email in normalized_accounts if email.lower() in customer_email_keys]
    subject_value = render_invoice_template(subject_text or invoice_email_subject(invoice, config), invoice)
    customer_message_value = render_invoice_template(customer_message or invoice_email_body(invoice, "customer"), invoice)
    accounts_message_value = render_invoice_template(invoice_email_body(invoice, "accounts"), invoice)
    preview_lines = []
    for line in invoice.get("line_rows", []):
        preview_lines.append({
            "job_date_label": line.get("job_date_label", ""),
            "farm_name": line.get("farm_name", ""),
            "field_name": invoice_line_field_label(line),
            "muck_type": line.get("muck_type", ""),
            "tons_display": format_tons(line.get("tons", 0)) if str(line.get("tons", "")).strip() != "" else "",
            "rate_display": format_money(line.get("rate_per_ton", 0)) if str(line.get("rate_per_ton", "")).strip() != "" else "",
            "line_total_display": format_money(line.get("line_total", 0)),
        })
    return {
        "invoice": invoice,
        "customer_email": customer_email,
        "customer_emails": customer_emails,
        "accounts_emails": deduped_accounts,
        "removed_accounts": removed_accounts,
        "subject": subject_value,
        "customer_message": customer_message_value,
        "accounts_message": accounts_message_value,
        "rate_override": invoice.get("rate_override_label", ""),
        "grand_total_display": format_money(invoice.get("grand_total", 0)),
        "preview_lines": preview_lines,
        "additional_fee_rows": invoice.get("additional_fee_rows", []),
        "xlsx_filename": invoice.get("filename", "invoice.xlsx"),
        "pdf_filename": invoice_pdf_filename(invoice),
    }


def format_invoice_number(number):
    try:
        return str(int(number))
    except Exception:
        return str(number or "")


def sanitize_filename_part(value):
    cleaned = clean_name(value)
    if not cleaned:
        return ""
    safe = []
    for char in cleaned:
        if char.isalnum() or char in [" ", "-", "&", "_", "."]:
            safe.append(char)
    return "".join(safe).strip().strip(".")


def next_invoice_number():
    highest_issued = 0
    for invoice in load_invoice_ledger():
        try:
            highest_issued = max(highest_issued, int(invoice.get("invoice_number", 0) or 0))
        except Exception:
            continue
    state = load_invoice_state()
    try:
        number = int(state.get("next_invoice_number", 1) or 1)
    except Exception:
        number = 1
    return max(1, number, highest_issued + 1)


def reserve_next_invoice_number(next_number):
    save_invoice_state({"next_invoice_number": max(1, int(next_number or 1))})


def resolve_invoice_number(requested_value="", existing_invoice_number=None):
    requested_text = str(requested_value or "").strip()
    if not requested_text:
        return next_invoice_number()
    try:
        requested_number = int(requested_text)
    except Exception:
        raise ValueError("Invoice number must be a whole number.")
    if requested_number < 1:
        raise ValueError("Invoice number must be at least 1.")

    used_numbers = set()
    for invoice in load_invoice_ledger():
        try:
            value = int(invoice.get("invoice_number", 0) or 0)
        except Exception:
            continue
        if value > 0:
            used_numbers.add(value)

    if requested_number in used_numbers and requested_number != int(existing_invoice_number or 0):
        raise ValueError("Invoice number %s has already been used." % requested_number)
    return requested_number


def invoiced_job_ids():
    used_ids = set()
    for invoice in load_invoice_ledger():
        for value in invoice.get("job_ids", []):
            try:
                used_ids.add(int(value))
            except Exception:
                continue
    return used_ids


def invoice_scope_jobs(customer_name, farm_name="", job_date_from="", master_rows=None):
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    job_date_from = str(job_date_from or "").strip()
    if not customer_name:
        return []

    used_ids = invoiced_job_ids()
    rows = []
    master_rows = master_rows if isinstance(master_rows, list) else load_customer_master_rows()
    for row in load_jobs():
        if not job_matches_invoice_scope(row, customer_name, farm_name, master_rows):
            continue
        try:
            row_id = int(row.get("id", 0) or 0)
        except Exception:
            row_id = 0
        if row_id and row_id in used_ids:
            continue
        row_job_date = str(row.get("job_date", "") or "").strip()
        if job_date_from and row_job_date and row_job_date < job_date_from:
            continue
        rows.append(dict(row))

    rows.sort(
        key=lambda row: (
            str(row.get("job_date", "")),
            clean_name(row.get("farm_name")).lower(),
            clean_name(row.get("field_name")).lower(),
            int(row.get("created_ts", 0) or 0),
        )
    )
    return rows


def last_invoice_for_scope(customer_name, farm_name=""):
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    matches = []
    for row in load_invoice_ledger():
        if clean_name(row.get("customer")).lower() != customer_name.lower():
            continue
        if clean_name(row.get("farm_name")).lower() != farm_name.lower():
            continue
        matches.append(row)
    if not matches:
        return None
    matches.sort(key=lambda row: int(row.get("created_ts", 0) or 0), reverse=True)
    return matches[0]


def invoice_address_lines(invoice):
    lines = []
    lines.append(clean_name(invoice.get("display_customer_name") or invoice.get("customer")))
    for key in ["customer_address_line_1", "customer_address_line_2", "customer_town", "customer_postcode"]:
        text = clean_name(invoice.get(key))
        if text:
            lines.append(text)
    return lines


def invoice_customer_label(invoice):
    return clean_name(invoice.get("display_customer_name") or invoice.get("customer"))


def invoice_line_field_label(line):
    if not isinstance(line, dict):
        return ""
    if line.get("is_extra_line"):
        return clean_name(line.get("field_name"))
    farm_name = clean_name(line.get("farm_name"))
    field_name = clean_name(line.get("field_name"))
    if farm_name and field_name:
        return "%s - %s" % (farm_name, field_name)
    return field_name or farm_name


def invoice_default_vat_rate(jobs):
    for job in jobs:
        vat_rate = parse_decimal_or_zero(job.get("vat_rate"))
        if vat_rate >= 0:
            return round(vat_rate, 2)
    return 0.0


def parse_additional_fee_lines(descriptions, amounts):
    lines = []
    descriptions = descriptions if isinstance(descriptions, list) else []
    amounts = amounts if isinstance(amounts, list) else []
    max_len = max(len(descriptions), len(amounts))
    index = 0
    while index < max_len:
        description = clean_name(descriptions[index] if index < len(descriptions) else "")
        amount_text = str(amounts[index] if index < len(amounts) else "" or "").strip()
        if not description and not amount_text:
            index += 1
            continue
        if not description:
            raise ValueError("Additional fee description is required.")
        try:
            amount = round(float(amount_text), 2)
        except Exception:
            raise ValueError("Additional fee amount must be a number.")
        lines.append({
            "description": description,
            "amount": amount,
            "vat_rate": 20.0,
        })
        index += 1
    return lines


def customer_scope_farm_aliases(master_rows, customer_name):
    customer_name = clean_name(customer_name)
    aliases = []
    if not customer_name or not isinstance(master_rows, list):
        return aliases
    for row in master_rows:
        if not isinstance(row, dict):
            continue
        row_customer = clean_name(row.get("customer_name"))
        farm_name = clean_name(row.get("farm_name"))
        if row_customer.lower() == customer_name.lower() and farm_name and farm_name not in aliases:
            aliases.append(farm_name)
    return aliases


def job_matches_invoice_scope(job_row, customer_name, farm_name="", master_rows=None):
    if not isinstance(job_row, dict):
        return False
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    row_customer = clean_name(job_row.get("customer"))
    row_farm_name = clean_name(job_row.get("farm_name"))
    if not customer_name or not row_customer:
        return False

    farm_aliases = customer_scope_farm_aliases(master_rows or [], customer_name)

    if farm_name:
        if row_customer.lower() == customer_name.lower() and row_farm_name.lower() == farm_name.lower():
            return True
        if row_customer.lower() == farm_name.lower():
            return True
        return False

    if row_customer.lower() == customer_name.lower():
        return True
    return any(row_customer.lower() == alias.lower() for alias in farm_aliases)


def build_invoice_payload(customer_name, farm_name="", rate_override="", additional_fee_descriptions=None, additional_fee_amounts=None, invoice_number_override="", invoice_date_override="", job_date_from="", payment_terms_days="", jobs_override=None, existing_invoice_number=None, customer_name_override="", address_line_1_override="", address_line_2_override="", town_override="", postcode_override=""):
    customer_name = clean_name(customer_name)
    farm_name = clean_name(farm_name)
    job_date_from_text = str(job_date_from or "").strip()
    if job_date_from_text:
        try:
            datetime.strptime(job_date_from_text, "%Y-%m-%d")
        except Exception:
            return {"error": "Job date from must be a valid date."}
    if isinstance(jobs_override, list):
        jobs = [dict(row) for row in jobs_override if isinstance(row, dict)]
        if job_date_from_text:
            jobs = [
                row for row in jobs
                if not str(row.get("job_date", "") or "").strip() or str(row.get("job_date", "") or "").strip() >= job_date_from_text
            ]
        jobs.sort(
            key=lambda row: (
                str(row.get("job_date", "")),
                clean_name(row.get("farm_name")).lower(),
                clean_name(row.get("field_name")).lower(),
                int(row.get("created_ts", 0) or 0),
            )
        )
    else:
        jobs = invoice_scope_jobs(customer_name, farm_name, job_date_from_text)
    if not jobs:
        return None
    master_rows = load_customer_master_rows()
    scope_master_record = find_customer_master_record(master_rows, customer_name, farm_name)

    try:
        invoice_number = resolve_invoice_number(invoice_number_override, existing_invoice_number)
    except ValueError as exc:
        return {"error": str(exc)}

    payment_terms_text = str(payment_terms_days or "14").strip()
    if payment_terms_text not in INVOICE_PAYMENT_TERMS_OPTIONS:
        return {"error": "Payment terms must be one of: %s days." % ", ".join(INVOICE_PAYMENT_TERMS_OPTIONS)}
    payment_terms_value = int(payment_terms_text)

    invoice_date_text = str(invoice_date_override or datetime.now().strftime("%Y-%m-%d")).strip()
    try:
        datetime.strptime(invoice_date_text, "%Y-%m-%d")
    except Exception:
        return {"error": "Invoice date must be a valid date."}

    try:
        rate_override_value = parse_optional_decimal(rate_override, "Rate per ton override")
    except ValueError as exc:
        return {"error": str(exc)}

    start_date = min(str(job.get("job_date", "")) for job in jobs)
    end_date = max(str(job.get("job_date", "")) for job in jobs)
    default_rate = 0.0
    for job in jobs:
        job_farm_name = clean_name(job.get("farm_name"))
        job_master_record = find_customer_master_record(master_rows, customer_name, job_farm_name) or scope_master_record
        default_rate = parse_decimal_or_zero(job.get("rate_per_ton"))
        if default_rate <= 0 and isinstance(job_master_record, dict):
            default_rate = parse_decimal_or_zero(job_master_record.get("rate_per_ton"))
        if default_rate > 0:
            break
    total_tons = 0.0
    subtotal = 0.0
    vat_total = 0.0
    line_rows = []
    customer_emails = []
    customer_email = ""
    customer_address_line_1 = ""
    customer_address_line_2 = ""
    customer_town = ""
    customer_postcode = ""

    for job in jobs:
        job_farm_name = clean_name(job.get("farm_name"))
        job_master_record = find_customer_master_record(master_rows, customer_name, job_farm_name) or scope_master_record
        tons = parse_decimal_or_zero(job.get("total_spreader_tons"))
        saved_rate = parse_decimal_or_zero(job.get("rate_per_ton"))
        if saved_rate <= 0 and isinstance(job_master_record, dict):
            saved_rate = parse_decimal_or_zero(job_master_record.get("rate_per_ton"))
        rate = rate_override_value if rate_override_value is not None else saved_rate
        vat_rate = parse_decimal_or_zero(job.get("vat_rate"))
        if vat_rate <= 0 and isinstance(job_master_record, dict):
            vat_rate = parse_decimal_or_zero(job_master_record.get("vat_rate"))
        if rate <= 0:
            return {"error": "Rate per ton is missing for one or more uninvoiced jobs in this scope."}

        if not customer_emails and isinstance(job_master_record, dict):
            customer_emails = normalize_email_list([job_master_record.get("email", ""), job_master_record.get("email_2", "")])
        if not customer_emails:
            customer_emails = normalize_email_list([job.get("customer_email", "")])
        if not customer_address_line_1:
            customer_address_line_1 = clean_name(job.get("customer_address_line_1"))
        if not customer_address_line_1 and isinstance(job_master_record, dict):
            customer_address_line_1 = clean_name(job_master_record.get("address_line_1"))
        if not customer_address_line_2:
            customer_address_line_2 = clean_name(job.get("customer_address_line_2"))
        if not customer_address_line_2 and isinstance(job_master_record, dict):
            customer_address_line_2 = clean_name(job_master_record.get("address_line_2"))
        if not customer_town:
            customer_town = clean_name(job.get("customer_town"))
        if not customer_town and isinstance(job_master_record, dict):
            customer_town = clean_name(job_master_record.get("town"))
        if not customer_postcode:
            customer_postcode = clean_name(job.get("customer_postcode"))
        if not customer_postcode and isinstance(job_master_record, dict):
            customer_postcode = clean_name(job_master_record.get("postcode"))

        line_total = round(tons * rate, 2)
        line_vat = round(line_total * (vat_rate / 100.0), 2)
        total_tons += tons
        subtotal += line_total
        vat_total += line_vat
        line_rows.append({
            "job_id": int(job.get("id", 0) or 0),
            "job_date": str(job.get("job_date", "") or ""),
            "job_date_label": format_invoice_date(job.get("job_date")),
            "farm_name": clean_name(job.get("farm_name")),
            "field_name": clean_name(job.get("field_name")),
            "muck_type": clean_name(job.get("muck_type")),
            "job_notes": clean_name(job.get("job_notes")),
            "tons": round(tons, 2),
            "rate_per_ton": round(rate, 2),
            "vat_rate": round(vat_rate, 2),
            "line_total": line_total,
            "line_vat": line_vat,
            "is_extra_line": False,
        })

    if not customer_emails:
        master_record = scope_master_record
        if isinstance(master_record, dict):
            customer_emails = normalize_email_list([master_record.get("email", ""), master_record.get("email_2", "")])
            customer_address_line_1 = customer_address_line_1 or clean_name(master_record.get("address_line_1"))
            customer_address_line_2 = customer_address_line_2 or clean_name(master_record.get("address_line_2"))
            customer_town = customer_town or clean_name(master_record.get("town"))
            customer_postcode = customer_postcode or clean_name(master_record.get("postcode"))

    if not customer_emails:
        return {"error": "Customer email is missing for this customer/farm scope."}
    customer_email = customer_emails[0]

    customer_name_override = clean_name(customer_name_override)
    address_line_1_override = clean_name(address_line_1_override)
    address_line_2_override = clean_name(address_line_2_override)
    town_override = clean_name(town_override)
    postcode_override = clean_name(postcode_override)
    if customer_name_override:
        display_customer_name = customer_name_override
    else:
        display_customer_name = customer_name
    customer_address_line_1 = address_line_1_override or customer_address_line_1
    customer_address_line_2 = address_line_2_override or customer_address_line_2
    customer_town = town_override or customer_town
    customer_postcode = postcode_override or customer_postcode

    try:
        extra_lines = parse_additional_fee_lines(additional_fee_descriptions, additional_fee_amounts)
    except ValueError as exc:
        return {"error": str(exc)}

    for extra in extra_lines:
        line_total = round(extra["amount"], 2)
        line_vat = round(line_total * (extra["vat_rate"] / 100.0), 2)
        subtotal += line_total
        vat_total += line_vat
        line_rows.append({
            "job_id": 0,
            "job_date": "",
            "job_date_label": "",
            "farm_name": "",
            "field_name": extra["description"],
            "muck_type": "Additional Fee",
            "job_notes": "",
            "tons": "",
            "rate_per_ton": "",
            "vat_rate": round(extra["vat_rate"], 2),
            "line_total": line_total,
            "line_vat": line_vat,
            "is_extra_line": True,
        })

    label_parts = [customer_name]
    if farm_name:
        label_parts.append(farm_name)
    invoice_label = " - ".join(label_parts)
    filename_label = sanitize_filename_part(invoice_label) or "Invoice"
    filename = "%s - %s.xlsx" % (format_invoice_number(invoice_number), filename_label)
    last_invoice = last_invoice_for_scope(customer_name, farm_name)

    return {
        "invoice_number": invoice_number,
        "invoice_number_label": format_invoice_number(invoice_number),
        "payment_terms_days": payment_terms_value,
        "invoice_date": invoice_date_text,
        "invoice_date_label": format_invoice_date(invoice_date_text),
        "job_date_from": job_date_from_text,
        "job_date_from_label": format_job_date(job_date_from_text) if job_date_from_text else "",
        "customer": customer_name,
        "customer_name_override": customer_name_override,
        "display_customer_name": display_customer_name,
        "farm_name": farm_name,
        "customer_email": customer_email,
        "customer_emails": customer_emails,
        "customer_address_line_1": customer_address_line_1,
        "customer_address_line_2": customer_address_line_2,
        "customer_town": customer_town,
        "customer_postcode": customer_postcode,
        "address_line_1_override": address_line_1_override,
        "address_line_2_override": address_line_2_override,
        "town_override": town_override,
        "postcode_override": postcode_override,
        "start_date": start_date,
        "end_date": end_date,
        "start_date_label": format_invoice_date(start_date),
        "end_date_label": format_invoice_date(end_date),
        "last_invoiced_end_date_label": format_invoice_date(last_invoice.get("end_date")) if isinstance(last_invoice, dict) and last_invoice.get("end_date") else "First invoice for this scope",
        "line_rows": line_rows,
        "job_ids": [row["job_id"] for row in line_rows if row.get("job_id")],
        "job_count": len([row for row in line_rows if not row.get("is_extra_line")]),
        "additional_fee_count": len([row for row in line_rows if row.get("is_extra_line")]),
        "total_tons": round(total_tons, 2),
        "subtotal": round(subtotal, 2),
        "vat_total": round(vat_total, 2),
        "grand_total": round(subtotal + vat_total, 2),
        "default_rate_per_ton": round(default_rate, 2) if default_rate > 0 else "",
        "rate_override_input": str(rate_override or "").strip(),
        "rate_override_label": format_money(rate_override_value if rate_override_value is not None else default_rate),
        "additional_fee_input_rows": [
            {"description": extra.get("description", ""), "amount": ("%.2f" % float(extra.get("amount", 0) or 0))}
            for extra in extra_lines
        ],
        "additional_fee_rows": [
            {"description": extra.get("description", ""), "amount": format_money(extra.get("amount", ""))}
            for extra in extra_lines
        ],
        "filename": filename,
    }


def format_job_date(value):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").strftime("%d %b %Y")
    except Exception:
        return str(value or "--")


def format_invoice_date(value):
    try:
        return datetime.strptime(str(value or ""), "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return str(value or "--")


def format_saved_time(value):
    try:
        return datetime.fromtimestamp(int(value)).strftime("%d %b %Y %H:%M:%S")
    except Exception:
        return "--"


def build_context(invoice_form=None, invoice_preview=None, status_msg_override=None, status_ok_override=None, invoice_page=False):
    jobs = load_jobs()
    master_rows = load_customer_master_rows()
    customers = load_customers()
    farms = load_farms()
    invoice_recipient_options = load_email_recipient_options()
    invoice_form = default_invoice_form(invoice_recipient_options, invoice_form)
    muck_types = load_muck_types(master_rows)
    field_map = load_field_map()
    customer_farm_map = build_customer_farm_map(master_rows, jobs, field_map)
    customer_rate_map = build_customer_rate_map(master_rows)
    customer_invoice_from_map = build_customer_invoice_from_map()
    job_invoice_status_map = invoice_status_map()
    invoice_unmark_choices = invoice_unmark_options(customers)
    try:
        jobs_page = max(1, int(str(request.args.get("jobs_page", "1") or "1")))
    except Exception:
        jobs_page = 1
    recent_jobs_limit = jobs_page * 20
    all_fields = []
    recent_jobs = []
    now = datetime.now()
    today_iso = now.strftime("%Y-%m-%d")
    today_human = now.strftime("%d %B %Y")
    current_year = now.year
    today_job_count = 0
    today_spreader = 0.0
    today_john_deere = 0.0
    year_job_count = 0
    year_spreader = 0.0
    year_john_deere = 0.0
    for customer_fields in field_map.values():
        if not isinstance(customer_fields, dict):
            continue
        for field_names in customer_fields.values():
            if not isinstance(field_names, list):
                continue
            for field_name in field_names:
                if field_name not in all_fields:
                    all_fields.append(field_name)
    all_fields.sort(key=lambda item: item.lower())

    for row in master_rows:
        customer_name = row.get("customer_name", "")
        farm_name = clean_name(row.get("farm_name"))
        if customer_name and customer_name not in customers:
            customers.append(customer_name)
        if farm_name and farm_name not in farms:
            farms.append(farm_name)

    for job in jobs:
        farm_name = clean_name(job.get("farm_name"))
        if farm_name and farm_name not in farms:
            farms.append(farm_name)
        muck_type = clean_name(job.get("muck_type"))
        if muck_type and muck_type not in muck_types:
            muck_types.append(muck_type)
    farms.sort(key=lambda item: item.lower())
    muck_types.sort(key=lambda item: item.lower())
    customers.sort(key=lambda item: item.lower())

    for job in jobs[:recent_jobs_limit]:
        row = dict(job)
        try:
            row_id = int(row.get("id", 0) or 0)
        except Exception:
            row_id = 0
        status_info = job_invoice_status_map.get(row_id, {})
        row["job_date_label"] = format_job_date(row.get("job_date"))
        row["spreader_tons_label"] = format_tons(row.get("total_spreader_tons"))
        row["john_deere_tons_label"] = format_tons(row.get("total_john_deere_tons"))
        row["saved_label"] = format_saved_time(row.get("created_ts"))
        row["job_notes"] = clean_name(row.get("job_notes"))
        row["issue_photo_items"] = issue_photo_items(row.get("issue_photos"))[:4]
        row["issue_photo_count"] = len(normalize_issue_photo_names(row.get("issue_photos")))
        row["invoice_status_label"] = str(status_info.get("label", "Uninvoiced") or "Uninvoiced")
        row["invoice_status_key"] = str(status_info.get("status_key", "open") or "open")
        recent_jobs.append(row)

    for job in jobs:
        if str(job.get("job_date")) != today_iso:
            pass
        else:
            today_job_count += 1
            try:
                today_spreader += float(job.get("total_spreader_tons", 0) or 0)
            except Exception:
                pass
            try:
                today_john_deere += float(job.get("total_john_deere_tons", 0) or 0)
            except Exception:
                pass

        try:
            job_year = datetime.strptime(str(job.get("job_date", "")), "%Y-%m-%d").year
        except Exception:
            job_year = None
        if job_year == current_year:
            year_job_count += 1
            try:
                year_spreader += float(job.get("total_spreader_tons", 0) or 0)
            except Exception:
                pass
            try:
                year_john_deere += float(job.get("total_john_deere_tons", 0) or 0)
            except Exception:
                pass

    form_job = {
        "id": "",
        "job_date": "",
        "job_date_label": today_human,
        "customer": "",
        "farm_name": "",
        "field_name": "",
        "muck_type": "",
        "total_spreader_tons": "",
        "total_john_deere_tons": "",
        "job_notes": "",
        "issue_photo_items": [],
    }
    is_editing = False
    form_title = "New Job"
    form_submit_label = "Save Job"

    edit_id = str(request.args.get("edit_id", "") or "").strip()
    seed_job = None
    if edit_id:
        seed_job = find_job_by_id(edit_id)
        if seed_job:
            is_editing = True
            form_title = "Edit Job"
            form_submit_label = "Update Job"

    if isinstance(seed_job, dict):
        form_job = {
            "id": seed_job.get("id") if is_editing else "",
            "job_date": seed_job.get("job_date", "") if is_editing else today_iso,
            "job_date_label": format_job_date(seed_job.get("job_date", "")) if is_editing else today_human,
            "customer": seed_job.get("customer", ""),
            "farm_name": seed_job.get("farm_name", ""),
            "field_name": seed_job.get("field_name", ""),
            "muck_type": seed_job.get("muck_type", ""),
            "total_spreader_tons": format_tons(seed_job.get("total_spreader_tons")) if str(seed_job.get("total_spreader_tons", "")).strip() else "",
            "total_john_deere_tons": format_tons(seed_job.get("total_john_deere_tons")) if str(seed_job.get("total_john_deere_tons", "")).strip() else "",
            "job_notes": seed_job.get("job_notes", ""),
            "issue_photo_items": issue_photo_items(seed_job.get("issue_photos")),
        }

    default_week_start, default_week_end = previous_full_week_range(now)

    return {
        "today_iso": today_iso,
        "today_human": today_human,
        "app_version": app_version_label(),
        "invoice_page": bool(invoice_page),
        "customers": customers,
        "customers_json": json.dumps(customers),
        "farms": farms,
        "muck_types": muck_types,
        "muck_types_json": json.dumps(muck_types),
        "all_fields": all_fields,
        "all_farms_json": json.dumps(farms),
        "all_fields_json": json.dumps(all_fields),
        "customer_farm_map_json": json.dumps(customer_farm_map),
        "customer_rate_map_json": json.dumps(customer_rate_map),
        "customer_invoice_from_map_json": json.dumps(customer_invoice_from_map),
        "field_map_json": json.dumps(field_map),
        "recent_jobs": recent_jobs,
        "today_job_count": today_job_count,
        "today_spreader_tons": format_tons(today_spreader),
        "today_john_deere_tons": format_tons(today_john_deere),
        "year_job_count": year_job_count,
        "year_spreader_tons": format_tons(year_spreader),
        "year_john_deere_tons": format_tons(year_john_deere),
        "form_job": form_job,
        "form_title": form_title,
        "form_submit_label": form_submit_label,
        "is_editing": is_editing,
        "status_msg": str(status_msg_override if status_msg_override is not None else request.args.get("msg", "") or "").strip(),
        "status_ok": bool(status_ok_override) if status_msg_override is not None else str(request.args.get("ok", "1")) == "1",
        "invoice_form": invoice_form,
        "invoice_preview": invoice_preview,
        "invoice_payment_terms_options": INVOICE_PAYMENT_TERMS_OPTIONS,
        "invoice_recipient_options": invoice_recipient_options,
        "invoice_unmark_options": invoice_unmark_choices,
        "invoice_unmark_options_json": json.dumps(invoice_unmark_choices),
        "jobs_page": jobs_page,
        "recent_jobs_limit": recent_jobs_limit,
        "has_more_recent_jobs": len(jobs) > recent_jobs_limit,
        "data_dir": DATA_DIR,
        "total_jobs": len(jobs),
    }


@app.route("/favicon.ico")
@app.route("/favicon.svg")
def favicon():
    return Response(ICON, mimetype="image/svg+xml")


@app.route("/app-icon-<int:size>.png")
def app_icon_png(size):
    return Response(app_icon_bytes(size), mimetype="image/png")


@app.route("/site.webmanifest")
def web_manifest():
    icon_size = app_icon_dimensions()
    return jsonify({
        "name": "A. Farrell Contracting Muck Spreading Jobs",
        "short_name": APP_SHORT_NAME,
        "display": "standalone",
        "background_color": APP_THEME_COLOR,
        "theme_color": APP_THEME_COLOR,
        "start_url": url_for("home"),
        "icons": [
            {
                "src": url_for("app_icon_png", size=180),
                "sizes": "%sx%s" % (icon_size, icon_size),
                "type": "image/png",
            },
        ],
    })


@app.route("/")
def home():
    ensure_data_dir()
    return render_template_string(DASHBOARD_HTML)


@app.route("/muck")
def muck_home():
    ensure_data_dir()
    return render_template_string(HTML, **build_context(invoice_page=False))


@app.route("/straw")
def straw_home():
    ensure_data_dir()
    return render_template("straw_app.html", farms=load_farms(), crops=settings_names("Straw Crops", ["Wheat", "Barley", "Spring Barley", "Oats", "Hay"]))


def normalize_straw_state(value):
    value = value if isinstance(value, dict) else {}
    return {
        "fields": value.get("fields") if isinstance(value.get("fields"), list) else [],
        "stocktakes": value.get("stocktakes") if isinstance(value.get("stocktakes"), list) else [],
        "loads": value.get("loads") if isinstance(value.get("loads"), list) else [],
        "stockMovements": value.get("stockMovements") if isinstance(value.get("stockMovements"), list) else [],
    }


@app.route("/api/straw/state", methods=["GET", "PUT"])
def straw_state():
    if request.method in ("GET", "HEAD"):
        response = jsonify(normalize_straw_state(read_json_file(STRAW_STATE_PATH, {})))
        response.headers["Cache-Control"] = "no-store"
        return response
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="Please submit valid Straw app data."), 400
    current = normalize_straw_state(read_json_file(STRAW_STATE_PATH, {}))
    updated = normalize_straw_state(payload)
    updated["loads"] = current["loads"]
    write_json_atomic(STRAW_STATE_PATH, updated)
    return jsonify(ok=True)


def migrate_legacy_straw_loads(connection):
    state = normalize_straw_state(read_json_file(STRAW_STATE_PATH, {}))
    for index, load in enumerate(state["loads"]):
        if not isinstance(load, dict):
            continue
        migration_key = "legacy-load:%s" % clean_name(load.get("id") or index)
        if connection.execute("SELECT 1 FROM straw_migrations WHERE migration_key=?", (migration_key,)).fetchone():
            continue
        date_text = str(load.get("date", "") or "").strip()
        try:
            parsed = datetime.fromisoformat(date_text.replace("Z", "+00:00"))
        except ValueError:
            parsed = datetime.now()
        try:
            bale_total = max(0, int(round(float(load.get("bales", 0) or 0))))
        except (TypeError, ValueError):
            bale_total = 0
        weight_total = str(load.get("weight", "") or "").strip()
        try:
            weight_total = ("%.2f" % float(weight_total)).rstrip("0").rstrip(".") if weight_total else ""
        except ValueError:
            weight_total = ""
        registration = clean_name(load.get("vehicleReg")).upper() or "NOT RECORDED"
        customer = clean_name(load.get("customer")) or "Customer not recorded"
        updated_at = str(load.get("updatedAt", "") or datetime.now().isoformat(timespec="seconds"))
        connection.execute(
            "INSERT INTO straw_deliveries (customer, registration, delivery_date, delivery_time, bale_total, weight_total, version, updated_at) VALUES (?, ?, ?, ?, ?, ?, 1, ?)",
            (customer, registration, parsed.strftime("%Y-%m-%d"), parsed.strftime("%H:%M"), bale_total, weight_total, updated_at),
        )
        connection.execute("INSERT INTO straw_migrations (migration_key, migrated_at) VALUES (?, ?)", (migration_key, datetime.now().isoformat(timespec="seconds")))


def straw_delivery_connection():
    os.makedirs(DATA_DIR, exist_ok=True)
    connection = sqlite3.connect(os.path.join(DATA_DIR, "straw_deliveries.sqlite3"), timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("""CREATE TABLE IF NOT EXISTS straw_deliveries (
        id INTEGER PRIMARY KEY,
        customer TEXT NOT NULL COLLATE NOCASE,
        registration TEXT NOT NULL COLLATE NOCASE,
        delivery_date TEXT NOT NULL,
        delivery_time TEXT NOT NULL,
        bale_total INTEGER NOT NULL,
        weight_total TEXT NOT NULL,
        version INTEGER NOT NULL,
        updated_at TEXT NOT NULL
    )""")
    connection.execute("""CREATE TABLE IF NOT EXISTS straw_customers (
        name TEXT PRIMARY KEY COLLATE NOCASE,
        updated_at TEXT NOT NULL
    )""")
    connection.execute("""CREATE TABLE IF NOT EXISTS straw_migrations (
        migration_key TEXT PRIMARY KEY,
        migrated_at TEXT NOT NULL
    )""")
    with connection:
        migrate_legacy_straw_loads(connection)
    connection.execute(
        "INSERT OR IGNORE INTO straw_customers (name, updated_at) SELECT DISTINCT customer, updated_at FROM straw_deliveries WHERE TRIM(customer) <> '' AND customer <> 'Customer not recorded' COLLATE NOCASE"
    )
    return connection


def save_straw_customer(connection, customer):
    customer = clean_name(customer)
    if not customer:
        return
    connection.execute(
        "INSERT INTO straw_customers (name, updated_at) VALUES (?, ?) ON CONFLICT(name) DO UPDATE SET name=excluded.name, updated_at=excluded.updated_at",
        (customer, datetime.now().isoformat(timespec="seconds")),
    )


@app.route("/api/straw/customers", methods=["GET", "POST"])
def straw_customers():
    connection = straw_delivery_connection()
    try:
        if request.method in ("GET", "HEAD"):
            rows = connection.execute("SELECT name FROM straw_customers ORDER BY name COLLATE NOCASE").fetchall()
            response = jsonify(customers=[row["name"] for row in rows])
            response.headers["Cache-Control"] = "no-store"
            return response
        payload = request.get_json(silent=True)
        customer = clean_name(payload.get("customer")) if isinstance(payload, dict) else ""
        if not customer or len(customer) > 160:
            return jsonify(error="Enter a customer name."), 400
        with connection:
            save_straw_customer(connection, customer)
        return jsonify(customer=customer), 201
    finally:
        connection.close()


@app.route("/api/straw/deliveries", methods=["GET", "POST"])
@app.route("/api/straw/deliveries/<int:delivery_id>", methods=["PUT"])
def straw_deliveries(delivery_id=None):
    if request.method == "GET":
        connection = straw_delivery_connection()
        try:
            rows = connection.execute("SELECT * FROM straw_deliveries ORDER BY delivery_date DESC, delivery_time DESC, id DESC").fetchall()
            response = jsonify(deliveries=[dict(row) for row in rows])
            response.headers["Cache-Control"] = "no-store"
            return response
        finally:
            connection.close()

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="Please submit a valid delivered load."), 400
    customer = clean_name(payload.get("customer"))
    registration = clean_name(payload.get("registration")).upper()
    delivery_date = str(payload.get("delivery_date", "") or "").strip()
    delivery_time = str(payload.get("delivery_time", "") or "").strip()
    weight_total = str(payload.get("weight_total", "") or "").strip()
    try:
        bale_total = int(str(payload.get("bale_total", "") or "").strip())
    except ValueError:
        return jsonify(error="Enter the bale total as a whole number."), 400
    if not customer or len(customer) > 160:
        return jsonify(error="Enter a customer name."), 400
    if not registration or len(registration) > 30:
        return jsonify(error="Enter the lorry or trailer registration."), 400
    try:
        if datetime.strptime(delivery_date, "%Y-%m-%d").strftime("%Y-%m-%d") != delivery_date:
            raise ValueError()
    except ValueError:
        return jsonify(error="Choose a valid delivery date."), 400
    if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", delivery_time):
        return jsonify(error="Choose a valid delivery time."), 400
    if bale_total < 0:
        return jsonify(error="Bale total cannot be negative."), 400
    if weight_total:
        try:
            weight_number = float(weight_total)
            if weight_number < 0:
                raise ValueError()
            weight_total = ("%.2f" % weight_number).rstrip("0").rstrip(".")
        except ValueError:
            return jsonify(error="Weight total must be a positive number."), 400
    if delivery_id is not None and type(payload.get("version")) is not int:
        return jsonify(error="Reopen this delivered load before saving it again."), 400

    values = (customer, registration, delivery_date, delivery_time, bale_total, weight_total)
    connection = straw_delivery_connection()
    try:
        with connection:
            updated_at = datetime.now().isoformat(timespec="seconds")
            save_straw_customer(connection, customer)
            if delivery_id is None:
                cursor = connection.execute(
                    "INSERT INTO straw_deliveries (customer, registration, delivery_date, delivery_time, bale_total, weight_total, version, updated_at) VALUES (?, ?, ?, ?, ?, ?, 1, ?)",
                    values + (updated_at,),
                )
                delivery_id = cursor.lastrowid
            else:
                cursor = connection.execute(
                    "UPDATE straw_deliveries SET customer=?, registration=?, delivery_date=?, delivery_time=?, bale_total=?, weight_total=?, version=version+1, updated_at=? WHERE id=? AND version=?",
                    values + (updated_at, delivery_id, payload["version"]),
                )
                if cursor.rowcount != 1:
                    return jsonify(error="This delivered load changed elsewhere. Reopen it before saving again."), 409
            row = connection.execute("SELECT * FROM straw_deliveries WHERE id=?", (delivery_id,)).fetchone()
        return jsonify(delivery=dict(row)), 201 if request.method == "POST" else 200
    finally:
        connection.close()


def maintenance_connection():
    os.makedirs(DATA_DIR, exist_ok=True)
    connection = sqlite3.connect(os.path.join(DATA_DIR, "maintenance.sqlite3"), timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("""CREATE TABLE IF NOT EXISTS maintenance (
        id INTEGER PRIMARY KEY,
        maintenance_date TEXT NOT NULL,
        machinery TEXT NOT NULL COLLATE NOCASE,
        company TEXT NOT NULL,
        machine_hours TEXT NOT NULL,
        work_completed TEXT NOT NULL,
        parts_used TEXT NOT NULL,
        cost TEXT NOT NULL,
        completed_by TEXT NOT NULL,
        next_service_date TEXT NOT NULL,
        notes TEXT NOT NULL,
        version INTEGER NOT NULL,
        updated_at TEXT NOT NULL
    )""")
    return connection


def load_maintenance_entries():
    connection = maintenance_connection()
    try:
        return [dict(row) for row in connection.execute("SELECT * FROM maintenance ORDER BY maintenance_date DESC, id DESC").fetchall()]
    finally:
        connection.close()


@app.route("/maintenance")
def maintenance_home():
    machinery = settings_names("Machinery", [])
    for row in load_maintenance_entries():
        name = clean_name(row.get("machinery"))
        if name and name not in machinery:
            machinery.append(name)
    machinery.sort(key=lambda value: value.lower())
    return render_template(
        "maintenance.html",
        machinery=machinery,
    )


@app.route("/api/maintenance", methods=["GET", "POST"])
@app.route("/api/maintenance/<int:entry_id>", methods=["PUT"])
def maintenance_entries(entry_id=None):
    if request.method == "GET":
        response = jsonify(entries=load_maintenance_entries())
        response.headers["Cache-Control"] = "no-store"
        return response
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="Please submit a valid maintenance record."), 400
    keys = ("maintenance_date", "machinery", "company", "machine_hours", "work_completed", "parts_used", "cost", "completed_by", "next_service_date", "notes")
    values = {}
    for key in keys:
        value = payload.get(key, "")
        if not isinstance(value, str):
            return jsonify(error="Invalid maintenance field: " + key), 400
        values[key] = value.strip() if key not in ("work_completed", "parts_used", "notes") else value.strip()
    try:
        if datetime.strptime(values["maintenance_date"], "%Y-%m-%d").strftime("%Y-%m-%d") != values["maintenance_date"]:
            raise ValueError()
    except ValueError:
        return jsonify(error="Choose a valid maintenance date."), 400
    if not values["machinery"] or len(values["machinery"]) > 160:
        return jsonify(error="Enter the machinery name."), 400
    if not values["work_completed"] or len(values["work_completed"]) > 20000:
        return jsonify(error="Enter the maintenance work completed."), 400
    if len(values["completed_by"]) > 120:
        return jsonify(error="Completed by must be no more than 120 characters."), 400
    for key, label in (("machine_hours", "Machine hours"), ("cost", "Cost")):
        if values[key]:
            try:
                number = float(values[key])
                if number < 0:
                    raise ValueError()
                values[key] = ("%.2f" % number).rstrip("0").rstrip(".")
            except ValueError:
                return jsonify(error=label + " must be a positive number."), 400
    if values["next_service_date"]:
        try:
            if datetime.strptime(values["next_service_date"], "%Y-%m-%d").strftime("%Y-%m-%d") != values["next_service_date"]:
                raise ValueError()
        except ValueError:
            return jsonify(error="Choose a valid next service date."), 400
    if len(values["parts_used"]) > 20000 or len(values["notes"]) > 20000:
        return jsonify(error="Please keep parts and notes below 20,000 characters."), 400
    if entry_id is not None and type(payload.get("version")) is not int:
        return jsonify(error="Reopen this maintenance record before saving."), 400

    connection = maintenance_connection()
    try:
        with connection:
            params = tuple(values[key] for key in keys)
            updated_at = datetime.now().isoformat(timespec="seconds")
            if entry_id is None:
                cursor = connection.execute(
                    "INSERT INTO maintenance (maintenance_date, machinery, company, machine_hours, work_completed, parts_used, cost, completed_by, next_service_date, notes, version, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)",
                    params + (updated_at,),
                )
                entry_id = cursor.lastrowid
            else:
                cursor = connection.execute(
                    "UPDATE maintenance SET maintenance_date=?, machinery=?, company=?, machine_hours=?, work_completed=?, parts_used=?, cost=?, completed_by=?, next_service_date=?, notes=?, version=version+1, updated_at=? WHERE id=? AND version=?",
                    params + (updated_at, entry_id, payload["version"]),
                )
                if cursor.rowcount != 1:
                    return jsonify(error="This record changed elsewhere. Reopen it before saving again."), 409
            row = dict(connection.execute("SELECT * FROM maintenance WHERE id=?", (entry_id,)).fetchone())
        try:
            names = settings_names("Machinery", [])
            if values["machinery"] not in names:
                save_settings_names("Machinery", names + [values["machinery"]])
        except Exception:
            app.logger.warning("Maintenance saved, but machinery could not be added to settings.xlsx")
        return jsonify(entry=row), 201 if request.method == "POST" else 200
    finally:
        connection.close()


@app.route("/timesheet")
def timesheet_home():
    session.pop("timesheet_staff", None)
    staff = timesheet_staff_records()
    return render_template(
        "timesheet_staff.html",
        staff=staff,
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


@app.route("/timesheet/unlock/<path:staff_name>")
def timesheet_unlock_form(staff_name):
    selected = next((row for row in timesheet_staff_records() if row["name"].casefold() == clean_name(staff_name).casefold()), None)
    if not selected:
        return redirect(url_for("timesheet_home", ok=0, msg="That staff member is not available."))
    configured_pin = str(selected.get("pin", "") or "").strip()
    return render_template(
        "timesheet_pin.html",
        staff_name=selected["name"],
        status_msg="",
        pin_mode="unlock" if re.fullmatch(r"\d{4,8}", configured_pin) else "setup",
    )


@app.route("/timesheet/setup-pin", methods=["POST"])
def timesheet_setup_pin():
    staff_name = clean_name(request.form.get("staff_name"))
    pin = str(request.form.get("new_pin", "") or "").strip()
    confirmation = str(request.form.get("confirm_pin", "") or "").strip()
    selected = next((row for row in timesheet_staff_records() if row["name"].casefold() == staff_name.casefold()), None)
    if not selected:
        return redirect(url_for("timesheet_home", ok=0, msg="That staff member is not available."))
    if re.fullmatch(r"\d{4,8}", str(selected.get("pin", "") or "").strip()):
        return redirect(url_for("timesheet_unlock_form", staff_name=selected["name"]))
    if not re.fullmatch(r"\d{4,8}", pin):
        return render_template("timesheet_pin.html", staff_name=selected["name"], pin_mode="setup", status_msg="Choose a PIN containing 4–8 digits."), 400
    if pin != confirmation:
        return render_template("timesheet_pin.html", staff_name=selected["name"], pin_mode="setup", status_msg="The two PINs do not match."), 400
    try:
        save_timesheet_staff_pin(selected["name"], pin)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        return render_template("timesheet_pin.html", staff_name=selected["name"], pin_mode="setup", status_msg=str(exc)), 500
    session["timesheet_staff"] = selected["name"]
    return redirect(url_for("timesheet_calendar", ok=1, msg="Your PIN has been set."))


@app.route("/timesheet/unlock", methods=["POST"])
def timesheet_unlock():
    staff_name = clean_name(request.form.get("staff_name"))
    pin = str(request.form.get("pin", "") or "").strip()
    selected = next((row for row in timesheet_staff_records() if row["name"].casefold() == staff_name.casefold()), None)
    if not selected:
        return redirect(url_for("timesheet_home", ok=0, msg="That staff member is not available."))
    now = int(time.time())
    failures = session.get("timesheet_pin_failures", {})
    failure = failures.get(selected["name"].casefold(), {}) if isinstance(failures, dict) else {}
    locked_until = int(failure.get("locked_until", 0) or 0) if isinstance(failure, dict) else 0
    if locked_until > now:
        return render_template("timesheet_pin.html", staff_name=selected["name"], pin_mode="unlock", status_msg="Too many incorrect attempts. Try again in %s seconds." % (locked_until - now)), 429
    configured_pin = str(selected.get("pin", "") or "").strip()
    if not re.fullmatch(r"\d{4,8}", configured_pin):
        return redirect(url_for("timesheet_unlock_form", staff_name=selected["name"]))
    if not hmac.compare_digest(pin, configured_pin):
        count = int(failure.get("count", 0) or 0) + 1 if isinstance(failure, dict) else 1
        failures[selected["name"].casefold()] = {"count": count, "locked_until": now + 60 if count >= 5 else 0}
        session["timesheet_pin_failures"] = failures
        return render_template("timesheet_pin.html", staff_name=selected["name"], pin_mode="unlock", status_msg="Incorrect PIN."), 403
    failures.pop(selected["name"].casefold(), None)
    session["timesheet_pin_failures"] = failures
    session["timesheet_staff"] = selected["name"]
    return redirect(url_for("timesheet_calendar"))


@app.route("/timesheet/change-pin", methods=["GET", "POST"])
def timesheet_change_pin():
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    if request.method == "GET":
        return render_template("timesheet_pin.html", staff_name=staff_name, pin_mode="change", status_msg="")
    current_pin = str(request.form.get("current_pin", "") or "").strip()
    new_pin = str(request.form.get("new_pin", "") or "").strip()
    confirmation = str(request.form.get("confirm_pin", "") or "").strip()
    selected = next((row for row in timesheet_staff_records() if row["name"].casefold() == staff_name.casefold()), None)
    if not selected or not hmac.compare_digest(current_pin, str(selected.get("pin", "") or "").strip()):
        return render_template("timesheet_pin.html", staff_name=staff_name, pin_mode="change", status_msg="Your current PIN is incorrect."), 403
    if not re.fullmatch(r"\d{4,8}", new_pin):
        return render_template("timesheet_pin.html", staff_name=staff_name, pin_mode="change", status_msg="Choose a new PIN containing 4–8 digits."), 400
    if new_pin != confirmation:
        return render_template("timesheet_pin.html", staff_name=staff_name, pin_mode="change", status_msg="The two new PINs do not match."), 400
    try:
        save_timesheet_staff_pin(staff_name, new_pin)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        return render_template("timesheet_pin.html", staff_name=staff_name, pin_mode="change", status_msg=str(exc)), 500
    return redirect(url_for("timesheet_calendar", ok=1, msg="Your PIN has been changed."))


@app.route("/timesheet/calendar")
def timesheet_calendar():
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    return render_template(
        "timesheet.html",
        companies=settings_names("Companies", TIMESHEET_COMPANIES),
        staff_names=[staff_name],
        staff_name=staff_name,
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


def timesheet_time_options():
    options = []
    for minutes in range(0, 24 * 60, 15):
        hour = minutes // 60
        minute = minutes % 60
        suffix = "am" if hour < 12 else "pm"
        display_hour = hour % 12 or 12
        options.append({"value": "%02d:%02d" % (hour, minute), "label": "%d:%02d %s" % (display_hour, minute, suffix)})
    return options


def timesheet_duration_label(duration_minutes):
    hours, minutes = divmod(int(duration_minutes), 60)
    parts = []
    if hours:
        parts.append("%d hour%s" % (hours, "" if hours == 1 else "s"))
    if minutes or not parts:
        parts.append("%d minute%s" % (minutes, "" if minutes == 1 else "s"))
    return " ".join(parts)


def timesheet_entry_dict(row):
    entry = dict(row)
    entry.update({"duration_minutes": None, "duration_label": "", "finish_date": entry.get("date", ""), "crosses_midnight": False, "daily_segments": []})
    start_text = str(entry.get("start", "") or "")
    finish_text = str(entry.get("finish", "") or "")
    if not start_text or not finish_text:
        return entry
    try:
        start_at = datetime.strptime("%s %s" % (entry["date"], start_text), "%Y-%m-%d %H:%M")
        finish_at = datetime.strptime("%s %s" % (entry["date"], finish_text), "%Y-%m-%d %H:%M")
    except (KeyError, ValueError):
        return entry
    if finish_at < start_at:
        finish_at += timedelta(days=1)
        entry["crosses_midnight"] = True
    duration_minutes = int((finish_at - start_at).total_seconds() // 60)
    entry["duration_minutes"] = duration_minutes
    entry["duration_label"] = timesheet_duration_label(duration_minutes)
    entry["finish_date"] = finish_at.strftime("%Y-%m-%d")
    entry["finish_date_label"] = finish_at.strftime("%A %d %B %Y")
    if entry["crosses_midnight"]:
        start_minutes = start_at.hour * 60 + start_at.minute
        first_minutes = (24 * 60) - start_minutes
        second_minutes = finish_at.hour * 60 + finish_at.minute
        if first_minutes:
            entry["daily_segments"].append({
                "date": entry["date"], "start": start_text, "finish": "00:00",
                "duration_minutes": first_minutes, "duration_label": timesheet_duration_label(first_minutes),
            })
        if second_minutes:
            entry["daily_segments"].append({
                "date": entry["finish_date"], "start": "00:00", "finish": finish_text,
                "duration_minutes": second_minutes, "duration_label": timesheet_duration_label(second_minutes),
            })
    else:
        entry["daily_segments"].append({
            "date": entry["date"], "start": start_text, "finish": finish_text,
            "duration_minutes": duration_minutes, "duration_label": entry["duration_label"],
        })
    return entry


@app.route("/timesheet/day/<day>")
def timesheet_day(day):
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    try:
        parsed = datetime.strptime(day, "%Y-%m-%d")
        if parsed.strftime("%Y-%m-%d") != day:
            raise ValueError()
    except ValueError:
        return "Invalid timesheet date", 404
    return render_template(
        "timesheet.html", companies=settings_names("Companies", TIMESHEET_COMPANIES), staff_names=[staff_name], staff_name=staff_name, time_options=timesheet_time_options(), day=day,
        day_label=parsed.strftime("%A %d %B %Y"),
        calendar_url=url_for("timesheet_calendar", month=day[:7]),
    )


def timesheet_connection():
    os.makedirs(DATA_DIR, exist_ok=True)
    connection = sqlite3.connect(os.path.join(DATA_DIR, "timesheets.sqlite3"), timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("""CREATE TABLE IF NOT EXISTS timesheets (
        id INTEGER PRIMARY KEY,
        name TEXT NOT NULL COLLATE NOCASE,
        date TEXT NOT NULL,
        company TEXT NOT NULL,
        start TEXT NOT NULL,
        finish TEXT NOT NULL,
        notes TEXT NOT NULL,
        version INTEGER NOT NULL,
        updated_at TEXT NOT NULL,
        UNIQUE(name, date, company)
    )""")
    return connection


def parse_timesheet_month(month):
    try:
        start = datetime.strptime(str(month or ""), "%Y-%m")
        if start.strftime("%Y-%m") != month:
            raise ValueError()
    except (TypeError, ValueError):
        raise ValueError("Choose a valid month")
    following = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
    return start, following


def format_timesheet_hours(minutes):
    hours = float(minutes or 0) / 60.0
    return ("%.2f" % hours).rstrip("0").rstrip(".")


def timesheet_month_report(month, staff_name=""):
    start, following = parse_timesheet_month(month)
    connection = timesheet_connection()
    try:
        if staff_name:
            query = "SELECT * FROM timesheets WHERE name=? COLLATE NOCASE ORDER BY date, name, company"
            source_rows = connection.execute(query, (staff_name,)).fetchall()
        else:
            source_rows = connection.execute("SELECT * FROM timesheets ORDER BY date, name, company").fetchall()
        entries = [timesheet_entry_dict(row) for row in source_rows]
    finally:
        connection.close()
    rows = []
    totals = {}
    total_minutes = 0
    for entry in entries:
        segments = entry.get("daily_segments", [])
        if not segments and start.strftime("%Y-%m-%d") <= entry.get("date", "") < following.strftime("%Y-%m-%d"):
            segments = [{"date": entry["date"], "start": entry.get("start", ""), "finish": entry.get("finish", ""), "duration_minutes": 0, "duration_label": ""}]
        for segment in segments:
            segment_date = str(segment.get("date", "") or "")
            if not (start.strftime("%Y-%m-%d") <= segment_date < following.strftime("%Y-%m-%d")):
                continue
            minutes = int(segment.get("duration_minutes", 0) or 0)
            key = (entry.get("name", ""), entry.get("company", ""))
            totals[key] = totals.get(key, 0) + minutes
            total_minutes += minutes
            rows.append({
                "date": segment_date,
                "date_label": datetime.strptime(segment_date, "%Y-%m-%d").strftime("%a %d %b %Y"),
                "name": entry.get("name", ""),
                "company": entry.get("company", ""),
                "start": segment.get("start", ""),
                "finish": segment.get("finish", ""),
                "hours": format_timesheet_hours(minutes),
                "duration_label": segment.get("duration_label", ""),
                "shift_total": entry.get("duration_label", ""),
                "source_date": entry.get("date", ""),
                "continued": entry.get("date", "") != segment_date,
                "notes": entry.get("notes", ""),
            })
    rows.sort(key=lambda row: (row["date"], row["name"].lower(), row["company"].lower(), row["start"]))
    total_rows = [
        {"name": key[0], "company": key[1], "minutes": minutes, "hours": format_timesheet_hours(minutes), "duration_label": timesheet_duration_label(minutes)}
        for key, minutes in sorted(totals.items(), key=lambda item: (item[0][0].lower(), item[0][1].lower()))
    ]
    company_minutes = {}
    for row in rows:
        company_minutes[row["company"]] = company_minutes.get(row["company"], 0) + int(round(float(row["hours"] or 0) * 60))
    company_totals = [
        {"company": company, "minutes": minutes, "hours": format_timesheet_hours(minutes), "duration_label": timesheet_duration_label(minutes)}
        for company, minutes in sorted(company_minutes.items(), key=lambda item: item[0].lower())
    ]
    return {
        "staff_name": staff_name,
        "month": month,
        "month_label": start.strftime("%B %Y"),
        "start_date": start.strftime("%Y-%m-%d"),
        "end_date": (following - timedelta(days=1)).strftime("%Y-%m-%d"),
        "rows": rows,
        "totals": total_rows,
        "company_totals": company_totals,
        "total_minutes": total_minutes,
        "total_hours": format_timesheet_hours(total_minutes),
        "total_duration_label": timesheet_duration_label(total_minutes),
    }


def build_timesheet_month_xlsx(report):
    rows = [[report.get("staff_name", "")], ["Company Hours Summary"], ["Company", "Hours"]]
    for total in report["company_totals"]:
        rows.append([total["company"], total["hours"]])
    rows.extend([["Overall Total", report["total_hours"]], [""], ["Daily Timesheet Details"], ["Date", "Name", "Company", "Start", "Finish", "Hours on Day", "Full Shift Total", "Shift Started", "Job Details / Notes"]])
    for row in report["rows"]:
        rows.append([row["date"], row["name"], row["company"], row["start"], row["finish"], row["hours"], row["shift_total"], row["source_date"], row["notes"]])
    rows.extend([[""], ["Monthly Totals"], ["Name", "Company", "Hours"]])
    for total in report["totals"]:
        rows.append([total["name"], total["company"], total["hours"]])
    rows.append(["Overall", "", report["total_hours"]])
    return build_basic_xlsx_bytes("Timesheet", "%s Timesheet" % report["month_label"], rows, [14, 18, 30, 11, 11, 13, 16, 14, 55])


def timesheet_month_filename(report, extension="xlsx"):
    staff_part = re.sub(r"[^a-z0-9]+", "_", str(report.get("staff_name", "") or "").lower()).strip("_")
    return "timesheet_%s%s.%s" % ((staff_part + "_") if staff_part else "", report["month"], extension)


def send_timesheet_month_email(report):
    recipients = load_timesheet_email_recipients()
    if not recipients:
        raise RuntimeError("No active timesheet email recipients are selected in settings.xlsx")
    config = load_email_config()
    if not email_sender_ready(config):
        raise RuntimeError("Email sending settings are incomplete")
    xlsx_bytes = build_timesheet_month_xlsx(report)
    msg = EmailMessage()
    msg["Subject"] = "%s Timesheet - %s - %s" % (config.get("subject_prefix", "A. Farrell Contracting"), report.get("staff_name", "Staff"), report["month_label"])
    msg["From"] = config["from_email"]
    msg["To"] = ", ".join(recipients)
    company_lines = ["%s: %s" % (total["company"], total["duration_label"]) for total in report["company_totals"]]
    msg.set_content(
        "Please find attached the completed timesheet for %s for %s.\n\nCompany hours:\n%s\n\nOverall total: %s"
        % (report.get("staff_name", "Staff"), report["month_label"], "\n".join(company_lines) or "No recorded hours", report["total_duration_label"])
    )
    msg.add_attachment(xlsx_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=timesheet_month_filename(report))
    pdf_bytes = convert_xlsx_bytes_to_pdf_bytes(xlsx_bytes, timesheet_month_filename(report))
    if pdf_bytes:
        msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=timesheet_month_filename(report, "pdf"))
    with smtplib.SMTP(config["smtp_host"], int(config["smtp_port"]), timeout=30) as server:
        server.ehlo()
        if config.get("use_tls", True):
            server.starttls()
            server.ehlo()
        if config.get("smtp_username"):
            server.login(config.get("smtp_username", ""), config.get("smtp_password", ""))
        server.send_message(msg)
    return recipients


@app.route("/timesheet/month/<month>/export.xlsx")
def timesheet_month_export(month):
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    try:
        report = timesheet_month_report(month, staff_name)
    except ValueError:
        return "Invalid timesheet month", 404
    response = Response(build_timesheet_month_xlsx(report), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Content-Disposition"] = 'attachment; filename="%s"' % timesheet_month_filename(report)
    return response


@app.route("/timesheet/month/<month>/print")
def timesheet_month_print(month):
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    try:
        report = timesheet_month_report(month, staff_name)
    except ValueError:
        return "Invalid timesheet month", 404
    return render_template("timesheet_month.html", report=report)


@app.route("/timesheet/month/complete", methods=["POST"])
def timesheet_month_complete():
    staff_name = current_timesheet_staff()
    if not staff_name:
        return redirect(url_for("timesheet_home", ok=0, msg="Choose your name and enter your PIN."))
    month = str(request.form.get("month", "") or "").strip()
    try:
        report = timesheet_month_report(month, staff_name)
        if not report["rows"]:
            raise RuntimeError("There are no timesheet entries for %s" % report["month_label"])
        recipients = send_timesheet_month_email(report)
        state = read_json_file(TIMESHEET_MONTH_STATE_PATH, {})
        if not isinstance(state, dict):
            state = {}
        state["%s|%s" % (staff_name.casefold(), month)] = {"staff_name": staff_name, "month": month, "completed_at": datetime.now().isoformat(timespec="seconds"), "sent_to": recipients, "total_minutes": report["total_minutes"]}
        write_json_atomic(TIMESHEET_MONTH_STATE_PATH, state)
        return redirect(url_for("timesheet_calendar", month=month, ok=1, msg="%s timesheet for %s emailed to %s." % (report["month_label"], staff_name, ", ".join(recipients))))
    except (ValueError, RuntimeError, OSError, smtplib.SMTPException) as exc:
        return redirect(url_for("timesheet_calendar", month=month, ok=0, msg=str(exc)))


@app.route("/api/timesheets", methods=["GET", "POST"])
@app.route("/api/timesheets/<int:entry_id>", methods=["PUT"])
def timesheet_entries(entry_id=None):
    staff_name = current_timesheet_staff()
    if not staff_name:
        return jsonify(error="Unlock a staff timesheet first."), 401
    if request.method == "GET":
        connection = timesheet_connection()
        try:
            rows = connection.execute("SELECT * FROM timesheets WHERE name=? COLLATE NOCASE ORDER BY date DESC, name, company", (staff_name,)).fetchall()
            response = jsonify(entries=[timesheet_entry_dict(row) for row in rows])
            response.headers["Cache-Control"] = "no-store"
            return response
        finally:
            connection.close()

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="Please submit a valid timesheet."), 400
    values = {}
    for key in ("name", "date", "company", "start", "finish", "notes"):
        value = payload.get(key, "")
        if not isinstance(value, str):
            return jsonify(error="Invalid timesheet field: " + key), 400
        values[key] = value.strip() if key != "notes" else value
    values["name"] = staff_name
    if values["company"] not in settings_names("Companies", TIMESHEET_COMPANIES):
        return jsonify(error="Choose a company."), 400
    try:
        if datetime.strptime(values["date"], "%Y-%m-%d").strftime("%Y-%m-%d") != values["date"]:
            raise ValueError()
    except ValueError:
        return jsonify(error="Choose a valid date."), 400
    for key in ("start", "finish"):
        if values[key] and not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", values[key]):
            return jsonify(error="Enter a valid " + key + " time."), 400
        if values[key] and int(values[key].split(":", 1)[1]) % 15:
            return jsonify(error="Choose a %s time in 15-minute increments." % key), 400
    if len(values["notes"]) > 50000:
        return jsonify(error="Please keep daily notes below 50,000 characters."), 400
    if entry_id is not None and type(payload.get("version")) is not int:
        return jsonify(error="Reopen this timesheet before saving."), 400

    connection = timesheet_connection()
    try:
        with connection:
            params = tuple(values[key] for key in ("name", "date", "company", "start", "finish", "notes"))
            updated_at = datetime.now().isoformat(timespec="seconds")
            if entry_id is None:
                cursor = connection.execute(
                    "INSERT INTO timesheets (name, date, company, start, finish, notes, version, updated_at) VALUES (?, ?, ?, ?, ?, ?, 1, ?)",
                    params + (updated_at,),
                )
                entry_id = cursor.lastrowid
            else:
                cursor = connection.execute(
                    "UPDATE timesheets SET name=?, date=?, company=?, start=?, finish=?, notes=?, version=version+1, updated_at=? WHERE id=? AND version=? AND name=? COLLATE NOCASE",
                    params + (updated_at, entry_id, payload["version"], staff_name),
                )
                if cursor.rowcount != 1:
                    return jsonify(error="This timesheet changed elsewhere. Copy your latest notes, then reopen the saved day before updating it."), 409
            row = connection.execute("SELECT * FROM timesheets WHERE id=?", (entry_id,)).fetchone()
        return jsonify(entry=timesheet_entry_dict(row)), 201 if request.method == "POST" else 200
    except sqlite3.IntegrityError:
        return jsonify(error="A timesheet already exists for this name, date and company. Copy any new notes, then open the saved day below to continue it."), 409
    finally:
        connection.close()


@app.route("/invoice")
def invoice_home():
    ensure_data_dir()
    return render_template_string(HTML, **build_context(invoice_page=True))


@app.route("/invoice/history")
def invoice_history():
    ensure_data_dir()
    return render_template_string(
        INVOICE_HISTORY_HTML,
        history_rows=invoice_history_rows(),
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


def build_invoice_form_from_history_entry(history_row, ledger_index):
    values = invoice_history_form_values(history_row, ledger_index)
    if not isinstance(values, dict):
        return None
    fee_rows = values.get("additional_fee_rows", [])
    return {
        "history_ledger_index": values.get("history_ledger_index", ""),
        "customer": values.get("customer", ""),
        "farm_name": values.get("farm_name", ""),
        "customer_name_override": values.get("customer_name_override", ""),
        "address_line_1_override": values.get("address_line_1_override", ""),
        "address_line_2_override": values.get("address_line_2_override", ""),
        "town_override": values.get("town_override", ""),
        "postcode_override": values.get("postcode_override", ""),
        "invoice_number": values.get("invoice_number", ""),
        "invoice_date": values.get("invoice_date", ""),
        "job_date_from": values.get("job_date_from", ""),
        "payment_terms_days": values.get("payment_terms_days", ""),
        "rate_override": values.get("rate_override", ""),
        "additional_fee_descriptions": [str(row.get("description", "") or "") for row in fee_rows if isinstance(row, dict)],
        "additional_fee_amounts": [str(row.get("amount", "") or "") for row in fee_rows if isinstance(row, dict)],
        "subject": values.get("subject", ""),
        "customer_message": values.get("customer_message", ""),
        "edit_reference_label": values.get("edit_reference_label", ""),
    }


def build_invoice_from_history_entry(history_row, ledger_index):
    invoice_form = build_invoice_form_from_history_entry(history_row, ledger_index)
    if not isinstance(invoice_form, dict):
        return None, "That invoice history entry could not be loaded."
    return build_invoice_from_form(invoice_form)


def invoice_archive_response(filename, mimetype, download=False):
    file_path = invoice_archive_path(filename)
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        with open(file_path, "rb") as handle:
            payload = handle.read()
    except OSError:
        return None
    if mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        repaired_payload = enforce_invoice_single_page_print_settings(payload)
        if repaired_payload != payload:
            payload = repaired_payload
            save_invoice_archive(filename, payload)
    response = Response(payload, mimetype=mimetype)
    disposition = "attachment" if download else "inline"
    response.headers["Content-Disposition"] = '%s; filename="%s"' % (disposition, str(filename))
    return response


def rebuild_and_archive_history_invoice(history_row, ledger_index):
    invoice, error_message = build_invoice_from_history_entry(history_row, ledger_index)
    if error_message:
        return None, b"", b"", error_message
    xlsx_bytes = build_invoice_xlsx_bytes(invoice)
    pdf_bytes = build_invoice_pdf_bytes(invoice, xlsx_bytes)
    xlsx_filename = str(invoice.get("filename", "invoice.xlsx") or "invoice.xlsx")
    pdf_filename = invoice_pdf_filename(invoice)
    save_invoice_archive(xlsx_filename, xlsx_bytes)
    save_invoice_archive(pdf_filename, pdf_bytes)

    ledger = load_invoice_ledger()
    if 0 <= ledger_index < len(ledger) and isinstance(ledger[ledger_index], dict):
        ledger[ledger_index]["xlsx_filename"] = xlsx_filename
        ledger[ledger_index]["pdf_filename"] = pdf_filename
        save_invoice_ledger(ledger)
    return invoice, xlsx_bytes, pdf_bytes, ""


@app.route("/invoice/history/<int:ledger_index>/edit")
def invoice_history_edit(ledger_index):
    ensure_data_dir()
    history_row, actual_index, _ = invoice_history_row_at(ledger_index)
    if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
        return redirect(url_for("invoice_history", ok=0, msg="That invoice history entry could not be edited"))
    invoice_form = invoice_history_form_values(history_row, actual_index)
    return render_template_string(
        HTML,
        **build_context(
            invoice_form=invoice_form,
            invoice_page=True,
            status_msg_override="Editing %s." % ((("Invoice %s" % format_invoice_number(history_row.get("invoice_number", ""))).strip()) or "invoice"),
            status_ok_override=True,
        )
    )


@app.route("/invoice/history/<int:ledger_index>/show")
def invoice_history_show_pdf(ledger_index):
    ensure_data_dir()
    history_row, actual_index, _ = invoice_history_row_at(ledger_index)
    if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
        return redirect(url_for("invoice_history", ok=0, msg="That invoice PDF is not available"))
    invoice, _, pdf_bytes, error_message = rebuild_and_archive_history_invoice(history_row, actual_index)
    if error_message:
        return redirect(url_for("invoice_history", ok=0, msg=error_message))
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'inline; filename="%s"' % invoice_pdf_filename(invoice)
    return response


@app.route("/invoice/history/<int:ledger_index>/download.pdf")
def invoice_history_download_pdf(ledger_index):
    ensure_data_dir()
    history_row, actual_index, _ = invoice_history_row_at(ledger_index)
    if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
        return redirect(url_for("invoice_history", ok=0, msg="That invoice PDF is not available"))
    invoice, _, pdf_bytes, error_message = rebuild_and_archive_history_invoice(history_row, actual_index)
    if error_message:
        return redirect(url_for("invoice_history", ok=0, msg=error_message))
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'attachment; filename="%s"' % invoice_pdf_filename(invoice)
    return response


@app.route("/invoice/history/<int:ledger_index>/download")
def invoice_history_download_xlsx(ledger_index):
    ensure_data_dir()
    history_row, actual_index, _ = invoice_history_row_at(ledger_index)
    if not isinstance(history_row, dict) or bool(history_row.get("manual_only", False)):
        return redirect(url_for("invoice_history", ok=0, msg="That invoice workbook is not available"))
    invoice, xlsx_bytes, _, error_message = rebuild_and_archive_history_invoice(history_row, actual_index)
    if error_message:
        archived = invoice_archive_response(
            history_row.get("xlsx_filename", ""),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            download=True,
        )
        if archived is not None:
            return archived
        return redirect(url_for("invoice_history", ok=0, msg=error_message))
    response = Response(
        xlsx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = 'attachment; filename="%s"' % str(invoice.get("filename", "invoice.xlsx"))
    return response


@app.route("/settings")
def settings_home():
    ensure_data_dir()
    return render_template_string(
        SETTINGS_HTML,
        settings=load_app_settings(),
        email_config=load_email_config(),
        invoice_payment_terms_options=INVOICE_PAYMENT_TERMS_OPTIONS,
        app_version=app_version_label(),
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


@app.route("/settings/save", methods=["POST"])
def settings_save():
    ensure_data_dir()
    smtp_host = str(request.form.get("smtp_host", "") or "").strip()
    smtp_port = str(request.form.get("smtp_port", "") or "").strip()
    smtp_username = str(request.form.get("smtp_username", "") or "").strip()
    smtp_password = str(request.form.get("smtp_password", "") or "")
    smtp_from_email = str(request.form.get("smtp_from_email", "") or "").strip()
    smtp_use_tls = "1" if request.form.get("smtp_use_tls") == "1" else "0"
    invoice_from_email_value = str(request.form.get("invoice_from_email", "") or "").strip()
    invoice_subject_template = str(request.form.get("invoice_subject_template", "") or "").strip()
    invoice_customer_message_template = str(request.form.get("invoice_customer_message_template", "") or "").strip()
    invoice_accounts_message_template = str(request.form.get("invoice_accounts_message_template", "") or "").strip()
    invoice_default_payment_terms_days = str(request.form.get("invoice_default_payment_terms_days", "") or "").strip()

    if not invoice_from_email_value:
        return redirect(url_for("settings_home", ok=0, msg="Invoice from email is required"))
    if not invoice_subject_template:
        return redirect(url_for("settings_home", ok=0, msg="Invoice subject template is required"))
    if not invoice_customer_message_template:
        return redirect(url_for("settings_home", ok=0, msg="Customer message template is required"))
    if not invoice_accounts_message_template:
        return redirect(url_for("settings_home", ok=0, msg="Accounts message template is required"))
    if invoice_default_payment_terms_days not in INVOICE_PAYMENT_TERMS_OPTIONS:
        return redirect(url_for("settings_home", ok=0, msg="Default payment terms are invalid"))
    try:
        if not smtp_host or not smtp_username or not smtp_from_email or not (1 <= int(smtp_port) <= 65535):
            raise ValueError
    except ValueError:
        return redirect(url_for("settings_home", ok=0, msg="SMTP host, port, username, and from email are required"))

    save_app_settings({
        "invoice_from_email": invoice_from_email_value,
        "invoice_subject_template": invoice_subject_template,
        "invoice_customer_message_template": invoice_customer_message_template,
        "invoice_accounts_message_template": invoice_accounts_message_template,
        "invoice_default_payment_terms_days": invoice_default_payment_terms_days,
    })
    if not update_email_settings_csv({
        "smtp_host": smtp_host,
        "smtp_port": smtp_port,
        "smtp_username": smtp_username,
        "smtp_password": smtp_password,
        "from_email": smtp_from_email,
        "use_tls": smtp_use_tls,
    }):
        return redirect(url_for("settings_home", ok=0, msg="Could not save email sending settings"))
    return redirect(url_for("settings_home", ok=1, msg="Settings saved"))


@app.route("/invoice/mark-existing", methods=["POST"])
def invoice_mark_existing():
    ensure_data_dir()
    customer = clean_name(request.form.get("customer"))
    farm_name = clean_name(request.form.get("farm_name"))
    through_date = str(request.form.get("through_date", "") or "").strip()
    note = str(request.form.get("note", "") or "").strip()
    count, error_message = manual_mark_jobs_invoiced(customer, farm_name, through_date, note)
    if error_message:
        return redirect(url_for("invoice_home", ok=0, msg=error_message))
    scope_label = customer
    if farm_name:
        scope_label = "%s / %s" % (scope_label, farm_name)
    return redirect(url_for("invoice_home", ok=1, msg="Marked %s jobs as already invoiced for %s." % (count, scope_label)))


@app.route("/invoice/unmark-from-date", methods=["POST"])
def invoice_unmark_from_date():
    ensure_data_dir()
    customer = clean_name(request.form.get("customer"))
    invoice_date = str(request.form.get("invoice_date", "") or "").strip()
    try:
        datetime.strptime(invoice_date, "%Y-%m-%d")
    except ValueError:
        return redirect(url_for("invoice_home", ok=0, msg="Choose a valid invoice date"))
    ledger = load_invoice_ledger()
    kept = []
    removed = 0
    restored_jobs = 0
    for row in ledger:
        row_customer = clean_name(row.get("customer")) if isinstance(row, dict) else ""
        row_date = str(row.get("invoice_date", "") or row.get("end_date", "") or "").strip() if isinstance(row, dict) else ""
        if row_customer.lower() == customer.lower() and row_date >= invoice_date:
            removed += 1
            restored_jobs += len(row.get("job_ids", []))
            continue
        kept.append(row)
    if not removed:
        return redirect(url_for("invoice_home", ok=0, msg="No invoice records found for that customer from the selected date"))
    save_invoice_ledger(kept)
    return redirect(
        url_for(
            "invoice_home",
            ok=1,
            msg="Unmarked %s invoice records and made %s jobs available again from %s." % (
                removed,
                restored_jobs,
                format_invoice_date(invoice_date),
            ),
        )
    )


@app.route("/admin")
def admin_home():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    customers = load_customers()
    muck_types = load_muck_types(master_rows)
    for row in master_rows:
        customer_name = clean_name(row.get("customer_name"))
        if customer_name and customer_name not in customers:
            customers.append(customer_name)
    customers.sort(key=lambda item: item.lower())
    return render_template_string(
        ADMIN_HTML,
        admin_tree=build_customer_field_admin_map(master_rows, jobs, field_map),
        customers=customers,
        customers_json=json.dumps(customers),
        muck_types=muck_types,
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


@app.route("/admin/customer-details/save", methods=["POST"])
def admin_save_customer_details():
    ensure_data_dir()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=load_customer_master_rows(),
        customers=load_customers(),
        jobs=load_jobs(),
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    customer_email = str(request.form.get("customer_email", "") or "").strip()
    rate_per_ton = str(request.form.get("rate_per_ton", "") or "").strip()
    ok, message = save_customer_master_customer_details(customer, customer_email, rate_per_ton)
    if not ok:
        return redirect(url_for("admin_home", ok=0, msg=message))
    return redirect(url_for("admin_home", ok=1, msg="%s customer details saved" % customer))


@app.route("/admin/customer-master-row/save", methods=["POST"])
def admin_save_customer_master_row():
    ensure_data_dir()
    customer_name = clean_name(request.form.get("customer_name"))
    row_number = str(request.form.get("row_number", "") or "").strip()
    field_values = {
        "customer_name": customer_name,
        "farm_name": clean_name(request.form.get("farm_name")),
        "email": str(request.form.get("email", "") or "").strip(),
        "address_line_1": clean_name(request.form.get("address_line_1")),
        "address_line_2": clean_name(request.form.get("address_line_2")),
        "town": clean_name(request.form.get("town")),
        "postcode": clean_name(request.form.get("postcode")),
        "rate_per_ton": str(request.form.get("rate_per_ton", "") or "").strip(),
        "vat_rate": str(request.form.get("vat_rate", "") or "").strip(),
        "active": "1" if request.form.get("active") == "1" else "0",
        "muck_type": clean_name(request.form.get("muck_type")),
    }
    ok, message = save_customer_master_customer_details(
        customer_name,
        field_values["email"],
        field_values["rate_per_ton"],
        row_number=row_number,
        field_values=field_values,
    )
    return redirect(url_for("admin_home", ok=1 if ok else 0, msg=message if not ok else "%s customer row saved" % customer_name))


@app.route("/jobs/save", methods=["POST"])
def save_job():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    edit_job_id = str(request.form.get("edit_job_id", "") or "").strip()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=load_jobs(),
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    farm_name = clean_name(request.form.get("farm_name"))
    field_name = clean_name(request.form.get("field_name"))
    muck_type = clean_name(request.form.get("muck_type"))
    job_notes = clean_name(request.form.get("job_notes"))
    today_iso = datetime.now().strftime("%Y-%m-%d")
    job_date = str(request.form.get("job_date", "") or "").strip()

    if not customer:
        return redirect(url_for("home", ok=0, msg="Customer is required"))
    if not field_name:
        return redirect(url_for("home", ok=0, msg="Field name is required"))
    if not muck_type:
        return redirect(url_for("home", ok=0, msg="Muck type is required"))
    try:
        datetime.strptime(job_date, "%Y-%m-%d")
    except ValueError:
        return redirect(url_for("home", ok=0, msg="Job date must be a valid date"))

    try:
        spreader_tons = parse_tons(request.form.get("spreader_tons"), "Total spreader tons")
        john_deere_tons = parse_tons(request.form.get("john_deere_tons"), "Total Ops Center tons")
    except ValueError as exc:
        return redirect(url_for("home", ok=0, msg=str(exc)))

    existing_job = None
    if edit_job_id:
        existing_job = find_job_by_id(edit_job_id)
        if not existing_job:
            return redirect(url_for("home", ok=0, msg="Saved job could not be found for editing"))

    record = {
        "id": int(existing_job.get("id")) if isinstance(existing_job, dict) else int(time.time() * 1000),
        "job_date": job_date,
        "customer": customer,
        "farm_name": farm_name,
        "field_name": field_name,
        "muck_type": muck_type,
        "job_notes": job_notes,
        "total_spreader_tons": spreader_tons,
        "total_john_deere_tons": john_deere_tons,
        "created_ts": int(existing_job.get("created_ts")) if isinstance(existing_job, dict) and str(existing_job.get("created_ts", "")).strip() else int(time.time()),
        "updated_ts": int(time.time()),
        "issue_photos": normalize_issue_photo_names(existing_job.get("issue_photos")) if isinstance(existing_job, dict) else [],
    }

    apply_customer_master_snapshot(record, master_rows, customer, farm_name)

    uploaded_photos = save_issue_photo_uploads(record["id"], request.files.getlist("issue_photos"))
    if uploaded_photos:
        record["issue_photos"] = normalize_issue_photo_names(record.get("issue_photos", []) + uploaded_photos)

    upsert_job(record)

    customers = load_customers()
    if customer not in customers:
        customers.append(customer)
        save_customers(customers)

    farms = load_farms()
    if farm_name and farm_name not in farms:
        farms.append(farm_name)
        save_farms(farms)

    muck_types = load_muck_types()
    if muck_type not in muck_types:
        muck_types.append(muck_type)
        save_muck_types(muck_types)

    field_map = load_field_map()
    customer_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    farm_bucket = customer_bucket.get(farm_name, [])
    if field_name not in farm_bucket:
        farm_bucket.append(field_name)
        farm_bucket.sort(key=lambda item: item.lower())
        customer_bucket[farm_name] = farm_bucket
        field_map[customer] = customer_bucket
        save_field_map(field_map)

    action_label = "Updated" if existing_job else "Saved"
    return redirect(url_for("home", ok=1, msg="%s job for %s - %s." % (action_label, customer, field_name)))


@app.route("/jobs/delete/<int:job_id>", methods=["POST"])
def delete_job(job_id):
    ensure_data_dir()
    if not delete_job_by_id(job_id):
        return redirect(url_for("home", ok=0, msg="Saved job could not be found"))
    return redirect(url_for("home", ok=1, msg="Deleted saved job"))


@app.route("/jobs/issue-photos/<path:filename>")
def job_issue_photo(filename):
    safe_name = sanitize_issue_photo_filename(filename)
    if not safe_name:
        return Response(status=404)
    file_path = os.path.join(ISSUE_PHOTOS_DIR, safe_name)
    if not os.path.isfile(file_path):
        return Response(status=404)
    mimetype, _ = mimetypes.guess_type(file_path)
    try:
        with open(file_path, "rb") as handle:
            data = handle.read()
    except OSError:
        return Response(status=404)
    response = Response(data, mimetype=mimetype or "application/octet-stream")
    response.headers["Cache-Control"] = "no-store"
    return response


@app.route("/admin/jobs/move", methods=["POST"])
def admin_move_jobs():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    from_customer = canonical_customer_name(
        request.form.get("from_customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    to_customer = canonical_customer_name(
        request.form.get("to_customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    from_farm_name = clean_name(request.form.get("from_farm_name"))
    to_farm_name = clean_name(request.form.get("to_farm_name"))
    field_name = clean_name(request.form.get("field_name"))

    if not from_customer or not to_customer:
        return redirect(url_for("admin_home", ok=0, msg="Both customer names are required"))
    if normalized_name_key(from_customer) == normalized_name_key(to_customer):
        return redirect(url_for("admin_home", ok=0, msg="The customer would not change"))

    updated_jobs = 0
    for row in jobs:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) != normalized_name_key(from_customer):
            continue
        if from_farm_name and normalized_name_key(row.get("farm_name")) != normalized_name_key(from_farm_name):
            continue
        if field_name and normalized_name_key(row.get("field_name")) != normalized_name_key(field_name):
            continue

        row["customer"] = to_customer
        if to_farm_name:
            row["farm_name"] = to_farm_name
        apply_customer_master_snapshot(row, master_rows, to_customer, row.get("farm_name", ""))
        updated_jobs += 1

    if updated_jobs:
        save_jobs(jobs)

    customers = load_customers()
    if to_customer not in customers:
        customers.append(to_customer)
        save_customers(customers)

    return redirect(
        url_for(
            "admin_home",
            ok=1,
            msg="Moved %s jobs from %s to %s." % (updated_jobs, from_customer, to_customer),
        )
    )


@app.route("/admin/customers/merge", methods=["POST"])
def admin_merge_customers():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    invoice_ledger = load_invoice_ledger()
    from_customer = canonical_customer_name(
        request.form.get("from_customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    to_customer = canonical_customer_name(
        request.form.get("to_customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    if not from_customer or not to_customer:
        return redirect(url_for("admin_home", ok=0, msg="Both customer names are required"))
    if normalized_name_key(from_customer) == normalized_name_key(to_customer):
        return redirect(url_for("admin_home", ok=0, msg="Choose a different destination customer"))

    try:
        summary = merge_customer_into_customer(from_customer, to_customer, master_rows, jobs, field_map, invoice_ledger)
    except ValueError as exc:
        return redirect(url_for("admin_home", ok=0, msg=str(exc)))

    return redirect(
        url_for(
            "admin_home",
            ok=1,
            msg="Merged %s into %s (%s jobs, %s invoice records)." % (
                from_customer,
                to_customer,
                summary.get("updated_jobs", 0),
                summary.get("updated_invoices", 0),
            ),
        )
    )


@app.route("/admin/farms/merge", methods=["POST"])
def admin_merge_farms():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    invoice_ledger = load_invoice_ledger()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    from_farm_name = clean_name(request.form.get("from_farm_name"))
    to_farm_name = clean_name(request.form.get("to_farm_name"))
    if not customer or not from_farm_name or not to_farm_name:
        return redirect(url_for("admin_home", ok=0, msg="Customer and both farm names are required"))
    if normalized_name_key(from_farm_name) == normalized_name_key(to_farm_name):
        return redirect(url_for("admin_home", ok=0, msg="Choose two different farm names"))

    source_bucket = {}
    for farm_name, field_names in (field_map.get(customer, {}) or {}).items():
        if normalized_name_key(farm_name) == normalized_name_key(from_farm_name):
            source_bucket = field_names if isinstance(field_names, list) else []
            break
    target_bucket = {}
    customer_fields = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    for farm_name, field_names in customer_fields.items():
        if normalized_name_key(farm_name) == normalized_name_key(to_farm_name):
            target_bucket = field_names if isinstance(field_names, list) else []
            break
    merged_fields = list(target_bucket)
    for field_name in source_bucket:
        if normalized_name_key(field_name) not in [normalized_name_key(item) for item in merged_fields]:
            merged_fields.append(field_name)
    merged_fields.sort(key=lambda item: item.lower())
    customer_fields = {
        farm_name: fields
        for farm_name, fields in customer_fields.items()
        if normalized_name_key(farm_name) not in [normalized_name_key(from_farm_name), normalized_name_key(to_farm_name)]
    }
    customer_fields[to_farm_name] = merged_fields
    field_map[customer] = customer_fields
    save_field_map(field_map)

    updated_jobs = 0
    for row in jobs:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) != normalized_name_key(customer):
            continue
        if normalized_name_key(row.get("farm_name")) != normalized_name_key(from_farm_name):
            continue
        row["farm_name"] = to_farm_name
        apply_customer_master_snapshot(row, master_rows, customer, to_farm_name)
        updated_jobs += 1
    if updated_jobs:
        save_jobs(jobs)

    updated_invoices = 0
    for row in invoice_ledger:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(from_farm_name):
            row["farm_name"] = to_farm_name
            updated_invoices += 1
    if updated_invoices:
        save_invoice_ledger(invoice_ledger)

    updated_master = 0
    for record in master_rows:
        if normalized_name_key(record.get("customer_name")) != normalized_name_key(customer):
            continue
        if normalized_name_key(record.get("farm_name")) != normalized_name_key(from_farm_name):
            continue
        record["farm_name"] = to_farm_name
        ok, _ = save_customer_master_customer_details(
            customer,
            record.get("email", ""),
            record.get("rate_per_ton", ""),
            row_number=record.get("_row_number"),
            field_values=record,
        )
        if ok:
            updated_master += 1

    sync_farms_store(master_rows, jobs, field_map)
    return redirect(
        url_for(
            "admin_home",
            ok=1,
            msg="Merged %s into %s: %s jobs, %s invoice records, %s spreadsheet rows." % (
                from_farm_name,
                to_farm_name,
                updated_jobs,
                updated_invoices,
                updated_master,
            ),
        )
    )


@app.route("/admin/farms/consolidate-name", methods=["POST"])
def admin_consolidate_farm_name():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    invoice_ledger = load_invoice_ledger()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    farm_name = clean_name(request.form.get("farm_name"))
    canonical_farm_name = clean_name(request.form.get("canonical_farm_name"))
    if not customer or not farm_name or not canonical_farm_name:
        return redirect(url_for("admin_home", ok=0, msg="Customer and both farm names are required"))

    customer_fields = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    matching_field_names = [
        name for name in customer_fields
        if normalized_name_key(name) == normalized_name_key(farm_name)
    ]
    merged_fields = []
    for name in matching_field_names:
        for field_name in customer_fields.get(name, []) if isinstance(customer_fields.get(name), list) else []:
            if normalized_name_key(field_name) not in [normalized_name_key(item) for item in merged_fields]:
                merged_fields.append(field_name)
    for row in jobs:
        if isinstance(row, dict) and normalized_name_key(row.get("customer")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(farm_name):
            field_name = clean_name(row.get("field_name"))
            if field_name and normalized_name_key(field_name) not in [normalized_name_key(item) for item in merged_fields]:
                merged_fields.append(field_name)
    merged_fields.sort(key=lambda item: item.lower())
    for name in matching_field_names:
        customer_fields.pop(name, None)
    customer_fields[canonical_farm_name] = merged_fields
    field_map[customer] = customer_fields
    save_field_map(field_map)

    updated_jobs = 0
    for row in jobs:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(farm_name):
            row["farm_name"] = canonical_farm_name
            apply_customer_master_snapshot(row, master_rows, customer, canonical_farm_name)
            updated_jobs += 1
    if updated_jobs:
        save_jobs(jobs)

    updated_invoices = 0
    for row in invoice_ledger:
        if isinstance(row, dict) and normalized_name_key(row.get("customer")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(farm_name):
            row["farm_name"] = canonical_farm_name
            updated_invoices += 1
    if updated_invoices:
        save_invoice_ledger(invoice_ledger)

    updated_master = 0
    for record in master_rows:
        if normalized_name_key(record.get("customer_name")) != normalized_name_key(customer):
            continue
        if normalized_name_key(record.get("farm_name")) != normalized_name_key(farm_name):
            continue
        record["farm_name"] = canonical_farm_name
        ok, _ = save_customer_master_customer_details(
            customer,
            record.get("email", ""),
            record.get("rate_per_ton", ""),
            row_number=record.get("_row_number"),
            field_values=record,
        )
        if ok:
            updated_master += 1

    sync_farms_store(master_rows, jobs, field_map)
    return redirect(
        url_for(
            "admin_home",
            ok=1,
            msg="Consolidated farm name to %s: %s jobs, %s invoice records, %s spreadsheet rows." % (
                canonical_farm_name,
                updated_jobs,
                updated_invoices,
                updated_master,
            ),
        )
    )


@app.route("/admin/fields/add", methods=["POST"])
def admin_add_field():
    ensure_data_dir()
    customer = canonical_customer_name(
        clean_name(request.form.get("new_customer")) or clean_name(request.form.get("customer")),
        master_rows=load_customer_master_rows(),
        customers=load_customers(),
        jobs=load_jobs(),
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    farm_name = clean_name(request.form.get("farm_name"))
    field_name = clean_name(request.form.get("field_name"))
    if not customer or not field_name:
        return redirect(url_for("admin_home", ok=0, msg="Customer and field name are required"))

    customers = load_customers()
    if customer not in customers:
        customers.append(customer)
        save_customers(customers)

    farms = load_farms()
    if farm_name and farm_name not in farms:
        farms.append(farm_name)
        save_farms(farms)

    field_map = load_field_map()
    customer_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    farm_bucket = customer_bucket.get(farm_name, [])
    if field_name not in farm_bucket:
        farm_bucket.append(field_name)
        farm_bucket.sort(key=lambda item: item.lower())
        customer_bucket[farm_name] = farm_bucket
        field_map[customer] = customer_bucket
        save_field_map(field_map)
    return redirect(url_for("admin_home", ok=1, msg="Field link saved"))


@app.route("/admin/fields/delete", methods=["POST"])
def admin_delete_field():
    ensure_data_dir()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=load_customer_master_rows(),
        customers=load_customers(),
        jobs=load_jobs(),
        field_map=load_field_map(),
        invoice_ledger=load_invoice_ledger(),
    )
    farm_name = clean_name(request.form.get("farm_name"))
    field_name = clean_name(request.form.get("field_name"))
    field_map = load_field_map()
    customer_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    farm_bucket = customer_bucket.get(farm_name, [])
    updated_bucket = [name for name in farm_bucket if clean_name(name).lower() != field_name.lower()]
    if updated_bucket:
        customer_bucket[farm_name] = updated_bucket
    elif farm_name in customer_bucket:
        del customer_bucket[farm_name]
    if customer_bucket:
        field_map[customer] = customer_bucket
    elif customer in field_map:
        del field_map[customer]
    save_field_map(field_map)
    return redirect(url_for("admin_home", ok=1, msg="Field removed"))


@app.route("/admin/fields/move", methods=["POST"])
def admin_move_field():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    invoice_ledger = load_invoice_ledger()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    to_customer = canonical_customer_name(
        request.form.get("to_customer") or request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    from_farm_name = clean_name(request.form.get("farm_name"))
    to_farm_name = clean_name(request.form.get("to_farm_name"))
    field_name = clean_name(request.form.get("field_name"))

    if not customer or not field_name:
        return redirect(url_for("admin_home", ok=0, msg="Customer and field name are required"))
    if not to_customer:
        return redirect(url_for("admin_home", ok=0, msg="Destination customer is required"))
    if not to_farm_name:
        return redirect(url_for("admin_home", ok=0, msg="Move to farm name is required"))
    if normalized_name_key(customer) == normalized_name_key(to_customer) and normalized_name_key(from_farm_name) == normalized_name_key(to_farm_name):
        return redirect(url_for("admin_home", ok=0, msg="Field is already on that farm"))

    source_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    from_bucket = source_bucket.get(from_farm_name, [])
    updated_from_bucket = [name for name in from_bucket if normalized_name_key(name) != normalized_name_key(field_name)]
    if updated_from_bucket:
        source_bucket[from_farm_name] = updated_from_bucket
    elif from_farm_name in source_bucket:
        del source_bucket[from_farm_name]

    if customer in field_map and source_bucket:
        field_map[customer] = source_bucket
    elif customer in field_map:
        del field_map[customer]

    target_bucket = field_map.get(to_customer, {}) if isinstance(field_map.get(to_customer, {}), dict) else {}
    to_bucket = target_bucket.get(to_farm_name, [])
    if field_name not in to_bucket:
        to_bucket.append(field_name)
        to_bucket.sort(key=lambda item: item.lower())
    target_bucket[to_farm_name] = to_bucket
    field_map[to_customer] = target_bucket
    save_field_map(field_map)

    customers = load_customers()
    if to_customer not in customers:
        customers.append(to_customer)
        save_customers(customers)

    updated_jobs = 0
    for row in jobs:
        if not isinstance(row, dict):
            continue
        if normalized_name_key(row.get("customer")) != normalized_name_key(customer):
            continue
        if normalized_name_key(row.get("farm_name")) != normalized_name_key(from_farm_name):
            continue
        if normalized_name_key(row.get("field_name")) != normalized_name_key(field_name):
            continue
        row["customer"] = to_customer
        row["farm_name"] = to_farm_name
        apply_customer_master_snapshot(row, master_rows, to_customer, to_farm_name)
        updated_jobs += 1
    if updated_jobs:
        save_jobs(jobs)

    sync_farms_store(master_rows, jobs, field_map)
    return redirect(
        url_for(
            "admin_home",
            ok=1,
            msg="Moved %s to %s / %s. Updated %s saved jobs." % (field_name, to_customer, to_farm_name, updated_jobs),
        )
    )


@app.route("/admin/farms/remove", methods=["POST"])
def admin_remove_farm():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    invoice_ledger = load_invoice_ledger()
    customer = canonical_customer_name(
        request.form.get("customer"),
        master_rows=master_rows,
        customers=load_customers(),
        jobs=jobs,
        field_map=field_map,
        invoice_ledger=invoice_ledger,
    )
    farm_name = clean_name(request.form.get("farm_name"))
    if not customer or not farm_name:
        return redirect(url_for("admin_home", ok=0, msg="Customer and farm name are required"))

    customer_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    had_field_link = farm_name in customer_bucket
    if had_field_link:
        del customer_bucket[farm_name]
        if customer_bucket:
            field_map[customer] = customer_bucket
        elif customer in field_map:
            del field_map[customer]
        save_field_map(field_map)

    master_refs = 0
    for row in master_rows:
        if normalized_name_key(row.get("customer_name")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(farm_name):
            master_refs += 1

    job_refs = 0
    for row in jobs:
        if normalized_name_key(row.get("customer")) == normalized_name_key(customer) and normalized_name_key(row.get("farm_name")) == normalized_name_key(farm_name):
            job_refs += 1

    sync_farms_store(master_rows, jobs, field_map)

    if master_refs or job_refs:
        return redirect(
            url_for(
                "admin_home",
                ok=1,
                msg="Removed %s field link, but %s master rows and %s saved jobs still use that farm." % (farm_name, master_refs, job_refs),
            )
        )
    if had_field_link:
        return redirect(url_for("admin_home", ok=1, msg="Removed farm %s." % farm_name))
    return redirect(url_for("admin_home", ok=1, msg="Farm %s is already removed from field links." % farm_name))


@app.route("/admin/muck-types/add", methods=["POST"])
def admin_add_muck_type():
    ensure_data_dir()
    muck_type = clean_name(request.form.get("muck_type"))
    if not muck_type:
        return redirect(url_for("admin_home", ok=0, msg="Muck type is required"))

    muck_types = load_muck_types()
    if muck_type not in muck_types:
        muck_types.append(muck_type)
        save_muck_types(muck_types)
        return redirect(url_for("admin_home", ok=1, msg="Muck type saved"))
    return redirect(url_for("admin_home", ok=1, msg="Muck type already exists"))


@app.route("/admin/muck-types/delete", methods=["POST"])
def admin_delete_muck_type():
    ensure_data_dir()
    muck_type = clean_name(request.form.get("muck_type"))
    muck_types = [name for name in load_muck_types() if clean_name(name).lower() != muck_type.lower()]
    save_muck_types(muck_types)
    return redirect(url_for("admin_home", ok=1, msg="Muck type removed"))


@app.route("/api/jobs")
def jobs_api():
    return jsonify({"ok": True, "jobs": load_jobs()})


@app.route("/api/customers")
def customers_api():
    return jsonify({"ok": True, "customers": load_customers()})


@app.route("/api/field-map")
def field_map_api():
    return jsonify({"ok": True, "field_map": load_field_map()})


@app.route("/api/email/weekly/preview")
def weekly_email_preview_api():
    summary = weekly_jobs_summary()
    return jsonify({
        "ok": True,
        "config_ready": email_config_ready(load_email_config()),
        "summary": summary,
        "subject": weekly_email_subject(summary, load_email_config()),
        "body": weekly_email_body(summary),
    })


@app.route("/api/email/weekly/send-now", methods=["POST"])
def weekly_email_send_now_api():
    config = load_email_config()
    if not email_config_ready(config):
        return jsonify({"ok": False, "error": "Email configuration is incomplete"}), 400
    summary = weekly_jobs_summary()
    try:
        send_weekly_summary_email(summary, config)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500
    return jsonify({"ok": True, "summary": summary})


@app.route("/api/email/monthly/preview")
def monthly_email_preview_api():
    summary = monthly_jobs_summary()
    return jsonify({
        "ok": True,
        "config_ready": email_config_ready(load_email_config()),
        "summary": summary,
        "subject": monthly_email_subject(summary, load_email_config()),
        "body": monthly_email_body(summary),
    })


@app.route("/api/email/monthly/send-now", methods=["POST"])
def monthly_email_send_now_api():
    config = load_email_config()
    if not email_config_ready(config):
        return jsonify({"ok": False, "error": "Email configuration is incomplete"}), 400
    summary = monthly_jobs_summary()
    try:
        send_monthly_summary_email(summary, config)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500
    return jsonify({"ok": True, "summary": summary})


@app.route("/invoice/preview.pdf", methods=["POST"])
def invoice_preview_pdf():
    ensure_data_dir()
    invoice_form = invoice_form_from_request(request)
    invoice, error_message = build_invoice_from_form(invoice_form)
    if error_message:
        return Response(error_message, status=400, mimetype="text/plain")
    xlsx_bytes = build_invoice_xlsx_bytes(invoice)
    pdf_bytes = build_invoice_pdf_bytes(invoice, xlsx_bytes)
    response = Response(pdf_bytes, mimetype="application/pdf")
    response.headers["Content-Disposition"] = 'inline; filename="%s"' % invoice_pdf_filename(invoice)
    return response


@app.route("/invoice/preview.xlsx", methods=["POST"])
def invoice_preview_xlsx():
    ensure_data_dir()
    invoice_form = invoice_form_from_request(request)
    invoice, error_message = build_invoice_from_form(invoice_form)
    if error_message:
        return Response(error_message, status=400, mimetype="text/plain")
    xlsx_bytes = build_invoice_xlsx_bytes(invoice)
    response = Response(
        xlsx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = 'attachment; filename="%s"' % str(invoice.get("filename", "invoice.xlsx"))
    return response


@app.route("/invoice/send", methods=["POST"])
def invoice_create_and_send():
    ensure_data_dir()
    config = load_email_config()
    if not email_sender_ready(config):
        return redirect(url_for("invoice_home", ok=0, msg="Email sender settings are incomplete"))

    customer = clean_name(request.form.get("customer"))
    farm_name = clean_name(request.form.get("farm_name"))
    invoice_number = str(request.form.get("invoice_number", "") or "").strip()
    invoice_date = str(request.form.get("invoice_date", "") or "").strip()
    job_date_from = str(request.form.get("job_date_from", "") or "").strip()
    payment_terms_days = str(request.form.get("payment_terms_days", "") or "").strip()
    history_ledger_index = str(request.form.get("history_ledger_index", "") or "").strip()
    rate_override = str(request.form.get("rate_override", "") or "").strip()
    additional_fee_descriptions = list(request.form.getlist("additional_fee_description"))
    additional_fee_amounts = list(request.form.getlist("additional_fee_amount"))
    subject_text = str(request.form.get("subject", "") or "").strip()
    customer_message = str(request.form.get("customer_message", "") or "").strip()
    if not customer:
        return redirect(url_for("invoice_home", ok=0, msg="Invoice customer is required"))

    invoice, error_message = build_invoice_from_form({
        "customer": customer,
        "farm_name": farm_name,
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "job_date_from": job_date_from,
        "payment_terms_days": payment_terms_days,
        "rate_override": rate_override,
        "additional_fee_descriptions": additional_fee_descriptions,
        "additional_fee_amounts": additional_fee_amounts,
        "subject": subject_text,
        "customer_message": customer_message,
        "history_ledger_index": history_ledger_index,
    })
    if error_message:
        return redirect(url_for("invoice_home", ok=0, msg=error_message))

    accounts_emails = invoice_accounts_copy_emails(load_email_recipient_options())
    customer_emails = normalize_email_list(invoice.get("customer_emails", [invoice.get("customer_email", "")]))
    customer_email_keys = {email.lower() for email in customer_emails}
    accounts_emails = [email for email in accounts_emails if email.lower() not in customer_email_keys]

    try:
        send_invoice_email(invoice, config, accounts_emails, subject_text, customer_message, history_ledger_index)
    except Exception as exc:
        return redirect(url_for("invoice_home", ok=0, msg="Invoice email failed: %s" % clean_name(exc)))

    invoice_scope = invoice.get("customer", "")
    if invoice.get("farm_name"):
        invoice_scope = "%s / %s" % (invoice_scope, invoice.get("farm_name"))
    return redirect(
        url_for(
            "invoice_home",
            ok=1,
            msg="Invoice %s emailed for %s." % (invoice.get("invoice_number_label", ""), invoice_scope),
        )
    )


@app.route("/app/update", methods=["POST"])
def update_app():
    wants_json = (
        request.headers.get("X-Requested-With") == "fetch"
        or "application/json" in (request.headers.get("Accept", "") or "")
    )

    def respond(ok, msg, status_key, restart=False):
        redirect_url = url_for("home", ok=1 if ok else 0, msg=msg)
        if wants_json:
            return jsonify(
                {
                    "ok": ok,
                    "msg": msg,
                    "status": status_key,
                    "restart": restart,
                    "redirect_url": redirect_url,
                }
            )
        return redirect(redirect_url)

    git_dir = os.path.join(APP_ROOT, ".git")
    if not os.path.isdir(git_dir):
        return respond(False, "This install is not a git repo", "error")

    branch_result = run_git_command(["git", "rev-parse", "--abbrev-ref", "HEAD"], timeout_seconds=30)
    if branch_result.returncode != 0:
        return respond(False, git_update_message(branch_result, "Could not read git branch"), "error")

    branch_name = clean_name(branch_result.stdout) or "main"
    fetch_result = run_git_command(["git", "fetch", "origin", branch_name], timeout_seconds=180)
    if fetch_result.returncode != 0:
        return respond(False, git_update_message(fetch_result, "Git fetch failed"), "error")

    head_result = run_git_command(["git", "rev-parse", "HEAD"], timeout_seconds=30)
    remote_result = run_git_command(["git", "rev-parse", "FETCH_HEAD"], timeout_seconds=30)
    if head_result.returncode != 0 or remote_result.returncode != 0:
        return respond(False, "Could not compare app versions", "error")

    local_rev = clean_name(head_result.stdout)
    remote_rev = clean_name(remote_result.stdout)
    if local_rev and remote_rev and local_rev == remote_rev:
        return respond(True, "App is already up to date", "up_to_date")

    pull_result = run_git_command(["git", "pull", "--ff-only", "origin", branch_name], timeout_seconds=180)
    if pull_result.returncode != 0:
        return respond(False, git_update_message(pull_result, "Git pull failed"), "error")

    restart_app_process(delay_seconds=6.0)
    if wants_json:
        return respond(True, "App updated", "updated", restart=True)

    return render_template_string(
        """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="refresh" content="12;url={{ url_for('home', ok=1, msg='App updated') }}">
  <title>Updating App</title>
  <style>
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      background: linear-gradient(180deg, #efe7d8 0%, #e6dcc9 100%);
      color: #272d21;
      font-family: Georgia, "Times New Roman", serif;
    }
    .card {
      max-width: 520px;
      padding: 28px;
      border-radius: 24px;
      background: rgba(250, 247, 240, 0.96);
      border: 1px solid rgba(82, 69, 42, 0.12);
      box-shadow: 0 18px 44px rgba(60, 49, 25, 0.12);
      text-align: center;
    }
    h1 {
      margin: 0 0 12px;
      font-size: 32px;
    }
    p {
      margin: 0;
      line-height: 1.5;
      color: #666653;
    }
    a {
      display: inline-flex;
      margin-top: 18px;
      min-height: 48px;
      padding: 0 18px;
      border-radius: 999px;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      font-weight: bold;
      background: rgba(60,95,70,0.1);
      color: #3c5f46;
      border: 1px solid rgba(60,95,70,0.12);
    }
  </style>
</head>
<body>
  <div class="card">
    <h1>Updating App</h1>
    <p>The latest update was installed. The app is restarting now and should reload automatically in a few seconds.</p>
    <a href="{{ url_for('home', ok=1, msg='App updated') }}">Return To App</a>
  </div>
  <nav class="bottom-fixed-nav" aria-label="Bottom navigation">
    <button class="bottom-nav-btn" type="button" onclick="window.history.back()"><strong>←</strong>Back</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/'"><strong>⌂</strong>Home</button>
    <button class="bottom-nav-btn" type="button" onclick="window.location.href='/settings'"><strong>⚙</strong>Settings</button>
  </nav>
  <script>
    (function () {
      const targetUrl = {{ url_for('home', ok=1, msg='App updated')|tojson }};
      function tryReturn() {
        window.location.replace(targetUrl);
      }
      window.setTimeout(tryReturn, 6500);
      window.setTimeout(tryReturn, 9000);
      window.setTimeout(tryReturn, 12000);
    })();
  </script>
</body>
</html>"""
    )


@app.route("/backup/export.zip")
def backup_export_zip():
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        candidate_paths = [
            ("app.py", os.path.join(APP_ROOT, "app.py")),
            ("settings_workbook.py", os.path.join(APP_ROOT, "settings_workbook.py")),
            ("settings.xlsx", SETTINGS_WORKBOOK_PATH),
            ("README.md", os.path.join(APP_ROOT, "README.md")),
            ("email_config.example.json", os.path.join(APP_ROOT, "email_config.example.json")),
            ("muckspreading-app.service", os.path.join(APP_ROOT, "muckspreading-app.service")),
            ("weekly_summary_layout_template.xlsx", os.path.join(APP_ROOT, "weekly_summary_layout_template.xlsx")),
        ]
        for archive_name, file_path in candidate_paths:
            if os.path.exists(file_path):
                archive.write(file_path, archive_name)
        if os.path.isdir(DATA_DIR):
            for name in sorted(os.listdir(DATA_DIR)):
                file_path = os.path.join(DATA_DIR, name)
                if os.path.isfile(file_path):
                    archive.write(file_path, os.path.join("data", name))
        if os.path.isdir(ISSUE_PHOTOS_DIR):
            for name in sorted(os.listdir(ISSUE_PHOTOS_DIR)):
                file_path = os.path.join(ISSUE_PHOTOS_DIR, name)
                if os.path.isfile(file_path):
                    archive.write(file_path, os.path.join("data", "job_issue_photos", name))
        if os.path.isdir(INVOICE_ARCHIVE_DIR):
            for name in sorted(os.listdir(INVOICE_ARCHIVE_DIR)):
                file_path = os.path.join(INVOICE_ARCHIVE_DIR, name)
                if os.path.isfile(file_path):
                    archive.write(file_path, os.path.join("data", "invoices", name))
    output.seek(0)
    filename = "muckspreading_backup_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S")
    return Response(
        output.getvalue(),
        mimetype="application/zip",
        headers={"Content-Disposition": 'attachment; filename="%s"' % filename},
    )


@app.route("/summary/export/<period_key>.xlsx")
def export_period_summary_xlsx(period_key):
    config = summary_export_config(period_key)
    if not isinstance(config, dict):
        return redirect(url_for("home", ok=0, msg="Unknown summary export period"))

    start_date, end_date = config["range"]
    summary = jobs_summary_for_range(start_date, end_date, config["title"])
    filename = "%s_%s_to_%s.xlsx" % (
        config["filename_prefix"],
        start_date.strftime("%Y%m%d"),
        end_date.strftime("%Y%m%d"),
    )
    return Response(
        build_xlsx_attachment_bytes(summary),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="%s"' % filename},
    )


@app.route("/jobs/export.xlsx")
def export_jobs_xlsx():
    rows = load_jobs()
    sheet_rows = [[
        "job_date",
        "customer",
        "farm_name",
        "field_name",
        "muck_type",
        "job_notes",
        "issue_photos",
        "total_spreader_tons",
        "total_john_deere_tons",
        "customer_email",
        "customer_address_line_1",
        "customer_address_line_2",
        "customer_town",
        "customer_postcode",
        "rate_per_ton",
        "vat_rate",
        "customer_master_match",
        "created_ts",
        "updated_ts",
    ]]
    for row in reversed(rows):
        sheet_rows.append([
            row.get("job_date", ""),
            row.get("customer", ""),
            row.get("farm_name", ""),
            row.get("field_name", ""),
            row.get("muck_type", ""),
            row.get("job_notes", ""),
            ", ".join(normalize_issue_photo_names(row.get("issue_photos"))),
            row.get("total_spreader_tons", ""),
            row.get("total_john_deere_tons", ""),
            row.get("customer_email", ""),
            row.get("customer_address_line_1", ""),
            row.get("customer_address_line_2", ""),
            row.get("customer_town", ""),
            row.get("customer_postcode", ""),
            row.get("rate_per_ton", ""),
            row.get("vat_rate", ""),
            row.get("customer_master_match", ""),
            row.get("created_ts", ""),
            row.get("updated_ts", ""),
        ])

    sheet_xml_rows = []
    row_index = 1
    for row in sheet_rows:
        cell_xml = []
        col_index = 1
        for value in row:
            cell_xml.append(xlsx_cell_xml(row_index, col_index, value))
            col_index += 1
        sheet_xml_rows.append('<row r="%s">%s</row>' % (row_index, "".join(cell_xml)))
        row_index += 1

    worksheet_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="6" width="20" customWidth="1"/>
    <col min="7" max="8" width="16" customWidth="1"/>
    <col min="9" max="14" width="24" customWidth="1"/>
    <col min="15" max="18" width="18" customWidth="1"/>
  </cols>
  <sheetData>%s</sheetData>
</worksheet>
""" % "".join(sheet_xml_rows)

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Jobs Export" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

    workbook_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

    timestamp = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Jobs Export</dc:title>
  <dc:creator>A. Farrell Contracting</dc:creator>
  <cp:lastModifiedBy>A. Farrell Contracting</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>
</cp:coreProperties>
""" % (timestamp, timestamp)

    app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python</Application>
</Properties>
"""

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)

    filename = "muckspreading_jobs_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="%s"' % filename},
    )


@app.route("/health")
def health():
    field_map = load_field_map()
    field_count = 0
    for customer_fields in field_map.values():
        if not isinstance(customer_fields, dict):
            continue
        for fields in customer_fields.values():
            if isinstance(fields, list):
                field_count += len(fields)
    return jsonify({
        "ok": True,
        "jobs": len(load_jobs()),
        "customers": len(load_customers()),
        "fields": field_count,
        "generated_at": int(time.time()),
    })


if __name__ == "__main__":
    ensure_settings_workbook()
    ensure_data_dir()
    port = int(os.environ.get("MUCKSPREADING_APP_PORT", "8093"))
    debug_mode = str(os.environ.get("MUCKSPREADING_APP_DEBUG", "") or "").strip().lower() in ["1", "true", "yes", "on"]
    start_background_workers(debug_mode=debug_mode)
    app.run(host="0.0.0.0", port=port, debug=debug_mode, use_reloader=debug_mode)
