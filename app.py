from flask import Flask, Response, jsonify, redirect, render_template_string, request, url_for
import csv
import io
import json
import os
import smtplib
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


app = Flask(__name__)
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_ROOT, "data")
JOBS_PATH = os.path.join(DATA_DIR, "jobs.ndjson")
CUSTOMERS_PATH = os.path.join(DATA_DIR, "customers.json")
FARMS_PATH = os.path.join(DATA_DIR, "farms.json")
MUCK_TYPES_PATH = os.path.join(DATA_DIR, "muck_types.json")
FIELD_MAP_PATH = os.path.join(DATA_DIR, "customer_fields.json")
EMAIL_CONFIG_PATH = os.path.join(DATA_DIR, "email_config.json")
EMAIL_STATE_PATH = os.path.join(DATA_DIR, "weekly_email_state.json")
CUSTOMER_MASTER_XLSX_PATH = os.path.join(APP_ROOT, "customer_master.xlsx")
CUSTOMER_MASTER_CSV_PATH = os.path.join(APP_ROOT, "customer_master.csv")
EMAIL_SETTINGS_CSV_PATH = os.path.join(APP_ROOT, "email_settings.csv")
WEEKLY_SUMMARY_TEMPLATE_PATH = os.path.join(APP_ROOT, "weekly_summary_layout_template.xlsx")
EMAIL_CHECK_INTERVAL_SECONDS = 300
XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CUSTOMER_MASTER_HEADERS = [
    "customer_name",
    "farm_name",
    "email",
    "address_line_1",
    "address_line_2",
    "town",
    "postcode",
    "rate_per_ton",
    "vat_rate",
    "active",
    "muck_type",
]

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
    "subject_prefix": "A. Farrell Contracting",
}

APP_SHORT_NAME = "Muck Jobs"
APP_THEME_COLOR = "#334d38"


def discover_custom_app_icon_path():
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
    :root {
      --bg: #e7decd;
      --panel: rgba(250, 247, 240, 0.96);
      --ink: #272d21;
      --muted: #666653;
      --line: #cabd9f;
      --green: #3c5f46;
      --gold: #ba9450;
      --red: #8b4738;
      --shadow: 0 18px 44px rgba(60, 49, 25, 0.12);
      --font-main: Georgia, "Times New Roman", serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font-main);
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(186,148,80,0.18), transparent 24%),
        linear-gradient(180deg, #efe7d8 0%, #e6dcc9 100%);
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
      text-align: center;
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
      margin-top: 18px;
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
    .field.full {
      grid-column: 1 / -1;
    }
    .field.notes-field {
      grid-column: 1 / 2;
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
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 18px;
      align-items: start;
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
    .bottom-export {
      margin-top: 18px;
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
      .hero {
        gap: 14px;
        margin-bottom: 14px;
      }
      .layout {
        gap: 14px;
      }
      .bottom-export {
        margin-top: 14px;
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
      }
      .mini-grid {
        grid-template-columns: 1fr;
      }
      .field-date {
        max-width: min(100%, 220px);
      }
      .actions {
        margin-top: 16px;
      }
      .button {
        width: 100%;
      }
      .actions-inline .button {
        width: auto;
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
    <section class="hero">
      <div class="card hero-title-card">
        <h1>
          <span class="title-line title-line-primary">A. Farrell Contracting</span>
          <span class="title-line title-line-secondary">Muck Spreading Records</span>
        </h1>
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
      <div class="card">
        <h2 class="panel-title">{{ form_title }}</h2>

        {% if status_msg %}
        <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
        {% endif %}

        <form method="post" action="{{ url_for('save_job') }}">
          <input type="hidden" name="edit_job_id" value="{{ form_job.id }}">
          <input type="hidden" name="job_date" value="{{ form_job.job_date or today_iso }}">
          <div class="form-grid">
            <div class="field field-date">
              <label for="job_date">Date</label>
              <input id="job_date" type="text" value="{{ form_job.job_date_label }}" readonly>
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
          </div>
          <div class="actions">
            <button class="button button-primary" type="submit">{{ form_submit_label }}</button>
            {% if is_editing %}
            <a class="button button-secondary" href="{{ url_for('home') }}">Cancel Edit</a>
            {% endif %}
          </div>
        </form>
      </div>

      <div class="card">
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
                  <th>Spreader Tons</th>
                  <th>Ops Center Tons</th>
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
                  <td>{{ job.spreader_tons_label }}</td>
                  <td>{{ job.john_deere_tons_label }}</td>
                  <td>{{ job.saved_label }}</td>
                  <td>
                    <div class="actions-inline">
                      <a class="button button-secondary button-small" href="{{ url_for('home', edit_id=job.id) }}">Edit</a>
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
                </div>
              </div>
              <div class="job-row-meta">
                <div class="job-row-date">{{ job.job_date_label }}</div>
                <div class="actions-inline">
                  <a class="button button-secondary button-small" href="{{ url_for('home', edit_id=job.id) }}">Edit</a>
                  <form method="post" action="{{ url_for('delete_job', job_id=job.id) }}" class="delete-job-form">
                    <button class="button button-danger button-small" type="submit">Delete</button>
                  </form>
                </div>
              </div>
            </div>
          </div>
          {% endfor %}
        </div>
        {% else %}
        <div class="empty">No jobs saved yet.</div>
        {% endif %}
      </div>
    </section>

    <section class="section-grid">
      <div class="card">
        <h2 class="panel-title">Tools</h2>
        <p class="copy">Use these for admin and backups.</p>
        <div class="actions">
          <a class="button button-secondary button-full" href="{{ url_for('admin_home') }}">Open Data Admin</a>
          <a class="button button-secondary button-full" href="{{ url_for('backup_export_zip') }}">Download Backup ZIP</a>
          <form method="post" action="{{ url_for('update_app') }}" class="button-full">
            <button class="button button-secondary button-full" type="submit">Update App</button>
          </form>
        </div>
      </div>
    </section>

    <section class="bottom-export">
      <div class="section-grid">
        <div class="card">
          <h2 class="panel-title">Summary Exports</h2>
          <p class="copy">Download a period summary in the same Excel layout used for the weekly summary email.</p>
          <div class="actions">
            <a class="button button-secondary button-full" href="{{ url_for('export_period_summary_xlsx', period_key='current-week') }}">Download Current Week Summary .xlsx</a>
            <a class="button button-secondary button-full" href="{{ url_for('export_period_summary_xlsx', period_key='current-month') }}">Download Current Month Summary .xlsx</a>
            <a class="button button-secondary button-full" href="{{ url_for('export_period_summary_xlsx', period_key='last-month') }}">Download Last Month Summary .xlsx</a>
          </div>
        </div>
        <div class="card">
          <h2 class="panel-title">Export Jobs</h2>
          <p class="copy">Download the full saved job list as an Excel file.</p>
          <a class="button button-secondary button-full" href="{{ url_for('export_jobs_xlsx') }}">Download Full Job List .xlsx</a>
        </div>
      </div>
    </section>
  </div>

  <script>
    const fieldMap = {{ field_map_json|safe }};
    const customerFarmMap = {{ customer_farm_map_json|safe }};
    const allCustomers = {{ customers_json|safe }};
    const customerInput = document.getElementById("customer");
    const customerSuggestions = document.getElementById("customer_suggestions");
    const farmInput = document.getElementById("farm_name");
    const farmSuggestions = document.getElementById("farm_suggestions");
    const fieldInput = document.getElementById("field_name");
    const fieldSuggestions = document.getElementById("field_suggestions");
    const muckTypeInput = document.getElementById("muck_type");
    const muckTypeSuggestions = document.getElementById("muck_type_suggestions");
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
      const farmOptions = currentFarmOptions();
      if (String(farmInput.value || "").trim() && !valueInOptions(farmInput.value, farmOptions)) {
        farmInput.value = "";
      }

      const fieldOptions = currentFieldOptions();
      if (String(fieldInput.value || "").trim() && !valueInOptions(fieldInput.value, fieldOptions)) {
        fieldInput.value = "";
      }
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
    muckTypeInput.addEventListener("input", showMuckTypeSuggestions);
    muckTypeInput.addEventListener("keyup", showMuckTypeSuggestions);
    muckTypeInput.addEventListener("focus", function () {
      showMuckTypeSuggestions();
    });

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
    :root {
      --bg: #e7decd;
      --panel: rgba(250, 247, 240, 0.96);
      --ink: #272d21;
      --muted: #666653;
      --line: #cabd9f;
      --green: #3c5f46;
      --gold: #ba9450;
      --red: #8b4738;
      --shadow: 0 18px 44px rgba(60, 49, 25, 0.12);
      --font-main: Georgia, "Times New Roman", serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font-main);
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(186,148,80,0.18), transparent 24%),
        linear-gradient(180deg, #efe7d8 0%, #e6dcc9 100%);
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
      border-radius: 999px;
      border: none;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
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
      padding: 14px;
      background: rgba(255,255,255,0.55);
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
      gap: 6px;
      padding: 8px 10px;
      border-radius: 999px;
      background: rgba(60,95,70,0.08);
      border: 1px solid rgba(60,95,70,0.1);
    }
    .top-links {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
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
      </div>

      <div class="card">
        <h2>Customer / Farm Links</h2>
        <div class="map-list">
          {% for customer in admin_tree %}
          <div class="map-customer">
            <h3>{{ customer.customer_name }}</h3>
            <div class="farm-list">
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
                    {% if farm.fields %}
                    <form method="post" action="{{ url_for('admin_clear_farm_fields') }}">
                      <input type="hidden" name="customer" value="{{ customer.customer_name }}">
                      <input type="hidden" name="farm_name" value="{{ farm.farm_name }}">
                      <button class="button button-danger button-small" type="submit">Clear Farm Fields</button>
                    </form>
                    {% endif %}
                  </div>
                  {% if farm.fields %}
                  <div class="field-tags">
                    {% for field_name in farm.fields %}
                    <div class="field-tag">
                      <span>{{ field_name }}</span>
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
          {% endfor %}
        </div>
      </div>
    </div>
  </div>
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


def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


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


def xlsx_shared_strings(archive):
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except Exception:
        return []

    values = []
    for item in root.findall("{%s}si" % XLSX_NS):
        text_parts = []
        text_node = item.find("{%s}t" % XLSX_NS)
        if text_node is not None and text_node.text is not None:
            text_parts.append(text_node.text)
        for run in item.findall("{%s}r" % XLSX_NS):
            run_text = run.find("{%s}t" % XLSX_NS)
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

            first_sheet = workbook_root.find("{%s}sheets/{%s}sheet" % (XLSX_NS, XLSX_NS))
            if first_sheet is None:
                return []
            rel_id = first_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
            target = relationships.get(rel_id, "worksheets/sheet1.xml")
            if not target.startswith("xl/"):
                target = "xl/%s" % target.lstrip("/")

            sheet_root = ET.fromstring(archive.read(target))
            shared_strings = xlsx_shared_strings(archive)
    except Exception:
        return []

    rows = []
    sheet_data = sheet_root.find("{%s}sheetData" % XLSX_NS)
    if sheet_data is None:
        return rows

    for row_node in sheet_data.findall("{%s}row" % XLSX_NS):
        row_values = []
        for cell in row_node.findall("{%s}c" % XLSX_NS):
            ref = cell.attrib.get("r", "")
            col_index = worksheet_ref_col_index(ref)
            while len(row_values) < max(col_index - 1, 0):
                row_values.append("")

            value = ""
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                inline_node = cell.find("{%s}is/{%s}t" % (XLSX_NS, XLSX_NS))
                if inline_node is not None and inline_node.text is not None:
                    value = inline_node.text
            else:
                value_node = cell.find("{%s}v" % XLSX_NS)
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
    if os.path.exists(CUSTOMER_MASTER_XLSX_PATH):
        rows = xlsx_first_sheet_rows(CUSTOMER_MASTER_XLSX_PATH)
        if rows:
            return rows
    if not os.path.exists(CUSTOMER_MASTER_CSV_PATH):
        return []
    try:
        with open(CUSTOMER_MASTER_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            return list(csv.reader(handle))
    except Exception:
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
            "customer_name": customer_name,
            "farm_name": farm_name,
            "muck_type": clean_name(row_dict.get("muck_type")),
            "email": str(row_dict.get("email", "") or "").strip(),
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


def customer_master_row_to_dict(header_row, raw_row):
    row_dict = {}
    i = 0
    while i < len(header_row):
        header = clean_name(header_row[i]).lower()
        if header:
            row_dict[header] = raw_row[i] if i < len(raw_row) else ""
        i += 1
    return row_dict


def build_customer_farm_map(master_rows):
    out = {}
    for row in master_rows:
        customer_name = row.get("customer_name", "")
        farm_name = clean_name(row.get("farm_name"))
        if not customer_name:
            continue
        bucket = out.setdefault(customer_name, [])
        if farm_name and farm_name not in bucket:
            bucket.append(farm_name)
    for customer_name in out:
        out[customer_name].sort(key=lambda item: item.lower())
    return out


def build_customer_field_admin_map(master_rows, jobs, field_map):
    tree = {}

    def ensure_bucket(customer_name, farm_name):
        customer_name = clean_name(customer_name)
        farm_name = clean_name(farm_name)
        if not customer_name:
            return None
        customer_bucket = tree.setdefault(customer_name, {})
        farm_bucket = customer_bucket.setdefault(farm_name, [])
        return farm_bucket

    for row in master_rows:
        ensure_bucket(row.get("customer_name"), row.get("farm_name"))

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
    text = str(value or "").strip()
    if not text:
        return int(default)
    try:
        return int(text)
    except Exception:
        return int(default)


def load_email_settings_csv():
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
        "subject_prefix": str(selected_row.get("subject_prefix", "") or "").strip(),
    }

    raw_to_emails = str(selected_row.get("to_emails", "") or "").strip()
    if raw_to_emails:
        parsed["to_emails"] = [email.strip() for email in raw_to_emails.split(",") if email.strip()]
    for email in recipient_emails:
        if email not in parsed.get("to_emails", []):
            parsed.setdefault("to_emails", []).append(email)

    return parsed


def load_email_config():
    data = read_json_file(EMAIL_CONFIG_PATH, {})
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
    merged["send_hour"] = max(0, min(23, int(merged.get("send_hour", 7) or 7)))
    merged["send_minute"] = max(0, min(59, int(merged.get("send_minute", 0) or 0)))
    merged["subject_prefix"] = str(merged.get("subject_prefix", "A. Farrell Contracting") or "A. Farrell Contracting").strip()
    return merged


def load_email_state():
    data = read_json_file(EMAIL_STATE_PATH, {})
    return data if isinstance(data, dict) else {}


def save_email_state(state):
    write_json_atomic(EMAIL_STATE_PATH, state if isinstance(state, dict) else {})


def email_config_ready(config):
    if not isinstance(config, dict):
        return False
    required = [
        str(config.get("smtp_host", "") or "").strip(),
        str(config.get("from_email", "") or "").strip(),
    ]
    if not all(required):
        return False
    to_emails = config.get("to_emails", [])
    if not isinstance(to_emails, list) or not to_emails:
        return False
    return True


def clean_name(value):
    return " ".join(str(value or "").strip().split())


def load_customers():
    data = read_json_file(CUSTOMERS_PATH, [])
    if not isinstance(data, list):
        return []
    names = []
    for item in data:
        name = clean_name(item)
        if name and name not in names:
            names.append(name)
    names.sort(key=lambda item: item.lower())
    return names


def save_customers(customers):
    cleaned = []
    for item in customers:
        name = clean_name(item)
        if name and name not in cleaned:
            cleaned.append(name)
    cleaned.sort(key=lambda item: item.lower())
    write_json_atomic(CUSTOMERS_PATH, cleaned)


def load_farms():
    data = read_json_file(FARMS_PATH, [])
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
    write_json_atomic(FARMS_PATH, cleaned)


def load_muck_types(master_rows=None):
    names = []

    master_rows = master_rows if isinstance(master_rows, list) else load_customer_master_rows()
    for raw_row in master_rows:
        if not isinstance(raw_row, dict):
            continue
        name = clean_name(raw_row.get("muck_type"))
        if name and name not in names:
            names.append(name)

    data = read_json_file(MUCK_TYPES_PATH, [])
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
    write_json_atomic(MUCK_TYPES_PATH, cleaned)


def load_field_map():
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


def weekly_email_subject(summary, config):
    prefix = str(config.get("subject_prefix", "A. Farrell Contracting") or "A. Farrell Contracting").strip()
    start_label = format_job_date(summary.get("start_date"))
    end_label = format_job_date(summary.get("end_date"))
    return "%s Weekly Jobs Summary: %s - %s" % (prefix, start_label, end_label)


def weekly_email_body(summary):
    return "\n".join([
        "Weekly Jobs Summary",
        "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date"))),
        "",
        "Jobs: %s" % summary.get("job_count", 0),
        "Spreader Tons: %s" % format_tons(summary.get("total_spreader_tons", 0)),
        "Ops Center Tons: %s" % format_tons(summary.get("total_john_deere_tons", 0)),
        "",
        "The full weekly summary is attached as XLSX and PDF files.",
    ])


def weekly_summary_attachment_filename(summary):
    return "weekly_jobs_summary_%s_to_%s.xlsx" % (
        str(summary.get("start_date", "")).replace("-", ""),
        str(summary.get("end_date", "")).replace("-", ""),
    )


def weekly_summary_pdf_attachment_filename(summary):
    return "weekly_jobs_summary_%s_to_%s.pdf" % (
        str(summary.get("start_date", "")).replace("-", ""),
        str(summary.get("end_date", "")).replace("-", ""),
    )


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
    customer_row_number = 10
    farm_row_number = None
    detail_row_number = None
    farm_total_row_number = None

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

        if not a_text and any([b_text, c_text, d_text, e_text, f_text]) and not detail_style_map:
            detail_style_map = row_style_map(row_number)
            detail_row_number = row_number
            continue

        if a_text and not any([b_text, c_text, d_text, e_text, f_text]) and not farm_style_map and row_number > 10:
            farm_style_map = row_style_map(row_number)
            farm_row_number = row_number

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
            "farm": farm_style_map or customer_style_map,
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
            "farm": row_layout(farm_row_number or customer_row_number),
            "detail": row_layout(detail_row_number or 11),
            "farm_total": row_layout(farm_total_row_number or 12),
        },
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
            farm_name = clean_name(row.get("farm_name")) or "Unassigned Farm"
            farm_groups.setdefault(farm_name, []).append(row)

        show_all_farm_labels = len(farm_groups) > 1

        for farm_name in sorted(farm_groups.keys(), key=lambda item: item.lower()):
            if show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower():
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

    xml_rows = []
    row_number = 1
    for kind, values in row_specs:
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
            farm_name = clean_name(row.get("farm_name")) or "Unassigned Farm"
            farm_groups.setdefault(farm_name, []).append(row)

        show_all_farm_labels = len(farm_groups) > 1

        for farm_name in sorted(farm_groups.keys(), key=lambda item: item.lower()):
            if show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower():
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
            farm_name = clean_name(row.get("farm_name")) or "Unassigned Farm"
            farm_groups.setdefault(farm_name, []).append(row)

        show_all_farm_labels = len(farm_groups) > 1

        for farm_name in sorted(farm_groups.keys(), key=lambda item: item.lower()):
            if show_all_farm_labels or clean_name(farm_name).lower() != clean_name(customer_name).lower():
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
        if row_kind == "farm":
            return "0.97 0.97 0.97"
        if row_kind == "farm_total":
            return "0.95 0.95 0.95"
        return None

    def row_font(row_kind):
        if row_kind in ["header", "customer", "farm", "farm_total"]:
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


def send_weekly_summary_email(summary, config):
    if not email_config_ready(config):
        raise RuntimeError("Email configuration is incomplete")

    msg = EmailMessage()
    msg["Subject"] = weekly_email_subject(summary, config)
    msg["From"] = config["from_email"]
    msg["To"] = ", ".join(config["to_emails"])
    msg.set_content(weekly_email_body(summary))
    msg.add_attachment(
        build_xlsx_attachment_bytes(summary),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=weekly_summary_attachment_filename(summary),
    )
    msg.add_attachment(
        build_pdf_attachment_bytes(summary),
        maintype="application",
        subtype="pdf",
        filename=weekly_summary_pdf_attachment_filename(summary),
    )

    with smtplib.SMTP(config["smtp_host"], int(config["smtp_port"]), timeout=30) as server:
        server.ehlo()
        if config.get("use_tls", True):
            server.starttls()
            server.ehlo()
        if config.get("smtp_username"):
            server.login(config.get("smtp_username", ""), config.get("smtp_password", ""))
        server.send_message(msg)


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
    if str(state.get("last_sent_period_key", "")) == period_key:
        return {"ok": False, "reason": "already_sent"}

    send_weekly_summary_email(summary, config)
    save_email_state({
        "last_sent_period_key": period_key,
        "last_sent_at": int(time.time()),
        "last_summary_job_count": summary.get("job_count", 0),
    })
    return {"ok": True, "summary": summary}


def weekly_email_worker():
    while True:
        try:
            maybe_send_weekly_summary()
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


def restart_app_process(delay_seconds=1.5):
    def _restart():
        time.sleep(delay_seconds)
        os.chdir(APP_ROOT)
        os.execv(sys.executable, [sys.executable, os.path.join(APP_ROOT, "app.py")])

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


def format_job_date(value):
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").strftime("%d %b %Y")
    except Exception:
        return str(value or "--")


def format_saved_time(value):
    try:
        return datetime.fromtimestamp(int(value)).strftime("%d %b %Y %H:%M:%S")
    except Exception:
        return "--"


def build_context():
    jobs = load_jobs()
    master_rows = load_customer_master_rows()
    customers = load_customers()
    farms = load_farms()
    muck_types = load_muck_types(master_rows)
    field_map = load_field_map()
    customer_farm_map = build_customer_farm_map(master_rows)
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

    for job in jobs[:20]:
        row = dict(job)
        row["job_date_label"] = format_job_date(row.get("job_date"))
        row["spreader_tons_label"] = format_tons(row.get("total_spreader_tons"))
        row["john_deere_tons_label"] = format_tons(row.get("total_john_deere_tons"))
        row["saved_label"] = format_saved_time(row.get("created_ts"))
        row["job_notes"] = clean_name(row.get("job_notes"))
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
        }

    default_week_start, default_week_end = previous_full_week_range(now)

    return {
        "today_iso": today_iso,
        "today_human": today_human,
        "customers": customers,
        "customers_json": json.dumps(customers),
        "farms": farms,
        "muck_types": muck_types,
        "muck_types_json": json.dumps(muck_types),
        "all_fields": all_fields,
        "all_farms_json": json.dumps(farms),
        "all_fields_json": json.dumps(all_fields),
        "customer_farm_map_json": json.dumps(customer_farm_map),
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
        "status_msg": str(request.args.get("msg", "") or "").strip(),
        "status_ok": str(request.args.get("ok", "1")) == "1",
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
    return render_template_string(HTML, **build_context())


@app.route("/admin")
def admin_home():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    jobs = load_jobs()
    field_map = load_field_map()
    customers = load_customers()
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
        status_msg=str(request.args.get("msg", "") or "").strip(),
        status_ok=str(request.args.get("ok", "1")) == "1",
    )


@app.route("/jobs/save", methods=["POST"])
def save_job():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    edit_job_id = str(request.form.get("edit_job_id", "") or "").strip()
    customer = clean_name(request.form.get("customer"))
    farm_name = clean_name(request.form.get("farm_name"))
    field_name = clean_name(request.form.get("field_name"))
    muck_type = clean_name(request.form.get("muck_type"))
    job_notes = clean_name(request.form.get("job_notes"))
    today_iso = datetime.now().strftime("%Y-%m-%d")

    if not customer:
        return redirect(url_for("home", ok=0, msg="Customer is required"))
    if not field_name:
        return redirect(url_for("home", ok=0, msg="Field name is required"))
    if not muck_type:
        return redirect(url_for("home", ok=0, msg="Muck type is required"))

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

    if isinstance(existing_job, dict):
        job_date = str(existing_job.get("job_date", "") or "").strip() or today_iso
    else:
        job_date = today_iso

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
    }

    master_record = find_customer_master_record(master_rows, customer, farm_name)
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


@app.route("/admin/fields/add", methods=["POST"])
def admin_add_field():
    ensure_data_dir()
    customer = clean_name(request.form.get("new_customer")) or clean_name(request.form.get("customer"))
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
    customer = clean_name(request.form.get("customer"))
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


@app.route("/admin/fields/clear-farm", methods=["POST"])
def admin_clear_farm_fields():
    ensure_data_dir()
    customer = clean_name(request.form.get("customer"))
    farm_name = clean_name(request.form.get("farm_name"))
    field_map = load_field_map()
    customer_bucket = field_map.get(customer, {}) if isinstance(field_map.get(customer, {}), dict) else {}
    if farm_name in customer_bucket:
        del customer_bucket[farm_name]
    if customer_bucket:
        field_map[customer] = customer_bucket
    elif customer in field_map:
        del field_map[customer]
    save_field_map(field_map)
    return redirect(url_for("admin_home", ok=1, msg="Farm fields cleared"))


@app.route("/api/jobs")
def jobs_api():
    return jsonify({"ok": True, "jobs": load_jobs()})


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


@app.route("/app/update", methods=["POST"])
def update_app():
    git_dir = os.path.join(APP_ROOT, ".git")
    if not os.path.isdir(git_dir):
        return redirect(url_for("home", ok=0, msg="This install is not a git repo"))

    branch_result = run_git_command(["git", "rev-parse", "--abbrev-ref", "HEAD"], timeout_seconds=30)
    if branch_result.returncode != 0:
        return redirect(url_for("home", ok=0, msg=git_update_message(branch_result, "Could not read git branch")))

    branch_name = clean_name(branch_result.stdout) or "main"
    fetch_result = run_git_command(["git", "fetch", "origin", branch_name], timeout_seconds=180)
    if fetch_result.returncode != 0:
        return redirect(url_for("home", ok=0, msg=git_update_message(fetch_result, "Git fetch failed")))

    head_result = run_git_command(["git", "rev-parse", "HEAD"], timeout_seconds=30)
    remote_result = run_git_command(["git", "rev-parse", "FETCH_HEAD"], timeout_seconds=30)
    if head_result.returncode != 0 or remote_result.returncode != 0:
        return redirect(url_for("home", ok=0, msg="Could not compare app versions"))

    local_rev = clean_name(head_result.stdout)
    remote_rev = clean_name(remote_result.stdout)
    if local_rev and remote_rev and local_rev == remote_rev:
        return redirect(url_for("home", ok=1, msg="App is already up to date"))

    pull_result = run_git_command(["git", "pull", "--ff-only", "origin", branch_name], timeout_seconds=180)
    if pull_result.returncode != 0:
        return redirect(url_for("home", ok=0, msg=git_update_message(pull_result, "Git pull failed")))

    restart_app_process()
    return render_template_string(
        """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="refresh" content="8;url={{ url_for('home') }}">
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
  </style>
</head>
<body>
  <div class="card">
    <h1>Updating App</h1>
    <p>The latest update was installed. The app is restarting now and should reload automatically in a few seconds.</p>
  </div>
</body>
</html>"""
    )


@app.route("/backup/export.zip")
def backup_export_zip():
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        candidate_paths = [
            ("app.py", os.path.join(APP_ROOT, "app.py")),
            ("README.md", os.path.join(APP_ROOT, "README.md")),
            ("customer_master.xlsx", os.path.join(APP_ROOT, "customer_master.xlsx")),
            ("customer_master.csv", os.path.join(APP_ROOT, "customer_master.csv")),
            ("customer_master.template.csv", os.path.join(APP_ROOT, "customer_master.template.csv")),
            ("email_settings.csv", os.path.join(APP_ROOT, "email_settings.csv")),
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
    ensure_data_dir()
    port = int(os.environ.get("MUCKSPREADING_APP_PORT", "8093"))
    debug_mode = str(os.environ.get("MUCKSPREADING_APP_DEBUG", "") or "").strip().lower() in ["1", "true", "yes", "on"]
    start_background_workers(debug_mode=debug_mode)
    app.run(host="0.0.0.0", port=port, debug=debug_mode, use_reloader=debug_mode)
