from flask import Flask, Response, jsonify, redirect, render_template_string, request, url_for
import csv
import io
import json
import os
import smtplib
import tempfile
import threading
import time
from datetime import datetime, timedelta
from email.message import EmailMessage


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
CUSTOMER_MASTER_CSV_PATH = os.path.join(APP_ROOT, "customer_master.csv")
EMAIL_CHECK_INTERVAL_SECONDS = 300
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
    "send_hour": 7,
    "send_minute": 0,
    "subject_prefix": "A. Farrell Contracting",
}

ICON = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
<rect width="64" height="64" rx="12" fill="#334d38"/>
<rect x="5" y="5" width="54" height="54" rx="10" fill="none" stroke="#d7bf7a" stroke-width="2.5"/>
<path d="M14 42c4-9 9-14 15-17 5-2 10-3 14-2 2 0 5 1 7 2-2 3-4 6-6 9-3 3-7 5-12 6-5 2-11 2-18 2z" fill="#d7bf7a"/>
</svg>"""

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>A. Farrell Contracting Muck Spreading Jobs</title>
  <link rel="icon" type="image/svg+xml" href="{{ url_for('favicon') }}">
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
    label {
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
      min-height: 56px;
      border-radius: 14px;
      border: 1px solid var(--line);
      background: #fffdfa;
      color: var(--ink);
      padding: 12px 14px;
      outline: none;
      font-size: 16px;
    }
    input:focus, select:focus {
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
    .button-full {
      width: 100%;
    }
    .bottom-export {
      margin-top: 18px;
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
    .job-card {
      padding: 16px;
      border-radius: 18px;
      border: 1px solid rgba(82, 69, 42, 0.1);
      background: rgba(255,255,255,0.68);
    }
    .job-card-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: flex-start;
      margin-bottom: 10px;
    }
    .job-card-title {
      font-size: 19px;
      font-weight: bold;
      line-height: 1.15;
    }
    .job-card-date {
      color: var(--muted);
      font-size: 14px;
      white-space: nowrap;
    }
    .job-card-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .job-card-item {
      padding: 11px 12px;
      border-radius: 14px;
      background: rgba(240,232,216,0.7);
    }
    .job-card-label {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #6a6657;
      margin-bottom: 4px;
    }
    .job-card-value {
      font-size: 16px;
      line-height: 1.25;
      word-break: break-word;
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
      .field-date {
        max-width: min(100%, 220px);
      }
      .actions {
        margin-top: 16px;
      }
      .button {
        width: 100%;
      }
      .table-wrap {
        display: none;
      }
      .mobile-jobs {
        display: grid;
      }
      .job-card-grid {
        grid-template-columns: 1fr;
      }
    }
    @media (max-width: 420px) {
      .job-card-head {
        display: block;
      }
      .job-card-date {
        display: block;
        margin-top: 4px;
        white-space: normal;
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
        <h2 class="panel-title">New Job</h2>

        {% if status_msg %}
        <div class="status {{ 'ok' if status_ok else 'error' }}">{{ status_msg }}</div>
        {% endif %}

        <form method="post" action="{{ url_for('save_job') }}">
          <div class="form-grid">
            <div class="field field-date">
              <label for="job_date">Date</label>
              <input id="job_date" name="job_date" type="date" value="{{ today_iso }}" required>
            </div>
            <div class="field">
              <label for="customer">Customer</label>
              <input id="customer" name="customer" type="text" placeholder="Start typing a customer name" autocomplete="off" required>
              <div id="customer_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="farm_name">Farm Name</label>
              <input id="farm_name" name="farm_name" type="text" placeholder="Start typing a farm name" autocomplete="off" required>
              <div id="farm_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="field_name">Field Name</label>
              <input id="field_name" name="field_name" type="text" placeholder="Start typing a field name" autocomplete="off" required>
              <div id="field_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="muck_type">Muck Type</label>
              <input id="muck_type" name="muck_type" type="text" placeholder="Start typing a muck type" autocomplete="off" required>
              <div id="muck_type_suggestions" class="suggestions"></div>
            </div>
            <div class="field">
              <label for="spreader_tons">Total Spreader Tons</label>
              <input id="spreader_tons" name="spreader_tons" type="number" inputmode="decimal" min="0" step="0.01" placeholder="0.00" required>
            </div>
            <div class="field">
              <label for="john_deere_tons">Total Ops Center Tons</label>
              <input id="john_deere_tons" name="john_deere_tons" type="number" inputmode="decimal" min="0" step="0.01" placeholder="0.00" required>
            </div>
          </div>
          <div class="actions">
            <button class="button button-primary" type="submit">Save Job</button>
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
                  <th>Spreader Tons</th>
                  <th>Ops Center Tons</th>
                  <th>Saved</th>
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
                  <td>{{ job.spreader_tons_label }}</td>
                  <td>{{ job.john_deere_tons_label }}</td>
                  <td>{{ job.saved_label }}</td>
                </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
        <div class="mobile-jobs">
          {% for job in recent_jobs %}
          <div class="job-card">
            <div class="job-card-head">
              <div class="job-card-title">{{ job.customer }}</div>
              <div class="job-card-date">{{ job.job_date_label }}</div>
            </div>
              <div class="job-card-grid">
                <div class="job-card-item">
                  <div class="job-card-label">Farm</div>
                  <div class="job-card-value">{{ job.farm_name }}</div>
                </div>
                <div class="job-card-item">
                  <div class="job-card-label">Field</div>
                  <div class="job-card-value">{{ job.field_name }}</div>
                </div>
                <div class="job-card-item">
                  <div class="job-card-label">Muck Type</div>
                  <div class="job-card-value">{{ job.muck_type }}</div>
                </div>
                <div class="job-card-item">
                  <div class="job-card-label">Spreader Tons</div>
                  <div class="job-card-value">{{ job.spreader_tons_label }}</div>
                </div>
              <div class="job-card-item">
                <div class="job-card-label">Ops Center Tons</div>
                <div class="job-card-value">{{ job.john_deere_tons_label }}</div>
              </div>
              <div class="job-card-item">
                <div class="job-card-label">Saved</div>
                <div class="job-card-value">{{ job.saved_label }}</div>
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

    <section class="bottom-export">
      <div class="card">
        <h2 class="panel-title">Export Jobs</h2>
        <p class="copy">Download the full saved job list as a CSV file.</p>
        <a class="button button-secondary button-full" href="{{ url_for('export_csv') }}">Download Full Job List CSV</a>
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
        button.addEventListener("pointerdown", function (event) {
          event.preventDefault();
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

    function currentFieldOptions() {
      const customerKey = findCustomerKey(customerInput.value);
      if (!customerKey) return allFields;

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

      return combined.length ? combined : allFields;
    }

    function showCustomerSuggestions() {
      openSuggestionBox(customerSuggestions, filterOptions(allCustomers, customerInput.value), function (value) {
        customerInput.value = value;
        if (!String(farmInput.value || "").trim()) {
          const farms = currentFarmOptions();
          if (farms.length === 1) {
            farmInput.value = farms[0];
          }
        }
      });
    }

    function showFarmSuggestions() {
      openSuggestionBox(farmSuggestions, filterOptions(currentFarmOptions(), farmInput.value), function (value) {
        farmInput.value = value;
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
    farmInput.addEventListener("input", showFarmSuggestions);
    farmInput.addEventListener("keyup", showFarmSuggestions);
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


def parse_csv_decimal(value):
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return ("%.2f" % float(text)).rstrip("0").rstrip(".")
    except Exception:
        return text


def load_customer_master_rows():
    if not os.path.exists(CUSTOMER_MASTER_CSV_PATH):
        return []

    rows = []
    try:
        with open(CUSTOMER_MASTER_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            raw_rows = list(csv.reader(handle))
    except Exception:
        return []

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


def customer_master_file_parts():
    prefix_rows = []
    header_row = list(CUSTOMER_MASTER_HEADERS)
    data_rows = []

    if not os.path.exists(CUSTOMER_MASTER_CSV_PATH):
        return prefix_rows, header_row, data_rows

    try:
        with open(CUSTOMER_MASTER_CSV_PATH, "r", newline="", encoding="utf-8-sig") as handle:
            raw_rows = list(csv.reader(handle))
    except Exception:
        return prefix_rows, header_row, data_rows

    header_index = None
    i = 0
    while i < len(raw_rows):
        candidate = [clean_name(cell).lower() for cell in raw_rows[i]]
        if "customer_name" in candidate:
            header_index = i
            break
        i += 1

    if header_index is None:
        return prefix_rows, header_row, data_rows

    prefix_rows = raw_rows[:header_index]
    header_row = raw_rows[header_index] if raw_rows[header_index] else list(CUSTOMER_MASTER_HEADERS)
    data_rows = raw_rows[header_index + 1:]
    return prefix_rows, header_row, data_rows


def customer_master_row_to_dict(header_row, raw_row):
    row_dict = {}
    i = 0
    while i < len(header_row):
        header = clean_name(header_row[i]).lower()
        if header:
            row_dict[header] = raw_row[i] if i < len(raw_row) else ""
        i += 1
    return row_dict


def customer_master_dict_to_row(header_row, row_dict):
    out = []
    i = 0
    while i < len(header_row):
        header = clean_name(header_row[i]).lower()
        if header:
            out.append(str(row_dict.get(header, "") or ""))
        else:
            out.append("")
        i += 1
    return out


def write_customer_master_file(prefix_rows, header_row, data_rows):
    parent = os.path.dirname(CUSTOMER_MASTER_CSV_PATH) or "."
    os.makedirs(parent, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(prefix="customer_master.", suffix=".tmp", dir=parent)
    with os.fdopen(fd, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        for row in prefix_rows:
            writer.writerow(row)
        writer.writerow(header_row if header_row else CUSTOMER_MASTER_HEADERS)
        for row in data_rows:
            writer.writerow(row)
    os.replace(temp_path, CUSTOMER_MASTER_CSV_PATH)


def sync_customer_master_from_job(job_record):
    prefix_rows, header_row, data_rows = customer_master_file_parts()
    if not header_row:
        header_row = list(CUSTOMER_MASTER_HEADERS)

    customer_name = clean_name(job_record.get("customer"))
    farm_name = clean_name(job_record.get("farm_name"))
    muck_type = clean_name(job_record.get("muck_type"))
    if not customer_name:
        return False

    exact_index = None
    empty_muck_index = None
    template_row = None
    i = 0
    while i < len(data_rows):
        row_dict = customer_master_row_to_dict(header_row, data_rows[i])
        row_customer = clean_name(row_dict.get("customer_name"))
        row_farm = clean_name(row_dict.get("farm_name"))
        row_muck = clean_name(row_dict.get("muck_type"))
        if row_customer.lower() == customer_name.lower():
            if template_row is None:
                template_row = row_dict
            if row_farm.lower() == farm_name.lower():
                template_row = row_dict
                if row_muck.lower() == muck_type.lower():
                    exact_index = i
                    break
                if not row_muck and muck_type:
                    empty_muck_index = i
        i += 1

    if exact_index is not None:
        return False

    if empty_muck_index is not None:
        row_dict = customer_master_row_to_dict(header_row, data_rows[empty_muck_index])
        row_dict["muck_type"] = muck_type
        if not clean_name(row_dict.get("active")):
            row_dict["active"] = "1"
        data_rows[empty_muck_index] = customer_master_dict_to_row(header_row, row_dict)
        write_customer_master_file(prefix_rows, header_row, data_rows)
        return True

    new_row = {
        "customer_name": customer_name,
        "farm_name": farm_name,
        "email": "",
        "address_line_1": "",
        "address_line_2": "",
        "town": "",
        "postcode": "",
        "rate_per_ton": "",
        "vat_rate": "",
        "active": "1",
        "muck_type": muck_type,
    }
    if isinstance(template_row, dict):
        for key in ["email", "address_line_1", "address_line_2", "town", "postcode", "rate_per_ton", "vat_rate", "active"]:
            value = str(template_row.get(key, "") or "")
            if value:
                new_row[key] = value
    data_rows.append(customer_master_dict_to_row(header_row, new_row))
    write_customer_master_file(prefix_rows, header_row, data_rows)
    return True


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


def load_email_config():
    data = read_json_file(EMAIL_CONFIG_PATH, {})
    merged = dict(DEFAULT_EMAIL_CONFIG)
    if isinstance(data, dict):
        merged.update(data)
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


def previous_full_week_range(now=None):
    now = now or datetime.now()
    today = now.date()
    current_week_start = today - timedelta(days=today.weekday())
    start_date = current_week_start - timedelta(days=7)
    end_date = current_week_start - timedelta(days=1)
    return start_date, end_date


def weekly_jobs_summary(start_date=None, end_date=None):
    if start_date is None or end_date is None:
        start_date, end_date = previous_full_week_range()

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
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "job_count": len(rows),
        "total_spreader_tons": round(total_spreader, 2),
        "total_john_deere_tons": round(total_john_deere, 2),
        "rows": rows,
    }


def weekly_email_subject(summary, config):
    prefix = str(config.get("subject_prefix", "A. Farrell Contracting") or "A. Farrell Contracting").strip()
    start_label = format_job_date(summary.get("start_date"))
    end_label = format_job_date(summary.get("end_date"))
    return "%s Weekly Jobs Summary: %s - %s" % (prefix, start_label, end_label)


def weekly_email_body(summary):
    lines = [
        "Weekly Jobs Summary",
        "%s to %s" % (format_job_date(summary.get("start_date")), format_job_date(summary.get("end_date"))),
        "",
        "Jobs: %s" % summary.get("job_count", 0),
        "Spreader Tons: %s" % format_tons(summary.get("total_spreader_tons", 0)),
        "Ops Center Tons: %s" % format_tons(summary.get("total_john_deere_tons", 0)),
        "",
        "Jobs Detail",
    ]

    rows = summary.get("rows", [])
    if not rows:
        lines.append("No jobs were recorded in this period.")
        return "\n".join(lines)

    for row in rows:
        lines.extend([
            "",
            "%s | %s" % (format_job_date(row.get("job_date")), row.get("customer", "")),
            "Farm: %s" % row.get("farm_name", ""),
            "Field: %s" % row.get("field_name", ""),
            "Muck Type: %s" % row.get("muck_type", ""),
            "Spreader Tons: %s" % format_tons(row.get("total_spreader_tons", 0)),
            "Ops Center Tons: %s" % format_tons(row.get("total_john_deere_tons", 0)),
        ])

    return "\n".join(lines)


def send_weekly_summary_email(summary, config):
    if not email_config_ready(config):
        raise RuntimeError("Email configuration is incomplete")

    msg = EmailMessage()
    msg["Subject"] = weekly_email_subject(summary, config)
    msg["From"] = config["from_email"]
    msg["To"] = ", ".join(config["to_emails"])
    msg.set_content(weekly_email_body(summary))

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
        "status_msg": str(request.args.get("msg", "") or "").strip(),
        "status_ok": str(request.args.get("ok", "1")) == "1",
        "data_dir": DATA_DIR,
        "total_jobs": len(jobs),
    }


@app.route("/favicon.ico")
@app.route("/favicon.svg")
def favicon():
    return Response(ICON, mimetype="image/svg+xml")


@app.route("/")
def home():
    ensure_data_dir()
    return render_template_string(HTML, **build_context())


@app.route("/jobs/save", methods=["POST"])
def save_job():
    ensure_data_dir()
    master_rows = load_customer_master_rows()
    job_date = str(request.form.get("job_date", "") or "").strip()
    customer = clean_name(request.form.get("customer"))
    farm_name = clean_name(request.form.get("farm_name"))
    field_name = clean_name(request.form.get("field_name"))
    muck_type = clean_name(request.form.get("muck_type"))

    if not job_date:
        return redirect(url_for("home", ok=0, msg="Date is required"))
    try:
        datetime.strptime(job_date, "%Y-%m-%d")
    except Exception:
        return redirect(url_for("home", ok=0, msg="Date must be in YYYY-MM-DD format"))
    if not customer:
        return redirect(url_for("home", ok=0, msg="Customer is required"))
    if not farm_name:
        return redirect(url_for("home", ok=0, msg="Farm name is required"))
    if not field_name:
        return redirect(url_for("home", ok=0, msg="Field name is required"))
    if not muck_type:
        return redirect(url_for("home", ok=0, msg="Muck type is required"))

    try:
        spreader_tons = parse_tons(request.form.get("spreader_tons"), "Total spreader tons")
        john_deere_tons = parse_tons(request.form.get("john_deere_tons"), "Total Ops Center tons")
    except ValueError as exc:
        return redirect(url_for("home", ok=0, msg=str(exc)))

    record = {
        "id": int(time.time() * 1000),
        "job_date": job_date,
        "customer": customer,
        "farm_name": farm_name,
        "field_name": field_name,
        "muck_type": muck_type,
        "total_spreader_tons": spreader_tons,
        "total_john_deere_tons": john_deere_tons,
        "created_ts": int(time.time()),
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

    append_json_line(JOBS_PATH, record)

    customers = load_customers()
    if customer not in customers:
        customers.append(customer)
        save_customers(customers)

    farms = load_farms()
    if farm_name not in farms:
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

    sync_error = ""
    try:
        sync_customer_master_from_job(record)
    except Exception:
        sync_error = " Customer master update failed."

    return redirect(url_for("home", ok=1, msg="Saved job for %s - %s.%s" % (customer, field_name, sync_error)))


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


@app.route("/jobs/export.csv")
def export_csv():
    rows = load_jobs()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "job_date",
        "customer",
        "farm_name",
        "field_name",
        "muck_type",
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
    ])
    for row in reversed(rows):
        writer.writerow([
            row.get("job_date", ""),
            row.get("customer", ""),
            row.get("farm_name", ""),
            row.get("field_name", ""),
            row.get("muck_type", ""),
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
        ])
    filename = "muckspreading_jobs_%s.csv" % datetime.now().strftime("%Y%m%d_%H%M%S")
    return Response(
        output.getvalue(),
        mimetype="text/csv",
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
