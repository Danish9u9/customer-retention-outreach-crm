"""
Customer outreach dashboard.

Run customer_scheduler.py first (it builds Customer_Outreach_Schedule.xlsx), then:
    python app.py
and open http://127.0.0.1:5000

Environment options:
    HOST=0.0.0.0     let other devices on your Wi-Fi open it (e.g. your phone)
    PORT=5000        change the port
    FLASK_DEBUG=1    auto-reload while editing

Contacted customers are saved in contacted.json next to this file.
"""
import json
import os
import re
import threading
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from flask import Flask, Response, jsonify, request

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Customer_Outreach_Schedule.xlsx"
CONTACTED_PATH = BASE_DIR / "contacted.json"  # who you already contacted (survives restarts)
TIMEZONE = ZoneInfo("Asia/Karachi")
COUNTRY_CODE = "92"

app = Flask(__name__)

_cache = {"mtime": None, "customers": []}
_lock = threading.Lock()
_contacted_lock = threading.Lock()


def read_contacted():
    """{customer key: {"date": "YYYY-MM-DD", "orderId": "#1234"}}"""
    try:
        data = json.loads(CONTACTED_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def write_contacted(data):
    tmp = CONTACTED_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(CONTACTED_PATH)


def normalize_phone(raw):
    digits = re.sub(r"\D", "", str(raw))
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith(COUNTRY_CODE) and len(digits) >= 12:
        return digits
    if digits.startswith("0") and len(digits) == 11:
        return COUNTRY_CODE + digits[1:]
    if len(digits) == 10 and digits.startswith("3"):
        return COUNTRY_CODE + digits
    return digits


def to_int(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def segment_from(row):
    seg = row.get("Segment", "")
    if seg in ("Priority", "Re-order", "Regular"):
        return seg
    label = row.get("Priority", "")  # files made by the older script
    if "Priority" in label:
        return "Priority"
    if "Re-order" in label:
        return "Re-order"
    return "Regular"


def load_customers():
    """Read the Excel file, re-reading only when it has changed on disk."""
    mtime = EXCEL_PATH.stat().st_mtime
    with _lock:
        if _cache["mtime"] == mtime:
            return _cache["customers"]

        df = pd.read_excel(EXCEL_PATH, dtype=str, keep_default_na=False)
        customers = []
        for row in df.to_dict(orient="records"):
            phone = row.get("Phone", "").strip()
            customers.append(
                {
                    "name": row.get("Customer Name", "").strip(),
                    "phone": phone,
                    "wa": row.get("WhatsApp Number", "").strip() or normalize_phone(phone),
                    "email": row.get("Email", "").strip(),
                    "lastOrderId": row.get("Last Order ID", "").strip(),
                    "lastOrderDate": row.get("Last Order Date", "").strip(),
                    "segment": segment_from(row),
                    "orders": to_int(row.get("Orders", row.get("Delivered Orders")), 1),
                    "nextDate": row.get("Next Contact Date", "").strip(),
                    "nextTime": row.get("Next Contact Time", "").strip(),
                    "nextText": row.get("Next Contact Schedule", "").strip(),
                    "projected": row.get("Projected Reorder Date", "").strip(),
                    "avgDays": to_float(row.get("Avg Days Between")),
                    "product": row.get("Most Favorable Product", "").strip(),
                    "ltv": to_float(row.get("Lifetime Value")),
                }
            )
        _cache["mtime"] = mtime
        _cache["customers"] = customers
        return customers


@app.get("/api/customers")
def api_customers():
    if not EXCEL_PATH.exists():
        return jsonify(error="missing"), 404
    try:
        customers = load_customers()
    except Exception as exc:  # noqa: BLE001 - show the problem in the UI
        return jsonify(error="unreadable", detail=str(exc)), 500

    built = datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime, TIMEZONE)
    return jsonify(
        customers=customers,
        today=datetime.now(TIMEZONE).date().isoformat(),
        builtAt=built.isoformat(),
        contacted=read_contacted(),
    )


@app.post("/api/contacted")
def api_contacted():
    """Body: {"set": {key: {"date": ..., "orderId": ...}}, "remove": [key, ...]}"""
    body = request.get_json(silent=True) or {}
    today = datetime.now(TIMEZONE).date().isoformat()
    with _contacted_lock:
        data = read_contacted()
        for key, rec in (body.get("set") or {}).items():
            if not isinstance(key, str) or not key or len(key) > 40 or not isinstance(rec, dict):
                continue
            data[key] = {
                "date": str(rec.get("date") or today)[:10],
                "orderId": str(rec.get("orderId") or "")[:60],
            }
        for key in body.get("remove") or []:
            data.pop(str(key), None)
        try:
            write_contacted(data)
        except OSError as exc:
            return jsonify(error="unwritable", detail=str(exc)), 500
    return jsonify(ok=True, count=len(data))


@app.get("/")
def home():
    return Response(PAGE, mimetype="text/html")


PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Customer outreach</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700&family=Instrument+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<script>
  (function () {
    var saved = null;
    try { saved = localStorage.getItem('outreach.theme'); } catch (e) {}
    var dark = saved ? saved === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  })();
</script>
<style>
  :root {
    --bg: #eef2f0;
    --panel: #ffffff;
    --soft: #f4f7f5;
    --ink: #14211d;
    --muted: #5a6c66;
    --line: #dce5e1;
    --brand: #0b6b5b;
    --brand-hover: #095a4d;
    --brand-ink: #ffffff;
    --brand-soft: #dcf0ea;
    --warn: #8f5509;
    --warn-soft: #fbebcf;
    --bad: #a63a28;
    --bad-soft: #fae3de;
    --shadow: 0 1px 2px rgba(20, 33, 29, .05), 0 12px 32px -16px rgba(20, 33, 29, .18);
    --ease: cubic-bezier(.22, 1, .36, 1);
    color-scheme: light;
  }
  [data-theme="dark"] {
    --bg: #0c1311;
    --panel: #131c19;
    --soft: #192521;
    --ink: #e4ede9;
    --muted: #8da29b;
    --line: #243430;
    --brand: #3fbfa0;
    --brand-hover: #5ccdb1;
    --brand-ink: #06211a;
    --brand-soft: #173229;
    --warn: #e8b35c;
    --warn-soft: #33270f;
    --bad: #f08c78;
    --bad-soft: #3a1e18;
    --shadow: 0 1px 2px rgba(0, 0, 0, .3), 0 12px 32px -16px rgba(0, 0, 0, .6);
    color-scheme: dark;
  }

  * { box-sizing: border-box; }
  html { -webkit-text-size-adjust: 100%; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--ink);
    font: 400 15px/1.5 'Instrument Sans', system-ui, sans-serif;
    transition: background-color .3s, color .3s;
  }
  h1, h2, .display { font-family: 'Bricolage Grotesque', 'Instrument Sans', sans-serif; }
  button, input, select, textarea { font: inherit; color: inherit; }
  :focus-visible { outline: 2px solid var(--brand); outline-offset: 2px; }

  .shell { max-width: 1240px; margin: 0 auto; padding: 28px 20px 80px; }

  /* Header */
  .top { display: flex; flex-wrap: wrap; gap: 16px; align-items: flex-end; justify-content: space-between; margin-bottom: 22px; }
  .top h1 { margin: 0; font-size: 30px; font-weight: 700; letter-spacing: -.02em; line-height: 1.1; }
  .top p { margin: 6px 0 0; color: var(--muted); font-size: 14px; }
  .top-actions { display: flex; gap: 8px; flex-wrap: wrap; }

  .btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 8px;
    height: 40px; padding: 0 14px; border-radius: 10px; border: 1px solid var(--line);
    background: var(--panel); cursor: pointer; font-weight: 500; font-size: 14px; text-decoration: none;
    transition: transform .2s var(--ease), border-color .2s, background-color .2s, box-shadow .2s;
  }
  .btn:hover { border-color: var(--muted); transform: translateY(-1px); }
  .btn:active { transform: translateY(0) scale(.98); }
  .btn svg, .icon-btn svg { width: 18px; height: 18px; flex: none; }
  .btn.primary { background: var(--brand); border-color: var(--brand); color: var(--brand-ink); }
  .btn.primary:hover { background: var(--brand-hover); border-color: var(--brand-hover); }
  .icon-btn {
    width: 40px; height: 40px; display: inline-grid; place-items: center; border-radius: 10px;
    border: 1px solid var(--line); background: var(--panel); cursor: pointer; color: var(--muted);
    text-decoration: none;
    transition: transform .2s var(--ease), color .2s, border-color .2s, background-color .2s;
  }
  .icon-btn:hover { color: var(--ink); border-color: var(--muted); transform: translateY(-1px); }
  .icon-btn:active { transform: scale(.94); }
  .icon-btn[aria-disabled="true"] { opacity: .35; pointer-events: none; }
  .spin svg { animation: spin .7s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* Hero */
  .panel { background: var(--panel); border: 1px solid var(--line); border-radius: 18px; box-shadow: var(--shadow); }
  .hero { display: grid; grid-template-columns: minmax(0, 1.4fr) minmax(0, 1fr); gap: 32px; padding: 30px 32px; margin-bottom: 18px; }
  .hero-main { min-width: 0; }
  .hero-num { display: flex; align-items: baseline; gap: 14px; flex-wrap: wrap; }
  .hero-num strong { font-family: 'Bricolage Grotesque', sans-serif; font-weight: 700; font-size: clamp(64px, 11vw, 104px); line-height: .9; letter-spacing: -.04em; font-variant-numeric: tabular-nums; }
  .hero-num span { font-size: 18px; font-weight: 500; max-width: 16ch; line-height: 1.25; }
  .progress { margin-top: 24px; }
  .bar { height: 10px; border-radius: 99px; background: var(--soft); border: 1px solid var(--line); overflow: hidden; }
  .bar i { display: block; height: 100%; width: 0; background: var(--brand); border-radius: 99px; transition: width .9s var(--ease); }
  .progress p { margin: 10px 0 0; font-size: 14px; color: var(--muted); }
  .stats { display: grid; grid-template-columns: 1fr 1fr; gap: 0; align-self: center; }
  .stat { padding: 14px 0 14px 20px; border-left: 1px solid var(--line); }
  .stat:nth-child(n+3) { border-top: 1px solid var(--line); }
  .stat b { display: block; font-family: 'Bricolage Grotesque', sans-serif; font-size: 30px; font-weight: 700; letter-spacing: -.02em; line-height: 1.1; font-variant-numeric: tabular-nums; }
  .stat small { color: var(--muted); font-size: 13px; }

  /* Toolbar */
  .toolbar {
    position: sticky; top: 0; z-index: 5; padding: 12px 0; margin-bottom: 6px;
    background: color-mix(in srgb, var(--bg) 88%, transparent); backdrop-filter: blur(10px);
    display: grid; grid-template-columns: minmax(220px, 1fr) auto auto; gap: 10px; align-items: center;
  }
  .search { position: relative; }
  .search svg { position: absolute; left: 13px; top: 50%; transform: translateY(-50%); width: 18px; height: 18px; color: var(--muted); pointer-events: none; }
  .search input {
    width: 100%; height: 42px; padding: 0 44px 0 40px; border-radius: 10px; border: 1px solid var(--line);
    background: var(--panel); outline: none; transition: border-color .2s, box-shadow .2s;
  }
  .search input:focus { border-color: var(--brand); box-shadow: 0 0 0 4px color-mix(in srgb, var(--brand) 18%, transparent); }
  .search kbd { position: absolute; right: 10px; top: 50%; transform: translateY(-50%); font: 500 12px 'Instrument Sans'; color: var(--muted); border: 1px solid var(--line); border-radius: 6px; padding: 1px 7px; background: var(--soft); }
  select {
    height: 42px; padding: 0 34px 0 12px; border-radius: 10px; border: 1px solid var(--line); background: var(--panel); cursor: pointer;
    appearance: none; -webkit-appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' fill='none' stroke='%238da29b' stroke-width='2'%3E%3Cpath d='m1 1.5 5 5 5-5'/%3E%3C/svg%3E");
    background-repeat: no-repeat; background-position: right 12px center;
  }
  .chips { grid-column: 1 / -1; display: flex; gap: 8px; flex-wrap: wrap; }
  .chip {
    display: inline-flex; align-items: center; gap: 8px; height: 36px; padding: 0 14px; border-radius: 99px;
    border: 1px solid var(--line); background: var(--panel); cursor: pointer; font-size: 14px; font-weight: 500;
    transition: background-color .2s, color .2s, border-color .2s, transform .2s var(--ease);
  }
  .chip:hover { border-color: var(--muted); }
  .chip:active { transform: scale(.96); }
  .chip em { font-style: normal; font-size: 12px; color: var(--muted); font-variant-numeric: tabular-nums; transition: color .2s; }
  .chip[aria-pressed="true"] { background: var(--ink); border-color: var(--ink); color: var(--bg); }
  .chip[aria-pressed="true"] em { color: color-mix(in srgb, var(--bg) 70%, transparent); }

  /* List */
  .cols { display: grid; grid-template-columns: minmax(200px, 1.5fr) minmax(150px, 1fr) minmax(120px, .8fr) minmax(120px, .8fr) minmax(150px, 1.2fr) 190px; gap: 18px; align-items: center; }
  .thead { padding: 14px 24px; border-bottom: 1px solid var(--line); color: var(--muted); font-size: 13px; font-weight: 500; }
  .list { list-style: none; margin: 0; padding: 0; }
  .list.swap { animation: swap .28s var(--ease); }
  @keyframes swap { from { opacity: .35; transform: translateY(4px); } to { opacity: 1; transform: none; } }
  .row { padding: 14px 24px; border-bottom: 1px solid var(--line); transition: background-color .2s, opacity .3s; position: relative; }
  .row:last-child { border-bottom: 0; }
  .row:hover { background: var(--soft); }
  .row::before { content: ''; position: absolute; left: 0; top: 10px; bottom: 10px; width: 3px; border-radius: 0 3px 3px 0; background: var(--brand); transform: scaleY(0); transition: transform .25s var(--ease); }
  .row:hover::before { transform: scaleY(1); }
  .row.enter { animation: rise .5s var(--ease) both; animation-delay: calc(var(--i) * 28ms + 250ms); }
  @keyframes rise { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
  .cell { min-width: 0; }
  .cell strong { display: block; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .cell small { display: block; color: var(--muted); font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .who { display: flex; align-items: center; gap: 12px; }
  .who > div { min-width: 0; }
  .avatar { flex: none; width: 40px; height: 40px; border-radius: 50%; display: grid; place-items: center; background: var(--brand-soft); color: var(--brand); font-weight: 600; font-size: 15px; transition: background-color .25s, color .25s; }
  .row:hover .avatar { background: var(--brand); color: var(--brand-ink); }

  .when strong { display: flex; align-items: center; gap: 8px; }
  .when strong::before { content: ''; width: 8px; height: 8px; border-radius: 50%; background: var(--muted); flex: none; }
  .when.now strong { color: var(--brand); }
  .when.now strong::before { background: var(--brand); animation: pulse 2s ease-out infinite; }
  .when.late strong { color: var(--bad); }
  .when.late strong::before { background: var(--bad); }
  .when.lapsed strong { color: var(--muted); font-weight: 500; }
  @keyframes pulse { 0% { box-shadow: 0 0 0 0 color-mix(in srgb, var(--brand) 45%, transparent); } 70%, 100% { box-shadow: 0 0 0 8px transparent; } }

  .pill { display: inline-block; padding: 3px 10px; border-radius: 99px; font-size: 13px; font-weight: 600; }
  .pill.Priority { background: var(--warn-soft); color: var(--warn); }
  .pill.Re-order { background: var(--brand-soft); color: var(--brand); }
  .pill.Regular { background: var(--soft); color: var(--muted); border: 1px solid var(--line); }

  .actions { display: flex; gap: 6px; justify-content: flex-end; }
  .actions .btn { height: 36px; padding: 0 12px; }
  .actions .icon-btn { width: 36px; height: 36px; }
  .btn.wa { background: var(--brand); border-color: var(--brand); color: var(--brand-ink); }
  .btn.wa:hover { background: var(--brand-hover); border-color: var(--brand-hover); }
  .btn.wa[aria-disabled="true"] { opacity: .4; pointer-events: none; }
  .icon-btn.done[aria-pressed="true"] { background: var(--brand-soft); border-color: var(--brand); color: var(--brand); }
  .row.is-done { opacity: .55; }
  .row.is-done:hover { opacity: 1; }

  /* Empty / error / loading */
  .state { padding: 70px 24px; text-align: center; }
  .state h2 { margin: 0 0 8px; font-size: 22px; letter-spacing: -.01em; }
  .state p { margin: 0 auto 18px; max-width: 46ch; color: var(--muted); }
  .state code { background: var(--soft); border: 1px solid var(--line); border-radius: 6px; padding: 1px 7px; font-size: 13px; }
  .skeleton { height: 70px; margin: 0 24px; border-bottom: 1px solid var(--line); display: flex; align-items: center; }
  .skeleton span { display: block; height: 14px; width: 60%; border-radius: 6px; background: linear-gradient(90deg, var(--soft) 25%, var(--line) 50%, var(--soft) 75%); background-size: 200% 100%; animation: shimmer 1.3s linear infinite; }
  @keyframes shimmer { to { background-position: -200% 0; } }
  .more-wrap { padding: 18px; text-align: center; border-top: 1px solid var(--line); }

  /* Dialog + toast */
  dialog { width: min(560px, calc(100vw - 32px)); max-height: calc(100vh - 32px); overflow: auto; padding: 0; border: 1px solid var(--line); border-radius: 18px; background: var(--panel); color: var(--ink); box-shadow: 0 30px 80px -20px rgba(0, 0, 0, .45); }
  dialog[open] { animation: pop .3s var(--ease); }
  dialog::backdrop { background: rgba(8, 14, 12, .5); backdrop-filter: blur(3px); animation: fade .3s; }
  @keyframes pop { from { opacity: 0; transform: translateY(14px) scale(.97); } to { opacity: 1; transform: none; } }
  @keyframes fade { from { opacity: 0; } }
  .dlg { padding: 24px; }
  .dlg h2 { margin: 0 0 6px; font-size: 22px; }
  .dlg h3 { margin: 18px 0 8px; font-size: 15px; font-weight: 600; }
  .dlg p { margin: 0 0 14px; color: var(--muted); font-size: 14px; }
  .dlg code { background: var(--soft); border: 1px solid var(--line); border-radius: 6px; padding: 1px 6px; font-size: 13px; }
  .dlg input[type="text"] { width: 100%; height: 42px; padding: 0 14px; margin-bottom: 10px; border-radius: 10px; border: 1px solid var(--line); background: var(--soft); outline: none; }
  .dlg input[type="text"]:focus { border-color: var(--brand); }
  textarea { width: 100%; min-height: 150px; resize: vertical; padding: 12px 14px; border-radius: 10px; border: 1px solid var(--line); background: var(--soft); outline: none; line-height: 1.5; }
  textarea:focus { border-color: var(--brand); }
  .preview { margin: 14px 0 0; padding: 12px 14px; border-radius: 10px; background: var(--brand-soft); font-size: 14px; white-space: pre-wrap; }
  .dlg-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 18px; }

  .toast { position: fixed; left: 50%; bottom: 24px; transform: translate(-50%, 24px); opacity: 0; pointer-events: none; z-index: 20; display: flex; align-items: center; gap: 14px; padding: 12px 14px 12px 18px; border-radius: 12px; background: var(--ink); color: var(--bg); font-size: 14px; box-shadow: 0 16px 40px -12px rgba(0, 0, 0, .5); transition: transform .35s var(--ease), opacity .25s; }
  .toast.show { transform: translate(-50%, 0); opacity: 1; pointer-events: auto; }
  .toast button { border: 0; background: transparent; color: inherit; font-weight: 600; cursor: pointer; text-decoration: underline; padding: 2px 4px; }

  .sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); }

  @media (max-width: 960px) {
    .hero { grid-template-columns: 1fr; padding: 24px; gap: 22px; }
    .stats { grid-template-columns: repeat(4, 1fr); }
    .stat { padding: 4px 0 4px 14px; }
    .stat:nth-child(n+3) { border-top: 0; }
    .stat b { font-size: 24px; }
    .thead { display: none; }
    .toolbar { grid-template-columns: 1fr 1fr; }
    .search { grid-column: 1 / -1; }
    select { width: 100%; }
    .row { padding: 16px 18px; }
    .row.cols { grid-template-columns: minmax(0, 1fr) auto; grid-template-areas: "who actions" "when seg" "last prod"; gap: 12px 14px; }
    .who { grid-area: who; } .actions { grid-area: actions; } .when { grid-area: when; } .seg { grid-area: seg; text-align: right; }
    .last { grid-area: last; } .prod { grid-area: prod; text-align: right; }
    .actions .btn span { display: none; }
    .actions .btn { width: 36px; padding: 0; }
  }
  @media (max-width: 520px) {
    .shell { padding: 20px 14px 80px; }
    .stats { grid-template-columns: 1fr 1fr; row-gap: 12px; }
    .stat:nth-child(n+3) { border-top: 0; }
    .top h1 { font-size: 26px; }
  }
  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation: none !important; transition: none !important; }
  }
</style>
</head>
<body>
<div class="shell">

  <header class="top">
    <div>
      <h1>Customer outreach</h1>
      <p id="subline">Loading your customers</p>
    </div>
    <div class="top-actions">
      <button class="btn" id="tplBtn" type="button">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
        Messages
      </button>
      <button class="btn" id="exportBtn" type="button">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></svg>
        Export CSV
      </button>
      <button class="icon-btn" id="refreshBtn" type="button" aria-label="Reload from Excel" title="Reload from Excel">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 3v6h-6"/></svg>
      </button>
      <button class="icon-btn" id="themeBtn" type="button" aria-label="Switch light or dark theme" title="Switch theme">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>
      </button>
    </div>
  </header>

  <section class="panel hero" aria-live="polite">
    <div class="hero-main">
      <div class="hero-num">
        <strong id="heroNum">0</strong>
        <span id="heroLabel">customers to message today</span>
      </div>
      <div class="progress">
        <div class="bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" id="barWrap"><i id="barFill"></i></div>
        <p id="heroSub">&nbsp;</p>
      </div>
    </div>
    <div class="stats">
      <div class="stat"><b id="stTotal">0</b><small>Customers</small></div>
      <div class="stat"><b id="stPriority">0</b><small>Priority, 3+ orders</small></div>
      <div class="stat"><b id="stReorder">0</b><small>Re-order, 2 orders</small></div>
      <div class="stat"><b id="stGap">-</b><small>Typical days between orders</small></div>
    </div>
  </section>

  <div class="toolbar">
    <label class="search">
      <span class="sr">Search customers</span>
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      <input id="search" type="search" placeholder="Search name, phone, email, order ID or product" autocomplete="off">
      <kbd>/</kbd>
    </label>
    <label><span class="sr">Segment</span>
      <select id="segment">
        <option value="all">All segments</option>
        <option value="Priority">Priority</option>
        <option value="Re-order">Re-order</option>
        <option value="Regular">Regular</option>
      </select>
    </label>
    <label><span class="sr">Sort by</span>
      <select id="sort">
        <option value="next">Sort: Next contact</option>
        <option value="orders">Sort: Most orders</option>
        <option value="recent">Sort: Latest order</option>
        <option value="ltv">Sort: Lifetime value</option>
        <option value="name">Sort: Name A to Z</option>
      </select>
    </label>
    <div class="chips" id="chips" role="group" aria-label="Filter by timing"></div>
  </div>

  <section class="panel">
    <div class="thead cols" id="thead">
      <div>Customer</div><div>Next contact</div><div>Last order</div><div>Segment</div><div>Favorite product</div><div></div>
    </div>
    <div id="skeletons"><div class="skeleton"><span></span></div><div class="skeleton"><span style="width:45%"></span></div><div class="skeleton"><span style="width:70%"></span></div><div class="skeleton"><span style="width:52%"></span></div></div>
    <ul class="list" id="list"></ul>
    <div class="state" id="state" hidden></div>
    <div class="more-wrap" id="moreWrap" hidden><button class="btn" id="moreBtn" type="button"></button></div>
  </section>
</div>

<dialog id="tplDialog">
  <form method="dialog" class="dlg" id="tplForm">
    <h2>Messages</h2>
    <p>Placeholders: <code>{name}</code> first name, <code>{product}</code> favorite product, <code>{whatsapp}</code> our WhatsApp link, <code>{call}</code> our call number.</p>

    <h3>WhatsApp</h3>
    <label class="sr" for="tplText">WhatsApp text</label>
    <textarea id="tplText" style="min-height:90px"></textarea>

    <h3>Email</h3>
    <label class="sr" for="emSubject">Email subject</label>
    <input type="text" id="emSubject" placeholder="Subject">
    <label class="sr" for="emBody">Email body</label>
    <textarea id="emBody"></textarea>

    <div class="preview" id="tplPreview"></div>
    <div class="dlg-actions">
      <button class="btn" type="button" id="tplReset">Reset</button>
      <button class="btn" type="button" id="tplCancel">Cancel</button>
      <button class="btn primary" type="submit" value="save">Save</button>
    </div>
  </form>
</dialog>

<div class="toast" id="toast" role="status"><span id="toastMsg"></span><button type="button" id="toastUndo" hidden>Undo</button></div>

<script>
(function () {
  'use strict';

  var PAGE = 40;

  /* ---------- our shop contact details ---------- */
  var SHOP_WA = '923000000000';          // WhatsApp number, digits only
  var SHOP_CALL = '+92 300 0000000';     // call number

  var DEFAULT_TPL = "Hi {name}! Is your {product} running low or the bottle empty? We also have new articles available. Message us here and we'll take care of it. Call: {call}";
  var DEFAULT_EM_SUBJECT = 'Is your perfume running low, {name}?';
  var DEFAULT_EM_BODY = "Hi {name},\n\nIs your {product} running low, or is the bottle empty? We also have new articles available at our store.\n\nContact us here:\nWhatsApp: {whatsapp}\nCall: {call}\n\nThank you,\nYour Store";

  var ICON_WA = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.4 8.4 0 0 1-12.4 7.4L3 20.5l1.7-5.4A8.4 8.4 0 1 1 21 11.5z"/></svg>';
  var ICON_CHECK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12.5 4.5 4.5L19 7.5"/></svg>';
  var ICON_COPY = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>';
  var ICON_MAIL = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>';

  var $ = function (s) { return document.querySelector(s); };
  var reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  var state = { all: [], view: [], today: '', q: '', timing: 'all', segment: 'all', sort: 'next', shown: PAGE, firstPaint: true, heroShown: false };

  /* ---------- storage ---------- */
  function load(key, fallback) { try { var v = localStorage.getItem(key); return v === null ? fallback : JSON.parse(v); } catch (e) { return fallback; } }
  function save(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) {} }
  var contacted = {};   // loaded from the server (contacted.json), so it survives restarts and other browsers
  var template = load('outreach.template', DEFAULT_TPL);
  var emSubject = load('outreach.emSubject', DEFAULT_EM_SUBJECT);
  var emBody = load('outreach.emBody', DEFAULT_EM_BODY);

  /* ---------- helpers ---------- */
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function dayNum(iso) { var p = iso.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / 864e5; }
  function fmtDate(iso) { return iso ? new Date(iso + 'T00:00:00Z').toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short', timeZone: 'UTC' }) : ''; }
  function fmtNum(n) { return Number(n).toLocaleString('en-US', { maximumFractionDigits: 0 }); }
  function plural(n, word) { return n + ' ' + word + (n === 1 ? '' : 's'); }
  function firstName(name) { var f = String(name).trim().split(/\s+/)[0] || 'there'; return f.charAt(0).toUpperCase() + f.slice(1).toLowerCase(); }
  function initials(name) { var parts = String(name).trim().split(/\s+/).filter(Boolean); var a = Array.from(parts[0] || '?')[0]; var b = parts.length > 1 ? Array.from(parts[parts.length - 1])[0] : ''; return (a + b).toUpperCase(); }
  // A customer stays marked as contacted until they place a NEW order.
  function isDone(c) { var r = contacted[c.key]; return !!r && r.orderId === c.lastOrderId; }
  function whoNote(c) { var r = contacted[c.key]; return esc(c.phone) + (isDone(c) ? ' \u00b7 Contacted ' + esc(fmtDate(r.date)) : ''); }
  function postContacted(set, remove) {
    return fetch('/api/contacted', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ set: set || {}, remove: remove || [] }) })
      .then(function (r) { if (!r.ok) throw new Error('save failed'); })
      .catch(function () { toast('Could not save to contacted.json. Is app.py still running?'); });
  }

  function enrich(c) {
    c.key = c.wa.slice(-10) || c.phone;
    var ref = c.projected || c.nextDate;
    c.days = ref && state.today ? dayNum(ref) - dayNum(state.today) : null;
    if (c.days === null) c.status = 'later';
    else if (c.days <= 0 && c.days >= -14) c.status = 'now';
    else if (c.days < -14) c.status = 'lapsed';
    else if (c.days <= 7) c.status = 'week';
    else c.status = 'later';
    c.hay = (c.name + ' ' + c.phone + ' ' + c.wa + ' ' + c.lastOrderId + ' ' + c.product + ' ' + c.email).toLowerCase();
    return c;
  }

  function whenLabel(c) {
    var d = c.days;
    if (d === null) return c.nextText || 'Not scheduled';
    if (c.status === 'lapsed') return 'Lapsed';
    if (d === 0) return 'Today';
    if (d === 1) return 'Tomorrow';
    if (d === -1) return 'Yesterday';
    if (d > 1) return 'In ' + d + ' days';
    return plural(-d, 'day') + ' overdue';
  }

  function whenNote(c) {
    if (c.status === 'lapsed') return 'Was due ' + esc(fmtDate(c.projected));
    if (c.status === 'now') return c.nextTime ? 'Best time ' + esc(c.nextTime) : 'Reach out now';
    return esc(fmtDate(c.nextDate)) + (c.nextTime ? ' at ' + esc(c.nextTime) : '');
  }

  /* ---------- message building ---------- */
  function fill(text, c) {
    var product = c.product && c.product !== 'N/A' ? c.product : 'perfume';
    return String(text)
      .replace(/\{name\}/g, firstName(c.name))
      .replace(/\{product\}/g, product)
      .replace(/\{whatsapp\}/g, 'https://wa.me/' + SHOP_WA)
      .replace(/\{call\}/g, SHOP_CALL);
  }
  function message(c) { return fill(template, c); }
  function mailHref(c) {
    if (!c.email) return '#';
    return 'mailto:' + encodeURIComponent(c.email).replace(/%40/g, '@') +
      '?subject=' + encodeURIComponent(fill(emSubject, c)) +
      '&body=' + encodeURIComponent(fill(emBody, c).replace(/\n/g, '\r\n'));
  }

  /* ---------- data ---------- */
  function fetchData(manual) {
    var btn = $('#refreshBtn');
    if (manual) btn.classList.add('spin');
    return fetch('/api/customers', { cache: 'no-store' })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, body: j }; }); })
      .then(function (res) {
        $('#skeletons').hidden = true;
        if (!res.ok) { showError(res.body); return; }
        state.today = res.body.today;
        contacted = res.body.contacted || {};
        state.all = res.body.customers.map(enrich);
        var legacy = load('outreach.contacted', null);
        if (legacy) {
          var migrate = {};
          state.all.forEach(function (c) {
            if (legacy[c.key] && !contacted[c.key]) { contacted[c.key] = migrate[c.key] = { date: legacy[c.key], orderId: c.lastOrderId }; }
          });
          if (Object.keys(migrate).length) postContacted(migrate, []);
          try { localStorage.removeItem('outreach.contacted'); } catch (e) {}
        }
        var built = new Date(res.body.builtAt);
        var withEmail = state.all.filter(function (c) { return c.email; }).length;
        $('#subline').textContent = plural(state.all.length, 'customer') + ' (' + withEmail + ' with email) from your Excel schedule, built ' + ago(built) + '.';
        buildChips();
        render(true);
        if (manual) toast('Reloaded from Excel');
      })
      .catch(function () { $('#skeletons').hidden = true; showError({ error: 'network' }); })
      .then(function () { setTimeout(function () { btn.classList.remove('spin'); }, 400); });
  }

  function ago(date) {
    var mins = Math.round((Date.now() - date.getTime()) / 60000);
    if (mins < 2) return 'just now';
    if (mins < 60) return mins + ' minutes ago';
    var hrs = Math.round(mins / 60);
    if (hrs < 24) return plural(hrs, 'hour') + ' ago';
    return plural(Math.round(hrs / 24), 'day') + ' ago';
  }

  function showError(err) {
    $('#list').innerHTML = ''; $('#moreWrap').hidden = true;
    var st = $('#state'); st.hidden = false;
    $('#subline').textContent = 'No data to show yet.';
    if (err.error === 'missing') st.innerHTML = '<h2>No schedule yet</h2><p>Build it from your Shopify export, then reload this page.</p><p><code>python customer_scheduler.py</code></p>';
    else if (err.error === 'unreadable') st.innerHTML = '<h2>The Excel file could not be read</h2><p>Close it in Excel if it is open, then run <code>python customer_scheduler.py</code> again.</p><p>' + esc(err.detail || '') + '</p>';
    else st.innerHTML = '<h2>Can\'t reach the server</h2><p>Check that <code>python app.py</code> is still running, then reload.</p>';
  }

  /* ---------- filtering ---------- */
  var TIMINGS = [['all', 'All'], ['now', 'Due now'], ['todo', 'To contact'], ['contacted', 'Contacted'], ['week', 'Next 7 days'], ['later', 'Later'], ['lapsed', 'Lapsed']];

  function buildChips() {
    var counts = { all: state.all.length };
    state.all.forEach(function (c) {
      counts[c.status] = (counts[c.status] || 0) + 1;
      if (isDone(c)) counts.contacted = (counts.contacted || 0) + 1;
      else if (c.status === 'now') counts.todo = (counts.todo || 0) + 1;
    });
    $('#chips').innerHTML = TIMINGS.map(function (t) {
      return '<button class="chip" type="button" data-t="' + t[0] + '" aria-pressed="' + (state.timing === t[0]) + '">' + t[1] + ' <em>' + fmtNum(counts[t[0]] || 0) + '</em></button>';
    }).join('');
  }

  var SORTS = {
    next: function (a, b) {
      var rank = { now: 0, week: 1, later: 2, lapsed: 3 };
      return (isDone(a) - isDone(b)) || rank[a.status] - rank[b.status] || (a.days === null ? 1e9 : a.days) - (b.days === null ? 1e9 : b.days);
    },
    orders: function (a, b) { return b.orders - a.orders || SORTS.next(a, b); },
    recent: function (a, b) { return (b.lastOrderDate || '').localeCompare(a.lastOrderDate || ''); },
    ltv: function (a, b) { return (b.ltv || 0) - (a.ltv || 0); },
    name: function (a, b) { return a.name.localeCompare(b.name); }
  };

  function applyFilters() {
    var q = state.q.trim().toLowerCase();
    var qDigits = q.replace(/\D/g, '');
    state.view = state.all.filter(function (c) {
      if (state.timing === 'contacted') { if (!isDone(c)) return false; }
      else if (state.timing === 'todo') { if (isDone(c) || c.status !== 'now') return false; }
      else if (state.timing !== 'all' && c.status !== state.timing) return false;
      if (state.segment !== 'all' && c.segment !== state.segment) return false;
      if (!q) return true;
      return c.hay.indexOf(q) !== -1 || (qDigits.length > 2 && (c.wa.indexOf(qDigits) !== -1 || c.phone.replace(/\D/g, '').indexOf(qDigits) !== -1));
    }).sort(SORTS[state.sort]);
  }

  /* ---------- rendering ---------- */
  function rowHTML(c, i, enter) {
    var done = isDone(c);
    var href = c.wa ? 'https://wa.me/' + c.wa + '?text=' + encodeURIComponent(message(c)) : '#';
    var lateCls = c.status === 'now' && c.days < 0 ? 'late' : c.status;
    var product = c.product && c.product !== 'N/A' ? c.product : 'No product on record';
    var gap = c.avgDays ? 'Orders every ' + Math.round(c.avgDays) + ' days' : 'One order so far';
    var mailTitle = c.email ? 'Email ' + esc(c.email) : 'No email on file';
    return '<li class="row cols' + (enter ? ' enter' : '') + (done ? ' is-done' : '') + '" data-key="' + esc(c.key) + '" style="--i:' + i + '">' +
      '<div class="cell who"><span class="avatar" aria-hidden="true">' + esc(initials(c.name)) + '</span><div><strong title="' + esc(c.name) + '">' + esc(c.name) + '</strong><small>' + whoNote(c) + '</small></div></div>' +
      '<div class="cell when ' + lateCls + '"><strong>' + whenLabel(c) + '</strong><small>' + whenNote(c) + '</small></div>' +
      '<div class="cell last"><strong>' + esc(c.lastOrderId || 'N/A') + '</strong><small>' + esc(fmtDate(c.lastOrderDate)) + '</small></div>' +
      '<div class="cell seg"><span class="pill ' + c.segment + '">' + c.segment + '</span><small>' + plural(c.orders, 'order') + '</small></div>' +
      '<div class="cell prod"><strong title="' + esc(product) + '">' + esc(product) + '</strong><small>' + gap + '</small></div>' +
      '<div class="actions">' +
        '<a class="btn wa" data-act="wa" target="_blank" rel="noopener" href="' + href + '"' + (c.wa ? '' : ' aria-disabled="true"') + '>' + ICON_WA + '<span>WhatsApp</span></a>' +
        '<a class="icon-btn" data-act="mail" href="' + esc(mailHref(c)) + '"' + (c.email ? '' : ' aria-disabled="true"') + ' aria-label="Send email" title="' + mailTitle + '">' + ICON_MAIL + '</a>' +
        '<button class="icon-btn done" type="button" data-act="done" aria-pressed="' + done + '" aria-label="Mark as contacted today" title="Mark as contacted">' + ICON_CHECK + '</button>' +
        '<button class="icon-btn" type="button" data-act="copy" aria-label="Copy phone number" title="Copy phone number">' + ICON_COPY + '</button>' +
      '</div></li>';
  }

  function renderList(animateSwap) {
    applyFilters();
    var list = $('#list'), st = $('#state');
    var slice = state.view.slice(0, state.shown);
    var enter = state.firstPaint;
    list.innerHTML = slice.map(function (c, i) { return rowHTML(c, i, enter && i < 14); }).join('');
    state.firstPaint = false;

    if (animateSwap && !reduceMotion) { list.classList.remove('swap'); void list.offsetWidth; list.classList.add('swap'); }

    st.hidden = state.view.length > 0;
    if (!st.hidden) {
      var filtered = state.q || state.timing !== 'all' || state.segment !== 'all';
      st.innerHTML = filtered
        ? '<h2>No customers match</h2><p>Try a different search or clear the filters.</p><button class="btn" type="button" id="clearBtn">Clear filters</button>'
        : '<h2>No customers yet</h2><p>Run <code>python customer_scheduler.py</code> to build the schedule.</p>';
    }
    var left = state.view.length - slice.length;
    $('#moreWrap').hidden = left <= 0;
    $('#moreBtn').textContent = 'Show ' + Math.min(PAGE, left) + ' more (' + fmtNum(left) + ' left)';
  }

  function countUp(el, to, ms) {
    var from = Number(el.dataset.v || 0);
    el.dataset.v = to;
    if (reduceMotion || from === to) { el.textContent = fmtNum(to); return; }
    var start = performance.now();
    (function tick(now) {
      var t = Math.min(1, (now - start) / ms);
      var eased = 1 - Math.pow(1 - t, 3);
      el.textContent = fmtNum(Math.round(from + (to - from) * eased));
      if (t < 1) requestAnimationFrame(tick);
    })(start);
  }

  function renderHero() {
    var due = state.all.filter(function (c) { return c.status === 'now'; });
    var done = due.filter(isDone).length;
    var left = due.length - done;
    var ms = state.heroShown ? 450 : 1100;
    state.heroShown = true;

    countUp($('#heroNum'), left, ms);
    $('#heroLabel').textContent = left === 1 ? 'customer to message today' : 'customers to message today';
    var pct = due.length ? Math.round(done / due.length * 100) : 0;
    $('#barFill').style.width = pct + '%';
    $('#barWrap').setAttribute('aria-valuenow', pct);
    $('#heroSub').textContent = !due.length ? 'Nobody is due right now.' : left === 0 ? 'All ' + due.length + ' done. Nice work.' : done + ' of ' + due.length + ' done';

    var total = state.all.length;
    countUp($('#stTotal'), total, ms);
    countUp($('#stPriority'), state.all.filter(function (c) { return c.segment === 'Priority'; }).length, ms);
    countUp($('#stReorder'), state.all.filter(function (c) { return c.segment === 'Re-order'; }).length, ms);
    var gaps = state.all.filter(function (c) { return c.avgDays; }).map(function (c) { return c.avgDays; }).sort(function (a, b) { return a - b; });
    $('#stGap').textContent = gaps.length ? Math.round(gaps[Math.floor(gaps.length / 2)]) : '-';
  }

  function render(first) {
    renderHero();
    renderList(false);
    if (first) buildChips();
  }

  /* ---------- toast ---------- */
  var toastTimer, undoFn = null;
  function toast(msg, undo) {
    $('#toastMsg').textContent = msg;
    undoFn = undo || null;
    $('#toastUndo').hidden = !undo;
    $('#toast').classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { $('#toast').classList.remove('show'); }, undo ? 5000 : 2400);
  }
  $('#toastUndo').addEventListener('click', function () { if (undoFn) undoFn(); $('#toast').classList.remove('show'); });

  /* ---------- actions ---------- */
  function setDone(c, value) {
    if (value) {
      contacted[c.key] = { date: state.today, orderId: c.lastOrderId };
      postContacted((function () { var o = {}; o[c.key] = contacted[c.key]; return o; })(), []);
    } else {
      delete contacted[c.key];
      postContacted({}, [c.key]);
    }
    var row = document.querySelector('.row[data-key="' + CSS.escape(c.key) + '"]');
    if (row) {
      row.classList.toggle('is-done', value);
      row.querySelector('[data-act="done"]').setAttribute('aria-pressed', value);
      row.querySelector('.who small').innerHTML = whoNote(c);
    }
    renderHero();
    buildChips();
  }

  $('#list').addEventListener('click', function (e) {
    var btn = e.target.closest('[data-act]');
    if (!btn) return;
    var row = btn.closest('.row');
    var c = state.all.find(function (x) { return x.key === row.dataset.key; });
    if (!c) return;
    var act = btn.dataset.act;

    if (act === 'wa' || act === 'mail') {
      if (!isDone(c)) { setDone(c, true); toast('Marked ' + firstName(c.name) + ' as contacted', function () { setDone(c, false); }); }
    } else if (act === 'done') {
      var next = !isDone(c);
      setDone(c, next);
      toast(next ? 'Marked ' + firstName(c.name) + ' as contacted' : 'Marked as not contacted', next ? function () { setDone(c, false); } : null);
    } else if (act === 'copy') {
      var text = c.wa ? '+' + c.wa : c.phone;
      (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(function () { toast('Copied ' + text); }, function () { toast('Copy is blocked here. Number: ' + text); });
    }
  });

  $('#chips').addEventListener('click', function (e) {
    var b = e.target.closest('.chip'); if (!b) return;
    state.timing = b.dataset.t; state.shown = PAGE;
    buildChips(); renderList(true);
  });

  var searchTimer;
  $('#search').addEventListener('input', function (e) {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(function () { state.q = e.target.value; state.shown = PAGE; renderList(false); }, 120);
  });
  $('#segment').addEventListener('change', function (e) { state.segment = e.target.value; state.shown = PAGE; renderList(true); });
  $('#sort').addEventListener('change', function (e) { state.sort = e.target.value; state.shown = PAGE; renderList(true); });
  $('#moreBtn').addEventListener('click', function () { state.shown += PAGE; renderList(false); });
  $('#state').addEventListener('click', function (e) {
    if (e.target.id !== 'clearBtn') return;
    state.q = ''; state.timing = 'all'; state.segment = 'all'; state.shown = PAGE;
    $('#search').value = ''; $('#segment').value = 'all';
    buildChips(); renderList(true);
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === '/' && !/^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName)) { e.preventDefault(); $('#search').focus(); }
  });

  $('#refreshBtn').addEventListener('click', function () { fetchData(true); });

  $('#themeBtn').addEventListener('click', function () {
    var next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem('outreach.theme', next); } catch (e) {}
  });

  /* ---------- export ---------- */
  $('#exportBtn').addEventListener('click', function () {
    var head = ['Customer Name', 'Phone', 'WhatsApp Number', 'Email', 'Segment', 'Orders', 'Last Order ID', 'Last Order Date', 'Next Contact Date', 'Next Contact Time', 'Favorite Product', 'Lifetime Value', 'Contacted On'];
    var rows = state.view.map(function (c) { return [c.name, c.phone, c.wa, c.email, c.segment, c.orders, c.lastOrderId, c.lastOrderDate, c.nextDate, c.nextTime, c.product, c.ltv == null ? '' : c.ltv, isDone(c) ? contacted[c.key].date : '']; });
    var csv = [head].concat(rows).map(function (r) { return r.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
    var url = URL.createObjectURL(new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' }));
    var a = document.createElement('a'); a.href = url; a.download = 'customers-' + state.today + '.csv';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    toast('Exported ' + plural(rows.length, 'customer'));
  });

  /* ---------- messages dialog ---------- */
  var dlg = $('#tplDialog'), tplText = $('#tplText'), emSubj = $('#emSubject'), emBodyEl = $('#emBody');
  function updatePreview() {
    var s = state.all[0] || { name: 'Ayesha', product: 'Your perfume' };
    $('#tplPreview').textContent = 'WhatsApp:\n' + fill(tplText.value, s) + '\n\nEmail:\n' + fill(emSubj.value, s) + '\n\n' + fill(emBodyEl.value, s);
  }
  $('#tplBtn').addEventListener('click', function () {
    tplText.value = template; emSubj.value = emSubject; emBodyEl.value = emBody;
    updatePreview(); dlg.showModal(); tplText.focus();
  });
  [tplText, emSubj, emBodyEl].forEach(function (el) { el.addEventListener('input', updatePreview); });
  $('#tplReset').addEventListener('click', function () {
    tplText.value = DEFAULT_TPL; emSubj.value = DEFAULT_EM_SUBJECT; emBodyEl.value = DEFAULT_EM_BODY; updatePreview();
  });
  $('#tplCancel').addEventListener('click', function () { dlg.close(); });
  dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });
  $('#tplForm').addEventListener('submit', function () {
    template = tplText.value.trim() || DEFAULT_TPL;
    emSubject = emSubj.value.trim() || DEFAULT_EM_SUBJECT;
    emBody = emBodyEl.value.trim() || DEFAULT_EM_BODY;
    save('outreach.template', template);
    save('outreach.emSubject', emSubject);
    save('outreach.emBody', emBody);
    renderList(false);
    toast('Messages saved');
  });

  fetchData(false);
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG") == "1",
    )