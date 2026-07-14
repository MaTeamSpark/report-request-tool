"""
flask_api.py — Marketing Analytics Report Automation Backend
Deploy on your VM. Run with: python flask_api.py
Requires: pip install flask flask-cors requests pyodbc pandas openpyxl
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import requests
import json
import os
import sys
import uuid
from datetime import datetime, timedelta

# ── Redirect stdout/stderr to both console and log file ──────────────────────
class _Tee:
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            try: f.write(obj); f.flush()
            except: pass
    def flush(self):
        for f in self.files:
            try: f.flush()
            except: pass

try:
    _log_file = open(
        r"E:\Marketing Analytics\Report Request Automation\flask_log.txt",
        "a", encoding="utf-8", buffering=1
    )
    sys.stdout = _Tee(sys.__stdout__, _log_file)
    sys.stderr = _Tee(sys.__stderr__, _log_file)
except Exception:
    pass
# ─────────────────────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app)

# ── Config ──────────────────────────────────────────────────────────────────
# ── Config — edit config.json instead of this file ──────────────────────────
CONFIG_FILE = r"E:\Marketing Analytics\Report Request Automation\config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

_cfg = load_config()

TEAMS_WEBHOOK_URL        = _cfg.get('TEAMS_WEBHOOK_URL', 'YOUR_WEBHOOK_URL_HERE')
TEAMS_WEBHOOK_USER_EMAIL = _cfg.get('TEAMS_WEBHOOK_USER_EMAIL', '')
ADMIN_EMAIL              = "mohd.qasim@bain.com"
ADMIN_PAGE_URL           = "https://ma-reports.bain.com/admin.html"
OUTPUT_FOLDER            = r"E:\Marketing Analytics\Report Request Automation\reports_output"

# ── Email config — from config.json ──────────────────────────────────────────
SMTP_SERVER   = _cfg.get('SMTP_SERVER', 'mail.apac.bain.com')
SMTP_PORT     = _cfg.get('SMTP_PORT', 25)
EMAIL_FROM    = _cfg.get('EMAIL_FROM', 'Marketing.Analytics@bain.com')
EMAIL_ENABLED = _cfg.get('EMAIL_ENABLED', True)

# ── SQL Server config ────────────────────────────────────────────────────────
SQL_SERVER        = "wtrdynsql2v"
SQL_DRIVER        = "ODBC Driver 18 for SQL Server"
SQL_TRUSTED       = "yes"
SQL_TRUST_CERT    = "yes"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ── Persistent request store — survives Flask restarts ───────────────────────
STORE_FILE = r"E:\Marketing Analytics\Report Request Automation\requests_store.json"

def load_store():
    if os.path.exists(STORE_FILE):
        try:
            with open(STORE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_store():
    try:
        with open(STORE_FILE, 'w', encoding='utf-8') as f:
            json.dump(requests_store, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[Store] Failed to save: {e}")

requests_store = load_store()
print(f"[Store] Loaded {len(requests_store)} existing requests from disk")


# ────────────────────────────────────────────────────────────────────────────
# SQL CONNECTION HELPER — uses SQLAlchemy to avoid pandas warning
# ────────────────────────────────────────────────────────────────────────────
def get_connection(database):
    from sqlalchemy import create_engine
    import urllib
    conn_str = (
        f"DRIVER={{{SQL_DRIVER}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={database};"
        f"Trusted_Connection={SQL_TRUSTED};"
        f"TrustServerCertificate={SQL_TRUST_CERT};"
    )
    params = urllib.parse.quote_plus(conn_str)
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
    return engine


# ────────────────────────────────────────────────────────────────────────────
# SQL PIPELINE — runs queries and saves CSVs
# ────────────────────────────────────────────────────────────────────────────
def run_sql_pipeline(req):
    import pandas as pd

    request_id   = req.get('requestId')
    report_name  = req.get('reportName', 'Report').replace(' ', '_')
    final_queries = req.get('finalQueries', {})
    timestamp    = datetime.now().strftime('%Y%m%d_%H%M%S')

    csv_files = {}
    errors    = []

    # Database map per source — all use MarketingSF
    db_map = {
        'ga':          'MarketingSF',
        'marketo':     'MarketingSF',
        'linkedin':    'MarketingSF',
        'url_mapping': 'MarketingSF',
        'li':          'MarketingSF',   # fallback
        'mk':          'MarketingSF',   # fallback
    }

    for source, query in final_queries.items():
        if not query or not query.strip():
            continue
        db = db_map.get(source)
        if not db:
            continue

        source_label = source.upper()
        filename     = f"{request_id}_{report_name}_{source_label}_{timestamp}.csv"
        filepath     = os.path.join(OUTPUT_FOLDER, filename)

        try:
            print(f"[SQL] Running {source_label} query on {db}...")
            engine = get_connection(db)
            with engine.connect() as conn:
                df = pd.read_sql(query, conn)
            row_count = len(df)
            # For URL_MAPPING — add metadata in G1/H1 and J1/K1
            if source == 'url_mapping':
                import csv as csv_mod
                slice_by = req.get('sliceBy', '—')
                pa_option = req.get('slide2', {}).get('accountListOption', 'default')
                pa_value  = 'Default' if pa_option == 'default' else 'Uploaded'
                with open(filepath, 'w', newline='', encoding='utf-8') as f:
                    writer = csv_mod.writer(f)
                    # Header row with metadata in G1/H1 and J1/K1
                    headers = list(df.columns) + ['', '', 'Data segmentation', slice_by, '', 'PA List', pa_value]
                    writer.writerow(headers)
                    # Data rows
                    for _, row in df.iterrows():
                        writer.writerow(list(row) + ['', '', '', '', '', '', ''])
            else:
                df.to_csv(filepath, index=False)
            csv_files[source] = {'filename': filename, 'rows': row_count}
            print(f"[SQL] {source_label} → {row_count} rows → {filename}")

        except Exception as e:
            # Extract clean error message
            import re
            err_str = str(e)
            match = re.search(r'\[SQL Server\](.*?)\.?\s*\(', err_str)
            clean_err = match.group(1).strip() if match else err_str[:120]

            errors.append(f"{source_label}: {clean_err}")
            print(f"[SQL] ERROR {source_label}: {clean_err}")
            # Always create CSV even on error — with error message as header
            import csv as csv_mod
            with open(filepath, 'w', newline='') as f:
                writer = csv_mod.writer(f)
                writer.writerow(['ERROR', clean_err])
            csv_files[source] = {'filename': filename, 'rows': 0, 'error': clean_err}
            print(f"[SQL] {source_label} → empty error CSV created")

    return csv_files, errors


# ────────────────────────────────────────────────────────────────────────────
# ROUTE 1: User submits form
# POST /submit-request
# ────────────────────────────────────────────────────────────────────────────
@app.route('/submit-request', methods=['POST'])
def submit_request():
    # Handle multipart form (file upload) or JSON
    if request.content_type and 'multipart' in request.content_type:
        data = json.loads(request.form.get('payload', '{}'))
        uploaded_file = request.files.get('accountFile')
        if uploaded_file and uploaded_file.filename:
            request_id = data.get('requestId', 'REQ-' + str(uuid.uuid4())[:8].upper())
            safe_name  = f"{request_id}_Priority_Accounts_{uploaded_file.filename}"
            save_path  = os.path.join(OUTPUT_FOLDER, safe_name)
            uploaded_file.save(save_path)
            data['accountInputFile'] = safe_name
            print(f"[Submit] Uploaded account file saved → {safe_name}")
    else:
        data = request.json

    request_id = data.get('requestId', 'REQ-' + str(uuid.uuid4())[:8].upper())

    # Save manual account list as CSV if no file uploaded
    if not data.get('accountInputFile'):
        account_list = data.get('slide2', {}).get('accountList', [])
        if account_list:
            import csv as csv_mod
            report_name  = data.get('reportName', 'Report').replace(' ', '_')
            acct_filename = f"{request_id}_{report_name}_Priority_Accounts.csv"
            acct_filepath = os.path.join(OUTPUT_FOLDER, acct_filename)
            with open(acct_filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv_mod.writer(f)
                writer.writerow(['Company Name'])
                for company in account_list:
                    if company.strip():
                        writer.writerow([company.strip()])
            data['accountInputFile'] = acct_filename
            print(f"[Submit] Manual account list saved → {len(account_list)} companies → {acct_filename}")

    data['status']      = 'pending'
    data['submittedAt'] = datetime.utcnow().isoformat()
    requests_store[request_id] = data
    save_store()

    _notify_admin_new_request(data)
    _email_user_received(data)

    return jsonify({ 'success': True, 'requestId': request_id }), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE 2: Admin approves — triggers SQL pipeline
# POST /admin/approve
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/approve', methods=['POST'])
def approve_request():
    data         = request.json
    request_id   = data.get('requestId')
    final_queries = data.get('finalQueries', {})

    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404

    req = requests_store[request_id]
    req['status']       = 'running'
    req['finalQueries'] = final_queries
    req['approvedAt']   = datetime.utcnow().isoformat()
    req.pop('forceStopped', None)
    save_store()

    # Run pipeline in background thread so API returns immediately
    import threading
    threading.Thread(target=_run_pipeline_thread, args=(request_id,), daemon=True).start()

    return jsonify({ 'success': True }), 200


def _run_pipeline_thread(request_id):
    req = requests_store.get(request_id)
    if not req:
        return

    try:
        csv_files, errors = run_sql_pipeline(req)

        # If admin force-stopped this request while the SQL was running, discard results
        req = requests_store.get(request_id)  # re-fetch in case it changed
        if not req or req.get('forceStopped'):
            print(f"[Pipeline] Discarding results for {request_id} — was force-stopped")
            return

        req['status']      = 'completed'
        req['completedAt'] = datetime.utcnow().isoformat()
        req['csvFiles']    = csv_files
        req['errors']      = errors

        # Build download URLs with row counts
        download_links = {}
        for source, info in csv_files.items():
            filename = info['filename'] if isinstance(info, dict) else info
            rows     = info['rows'] if isinstance(info, dict) else '?'
            download_links[source] = {
                'url':      f"https://ma-reports.bain.com/downloads/{filename}",
                'rows':     rows,
                'filename': filename
            }

        req['downloadLinks'] = download_links

        # Add priority accounts input file
        account_input = req.get('accountInputFile')
        if account_input and os.path.exists(os.path.join(OUTPUT_FOLDER, account_input)):
            # Count rows in the file
            try:
                with open(os.path.join(OUTPUT_FOLDER, account_input), 'r', encoding='utf-8') as f:
                    row_count = sum(1 for line in f) - 1  # subtract header
            except:
                row_count = '—'
            req['downloadLinks']['input_accounts'] = {
                'url':      f"https://ma-reports.bain.com/downloads/{account_input}",
                'rows':     row_count,
                'filename': account_input
            }
            print(f"[Pipeline] Priority accounts file added → {row_count} companies")

        # Generate Request Summary CSV — all input fields in 2 columns
        report_name_safe = req.get('reportName', 'Report').replace(' ', '_')
        summary_timestamp = datetime.utcnow().strftime('%Y%m%d_%H%M%S')
        summary_filename = f"{request_id}_{report_name_safe}_Request_Summary_{summary_timestamp}.csv"
        summary_filepath = os.path.join(OUTPUT_FOLDER, summary_filename)
        try:
            import csv as csv_mod
            user = req.get('user', {})
            slide2 = req.get('slide2', {})
            pa_value = 'Default' if slide2.get('accountListOption', 'default') == 'default' else 'Uploaded'
            chapter_urls = req.get('chapterUrls', [])

            summary_rows = [
                ('Field', 'Value'),
                ('Request ID', req.get('requestId', '—')),
                ('Report Name', req.get('reportName', '—')),
                ('Requested By', f"{user.get('name', '—')} ({user.get('email', '—')})"),
                ('Landing Page URL', req.get('landingUrl', '—')),
                ('Chapter Scope', req.get('chapterScope', '—')),
                ('Chapter URLs', '; '.join(chapter_urls) if chapter_urls else '—'),
                ('Report Publish Date', req.get('publishDate', '—')),
                ('Analysis Period', _format_period(req.get('period', {}))),
                ('Data Sources', ', '.join(req.get('dataSources', []))),
                ('Data Segmentation', req.get('sliceBy', '—')),
                ('Priority Account List', pa_value),
                ('Submitted At', _to_ist_str(req.get('submittedAt', ''))),
                ('Status', req.get('status', '—')),
            ]
            with open(summary_filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv_mod.writer(f)
                writer.writerows(summary_rows)

            req['downloadLinks']['request_summary'] = {
                'url':      f"https://ma-reports.bain.com/downloads/{summary_filename}",
                'rows':     len(summary_rows) - 1,
                'filename': summary_filename
            }
            print(f"[Pipeline] Request Summary CSV generated → {summary_filename}")
        except Exception as e:
            print(f"[Pipeline] Request Summary CSV failed: {e}")

        save_store()  # Persist completed state

        # Notify admin on Teams that report is ready
        _notify_report_ready(req)
        # Email report ready removed — MA team sends final Excel manually
        print(f"[Pipeline] Completed for {request_id}: {csv_files}")

    except Exception as e:
        req['status'] = 'error'
        req['error']  = str(e)
        print(f"[Pipeline] FAILED for {request_id}: {e}")


# ────────────────────────────────────────────────────────────────────────────
# ROUTE 3: Admin sends modification feedback
# POST /admin/modify
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/modify', methods=['POST'])
def modify_request():
    data           = request.json
    request_id     = data.get('requestId')
    note           = data.get('note', '')
    flagged_fields = data.get('flaggedFields', [])

    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404

    req = requests_store[request_id]
    req['status']           = 'needs_modification'
    req['modificationNote'] = note
    req['flaggedFields']    = flagged_fields
    save_store()

    # Email user about modification needed
    _email_user_needs_modification(req, note, flagged_fields)

    return jsonify({ 'success': True }), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE 4: Get all requests (admin dashboard)
# GET /admin/requests
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/requests', methods=['GET'])
def get_requests():
    sorted_reqs = sorted(
        requests_store.values(),
        key=lambda r: r.get('submittedAt', ''),
        reverse=True
    )
    return jsonify(sorted_reqs), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE 5: Get single request by ID (for form pre-fill)
# GET /request/<request_id>
# ────────────────────────────────────────────────────────────────────────────
@app.route('/request/<request_id>', methods=['GET'])
def get_request(request_id):
    req = requests_store.get(request_id)
    if not req:
        return jsonify({ 'error': 'Request not found' }), 404
    return jsonify(req), 200


# ────────────────────────────────────────────────────────────────────────────
# TEAMS NOTIFICATIONS
# ────────────────────────────────────────────────────────────────────────────
def _now_ist():
    """Current time in IST (UTC+5:30)."""
    return datetime.utcnow() + timedelta(hours=5, minutes=30)


def _to_ist_str(iso_str, fmt='%d %b %Y %H:%M IST'):
    """Convert a stored UTC ISO string to IST display string."""
    if not iso_str:
        return '—'
    try:
        dt = datetime.fromisoformat(iso_str[:19])
        dt_ist = dt + timedelta(hours=5, minutes=30)
        return dt_ist.strftime(fmt)
    except Exception:
        return iso_str


def _format_period(period):
    if not period:
        return '—'
    if period.get('type') == 'from_publish':
        return f"First {period.get('days')} days ({period.get('startDate', '')} → {period.get('endDate', '')})"
    elif period.get('type') == 'custom':
        return f"Custom ({period.get('startDate', '')} → {period.get('endDate', '')})"
    else:
        return f"Last {period.get('days')} days"


def _send_teams_message(card_payload):
    """Send an Adaptive Card to the Teams channel via webhook."""
    try:
        resp = requests.post(
            TEAMS_WEBHOOK_URL,
            json=card_payload,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        print(f"[Teams] Status: {resp.status_code} Response: {resp.text[:200]}")
    except Exception as e:
        print(f"[Teams] Failed to send notification: {e}")


def _notify_request_closed(req):
    """Notify admin on Teams that request has been closed."""
    user = req.get('user', {})
    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    { "type": "TextBlock", "text": "🔒 Request Closed", "weight": "Bolder", "size": "Medium", "color": "Good" },
                    {
                        "type": "FactSet",
                        "facts": [
                            { "title": "Report",      "value": req.get('reportName', '—') },
                            { "title": "Request ID",  "value": req.get('requestId', '—') },
                            { "title": "Requested by","value": f"{user.get('name', '—')} ({user.get('email', '—')})" },
                            { "title": "Period",      "value": _format_period(req.get('period', {})) },
                            { "title": "Closed at",   "value": _now_ist().strftime('%d %b %Y %H:%M IST') }
                        ]
                    },
                    { "type": "TextBlock", "text": "This request has been reviewed, verified and closed.", "wrap": True, "isSubtle": True }
                ]
            }
        }]
    }
    _send_teams_message(card)


def _notify_report_ready(req):
    """Notify admin on Teams that report is ready."""
    csv_files  = req.get('csvFiles', {})
    errors     = req.get('errors', [])
    all_links  = req.get('downloadLinks', {})
    file_count = len(all_links) if all_links else len(csv_files)

    body = [
        { "type": "TextBlock", "text": "✅ Report Ready!", "weight": "Bolder", "size": "Medium", "color": "Good" },
        {
            "type": "FactSet",
            "facts": [
                { "title": "Report",     "value": req.get('reportName', '—') },
                { "title": "Request ID", "value": req.get('requestId', '—') },
                { "title": "Period",     "value": _format_period(req.get('period', {})) },
                { "title": "Files",      "value": f"{file_count} CSV file(s) generated" },
                { "title": "Completed",  "value": _now_ist().strftime('%d %b %Y %H:%M IST') }
            ]
        },
        {
            "type": "TextBlock",
            "text": "Files are ready to download from the Admin Dashboard.",
            "wrap": True, "isSubtle": True, "spacing": "Medium"
        }
    ]

    if errors:
        import re
        clean_errors = []
        for e in errors:
            match = re.search(r'\[SQL Server\](.*?)\.?\s*\(', e)
            clean_errors.append(match.group(1).strip() if match else e[:100])
        body.append({ "type": "TextBlock", "text": "⚠ Query errors: " + "; ".join(clean_errors), "wrap": True, "color": "Attention" })

    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": body,
                "actions": [{
                    "type": "Action.OpenUrl",
                    "title": "Open Admin Dashboard →",
                    "url": ADMIN_PAGE_URL,
                    "style": "positive"
                }]
            }
        }]
    }
    _send_teams_message(card)


def _notify_admin_new_request(req):
    user       = req.get('user', {})
    sources    = ', '.join(req.get('dataSources', []))
    slice_by   = req.get('sliceBy', '—')
    pa_option  = req.get('slide2', {}).get('accountListOption', 'default')
    pa_list    = 'Default' if pa_option == 'default' else 'Uploaded'
    period     = req.get('period', {})
    period_str = _format_period(period)

    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    {
                        "type": "Container",
                        "style": "emphasis",
                        "items": [{
                            "type": "ColumnSet",
                            "columns": [
                                { "type": "Column", "width": "stretch", "items": [
                                    { "type": "TextBlock", "text": "📊 New Report Request", "weight": "Bolder", "size": "Medium" },
                                    { "type": "TextBlock", "text": f"Request ID: **{req.get('requestId')}**", "spacing": "None", "isSubtle": True }
                                ]},
                                { "type": "Column", "width": "auto", "items": [
                                    { "type": "TextBlock", "text": "PENDING REVIEW", "color": "Warning", "weight": "Bolder", "size": "Small" }
                                ]}
                            ]
                        }]
                    },
                    {
                        "type": "FactSet",
                        "facts": [
                            { "title": "Report",       "value": req.get('reportName', '—') },
                            { "title": "Requested by", "value": f"{user.get('name')} ({user.get('email')})" },
                            { "title": "Period",       "value": period_str },
                            { "title": "Sources",      "value": sources },
                            { "title": "Data segmentation",     "value": slice_by },
                            { "title": "Priority account list", "value": pa_list },
                            { "title": "URL",          "value": req.get('landingUrl', '—') },
                            { "title": "Submitted",    "value": _to_ist_str(req.get('submittedAt', '')) }
                        ]
                    },
                    { "type": "TextBlock", "text": "SQL queries auto-generated. Click below to review and approve.", "wrap": True, "isSubtle": True, "spacing": "Medium" }
                ],
                "actions": [{
                    "type": "Action.OpenUrl",
                    "title": "Review & Approve →",
                    "url": ADMIN_PAGE_URL,
                    "style": "positive"
                }]
            }
        }]
    }
    _send_teams_message(card)



def _send_email(to_email, subject, html_body):
    """Send email via Bain internal SMTP."""
    if not EMAIL_ENABLED:
        print(f"[Email] Skipped (disabled) → {to_email}")
        return
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From']    = f"Marketing Analytics <{EMAIL_FROM}>"
        msg['To']      = to_email
        msg['Cc']      = 'Marketing.Analytics@bain.com'
        msg.attach(MIMEText(html_body, 'html'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.sendmail(EMAIL_FROM, [to_email, 'Marketing.Analytics@bain.com'], msg.as_string())
        print(f"[Email] Sent → {to_email} | {subject}")
    except Exception as e:
        print(f"[Email] Failed → {to_email} | {e}")


def _email_user_received(req):
    """Email user when request is received."""
    user = req.get('user', {})
    to   = user.get('email', '')
    if not to:
        return
    subject = f"Report Request Received — {req.get('reportName', '')} [{req.get('requestId', '')}]"
    period_str = _format_period(req.get('period', {}))
    html = f"""
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#CC0000;padding:16px 24px;">
        <h2 style="color:white;margin:0;font-size:18px;">Marketing Analytics · Report Request</h2>
      </div>
      <div style="padding:24px;background:#ffffff;border:1px solid #e0d9d2;">
        <p style="font-size:15px;">Hi {user.get('name', 'there').split()[0]},</p>
        <p style="margin-top:12px;">Your report request has been received and is currently under review by the MA team.</p>
        <div style="background:#F8F5F0;border-radius:8px;padding:16px;margin:20px 0;">
          <table style="font-size:14px;width:100%;">
            <tr><td style="color:#888;padding:4px 0;width:160px;">Request ID</td><td><strong>{req.get('requestId', '')}</strong></td></tr>
            <tr><td style="color:#888;padding:4px 0;">Report</td><td>{req.get('reportName', '')}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Reporting period</td><td>{period_str}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Data sources</td><td>{', '.join(req.get('dataSources', []))}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Data segmentation</td><td>{req.get('sliceBy', '—')}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Priority account list</td><td>{'Default' if req.get('slide2', {}).get('accountListOption', 'default') == 'default' else 'Uploaded'}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Submitted at</td><td>{_to_ist_str(req.get('submittedAt', ''))}</td></tr>
            <tr><td style="color:#888;padding:4px 0;">Status</td><td><span style="color:#7A5700;font-weight:600;">Under Review</span></td></tr>
          </table>
        </div>
        <p style="font-size:13px;color:#555;">You'll receive another email once your report is ready. Please keep your Request ID for reference.</p>
      </div>
      <div style="padding:12px 24px;background:#F2EEE9;font-size:12px;color:#888;">
        Marketing Analytics Team · Bain &amp; Company
      </div>
    </div>"""
    _send_email(to, subject, html)



def _email_user_needs_modification(req, note, flagged_fields):
    """Email user when request needs modification."""
    user = req.get('user', {})
    to   = user.get('email', '')
    if not to:
        return
    subject = f"Action Required — {req.get('reportName', '')} [{req.get('requestId', '')}]"
    fields_list = ''.join([f"<li style='padding:3px 0;'>{f}</li>" for f in flagged_fields]) if flagged_fields else ''
    html = f"""
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#CC0000;padding:16px 24px;">
        <h2 style="color:white;margin:0;font-size:18px;">Marketing Analytics · Action Required</h2>
      </div>
      <div style="padding:24px;background:#ffffff;border:1px solid #e0d9d2;">
        <p style="font-size:15px;">Hi {user.get('name', 'there').split()[0]},</p>
        <p style="margin-top:12px;">Your report request needs a small update before we can process it.</p>
        <div style="background:#FFF8E7;border:1px solid #F5DFA0;border-radius:8px;padding:16px;margin:20px 0;">
          <p style="font-weight:600;color:#7A5700;margin-bottom:8px;">Fields to review:</p>
          <ul style="font-size:13px;color:#7A5700;">{fields_list}</ul>
          {f'<p style="margin-top:10px;font-size:13px;color:#555;"><strong>Note:</strong> {note}</p>' if note else ''}
        </div>
        <p style="font-size:13px;color:#555;">Please resubmit your request using the link below and enter your Request ID <strong>{req.get('requestId', '')}</strong> to pre-fill your previous inputs.</p>
        <div style="margin-top:20px;">
          <a href="https://ma-reports.bain.com/" 
             style="background:#CC0000;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:600;font-size:14px;">
            Update Request →
          </a>
        </div>
      </div>
      <div style="padding:12px 24px;background:#F2EEE9;font-size:12px;color:#888;">
        Marketing Analytics Team · Bain &amp; Company
      </div>
    </div>"""
    _send_email(to, subject, html)


# ────────────────────────────────────────────────────────────────────────────
# ROUTE: Discard request
# POST /admin/discard
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/discard', methods=['POST'])
def discard_request():
    data = request.json
    request_id = data.get('requestId')
    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404
    req = requests_store[request_id]
    req['status']        = 'discarded'
    req['discardedAt']   = datetime.utcnow().isoformat()
    req['discardReason'] = data.get('reason', '')
    save_store()
    print(f"[Discard] Request {request_id} discarded")
    return jsonify({ 'success': True }), 200
# POST /admin/close
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/close', methods=['POST'])
def close_request():
    data = request.json
    request_id = data.get('requestId')
    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404
    req = requests_store[request_id]
    req['status']   = 'closed'
    req['closedAt'] = datetime.utcnow().isoformat()
    save_store()
    _notify_request_closed(req)
    print(f"[Close] Request {request_id} closed")
    return jsonify({ 'success': True }), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE: Force stop a running pipeline
# POST /admin/force-stop
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/force-stop', methods=['POST'])
def force_stop_request():
    data = request.json
    request_id = data.get('requestId')
    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404
    req = requests_store[request_id]
    req['status']       = 'pending'
    req['forceStopped'] = True
    save_store()
    print(f"[Force Stop] Request {request_id} reset to pending — any in-flight SQL will be discarded on completion")
    return jsonify({ 'success': True }), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE: Rerun request — reset to pending
# POST /admin/rerun
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/rerun', methods=['POST'])
def rerun_request():
    data = request.json
    request_id = data.get('requestId')
    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404
    req = requests_store[request_id]
    req['status'] = 'pending'
    req.pop('completedAt', None)
    req.pop('downloadLinks', None)
    req.pop('csvFiles', None)
    req.pop('errors', None)
    save_store()
    print(f"[Rerun] Request {request_id} reset to pending")
    return jsonify({ 'success': True }), 200


# ────────────────────────────────────────────────────────────────────────────
# ROUTE: Mark complete manually
# POST /admin/mark-complete
# ────────────────────────────────────────────────────────────────────────────
@app.route('/admin/mark-complete', methods=['POST'])
def mark_complete():
    data = request.json
    request_id = data.get('requestId')
    if request_id not in requests_store:
        return jsonify({ 'error': 'Request not found' }), 404
    req = requests_store[request_id]
    req['status'] = 'completed'
    req['completedAt'] = datetime.utcnow().isoformat()
    if 'downloadLinks' not in req:
        req['downloadLinks'] = {}
    save_store()
    print(f"[Mark Complete] Request {request_id} marked complete manually")
    return jsonify({ 'success': True }), 200


@app.route('/templates/<filename>')
def serve_template(filename):
    template_folder = r"E:\Marketing Analytics\Report Request Automation\templates"
    return send_from_directory(template_folder, filename, as_attachment=True)


@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

@app.route('/admin.html')
def serve_admin():
    return send_from_directory('.', 'admin.html')

@app.route('/downloads/<path:filename>')
def serve_download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)


# ────────────────────────────────────────────────────────────────────────────
# HEALTH CHECK
# ────────────────────────────────────────────────────────────────────────────
@app.route('/health')
def health():
    return jsonify({ 'status': 'ok', 'requests': len(requests_store) }), 200


if __name__ == '__main__':
    print("=" * 55)
    print("  MA Report Automation API")
    print("  Running on http://0.0.0.0:5000")
    print("  Admin page: https://ma-reports.bain.com/admin.html")
    print("=" * 55)
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
