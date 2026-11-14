from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json

from flask import Flask, flash, redirect, render_template_string, request, send_from_directory, url_for
from werkzeug.utils import secure_filename

from .checkin import format_checkin
from .docx_text import docx_body_blocks
from .run_of_show import build_checklist, parse_sessions, write_checklist_outputs


PARKING_URL = "https://transportation.ucsd.edu/departments/event.html#submit-a-space-request-form"

PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Event Manager</title>
  <style>
    :root { color-scheme: light; --ink:#172b4d; --line:#d0d7e2; --accent:#1769aa; --soft:#f3f6fa; }
    * { box-sizing: border-box; }
    body { margin:0; font-family: system-ui, -apple-system, sans-serif; color:var(--ink); background:#f7f8fb; }
    main { max-width: 1240px; margin:0 auto; padding:32px 20px 56px; }
    header { display:flex; justify-content:space-between; align-items:center; gap:16px; margin-bottom:28px; }
    h1 { font-size:24px; margin:0; letter-spacing:0; }
    .tag { font-size:12px; color:#54606f; background:#e7edf5; padding:4px 8px; border-radius:4px; }
    .workspace { display:grid; grid-template-columns: minmax(0, 1.25fr) minmax(280px, .75fr); gap:28px; align-items:start; }
    .chat { background:#fff; border:1px solid var(--line); border-radius:8px; padding:20px; }
    .messages { display:grid; gap:18px; }
    .msg { display:grid; grid-template-columns:36px minmax(0,1fr); gap:12px; align-items:start; }
    .msg.user { grid-template-columns:minmax(0,1fr) 36px; }
    .avatar { width:32px; height:32px; border-radius:50%; display:grid; place-items:center; font-size:12px; font-weight:700; color:#fff; background:var(--accent); }
    .msg.user .avatar { background:#22795a; }
    .bubble { background:var(--soft); border-radius:8px; padding:14px; font-size:14px; line-height:1.45; }
    .msg.user .bubble { background:#e8f5ee; }
    .bubble .field { margin-top:12px; }
    .roadmap { background:#fff; border:1px solid var(--line); border-radius:6px; padding:20px; }
    .roadmap h2 { margin-bottom:14px; }
    .roadmap .list { margin:0; padding:0; list-style:none; }
    .roadmap li { display:grid; grid-template-columns:14px minmax(0,1fr); gap:9px; padding:8px 0; border-bottom:1px solid #eef1f5; font-size:13px; line-height:1.45; }
    .roadmap li:last-child { border-bottom:0; }
    .dot { width:8px; height:8px; border-radius:50%; margin-top:5px; }
    .done { background:#22795a; }
    .active { background:#1769aa; }
    .planned { background:#a7b0bd; }
    .item-title { font-weight:600; color:#253858; }
    .item-detail { color:#64748b; margin-top:2px; }
    .roadmap-note { margin-top:14px; padding-top:12px; border-top:1px solid #eef1f5; font-size:12px; color:#64748b; }
    h2 { font-size:14px; margin:0 0 12px; color:#334155; }
    .grid { display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:16px; }
    .field { display:flex; flex-direction:column; gap:6px; }
    label { font-size:13px; font-weight:600; color:#42526e; }
    input[type=text], textarea, input[type=file] { width:100%; padding:8px 10px; border:1px solid #b6c2d1; border-radius:4px; font-size:14px; background:#fff; }
    textarea { min-height:84px; resize:vertical; }
    .hint { font-size:12px; color:#64748b; }
    button { background:var(--accent); color:#fff; border:0; border-radius:4px; padding:10px 16px; font-size:14px; font-weight:600; cursor:pointer; }
    button:hover { background:#125a92; }
    .flash { margin-bottom:18px; padding:10px 12px; background:#fff2f2; border-left:3px solid #d33; color:#8a2323; font-size:13px; }
    .note { margin-top:18px; padding:10px 12px; background:var(--soft); border-left:3px solid var(--accent); font-size:13px; }
    @media (max-width:900px) { .workspace { grid-template-columns:1fr; } }
    @media (max-width:700px) { .grid { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <main>
    {% for message in get_flashed_messages() %}
      <div class="flash">{{ message }}</div>
    {% endfor %}
    <header>
      <h1>Event Manager</h1>
      <span class="tag">MVP</span>
    </header>
    <div class="workspace">
      <form method="post" action="/generate" enctype="multipart/form-data">
        <div class="chat">
          <div class="messages">
            <div class="msg">
              <div class="avatar">EM</div>
              <div class="bubble">
                Hi! What event are we planning?
                <div class="grid">
                  <div class="field">
                    <label for="event_name">Event name</label>
                    <input type="text" id="event_name" name="event_name" value="Dialogue for Peace Staff Fellowship 2026">
                  </div>
                  <div class="field">
                    <label for="event_dates">Event dates</label>
                    <input type="text" id="event_dates" name="event_dates" value="Sep 14-16, 2026">
                  </div>
                </div>
              </div>
            </div>

            <div class="msg">
              <div class="avatar">EM</div>
              <div class="bubble">
                Great. Upload the event document you want me to read. A run-of-show works well.
                <div class="field">
                  <label for="info_doc">Event document</label>
                  <input type="file" id="info_doc" name="info_doc" accept=".docx">
                </div>
              </div>
            </div>

            <div class="msg">
              <div class="avatar">EM</div>
              <div class="bubble">
                Optional: upload a raw attendee registration report if you want me to format the check-in list.
                <div class="field">
                  <label for="raw_report">Raw attendee registration report</label>
                  <input type="file" id="raw_report" name="raw_report" accept=".xlsx">
                </div>
              </div>
            </div>

            <div class="msg user">
              <div class="bubble">
                Generate the event checklist and outputs.
                <div style="margin-top:12px;"><button type="submit">Generate</button></div>
              </div>
              <div class="avatar">You</div>
            </div>
          </div>
        </div>
      </form>

      <aside class="roadmap">
        <h2>What it does now</h2>
        <ul class="list">
          <li><span class="dot done"></span><span><span class="item-title">Run-of-show checklist</span><div class="item-detail">Reads facilitators, catering, space, AV, alumni, and day-of contacts.</div></span></li>
          <li><span class="dot done"></span><span><span class="item-title">Vendor and support tracker</span><div class="item-detail">Creates a session-by-session checklist with contacts and due dates.</div></span></li>
          <li><span class="dot done"></span><span><span class="item-title">Check-in formatting</span><div class="item-detail">Turns a raw registration report into the existing check-in workbook.</div></span></li>
          <li><span class="dot done"></span><span><span class="item-title">Saltaire menu reference</span><div class="item-detail">Uses the public menu page as the catering source.</div></span></li>
          <li><span class="dot done"></span><span><span class="item-title">Parking request link</span><div class="item-detail">Stores the UCSD event transportation request page.</div></span></li>
        </ul>

        <h2 style="margin-top:22px;">Second development stage</h2>
        <ul class="list">
          <li><span class="dot planned"></span><span><span class="item-title">TRES availability</span><div class="item-detail">Read event space availability once an approved integration or export is available.</div></span></li>
          <li><span class="dot planned"></span><span><span class="item-title">Faculty Club spaces</span><div class="item-detail">Confirm rooms, layouts, and AV availability through a live source.</div></span></li>
          <li><span class="dot planned"></span><span><span class="item-title">myEmma invitations</span><div class="item-detail">Create and review invitation drafts using vcotevents@ucsd.edu.</div></span></li>
          <li><span class="dot planned"></span><span><span class="item-title">Canva templates</span><div class="item-detail">Prefill the agenda, flyer with QR code, and foam-board sign templates.</div></span></li>
          <li><span class="dot planned"></span><span><span class="item-title">Campus notices</span><div class="item-detail">Copy and update the blank Google Doc template for each event.</div></span></li>
          <li><span class="dot planned"></span><span><span class="item-title">Order form intake</span><div class="item-detail">Read or ingest catering and attendee order submissions.</div></span></li>
        </ul>
        <div class="roadmap-note">Blocked items need UCSD access, API credentials, or an approved integration before they can be activated.</div>
      </aside>
    </div>
    <div class="note">The app creates local copies only. The blank check-in template is built in, so you do not need to upload it.</div>
  </main>
</body>
</html>"""

RESULT_PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Event Manager Results</title>
  <style>
    body { margin:0; font-family: system-ui, sans-serif; background:#f7f8fb; color:#172b4d; }
    main { max-width:800px; margin:0 auto; padding:32px 20px; }
    h1 { font-size:24px; margin-bottom:18px; }
    ul { list-style:none; padding:0; margin:0 0 24px; }
    li { background:#fff; border:1px solid #d0d7e2; border-radius:6px; padding:14px 16px; margin-bottom:10px; }
    a { color:#1769aa; font-weight:600; text-decoration:none; }
    a:hover { text-decoration:underline; }
    .small { font-size:13px; color:#54606f; margin-top:4px; }
    a.back { display:inline-block; margin-top:8px; color:#334155; font-weight:500; }
  </style>
</head>
<body>
  <main>
    <h1>Done</h1>
    {% for item in outputs %}
      <li><a href="{{ item.url }}">{{ item.label }}</a><div class="small">{{ item.detail }}</div></li>
    {% endfor %}
    <a class="back" href="/">Start another event</a>
  </main>
</body>
</html>"""


app = Flask(__name__)
app.secret_key = "event-manager-mvp-local"
UPLOAD_ROOT = Path(__file__).resolve().parents[1] / "data" / "generated" / "web"
BLANK_CHECKIN_TEMPLATE = Path(__file__).resolve().parent / "assets" / "DFP_Staff_Check-In_Blank.xlsx"


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.post("/generate")
def generate():
    event_name = request.form.get("event_name", "Event").strip() or "Event"
    event_dates = request.form.get("event_dates", "").strip()
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = UPLOAD_ROOT / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    outputs = []

    info_doc = request.files.get("info_doc")
    if info_doc and info_doc.filename:
        if not info_doc.filename.lower().endswith(".docx"):
            flash("The event document must be a .docx file.")
            return redirect(url_for("index"))

        run_path = run_dir / "inputs" / secure_filename(info_doc.filename)
        run_path.parent.mkdir(parents=True, exist_ok=True)
        info_doc.save(run_path)

        try:
            blocks = docx_body_blocks(run_path)
            info_path = run_dir / "checklists" / "document_info.json"
            info_path.parent.mkdir(parents=True, exist_ok=True)
            info_path.write_text(json.dumps({
                "event_name": event_name,
                "event_dates": event_dates,
                "blocks": blocks,
            }, indent=2), encoding="utf-8")
            outputs.append({"url": f"/download/{run_id}/checklists/document_info.json",
                            "label": "Event document structure (JSON)",
                            "detail": "Extracted text and tables from the uploaded document."})
        except Exception:
            flash("I could not read that event document. Please upload a valid .docx file.")
            return redirect(url_for("index"))

        try:
            sessions = parse_sessions(run_path)
            if sessions:
                checklist = build_checklist(sessions, PARKING_URL)
                write_checklist_outputs(sessions, checklist, run_dir / "checklists")
                outputs.append({"url": f"/download/{run_id}/checklists/dfp_staff_fellowship_2026_checklist.csv",
                                "label": "Vendor/support checklist (CSV)",
                                "detail": f"{len(checklist)} items from {len(sessions)} sessions."})
                outputs.append({"url": f"/download/{run_id}/checklists/run_of_show.json",
                                "label": "Run-of-show structure (JSON)",
                                "detail": "Structured data for later automation."})
        except Exception:
            flash("I read the document, but could not build the checklist from it.")
            return redirect(url_for("index"))

    raw_report = request.files.get("raw_report")
    if raw_report and raw_report.filename:
        if not raw_report.filename.lower().endswith(".xlsx"):
            flash("The raw attendee registration report must be an .xlsx file.")
            return redirect(url_for("index"))

        raw_path = run_dir / "inputs" / secure_filename(raw_report.filename)
        raw_report.save(raw_path)
        try:
            result = format_checkin(raw_path, BLANK_CHECKIN_TEMPLATE, run_dir / "checkin" / "Formatted_DFP_Staff_Check-In.xlsx")
            outputs.append({"url": f"/download/{run_id}/checkin/Formatted_DFP_Staff_Check-In.xlsx",
                            "label": "Formatted check-in workbook",
                            "detail": f"{result['attendees']} attendees formatted."})
        except Exception:
            flash("I could not read that registration report. Please upload the raw report in the expected .xlsx format.")
            return redirect(url_for("index"))

    if not outputs:
        flash("Please upload an event information document or a raw attendee registration report.")
        return redirect(url_for("index"))

    return render_template_string(RESULT_PAGE, outputs=outputs)


@app.get("/download/<run_id>/<path:subpath>")
def download(run_id: str, subpath: str):
    safe_run_id = "".join(ch for ch in run_id if ch.isalnum() or ch in "-_")
    return send_from_directory(UPLOAD_ROOT / safe_run_id, subpath, as_attachment=False)
