from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import csv
import json
import re

from .docx_text import docx_body_blocks


SESSION_RE = re.compile(r"^Session\s+(\d+):\s*(.+)$", re.I)
HEADER_LABELS = {
    "Location",
    "Topic",
    "Facilitators",
    "Catering Information",
    "Space and AV Contacts",
    "Team Day-of Contact Information",
    "Alumni in Attendance",
}


def parse_sessions(path: str | Path) -> list[dict]:
    """Parse the DFP run-of-show document into structured session data."""
    blocks = docx_body_blocks(path)
    sessions: list[dict] = []
    current = None

    for block in blocks:
        if block["type"] == "paragraph":
            match = SESSION_RE.match(block["text"])
            if match:
                current = {
                    "session_number": int(match.group(1)),
                    "session_title": match.group(2).strip(),
                    "metadata": {},
                    "schedule": [],
                }
                sessions.append(current)
                continue

            if current and ":" in block["text"]:
                label, value = block["text"].split(":", 1)
                label = label.strip()
                if label in HEADER_LABELS:
                    current["metadata"][label] = value.strip()

        elif block["type"] == "table" and current:
            rows = block["rows"]
            if not rows:
                continue
            headers = [cell.strip() for cell in rows[0]]
            if not headers or headers[0].lower() != "time:":
                continue

            for row in rows[1:]:
                row += [""] * (len(headers) - len(row))
                item = dict(zip([h.rstrip(":") for h in headers], row))
                if item.get("Time"):
                    current["schedule"].append(item)

    for session in sessions:
        session["date"] = _session_date(session["session_title"])
        session["metadata"]["Session"] = session["session_title"]

    return sessions


def _session_date(session_title: str) -> str | None:
    match = re.search(r"(Monday|Tuesday|Wednesday|Thursday|Friday),\s+([A-Z][a-z]+)\s+(\d{1,2}),\s+(\d{4})", session_title)
    if not match:
        return None
    _, month_name, day, year = match.groups()
    month = datetime.strptime(month_name, "%B").month
    return date(int(year), month, int(day)).isoformat()


def build_checklist(sessions: list[dict], parking_url: str | None = None) -> list[dict]:
    """Build a vendor/support checklist from the parsed session headers."""
    checklist = []

    for session in sessions:
        meta = session["metadata"]
        event_date = date.fromisoformat(session["date"]) if session["date"] else None
        due = event_date - timedelta(days=7) if event_date else None
        due_label = due.isoformat() if due else "Not set"

        def add(task: str, vendor: str, contact: str, notes: str = "") -> None:
            checklist.append({
                "event": "Dialogue for Peace Staff Fellowship 2026",
                "session": f"Session {session['session_number']}",
                "date": session["date"],
                "task": task,
                "vendor_support": vendor,
                "contact": contact,
                "due": due_label,
                "status": "Not started",
                "notes": notes,
            })

        catering = meta.get("Catering Information", "")
        add("Confirm catering order, delivery time, and setup", "Saltair / HDH Catering", catering,
            "Vendor must start service at the first meal time shown in the run-of-show.")
        add("Confirm space reservation, room setup, and AV support", "Faculty Club / Space and AV", meta.get("Space and AV Contacts", ""),
            "Include Conference Rooms 1-3 and Cecil's Lounge requirements.")

        facilitators = meta.get("Facilitators", "")
        if facilitators:
            add("Confirm facilitator arrival, materials, and session flow", "Facilitators", facilitators)

        alumni = meta.get("Alumni in Attendance", "")
        if alumni and alumni.upper() != "TBD":
            add("Confirm alumni arrival time and speaking slot", "DFP Alumni", alumni)

        day_of = meta.get("Team Day-of Contact Information", "")
        add("Confirm day-of staff coverage and handoff plan", "Event team", day_of)
        add("Complete setup, breakdown, and space lockdown", "Event team", day_of,
            "Verify the run-of-show setup and conclusion rows for exact times.")

    checklist.append({
        "event": "Dialogue for Peace Staff Fellowship 2026",
        "session": "All sessions",
        "date": "",
        "task": "Reserve parking or submit a space request for event attendees",
        "vendor_support": "UCSD Transportation",
        "contact": "https://transportation.ucsd.edu/departments/event.html#submit-a-space-request-form",
        "due": "Two weeks before event",
        "status": "Not started",
        "notes": parking_url or "Use the UCSD event transportation page.",
    })

    return checklist


def write_checklist_outputs(sessions: list[dict], checklist: list[dict], output_dir: str | Path) -> dict:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = output_dir / "run_of_show.json"
    checklist_path = output_dir / "dfp_staff_fellowship_2026_checklist.json"
    csv_path = output_dir / "dfp_staff_fellowship_2026_checklist.csv"

    json_path.write_text(json.dumps(sessions, indent=2), encoding="utf-8")
    checklist_path.write_text(json.dumps(checklist, indent=2), encoding="utf-8")

    fieldnames = ["event", "session", "date", "task", "vendor_support", "contact", "due", "status", "notes"]
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(checklist)

    return {"run_of_show": json_path, "checklist_json": checklist_path, "checklist_csv": csv_path}
