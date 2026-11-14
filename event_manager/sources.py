from pathlib import Path
import json


DEFAULT_SOURCES = {
    "event": "Dialogue for Peace Staff Fellowship 2026",
    "saltaire_menu_url": "https://saltaire.ucsd.edu/menus/index.html",
    "order_form_url": None,
    "tres": {
        "status": "manual_export",
        "reason": "No verified public API or harness-provided TRES connector is available yet.",
        "next_step": "Check with UCSD events/IT for a read-only integration or export endpoint.",
    },
    "faculty_club": {
        "status": "manual_export",
        "reason": "No public availability page has been identified yet.",
        "next_step": "Use Sharon Nelson or the Faculty Club team to confirm current space and AV availability.",
    },
    "parking_request_url": "https://transportation.ucsd.edu/departments/event.html#submit-a-space-request-form",
    "campus_notice_template_url": "https://docs.google.com/document/d/1wUZR7O90hIMt4uEXVevh9yZ7gUUJibvGaIQ7mU7k6yA/edit?usp=sharing",
    "campus_notice_fields": [
        " issuing offices",
        " notice date",
        " audience",
        " event title",
        " event description",
        " speaker",
        " event date/time",
        " location",
        " question form",
        " registration deadline",
        " catering note",
        " parking note",
        " learn-more links",
        " signatures and titles",
    ],
    "canva_templates": {
        "agenda": "https://canva.link/8rbbn6r4vpz4877",
        "flyer": "https://canva.link/r2hzoe76ezx1022",
        "foam_board": "https://canva.link/rndxott2bzvbopf",
    },
    "myemma": {
        "status": "pending_access",
        "login": "vcotevents@ucsd.edu",
        "invitation_example_url": "https://t.e2ma.net/webview/mahxjl/b30e0a6880ea4cdd85b2bac3fa1eeea6",
        "next_step": "Obtain myEmma API/account access in the second development round.",
    },
}


def write_sources(path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(DEFAULT_SOURCES, indent=2), encoding="utf-8")
    return path
