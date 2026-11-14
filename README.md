# Event Manager Agent MVP

This is the first working slice of the event operations agent. It parses a run-of-show document into a vendor/support checklist and creates a raw attendee registration template that can be formatted into the existing DFP check-in workbook.

## Current Features

- Parses session metadata, facilitators, catering, space/AV, day-of contacts, alumni, and schedule tables from `Daily Schedule-DFPStaff.docx`.
- Builds a structured JSON/CSV/XLSX vendor and support checklist.
- Creates a raw attendee registration template.
- Formats a raw registration report into the existing `DFP Staff Check-In 2026.xlsx` workbook structure.
- Records integration sources, parking link, campus-notice template link, Canva links, and pending integrations.

## Install and Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python scripts/build_mvp.py \
  --run-of-show "/Users/akshetha/Downloads/Daily Schedule-DFPStaff.docx" \
  --checkin-template "/Users/akshetha/Downloads/DFP Staff Check-In 2026.xlsx" \
  --output-dir data/generated
```

To use the browser interface instead of the CLI:

```bash
python scripts/run_ui.py
```

Then open `http://127.0.0.1:8000`. The first UI version supports run-of-show checklist generation and check-in formatting.

To format a completed raw report after filling in the template:

```bash
python scripts/build_mvp.py \
  --run-of-show "/Users/akshetha/Downloads/Daily Schedule-DFPStaff.docx" \
  --checkin-template "/Users/akshetha/Downloads/DFP Staff Check-In 2026.xlsx" \
  --raw-report data/generated/templates/Raw_Attendee_Registration_Template.xlsx \
  --output-dir data/generated
```

## File Sources

- Saltaire menu: https://saltaire.ucsd.edu/menus/index.html
- Parking/space request: https://transportation.ucsd.edu/departments/event.html#submit-a-space-request-form
- Campus notice template: https://docs.google.com/document/d/1wUZR7O90hIMt4uEXVevh9yZ7gUUJibvGaIQ7mU7k6yA/edit?usp=sharing
- myEmma invitation example: https://t.e2ma.net/webview/mahxjl/b30e0a6880ea4cdd85b2bac3fa1eeea6
- Canva agenda: https://canva.link/8rbbn6r4vpz4877
- Canva flyer: https://canva.link/r2hzoe76ezx1022
- Canva foam board: https://canva.link/rndxott2bzvbopf

See [Microsoft Graph Excel setup](docs/MICROSOFT_GRAPH_EXCEL.md) and [second-round notes](docs/SECOND_ROUND.md).
