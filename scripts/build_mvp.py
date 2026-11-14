#!/usr/bin/env python3
from __future__ import annotations

from argparse import ArgumentParser
from pathlib import Path
import sys
import json

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from event_manager.run_of_show import parse_sessions, build_checklist, write_checklist_outputs
from event_manager.checkin import create_raw_registration_template, format_checkin
from event_manager.sources import write_sources


PARKING_URL = "https://transportation.ucsd.edu/departments/event.html#submit-a-space-request-form"


def write_checklist_xlsx(checklist: list[dict], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Checklist"
    headers = ["Session", "Date", "Task", "Vendor/Support", "Contact", "Due", "Status", "Notes"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for item in checklist:
        sheet.append([
            item["session"], item["date"], item["task"], item["vendor_support"],
            item["contact"], item["due"], item["status"], item["notes"],
        ])

    widths = [16, 14, 55, 32, 48, 18, 18, 72]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:H{sheet.max_row}"
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    workbook.save(path)


def main() -> None:
    parser = ArgumentParser(description="Build DFP event-manager MVP artifacts.")
    parser.add_argument("--run-of-show", required=True, help="Path to Daily Schedule-DFPStaff.docx")
    parser.add_argument("--checkin-template", required=True, help="Path to DFP Staff Check-In 2026.xlsx")
    parser.add_argument("--raw-report", help="Optional raw attendee registration XLSX")
    parser.add_argument("--output-dir", default="data/generated", help="Output directory")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    sessions = parse_sessions(args.run_of_show)
    checklist = build_checklist(sessions, PARKING_URL)
    outputs = write_checklist_outputs(sessions, checklist, output_dir / "checklists")
    sources_path = write_sources(output_dir / "sources.json")
    raw_template_path = output_dir / "templates" / "Raw_Attendee_Registration_Template.xlsx"
    raw_template_path.parent.mkdir(parents=True, exist_ok=True)
    create_raw_registration_template(raw_template_path)
    write_checklist_xlsx(checklist, output_dir / "checklists" / "DFP_Staff_Fellowship_Checklist.xlsx")

    formatted = None
    if args.raw_report:
        formatted = format_checkin(
            args.raw_report,
            args.checkin_template,
            output_dir / "checkin" / "Formatted_DFP_Staff_Check-In.xlsx",
        )

    print(json.dumps({
        "sessions_parsed": len(sessions),
        "checklist_items": len(checklist),
        "outputs": {str(key): str(value) for key, value in outputs.items()},
        "sources": str(sources_path),
        "raw_template": str(raw_template_path),
        "formatted_checkin": formatted,
    }, indent=2))


if __name__ == "__main__":
    main()
