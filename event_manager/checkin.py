from __future__ import annotations

from copy import copy
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


RAW_COLUMNS = [
    "Registration Timestamp",
    "Registration Status",
    "First Name",
    "Last Name",
    "Email",
    "Phone",
    "Title",
    "Department",
    "Session(s)",
    "Proposed Project Name/Brief Description",
    "Consultation Time",
    "Assigned Consultant",
    "Presentation Order",
    "Notes",
]

SESSION_MAP = {
    "9/14": "Sep 14",
    "september 14": "Sep 14",
    "sep 14": "Sep 14",
    "9/15": "Sep 15",
    "september 15": "Sep 15",
    "sep 15": "Sep 15",
    "9/16": "Sep 16",
    "september 16": "Sep 16",
    "sep 16": "Sep 16",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def create_raw_registration_template(path: str | Path) -> None:
    """Create a raw-registration template with the columns the formatter expects."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Raw Registration"
    instructions = workbook.create_sheet("Instructions")

    sheet.append(RAW_COLUMNS)
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    example = [
        "2026-08-01 10:00", "Registered", "FirstName", "LastName",
        "email@ucsd.edu", "(000) 000-0000", "Title", "Department",
        "Sep 14; Sep 15; Sep 16", "Example project description", "2:30 PM", "Karen", "", "Example notes",
    ]
    sheet.append([""] * len(RAW_COLUMNS))
    sheet.append(example)

    widths = [22, 20, 18, 18, 34, 20, 30, 40, 26, 52, 18, 24, 22, 38]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(RAW_COLUMNS))}3"

    instructions["A1"] = "Raw attendee registration template"
    instructions["A2"] = "Keep the Raw Registration headers unchanged."
    instructions["A3"] = "One attendee per row. In Session(s), use Sep 14, Sep 15, and/or Sep 16 separated by semicolons."
    instructions["A4"] = "The Sep 15 consultation and project columns are optional for attendees not attending that session."
    instructions["A5"] = "Delete the example row before submitting a real report."
    instructions["A1"].font = Font(size=14, bold=True)
    instructions.column_dimensions["A"].width = 110

    workbook.save(path)


def _norm(value) -> str:
    return str(value or "").strip()


def _sessions(value: str) -> set[str]:
    found = set()
    lowered = value.lower()
    for token, session in SESSION_MAP.items():
        if token in lowered:
            found.add(session)
    if not found:
        raise ValueError(f"Could not identify sessions in '{value}'. Use Sep 14, Sep 15, or Sep 16.")
    return found


def _copy_style(source_cell, target_cell) -> None:
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format


def _clear_data_area(sheet: "object", last_data_row: int, first_data_row: int = 4, max_col: int = 4) -> None:
    for row in sheet.iter_rows(min_row=first_data_row, max_row=last_data_row, max_col=max_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.hyperlink = None


def format_checkin(raw_path: str | Path, template_path: str | Path, output_path: str | Path) -> dict:
    """Format a raw registration report into the DFP check-in workbook structure."""
    raw_wb = load_workbook(raw_path, data_only=True)
    raw = raw_wb.active
    headers = [_norm(cell.value) for cell in raw[1]]
    if not headers or headers[0] != RAW_COLUMNS[0]:
        raise ValueError("Raw report must start with the columns from create_raw_registration_template().")

    records = []
    for values in raw.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, values))
        first = _norm(row.get("First Name"))
        last = _norm(row.get("Last Name"))
        email = _norm(row.get("Email"))
        sessions_field = _norm(row.get("Session(s)"))
        if not any([first, last, email, sessions_field]):
            continue
        if not first or not last or not email:
            raise ValueError(f"Missing name or email in raw row: {row}")
        records.append({
            **row,
            "_sessions": _sessions(sessions_field),
        })

    if not records:
        raise ValueError("No attendee rows found in the raw report.")

    output_wb = load_workbook(template_path, data_only=False)
    directory = output_wb["Directory"]
    _clear_data_area(directory, directory.max_row, 2, 6)

    records.sort(key=lambda item: (_norm(item["Last Name"]).lower(), _norm(item["First Name"]).lower()))
    for row_num, record in enumerate(records, start=2):
        directory.cell(row_num, 1, _norm(record["Last Name"]))
        directory.cell(row_num, 2, _norm(record["First Name"]))
        directory.cell(row_num, 3, _norm(record["Email"]))
        directory.cell(row_num, 4, _norm(record["Phone"]))
        directory.cell(row_num, 5, _norm(record["Title"]))
        directory.cell(row_num, 6, _norm(record["Department"]))

    session_sheets = ["Sep 14", "Sep 15", "Sep 16"]
    for session_name in session_sheets:
        sheet = output_wb[session_name]
        data_rows = []
        for row_num in range(4, sheet.max_row + 1):
            if any(_norm(sheet.cell(row_num, col).value) for col in range(1, 4)):
                data_rows.append(row_num)
        last_data_row = max(data_rows, default=3)

        if session_name == "Sep 15":
            _clear_data_area(sheet, last_data_row, 4, 7)
        else:
            _clear_data_area(sheet, last_data_row, 4, 4)

        selected = [record for record in records if session_name in record["_sessions"]]
        for index, record in enumerate(selected):
            row_num = 4 + index
            sheet.cell(row_num, 1, _norm(record["Last Name"]))
            sheet.cell(row_num, 2, _norm(record["First Name"]))
            sheet.cell(row_num, 3, _norm(record["Email"]))
            if session_name == "Sep 15":
                sheet.cell(row_num, 5, _norm(record.get("Proposed Project Name/Brief Description")))
                sheet.cell(row_num, 6, _norm(record.get("Consultation Time")))
                sheet.cell(row_num, 7, _norm(record.get("Assigned Consultant")))

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_wb.save(output_path)
    return {"output": str(output_path), "attendees": len(records)}


def create_blank_checkin_template(source_path: str | Path, output_path: str | Path) -> Path:
    """Create a clean template copy of the check-in workbook with attendee rows removed."""
    workbook = load_workbook(source_path, data_only=False)
    for sheet in workbook.worksheets:
        if sheet.title == "Directory":
            _clear_data_area(sheet, sheet.max_row, 2, 6)
        elif sheet.title == "Sep 15":
            _clear_data_area(sheet, sheet.max_row, 4, 7)
        else:
            _clear_data_area(sheet, sheet.max_row, 4, 4)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path
