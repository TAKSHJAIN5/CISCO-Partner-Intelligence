"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
build_final_excel.py

Purpose:
Presentation layer. Transforms the collected partner details
and detected changes into the final, management-ready Excel
workbook that is delivered to marketing teams and leadership.

The workbook is the only deliverable produced by this script.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    AUTO_FILTER,
    AUTO_FIT_COLUMNS,
    CHANGE_LOG_COLUMNS,
    CHANGE_LOG_FILE,
    CHANGE_TYPE_NEW_PARTNER,
    CHANGE_TYPE_REMOVED_PARTNER,
    CHANGES_SHEET,
    COLOR_HEADER,
    COLOR_TITLE,
    FREEZE_HEADER_ROW,
    FREEZE_NAME_COLUMN,
    LOGS_SHEET,
    OUTPUT_EXCEL_FILE,
    PARTNER_DETAILS_FILE,
    PARTNERS_SHEET,
    PARTNERS_SHEET_COLUMNS,
    PREFERRED_CLOUD_AI_PORTFOLIO,
    PREFERRED_KEYWORD,
    PREFERRED_NETWORKING_PORTFOLIO,
    PREFERRED_SECURITY_PORTFOLIO,
    PROJECT_NAME,
    QUALIFICATION_SEPARATOR,
    REPORT_SUBTITLE,
    REPORT_TITLE,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SUMMARY_BORDER,
    SUMMARY_SHEET,
    TIMESTAMP_FORMAT,
    TRACKED_CHANGE_FIELDS,
    VERSION,
)
from logger import log

# Snapshot column names read from partner_details.csv.
PARTNER_NAME_COLUMN = "partner_name"
PORTFOLIOS_COLUMN = "portfolios"
COMPETENCIES_COLUMN = "competencies"
LOCATION_COLUMN = "location"

# Change log column that stores the type of each change.
CHANGE_TYPE_COLUMN = "change_type"

# Columns required from each input dataset.
REQUIRED_PARTNER_COLUMNS = [
    PARTNER_NAME_COLUMN,
    PORTFOLIOS_COLUMN,
    COMPETENCIES_COLUMN,
    LOCATION_COLUMN,
]

# Presentation labels.
YES_LABEL = "Yes"
NO_LABEL = "No"
LOGS_SUBTITLE = "Execution Metadata"
HEADER_FONT_COLOR = "FFFFFF"

# Column auto-fit bounds (in Excel character units).
MIN_COLUMN_WIDTH = 12
MAX_COLUMN_WIDTH = 60
COLUMN_WIDTH_PADDING = 2

# Reusable cell styles.
_HEADER_FILL = PatternFill(
    start_color=COLOR_HEADER, end_color=COLOR_HEADER, fill_type="solid"
)
_HEADER_FONT = Font(bold=True, color=HEADER_FONT_COLOR)
_TITLE_FILL = PatternFill(
    start_color=COLOR_TITLE, end_color=COLOR_TITLE, fill_type="solid"
)
_TITLE_FONT = Font(bold=True, size=14)
_SUBTITLE_FONT = Font(italic=True, size=11)
_LABEL_FONT = Font(bold=True)
_CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_DATA_ALIGNMENT = Alignment(horizontal="left", vertical="center")


# =========================================================
# Value Helpers
# =========================================================

def _format_duration(seconds: float) -> str:
    """Render a duration in seconds as a compact ``h m s`` string."""
    total_seconds = int(round(seconds))
    hours, remainder = divmod(total_seconds, SECONDS_PER_HOUR)
    minutes, secs = divmod(remainder, SECONDS_PER_MINUTE)
    if hours:
        return f"{hours}h {minutes}m {secs}s"
    if minutes:
        return f"{minutes}m {secs}s"
    return f"{secs}s"


def _clean(value: Any) -> str:
    """Return a trimmed string, treating ``None`` and NaN as empty."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _cell_value(value: Any) -> Any:
    """Return a spreadsheet-safe value, replacing NaN with empty text."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return value


def _yes_no(flag: bool) -> str:
    """Render a boolean flag as a ``Yes`` / ``No`` label."""
    return YES_LABEL if flag else NO_LABEL


def _split_portfolios(portfolios_text: str) -> List[str]:
    """Split a serialized portfolio string into individual portfolios."""
    separator = QUALIFICATION_SEPARATOR.strip()
    return [
        portfolio.strip()
        for portfolio in portfolios_text.split(separator)
        if portfolio.strip()
    ]


def _classify_portfolios(
    portfolio_list: List[str],
) -> Tuple[List[str], List[str]]:
    """Partition portfolios into preferred and non-preferred lists."""
    preferred = [
        portfolio
        for portfolio in portfolio_list
        if PREFERRED_KEYWORD in portfolio
    ]
    non_preferred = [
        portfolio
        for portfolio in portfolio_list
        if PREFERRED_KEYWORD not in portfolio
    ]
    return preferred, non_preferred


# =========================================================
# Filesystem & Validation
# =========================================================

def _validate_file_exists(file_path: str, label: str) -> None:
    """Raise if a required input file is missing."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(
            f"The {label} file was not found: {path}"
        )


def _validate_columns(
    frame: pd.DataFrame,
    required_columns: List[str],
    label: str,
) -> None:
    """Raise if an input dataset is missing any required column."""
    missing = [
        column
        for column in required_columns
        if column not in frame.columns
    ]
    if missing:
        raise ValueError(
            f"The {label} file is missing required columns: "
            f"{', '.join(missing)}"
        )


# =========================================================
# Loading
# =========================================================

def load_dataset(
    file_path: str,
    label: str,
    required_columns: List[str],
) -> pd.DataFrame:
    """Load, validate and return one input dataset."""
    _validate_file_exists(file_path, label)
    try:
        frame = pd.read_csv(file_path)
    except Exception as exc:  # noqa: BLE001 - re-raised with context
        raise ValueError(
            f"The {label} file could not be read: {file_path} ({exc})"
        ) from exc

    _validate_columns(frame, required_columns, label)
    log(f"Loaded {label}: {len(frame)} rows from {file_path}")
    return frame


# =========================================================
# Business Transformations
# =========================================================

def _build_partner_row(row: pd.Series) -> Dict[str, Any]:
    """Transform one raw partner record into the Partners sheet schema."""
    portfolios = _split_portfolios(_clean(row.get(PORTFOLIOS_COLUMN)))
    preferred, non_preferred = _classify_portfolios(portfolios)

    preferred_networking = PREFERRED_NETWORKING_PORTFOLIO in preferred
    preferred_security = PREFERRED_SECURITY_PORTFOLIO in preferred
    preferred_cloud_ai = PREFERRED_CLOUD_AI_PORTFOLIO in preferred

    sec_net_spez = preferred_networking and preferred_security
    sec_ai_infra_spez = preferred_security and preferred_cloud_ai

    return {
        "Name": _clean(row.get(PARTNER_NAME_COLUMN)),
        "Preferred Portfolios": ", ".join(preferred),
        "Non Preferred Portfolios": ", ".join(non_preferred),
        "Preferred Networking": _yes_no(preferred_networking),
        "Preferred Security": _yes_no(preferred_security),
        "Preferred Cloud + AI": _yes_no(preferred_cloud_ai),
        "Sec Net Spez": _yes_no(sec_net_spez),
        "Sec AI Infra Spez": _yes_no(sec_ai_infra_spez),
        "Competencies": _clean(row.get(COMPETENCIES_COLUMN)),
        "Location": _clean(row.get(LOCATION_COLUMN)),
    }


def build_partner_dataframe(details_df: pd.DataFrame) -> pd.DataFrame:
    """Build the business-friendly Partners dataset."""
    rows = [_build_partner_row(row) for _, row in details_df.iterrows()]
    return pd.DataFrame(rows, columns=PARTNERS_SHEET_COLUMNS)


def summarize_locations(details_df: pd.DataFrame) -> Tuple[int, int]:
    """Count distinct countries and cities from partner locations.

    Returns:
        A ``(countries, cities)`` tuple.
    """
    cities: set = set()
    countries: set = set()

    if LOCATION_COLUMN not in details_df.columns:
        return 0, 0

    for location in details_df[LOCATION_COLUMN]:
        text = _clean(location)
        if not text:
            continue
        parts = [part.strip() for part in text.split(",")]
        if len(parts) >= 2:
            cities.add(parts[0])
            countries.add(parts[-1])

    return len(countries), len(cities)


def summarize_changes(changes_df: pd.DataFrame) -> Dict[str, int]:
    """Derive change statistics from the change log (its source of truth)."""
    if changes_df.empty or CHANGE_TYPE_COLUMN not in changes_df.columns:
        counts: Dict[str, int] = {}
    else:
        counts = changes_df[CHANGE_TYPE_COLUMN].value_counts().to_dict()

    return {
        "new": counts.get(CHANGE_TYPE_NEW_PARTNER, 0),
        "removed": counts.get(CHANGE_TYPE_REMOVED_PARTNER, 0),
        "portfolio": counts.get(TRACKED_CHANGE_FIELDS["portfolios"], 0),
        "competency": counts.get(TRACKED_CHANGE_FIELDS["competencies"], 0),
        "location": counts.get(TRACKED_CHANGE_FIELDS["location"], 0),
        "total": len(changes_df),
    }


def build_summary_rows(
    total_partners: int,
    countries: int,
    cities: int,
    change_stats: Dict[str, int],
    generated_time: str,
) -> List[Tuple[str, Any]]:
    """Assemble the Executive Summary key/value rows."""
    return [
        ("Generated Time", generated_time),
        ("Project Name", PROJECT_NAME),
        ("Version", VERSION),
        ("Total Partners", total_partners),
        ("Countries", countries),
        ("Cities", cities),
        ("New Partners", change_stats["new"]),
        ("Removed Partners", change_stats["removed"]),
        ("Portfolio Changes", change_stats["portfolio"]),
        ("Competency Changes", change_stats["competency"]),
        ("Location Changes", change_stats["location"]),
        ("Total Change Records", change_stats["total"]),
    ]


def build_logs_rows(
    total_partners: int,
    total_changes: int,
    generated_time: str,
) -> List[Tuple[str, Any]]:
    """Assemble the Logs key/value rows."""
    return [
        ("Generation Time", generated_time),
        ("Project Version", VERSION),
        ("Source File (Partners)", PARTNER_DETAILS_FILE),
        ("Source File (Changes)", CHANGE_LOG_FILE),
        ("Output File", OUTPUT_EXCEL_FILE),
        ("Total Partners", total_partners),
        ("Total Changes", total_changes),
    ]


# =========================================================
# Workbook Builder
# =========================================================

def _style_header_row(ws: Worksheet, num_columns: int) -> None:
    """Apply the standard header styling to the first row."""
    for column_index in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=column_index)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER_ALIGNMENT


def _autofit_columns(ws: Worksheet) -> None:
    """Size every column to fit its widest value within sensible bounds."""
    for column_cells in ws.columns:
        longest = max(
            (
                len(str(cell.value))
                for cell in column_cells
                if cell.value not in (None, "")
            ),
            default=0,
        )
        letter = get_column_letter(column_cells[0].column)
        width = min(
            max(longest + COLUMN_WIDTH_PADDING, MIN_COLUMN_WIDTH),
            MAX_COLUMN_WIDTH,
        )
        ws.column_dimensions[letter].width = width


def _freeze_panes(ws: Worksheet, freeze_first_column: bool) -> None:
    """Freeze the header row and, optionally, the first column."""
    freeze_column = freeze_first_column and FREEZE_NAME_COLUMN
    if not FREEZE_HEADER_ROW and not freeze_column:
        return
    column_letter = "B" if freeze_column else "A"
    row_number = 2 if FREEZE_HEADER_ROW else 1
    ws.freeze_panes = f"{column_letter}{row_number}"


def _apply_auto_filter(ws: Worksheet, row_count: int, column_count: int) -> None:
    """Apply an auto filter across the populated table range."""
    if not AUTO_FILTER or column_count == 0:
        return
    last_column = get_column_letter(column_count)
    last_row = row_count + 1
    ws.auto_filter.ref = f"A1:{last_column}{last_row}"


def _write_table_sheet(
    ws: Worksheet,
    frame: pd.DataFrame,
    freeze_first_column: bool,
) -> None:
    """Write a tabular DataFrame to a worksheet with full formatting."""
    headers = list(frame.columns)
    ws.append(headers)

    for record in frame.fillna("").itertuples(index=False, name=None):
        ws.append([_cell_value(value) for value in record])

    _style_header_row(ws, len(headers))

    for row_cells in ws.iter_rows(min_row=2, max_col=len(headers)):
        for cell in row_cells:
            cell.alignment = _DATA_ALIGNMENT

    if AUTO_FIT_COLUMNS:
        _autofit_columns(ws)
    _freeze_panes(ws, freeze_first_column)
    _apply_auto_filter(ws, len(frame), len(headers))


def _write_key_value_sheet(
    ws: Worksheet,
    title: str,
    subtitle: str,
    rows: List[Tuple[str, Any]],
) -> None:
    """Write a titled key/value dashboard-style worksheet."""
    ws.append([title])
    ws.append([subtitle])
    ws.append([])
    for label, value in rows:
        ws.append([label, _cell_value(value)])

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)

    title_cell = ws.cell(row=1, column=1)
    title_cell.font = _TITLE_FONT
    title_cell.fill = _TITLE_FILL
    title_cell.alignment = _CENTER_ALIGNMENT

    subtitle_cell = ws.cell(row=2, column=1)
    subtitle_cell.font = _SUBTITLE_FONT
    subtitle_cell.alignment = _CENTER_ALIGNMENT

    for row_cells in ws.iter_rows(min_row=4, min_col=1, max_col=1):
        for cell in row_cells:
            cell.font = _LABEL_FONT

    if AUTO_FIT_COLUMNS:
        _autofit_columns(ws)


def build_workbook(
    partners_df: pd.DataFrame,
    changes_df: pd.DataFrame,
    summary_rows: List[Tuple[str, Any]],
    logs_rows: List[Tuple[str, Any]],
) -> Workbook:
    """Assemble the complete workbook with all configured sheets."""
    workbook = Workbook()
    workbook.remove(workbook.active)

    summary_sheet = workbook.create_sheet(SUMMARY_SHEET)
    partners_sheet = workbook.create_sheet(PARTNERS_SHEET)
    changes_sheet = workbook.create_sheet(CHANGES_SHEET)
    logs_sheet = workbook.create_sheet(LOGS_SHEET)

    _write_key_value_sheet(
        summary_sheet, REPORT_TITLE, REPORT_SUBTITLE, summary_rows
    )
    _write_table_sheet(partners_sheet, partners_df, freeze_first_column=True)
    _write_table_sheet(changes_sheet, changes_df, freeze_first_column=False)
    _write_key_value_sheet(
        logs_sheet, REPORT_TITLE, LOGS_SUBTITLE, logs_rows
    )

    return workbook


def save_workbook(workbook: Workbook) -> None:
    """Persist the workbook to the configured output path."""
    output_path = Path(OUTPUT_EXCEL_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


# =========================================================
# Reporting
# =========================================================

def print_summary(
    rows_exported: int,
    countries: int,
    cities: int,
    total_changes: int,
    duration: float,
) -> None:
    """Log the final workbook-build summary block."""
    log(SUMMARY_BORDER)
    log("Final Excel Workbook Complete")
    log(f"Rows Exported   : {rows_exported}")
    log(f"Countries       : {countries}")
    log(f"Cities          : {cities}")
    log(f"Total Changes   : {total_changes}")
    log(f"Duration        : {_format_duration(duration)}")
    log(f"Output Workbook : {OUTPUT_EXCEL_FILE}")
    log(SUMMARY_BORDER)


# =========================================================
# Orchestration
# =========================================================

def main() -> None:
    """Build and save the final Excel deliverable."""
    start_time = time.perf_counter()
    start_timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    log(SUMMARY_BORDER)
    log(PROJECT_NAME)
    log(f"Version      : {VERSION}")
    log(f"Start Time   : {start_timestamp}")
    log(f"Partner File : {PARTNER_DETAILS_FILE}")
    log(f"Changes File : {CHANGE_LOG_FILE}")
    log(f"Output File  : {OUTPUT_EXCEL_FILE}")
    log(SUMMARY_BORDER)

    details_df = load_dataset(
        PARTNER_DETAILS_FILE, "partner details", REQUIRED_PARTNER_COLUMNS
    )
    changes_df = load_dataset(
        CHANGE_LOG_FILE, "changes log", CHANGE_LOG_COLUMNS
    )

    partners_df = build_partner_dataframe(details_df)
    countries, cities = summarize_locations(details_df)
    change_stats = summarize_changes(changes_df)

    generated_time = datetime.now().strftime(TIMESTAMP_FORMAT)
    summary_rows = build_summary_rows(
        len(details_df), countries, cities, change_stats, generated_time
    )
    logs_rows = build_logs_rows(
        len(details_df), change_stats["total"], generated_time
    )

    workbook = build_workbook(
        partners_df, changes_df, summary_rows, logs_rows
    )
    save_workbook(workbook)

    duration = time.perf_counter() - start_time
    print_summary(
        len(partners_df),
        countries,
        cities,
        change_stats["total"],
        duration,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - logged before fatal re-raise
        log("Fatal error while building the final Excel workbook.", "ERROR")
        log(str(exc), "ERROR")
        raise
