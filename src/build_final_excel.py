"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
build_final_excel.py

Purpose:
Presentation layer. Transforms the collected partner details
and detected changes into the final, management-ready Excel
workbook, applies all professional formatting, and saves the
workbook in a single pass.

The workbook is the only deliverable produced by this script.
There is no separate formatting step and no post-processing.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
    COLOR_ALTERNATE_ROW,
    COLOR_BORDER,
    COLOR_HEADER,
    COLOR_HEADER_FONT,
    COLOR_TITLE,
    COLUMN_WIDTH_PADDING,
    ENABLE_ALTERNATE_ROWS,
    ENABLE_BORDERS,
    FILE_ENCODING,
    FREEZE_HEADER_ROW,
    FREEZE_NAME_COLUMN,
    HEADER_FONT_SIZE,
    LOGS_SHEET,
    LOGS_SUBTITLE,
    MAX_COLUMN_WIDTH,
    MIN_COLUMN_WIDTH,
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
    REPORT_SUMMARY_FILE,
    REPORT_SUBTITLE,
    REPORT_TITLE,
    RUN_BY,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SHOW_GRIDLINES,
    SUBTITLE_FONT_SIZE,
    SUMMARY_BORDER,
    SUMMARY_SHEET,
    TIMESTAMP_FORMAT,
    TITLE_FONT_SIZE,
    TRACKED_CHANGE_FIELDS,
    VERSION,
    WORKSHEET_ORDER,
)
from logger import log

# Snapshot column names read from partner_details.csv.
PARTNER_NAME_COLUMN = "partner_name"
PORTFOLIOS_COLUMN = "portfolios"
COMPETENCIES_COLUMN = "competencies"
LOCATION_COLUMN = "location"
WEBSITE_COLUMN = "website"
PARTNER_SINCE_COLUMN = "partner_since"
PHONE_COLUMN = "phone"
FAX_COLUMN = "fax"
DESCRIPTION_COLUMN = "description"
ADDITIONAL_LOCATIONS_COLUMN = "additional_locations"
INDUSTRY_KEYWORDS_COLUMN = "industry_keywords"
TECHNOLOGY_KEYWORDS_COLUMN = "technology_keywords"
COMPANY_SIZE_KEYWORDS_COLUMN = "company_size_keywords"

# Labels used when composing the presentation-only Contact Details column.
CONTACT_PHONE_PREFIX = "Phone: "
CONTACT_FAX_PREFIX = "Fax: "

# Change log column that stores the type of each change.
CHANGE_TYPE_COLUMN = "change_type"

# Columns required from each input dataset.
REQUIRED_PARTNER_COLUMNS = [
    PARTNER_NAME_COLUMN,
    PORTFOLIOS_COLUMN,
    COMPETENCIES_COLUMN,
    LOCATION_COLUMN,
]

# Presentation labels for boolean partner attributes.
YES_LABEL = "Yes"
NO_LABEL = "No"

# Worksheet layout kinds, used to select the correct formatting.
SHEET_KIND_TABLE = "table"
SHEET_KIND_KEY_VALUE = "key_value"

# Machine-readable report summary (for Power Automate and future
# integrations). The schema version is independent of the platform
# version and is bumped only when the JSON contract changes.
REPORT_SUMMARY_VERSION = "1.1"
EXECUTION_WORKFLOW = "Cisco Partner Intelligence"
EXECUTION_GENERATED_BY = "build_final_excel.py"

# Comparison state, used when no previous snapshot exists yet (the
# initial pipeline run, before detect_changes has produced a log).
COMPARISON_STATUS_INITIAL = "initial_run"
COMPARISON_STATUS_COMPARED = "compared"
SUMMARY_COMPARISON_LABEL = "Comparison Status"
SUMMARY_COMPARISON_INITIAL = "Initial run - no previous snapshot available"
SUMMARY_COMPARISON_COMPARED = "Compared against previous snapshot"
CHANGES_INITIAL_RUN_MESSAGE = (
    "No previous snapshot available. "
    "Change detection will begin from the next execution."
)
CHANGES_NO_CHANGES_MESSAGE = "No changes detected in this run."

# Key/value sheet layout: title, subtitle, blank spacer, then data.
KEY_VALUE_TITLE_ROW = 1
KEY_VALUE_SUBTITLE_ROW = 2
KEY_VALUE_DATA_START_ROW = 4
KEY_VALUE_COLUMN_COUNT = 2

# Table sheet layout: header row 1, data from row 2.
TABLE_HEADER_ROW = 1
TABLE_DATA_START_ROW = 2

# Reusable cell styles derived from configuration.
_THIN_SIDE = Side(style="thin", color=COLOR_BORDER)
_THIN_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE
)
_HEADER_FILL = PatternFill(fill_type="solid", fgColor=COLOR_HEADER)
_HEADER_FONT = Font(color=COLOR_HEADER_FONT, bold=True, size=HEADER_FONT_SIZE)
_TITLE_FILL = PatternFill(fill_type="solid", fgColor=COLOR_TITLE)
_TITLE_FONT = Font(bold=True, size=TITLE_FONT_SIZE)
_SUBTITLE_FONT = Font(italic=True, size=SUBTITLE_FONT_SIZE)
_LABEL_FONT = Font(bold=True)
_ALTERNATE_FILL = PatternFill(fill_type="solid", fgColor=COLOR_ALTERNATE_ROW)
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_DATA_ALIGNMENT = Alignment(vertical="top", wrap_text=True)
_LABEL_ALIGNMENT = Alignment(horizontal="left", vertical="center")
_MESSAGE_FONT = Font(italic=True)
_MESSAGE_ALIGNMENT = Alignment(
    horizontal="center", vertical="center", wrap_text=True
)


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


def _format_contact_details(phone: str, fax: str) -> str:
    """Combine phone and fax into a single presentation value.

    Renders one labeled line per available value ("Phone: ..." / "Fax: ...").
    If the fax is empty only the phone line is shown, and if both are empty
    the result is an empty string so the cell stays blank. This formatting is
    presentation-only; the master CSV keeps phone and fax as separate columns.
    """
    lines: List[str] = []
    if phone:
        lines.append(f"{CONTACT_PHONE_PREFIX}{phone}")
    if fax:
        lines.append(f"{CONTACT_FAX_PREFIX}{fax}")
    return "\n".join(lines)


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


def load_change_log() -> Tuple[pd.DataFrame, bool]:
    """Load the change log, tolerating its absence on the first run.

    ``detect_changes.py`` only produces the change log once a previous
    snapshot exists, so on the initial pipeline run the file is missing.
    In that case an empty log with the canonical schema is returned and
    the pipeline continues rather than failing.

    Returns
    -------
    tuple
        A ``(changes_df, previous_snapshot_available)`` pair. The flag is
        ``False`` on the initial run and ``True`` once a change log exists.
    """
    if Path(CHANGE_LOG_FILE).exists():
        changes_df = load_dataset(
            CHANGE_LOG_FILE, "changes log", CHANGE_LOG_COLUMNS
        )
        return changes_df, True

    log(
        "No change log found; treating this as the initial pipeline run.",
        "WARNING",
    )
    return pd.DataFrame(columns=CHANGE_LOG_COLUMNS), False


# =========================================================
# Business Transformations
# =========================================================

def _build_partner_row(row: pd.Series) -> Dict[str, Any]:
    """Transform one raw partner record into the Partners sheet schema.

    Enriched fields are read defensively from the master row, so a partner
    details file produced before enrichment (without the new columns) still
    renders with blank cells rather than failing.
    """
    portfolios = _split_portfolios(_clean(row.get(PORTFOLIOS_COLUMN)))
    preferred, non_preferred = _classify_portfolios(portfolios)

    preferred_networking = PREFERRED_NETWORKING_PORTFOLIO in preferred
    preferred_security = PREFERRED_SECURITY_PORTFOLIO in preferred
    preferred_cloud_ai = PREFERRED_CLOUD_AI_PORTFOLIO in preferred

    sec_net_spez = preferred_networking and preferred_security
    sec_ai_infra_spez = preferred_security and preferred_cloud_ai

    contact_details = _format_contact_details(
        _clean(row.get(PHONE_COLUMN)),
        _clean(row.get(FAX_COLUMN)),
    )

    return {
        "Partner Name": _clean(row.get(PARTNER_NAME_COLUMN)),
        "Website": _clean(row.get(WEBSITE_COLUMN)),
        "Partner Since": _clean(row.get(PARTNER_SINCE_COLUMN)),
        "Contact Details": contact_details,
        "Description": _clean(row.get(DESCRIPTION_COLUMN)),
        "Additional Locations": _clean(row.get(ADDITIONAL_LOCATIONS_COLUMN)),
        "Industry Keywords": _clean(row.get(INDUSTRY_KEYWORDS_COLUMN)),
        "Technology Keywords": _clean(row.get(TECHNOLOGY_KEYWORDS_COLUMN)),
        "Company Size Keywords": _clean(row.get(COMPANY_SIZE_KEYWORDS_COLUMN)),
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
    previous_snapshot: bool,
) -> List[Tuple[str, Any]]:
    """Assemble the Executive Summary key/value rows."""
    comparison_status = (
        SUMMARY_COMPARISON_COMPARED
        if previous_snapshot
        else SUMMARY_COMPARISON_INITIAL
    )
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
        (SUMMARY_COMPARISON_LABEL, comparison_status),
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

def _write_table(ws: Worksheet, frame: pd.DataFrame) -> Tuple[int, int]:
    """Write a tabular DataFrame (header plus rows) to a worksheet.

    Returns:
        A ``(row_count, column_count)`` tuple describing the data written.
    """
    headers = list(frame.columns)
    ws.append(headers)
    for record in frame.fillna("").itertuples(index=False, name=None):
        ws.append([_cell_value(value) for value in record])
    return len(frame), len(headers)


def _write_changes_sheet(
    ws: Worksheet,
    changes_df: pd.DataFrame,
    empty_message: str,
) -> Tuple[int, int, Optional[int]]:
    """Write the Changes sheet, showing a message when there are no records.

    Returns:
        A ``(row_count, column_count, message_row)`` tuple. ``message_row``
        is the row holding the empty-state message, or ``None`` when change
        records were written.
    """
    if not changes_df.empty:
        row_count, column_count = _write_table(ws, changes_df)
        return row_count, column_count, None

    headers = list(changes_df.columns)
    ws.append(headers)
    message_row = TABLE_DATA_START_ROW
    ws.append([empty_message])
    ws.merge_cells(
        start_row=message_row,
        start_column=1,
        end_row=message_row,
        end_column=len(headers),
    )
    return 0, len(headers), message_row


def _write_key_value(
    ws: Worksheet,
    title: str,
    subtitle: str,
    rows: List[Tuple[str, Any]],
) -> Tuple[int, int]:
    """Write a titled key/value dashboard layout to a worksheet.

    Returns:
        A ``(row_count, column_count)`` tuple describing the data written.
    """
    ws.append([title])
    ws.append([subtitle])
    ws.append([])
    for label, value in rows:
        ws.append([label, _cell_value(value)])

    ws.merge_cells(
        start_row=KEY_VALUE_TITLE_ROW,
        start_column=1,
        end_row=KEY_VALUE_TITLE_ROW,
        end_column=KEY_VALUE_COLUMN_COUNT,
    )
    ws.merge_cells(
        start_row=KEY_VALUE_SUBTITLE_ROW,
        start_column=1,
        end_row=KEY_VALUE_SUBTITLE_ROW,
        end_column=KEY_VALUE_COLUMN_COUNT,
    )
    return len(rows), KEY_VALUE_COLUMN_COUNT


def build_workbook(
    partners_df: pd.DataFrame,
    changes_df: pd.DataFrame,
    summary_rows: List[Tuple[str, Any]],
    logs_rows: List[Tuple[str, Any]],
    changes_empty_message: str,
) -> Tuple[Workbook, List[Dict[str, Any]]]:
    """Create the workbook, write every sheet, and describe each for formatting.

    Returns:
        The populated (unformatted) workbook and a list of sheet
        specifications consumed by :func:`format_workbook`.
    """
    workbook = Workbook()
    workbook.remove(workbook.active)
    worksheets = {
        name: workbook.create_sheet(name) for name in WORKSHEET_ORDER
    }

    specs: List[Dict[str, Any]] = []

    summary_ws = worksheets[SUMMARY_SHEET]
    rows, columns = _write_key_value(
        summary_ws, REPORT_TITLE, REPORT_SUBTITLE, summary_rows
    )
    specs.append(_sheet_spec(summary_ws, SHEET_KIND_KEY_VALUE, rows, columns))

    partners_ws = worksheets[PARTNERS_SHEET]
    rows, columns = _write_table(partners_ws, partners_df)
    specs.append(
        _sheet_spec(
            partners_ws, SHEET_KIND_TABLE, rows, columns,
            freeze_first_column=True,
        )
    )

    changes_ws = worksheets[CHANGES_SHEET]
    rows, columns, message_row = _write_changes_sheet(
        changes_ws, changes_df, changes_empty_message
    )
    specs.append(
        _sheet_spec(
            changes_ws, SHEET_KIND_TABLE, rows, columns,
            message_row=message_row,
        )
    )

    logs_ws = worksheets[LOGS_SHEET]
    rows, columns = _write_key_value(
        logs_ws, REPORT_TITLE, LOGS_SUBTITLE, logs_rows
    )
    specs.append(_sheet_spec(logs_ws, SHEET_KIND_KEY_VALUE, rows, columns))

    return workbook, specs


def _sheet_spec(
    worksheet: Worksheet,
    kind: str,
    row_count: int,
    column_count: int,
    freeze_first_column: bool = False,
    message_row: Optional[int] = None,
) -> Dict[str, Any]:
    """Build a formatting specification for a single worksheet."""
    return {
        "worksheet": worksheet,
        "kind": kind,
        "row_count": row_count,
        "column_count": column_count,
        "freeze_first_column": freeze_first_column,
        "message_row": message_row,
    }


def save_workbook(workbook: Workbook) -> None:
    """Persist the workbook to the configured output path (once)."""
    output_path = Path(OUTPUT_EXCEL_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_report_summary(
    total_partners: int,
    countries: int,
    cities: int,
    change_stats: Dict[str, int],
    previous_snapshot: bool,
) -> None:
    """Write the machine-readable JSON report summary.

    This artifact is consumed only by downstream integrations
    (Power Automate, Teams, and future workflows); it is not a
    user-facing report. All values are reused from the current run's
    in-memory results, so no CSV or workbook is re-read.

    Parameters
    ----------
    total_partners:
        Number of partners in the current details snapshot.
    countries:
        Count of distinct countries.
    cities:
        Count of distinct cities.
    change_stats:
        Change counts as returned by :func:`summarize_changes`.
    previous_snapshot:
        ``True`` when a previous snapshot was compared, ``False`` on the
        initial pipeline run when no change log existed yet.
    """
    comparison_status = (
        COMPARISON_STATUS_COMPARED
        if previous_snapshot
        else COMPARISON_STATUS_INITIAL
    )
    summary: Dict[str, Any] = {
        "report": {
            "name": REPORT_TITLE,
            "generated_on": datetime.now().isoformat(timespec="seconds"),
            "version": REPORT_SUMMARY_VERSION,
        },
        "summary": {
            "total_partners": total_partners,
            "countries": countries,
            "cities": cities,
        },
        "changes": {
            "new_partners": change_stats["new"],
            "removed_partners": change_stats["removed"],
            "portfolio_changes": change_stats["portfolio"],
            "competency_changes": change_stats["competency"],
            "location_changes": change_stats["location"],
            "total_change_records": change_stats["total"],
        },
        "comparison": {
            "status": comparison_status,
            "previous_snapshot": previous_snapshot,
        },
        "files": {
            "excel_report": Path(OUTPUT_EXCEL_FILE).name,
        },
        "execution": {
            "environment": RUN_BY,
            "workflow": EXECUTION_WORKFLOW,
            "generated_by": EXECUTION_GENERATED_BY,
        },
        "status": {
            "success": True,
        },
    }

    output_path = Path(REPORT_SUMMARY_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding=FILE_ENCODING) as summary_file:
        json.dump(summary, summary_file, indent=4)

    log(f"Report summary written: {REPORT_SUMMARY_FILE}")


# =========================================================
# Formatting
# =========================================================

def apply_header_style(ws: Worksheet, column_count: int) -> None:
    """Style a table's header row: fill, font, alignment and border."""
    for column_index in range(1, column_count + 1):
        cell = ws.cell(row=TABLE_HEADER_ROW, column=column_index)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _HEADER_ALIGNMENT
        if ENABLE_BORDERS:
            cell.border = _THIN_BORDER


def apply_alignment(
    ws: Worksheet,
    column_count: int,
    start_row: int,
    alignment: Alignment,
) -> None:
    """Apply a shared alignment to a block of data cells."""
    for row_cells in ws.iter_rows(min_row=start_row, max_col=column_count):
        for cell in row_cells:
            cell.alignment = alignment


def apply_alternating_rows(
    ws: Worksheet,
    column_count: int,
    start_row: int,
    end_row: int,
) -> None:
    """Apply borders to data cells and shade alternating rows."""
    if end_row < start_row:
        return
    for row_index in range(start_row, end_row + 1):
        shade = ENABLE_ALTERNATE_ROWS and row_index % 2 == 0
        for column_index in range(1, column_count + 1):
            cell = ws.cell(row=row_index, column=column_index)
            if ENABLE_BORDERS:
                cell.border = _THIN_BORDER
            if shade:
                cell.fill = _ALTERNATE_FILL


def apply_freeze_panes(ws: Worksheet, freeze_first_column: bool) -> None:
    """Freeze the header row and, optionally, the first column."""
    freeze_column = freeze_first_column and FREEZE_NAME_COLUMN
    if not FREEZE_HEADER_ROW and not freeze_column:
        return
    column_letter = "B" if freeze_column else "A"
    row_number = 2 if FREEZE_HEADER_ROW else 1
    ws.freeze_panes = f"{column_letter}{row_number}"


def apply_auto_filter(ws: Worksheet, row_count: int, column_count: int) -> None:
    """Apply an auto filter across the populated table range."""
    if not AUTO_FILTER or column_count == 0:
        return
    last_column = get_column_letter(column_count)
    last_row = row_count + 1
    ws.auto_filter.ref = f"A1:{last_column}{last_row}"


def apply_column_widths(ws: Worksheet) -> None:
    """Size every column to fit its widest value within configured bounds."""
    if not AUTO_FIT_COLUMNS:
        return
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
        if longest < MIN_COLUMN_WIDTH:
            width = MIN_COLUMN_WIDTH
        elif longest > MAX_COLUMN_WIDTH:
            width = MAX_COLUMN_WIDTH
        else:
            width = longest + COLUMN_WIDTH_PADDING
        ws.column_dimensions[letter].width = width


def _apply_key_value_style(ws: Worksheet) -> None:
    """Style a key/value sheet: title banner, subtitle and bold labels."""
    title_cell = ws.cell(row=KEY_VALUE_TITLE_ROW, column=1)
    title_cell.font = _TITLE_FONT
    title_cell.fill = _TITLE_FILL
    title_cell.alignment = _HEADER_ALIGNMENT

    subtitle_cell = ws.cell(row=KEY_VALUE_SUBTITLE_ROW, column=1)
    subtitle_cell.font = _SUBTITLE_FONT
    subtitle_cell.alignment = _HEADER_ALIGNMENT

    for row_index in range(KEY_VALUE_DATA_START_ROW, ws.max_row + 1):
        label_cell = ws.cell(row=row_index, column=1)
        value_cell = ws.cell(row=row_index, column=2)
        label_cell.font = _LABEL_FONT
        label_cell.alignment = _LABEL_ALIGNMENT
        value_cell.alignment = _LABEL_ALIGNMENT
        if ENABLE_BORDERS:
            label_cell.border = _THIN_BORDER
            value_cell.border = _THIN_BORDER


def _style_message_row(ws: Worksheet, row: int, column_count: int) -> None:
    """Style an empty-state message row: italic, centered, bordered."""
    for column_index in range(1, column_count + 1):
        cell = ws.cell(row=row, column=column_index)
        cell.font = _MESSAGE_FONT
        cell.alignment = _MESSAGE_ALIGNMENT
        if ENABLE_BORDERS:
            cell.border = _THIN_BORDER


def _apply_worksheet_view(ws: Worksheet) -> None:
    """Apply workbook view options (gridline visibility)."""
    ws.sheet_view.showGridLines = SHOW_GRIDLINES


def format_sheet(spec: Dict[str, Any]) -> None:
    """Apply the full formatting appropriate to one worksheet."""
    ws = spec["worksheet"]
    column_count = spec["column_count"]
    row_count = spec["row_count"]
    message_row = spec.get("message_row")

    if spec["kind"] == SHEET_KIND_TABLE:
        apply_header_style(ws, column_count)
        apply_alignment(ws, column_count, TABLE_DATA_START_ROW, _DATA_ALIGNMENT)
        apply_alternating_rows(
            ws, column_count, TABLE_DATA_START_ROW, row_count + 1
        )
        apply_column_widths(ws)
        apply_freeze_panes(ws, spec["freeze_first_column"])
        apply_auto_filter(ws, row_count, column_count)
        if message_row is not None:
            _style_message_row(ws, message_row, column_count)
    else:
        _apply_key_value_style(ws)
        apply_column_widths(ws)

    _apply_worksheet_view(ws)


def format_workbook(specs: List[Dict[str, Any]]) -> None:
    """Apply formatting to every worksheet in the workbook."""
    for spec in specs:
        format_sheet(spec)


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
    """Build, format and save the final Excel deliverable in one pass."""
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
    changes_df, previous_snapshot = load_change_log()

    partners_df = build_partner_dataframe(details_df)
    countries, cities = summarize_locations(details_df)
    change_stats = summarize_changes(changes_df)

    generated_time = datetime.now().strftime(TIMESTAMP_FORMAT)
    summary_rows = build_summary_rows(
        len(details_df), countries, cities, change_stats,
        generated_time, previous_snapshot,
    )
    logs_rows = build_logs_rows(
        len(details_df), change_stats["total"], generated_time
    )

    changes_empty_message = (
        CHANGES_NO_CHANGES_MESSAGE
        if previous_snapshot
        else CHANGES_INITIAL_RUN_MESSAGE
    )

    workbook, specs = build_workbook(
        partners_df, changes_df, summary_rows, logs_rows,
        changes_empty_message,
    )
    format_workbook(specs)
    save_workbook(workbook)
    write_report_summary(
        len(details_df), countries, cities, change_stats, previous_snapshot
    )

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
