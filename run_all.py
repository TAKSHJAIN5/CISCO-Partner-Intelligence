"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
run_all.py

Purpose:
Main entry point for the platform. Executes the complete
partner intelligence pipeline in order and records every run
in the run-history workbook.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

sys.path.append(str(Path(__file__).parent / "src"))

from config import (  # noqa: E402 - import after src path is registered
    COLOR_HEADER_FONT,
    COLUMN_WIDTH_PADDING,
    PIPELINE_SCRIPTS,
    PROJECT_NAME,
    RUN_HISTORY_FAILURE_COLOR,
    RUN_HISTORY_FILE,
    RUN_HISTORY_HEADER_COLOR,
    RUN_HISTORY_HEADERS,
    RUN_HISTORY_SHEET_NAME,
    RUN_HISTORY_SUCCESS_COLOR,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SUMMARY_BORDER,
    TIMESTAMP_FORMAT,
    VERSION,
)
from logger import log  # noqa: E402 - import after src path is registered

# Run status labels recorded in the run-history workbook.
STATUS_SUCCESS = "SUCCESS"
STATUS_FAILED = "FAILED"

# The Status column is the second column in the run history.
STATUS_COLUMN_INDEX = 2

# Minimum column width for the run-history workbook.
RUN_HISTORY_MIN_COLUMN_WIDTH = 18


# =========================================================
# Helpers
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


def _elapsed_minutes(start_time: float) -> float:
    """Return minutes elapsed since a ``perf_counter`` start value."""
    return (time.perf_counter() - start_time) / SECONDS_PER_MINUTE


# =========================================================
# Run History
# =========================================================

def _style_run_history_header(ws: Worksheet) -> None:
    """Write and style the run-history header row."""
    ws.append(RUN_HISTORY_HEADERS)

    header_fill = PatternFill(
        fill_type="solid", fgColor=RUN_HISTORY_HEADER_COLOR
    )
    header_font = Font(color=COLOR_HEADER_FONT, bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"
    last_column = get_column_letter(len(RUN_HISTORY_HEADERS))
    ws.auto_filter.ref = f"A1:{last_column}1"


def _open_run_history() -> Tuple[Workbook, Worksheet]:
    """Open the run-history workbook, creating and initializing it if absent."""
    history_file = Path(RUN_HISTORY_FILE)
    history_file.parent.mkdir(parents=True, exist_ok=True)

    if history_file.exists():
        workbook = load_workbook(history_file)
        worksheet = workbook.active
        return workbook, worksheet

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = RUN_HISTORY_SHEET_NAME
    _style_run_history_header(worksheet)
    return workbook, worksheet


def _style_status_cell(ws: Worksheet, row: int, status: str) -> None:
    """Color the status cell green for success and red for failure."""
    color = (
        RUN_HISTORY_SUCCESS_COLOR
        if status == STATUS_SUCCESS
        else RUN_HISTORY_FAILURE_COLOR
    )
    column_letter = get_column_letter(STATUS_COLUMN_INDEX)
    ws[f"{column_letter}{row}"].font = Font(color=color, bold=True)


def _autofit_run_history(ws: Worksheet) -> None:
    """Size run-history columns to fit their contents."""
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
        ws.column_dimensions[letter].width = max(
            RUN_HISTORY_MIN_COLUMN_WIDTH, longest + COLUMN_WIDTH_PADDING
        )


def update_run_history(status: str, duration_minutes: float) -> None:
    """Append the current run to the run-history workbook.

    The workbook is created and styled on first use.
    """
    workbook, worksheet = _open_run_history()

    worksheet.append(
        [
            datetime.now().strftime(TIMESTAMP_FORMAT),
            status,
            round(duration_minutes, 2),
        ]
    )
    _style_status_cell(worksheet, worksheet.max_row, status)
    _autofit_run_history(worksheet)

    workbook.save(Path(RUN_HISTORY_FILE))
    log(f"Run history updated: {RUN_HISTORY_FILE}")


# =========================================================
# Pipeline Execution
# =========================================================

def run_script(script: str) -> None:
    """Execute a single pipeline script as a subprocess.

    Raises:
        RuntimeError: If the script exits with a non-zero status.
    """
    log(f"Running: {script}")
    script_start = time.perf_counter()

    result = subprocess.run([sys.executable, script])

    elapsed = time.perf_counter() - script_start
    if result.returncode != 0:
        raise RuntimeError(
            f"{script} failed with exit code {result.returncode}."
        )

    log(f"Completed: {script} ({_format_duration(elapsed)})")


def run_pipeline() -> None:
    """Execute every pipeline script in order."""
    for script in PIPELINE_SCRIPTS:
        run_script(script)


# =========================================================
# Reporting
# =========================================================

def _log_startup(start_timestamp: str) -> None:
    """Log the orchestration startup banner."""
    log(SUMMARY_BORDER)
    log(PROJECT_NAME)
    log(f"Version    : {VERSION}")
    log(f"Start Time : {start_timestamp}")
    log(f"Pipeline   : {len(PIPELINE_SCRIPTS)} scripts")
    log(SUMMARY_BORDER)


def _log_completion(duration_minutes: float) -> None:
    """Log the successful-completion summary block."""
    log(SUMMARY_BORDER)
    log("Pipeline Completed Successfully")
    log(f"Scripts Run : {len(PIPELINE_SCRIPTS)}")
    log(f"Total Time  : {duration_minutes:.2f} mins")
    log(f"Run History : {RUN_HISTORY_FILE}")
    log(SUMMARY_BORDER)


def _log_failure(message: str) -> None:
    """Log the failure summary block."""
    log(SUMMARY_BORDER)
    log("Pipeline Failed", "ERROR")
    log(message, "ERROR")
    log(SUMMARY_BORDER)


# =========================================================
# Orchestration
# =========================================================

def main() -> None:
    """Run the full pipeline and record the outcome in run history."""
    start_time = time.perf_counter()
    start_timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    _log_startup(start_timestamp)

    try:
        run_pipeline()
    except Exception as exc:  # noqa: BLE001 - recorded, then run aborts
        duration_minutes = _elapsed_minutes(start_time)
        _log_failure(str(exc))
        update_run_history(STATUS_FAILED, duration_minutes)
        sys.exit(1)

    duration_minutes = _elapsed_minutes(start_time)
    update_run_history(STATUS_SUCCESS, duration_minutes)
    _log_completion(duration_minutes)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 - logged before fatal re-raise
        log("Unexpected error during pipeline orchestration.", "ERROR")
        log(str(exc), "ERROR")
        raise
