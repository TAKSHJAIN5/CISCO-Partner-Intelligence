"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
run_all.py

Purpose:
Main entry point for the application.
Executes the complete Cisco Partner Automation pipeline.

Author:
Taksh Jain

Version:
1.1.0
===========================================================
"""

import json
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.append(
    str(Path(__file__).parent / "src")
)

from config import (
    PROJECT_NAME,
    VERSION,
    RUN_HISTORY_FILE,
    REPORT_SUMMARY_FILE
)

from logger import log


SCRIPTS = [
    "src/collect_partner_inventory.py",
    "src/get_partner_details.py",
    "src/detect_changes.py",
    "src/build_final_excel.py",
    "src/format_excel.py"
]


def update_run_history(status, duration):
    """
    Appends the current run to the Run_History.xlsx workbook.
    Creates the workbook if it does not already exist.
    """

    history_file = Path(RUN_HISTORY_FILE)

    history_file.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    if history_file.exists():

        wb = load_workbook(history_file)
        ws = wb.active

    else:

        wb = Workbook()
        ws = wb.active
        ws.title = "Run History"

        headers = [
            "Timestamp",
            "Status",
            "Duration (mins)"
        ]

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        ws.append(headers)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:C1"

    ws.append([
        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        status,
        round(duration, 2)
    ])

    current_row = ws.max_row

    status_cell = ws[f"B{current_row}"]

    if status == "SUCCESS":

        status_cell.font = Font(
            color="008000",
            bold=True
        )

    else:

        status_cell.font = Font(
            color="FF0000",
            bold=True
        )

    for column in ws.columns:

        column_letter = get_column_letter(
            column[0].column
        )

        max_length = 0

        for cell in column:

            if cell.value:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = max(
            18,
            max_length + 3
        )

    wb.save(history_file)


def update_report_summary(status, runtime_minutes):
    """
    Updates latest_report.json with
    execution status and runtime.
    """

    summary_file = Path(
        REPORT_SUMMARY_FILE
    )

    if not summary_file.exists():

        return

    def update_report_summary(status, runtime_minutes):
    """
    Updates latest_report.json with
    execution status and runtime.
    """

    summary_file = Path(
        REPORT_SUMMARY_FILE
    )

    if not summary_file.exists():

        return

    try:

        with open(
            summary_file,
            "r",
            encoding="utf-8"
        ) as f:

            summary = json.load(f)

        summary["status"] = status

        summary["runtime_minutes"] = round(
            runtime_minutes,
            2
        )

        with open(
            summary_file,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                summary,
                f,
                indent=4
            )

        log(
            "Report summary updated."
        )

    except Exception as e:

        log(
            f"Could not update report summary: {e}",
            "ERROR"
        )

    summary["status"] = status

    summary["runtime_minutes"] = round(
        runtime_minutes,
        2
    )

    with open(
        summary_file,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            summary,
            f,
            indent=4
        )

    log(
        "Report summary updated."
    )

def main():

    start = time.time()

    log("=" * 60)
    log(PROJECT_NAME)
    log(f"Version : {VERSION}")
    log(
        f"Started : "
        f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    log("=" * 60)

    for script in SCRIPTS:

        script_start = time.time()

        log(f"Running : {script}")

        result = subprocess.run(
            [sys.executable, script]
        )

        elapsed = round(
            time.time() - script_start,
            2
        )

        if result.returncode != 0:

            total = (
                time.time() - start
            ) / 60

            log(
                f"{script} FAILED",
                "ERROR"
            )

            update_run_history(
                "FAILED",
                total
            )

            update_report_summary(
                "FAILED",
                total
            )

            log("=" * 60)
            log(
                "PROCESS FAILED",
                "ERROR"
            )
            log("=" * 60)

            sys.exit(1)

        log(
            f"Completed ({elapsed} sec)"
        )

    total = (
        time.time() - start
    ) / 60

    update_run_history(
        "SUCCESS",
        total
    )

    update_report_summary(
        "SUCCESS",
        total
    )

    log("=" * 60)
    log("PROCESS COMPLETED SUCCESSFULLY")
    log(
        f"Total Time : "
        f"{round(total,2)} mins"
    )
    log("=" * 60)


if __name__ == "__main__":

    try:

        main()

    except Exception as e:

        log(
            "Unexpected error occurred.",
            "ERROR"
        )

        log(
            str(e),
            "ERROR"
        )

        raise
