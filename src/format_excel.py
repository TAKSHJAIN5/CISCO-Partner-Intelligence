"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
format_excel.py

Purpose:
Applies professional formatting to the generated Excel
workbook for business users.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side
)
from openpyxl.utils import get_column_letter

from config import (
    OUTPUT_EXCEL_FILE,
    FREEZE_HEADER_ROW,
    FREEZE_NAME_COLUMN,
    AUTO_FILTER,
    AUTO_FIT_COLUMNS
)

from logger import log


try:

    log("Applying Excel formatting...")

    wb = load_workbook(
        OUTPUT_EXCEL_FILE
    )

    ws = wb.active

    # =====================================================
    # Header Style
    # =====================================================

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
        size=11
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        cell.border = thin_border

    # =====================================================
    # Freeze Panes
    # =====================================================

    if FREEZE_HEADER_ROW and FREEZE_NAME_COLUMN:

        ws.freeze_panes = "B2"

    elif FREEZE_HEADER_ROW:

        ws.freeze_panes = "A2"

    elif FREEZE_NAME_COLUMN:

        ws.freeze_panes = "B1"

    # =====================================================
    # Filters
    # =====================================================

    if AUTO_FILTER:

        ws.auto_filter.ref = ws.dimensions

    # =====================================================
    # Alternate Row Formatting
    # =====================================================

    alternate_fill = PatternFill(
        fill_type="solid",
        fgColor="F7F9FC"
    )

    for row in range(2, ws.max_row + 1):

        for col in range(1, ws.max_column + 1):

            cell = ws.cell(
                row=row,
                column=col
            )

            cell.border = thin_border

            if row % 2 == 0:

                cell.fill = alternate_fill

    # =====================================================
    # Column Widths
    # =====================================================

    if AUTO_FIT_COLUMNS:

        for column in ws.columns:

            max_length = 0

            column_letter = get_column_letter(
                column[0].column
            )

            for cell in column:

                try:

                    if cell.value is not None:

                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                except Exception:

                    continue

            if max_length < 15:

                width = 15

            elif max_length > 60:

                width = 60

            else:

                width = max_length + 3

            ws.column_dimensions[
                column_letter
            ].width = width

    # =====================================================
    # Cell Alignment
    # =====================================================

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row
    ):

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    # =====================================================
    # Workbook View
    # =====================================================

    ws.sheet_view.showGridLines = False

    # =====================================================
    # Save Workbook
    # =====================================================

    wb.save(
        OUTPUT_EXCEL_FILE
    )

    log("=" * 60)

    log("Excel formatting completed successfully.")

    log(f"Workbook : {OUTPUT_EXCEL_FILE}")

    log("=" * 60)

except Exception as e:

    log(
        "Fatal error while formatting Excel.",
        "ERROR"
    )

    log(
        str(e),
        "ERROR"
    )

    raise