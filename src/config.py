"""
===========================================================
Cisco Partner Intelligence Automation Platform
===========================================================

Version : 1.0.0

Author  : Taksh Jain
Company : NetCom Learning

Description:
Central configuration file for the Cisco Partner
Intelligence Automation Platform.

This file contains every configurable setting used
throughout the project.

All scripts should import values from here instead
of using hardcoded constants.

===========================================================
"""

# =========================================================
# Project Information
# =========================================================

PROJECT_NAME = "Cisco Partner Intelligence Automation Platform"

VERSION = "2.1.0"

AUTHOR = "Taksh Jain"

COMPANY = "NetCom Learning"

RELEASE_DATE = "July 2026"

STATUS = "Production"

REFRESH_FREQUENCY = "Bi-Weekly"

RUN_BY = "GitHub Actions"

TIMEZONE = "Asia/Kolkata"

# =========================================================
# Report Information
# =========================================================

REPORT_TITLE = "Cisco Partner Intelligence Report"

REPORT_SUBTITLE = (
    "Automated Cisco Partner Intelligence & Change Report"
)

# =========================================================
# Folder Paths
# =========================================================

DATA_FOLDER = "data"

LOG_FOLDER = f"{DATA_FOLDER}/logs"

OUTPUT_FOLDER = "output"

# =========================================================
# Data Files
# =========================================================

PARTNER_INVENTORY_FILE = (
    f"{DATA_FOLDER}/partner_inventory.csv"
)

PARTNER_DETAILS_FILE = (
    f"{DATA_FOLDER}/partner_details.csv"
)

PARTNER_DETAILS_PREVIOUS_FILE = (
    f"{DATA_FOLDER}/partner_details_previous.csv"
)

# =========================================================
# Log Files
# =========================================================

FAILED_PARTNERS_FILE = (
    f"{LOG_FOLDER}/failed_partners.csv"
)

CHANGE_LOG_FILE = (
    f"{LOG_FOLDER}/changes_log.csv"
)

RUN_LOG_FILE = (
    f"{LOG_FOLDER}/run_log.txt"
)

RUN_HISTORY_FILE = (
    f"{LOG_FOLDER}/Run_History.xlsx"
)

# =========================================================
# Output Files
# =========================================================

OUTPUT_EXCEL_FILE = (
    f"{OUTPUT_FOLDER}/Cisco_Partner_Intelligence_Report.xlsx"
)

# =========================================================
# Excel Sheet Names
# =========================================================

SUMMARY_SHEET = "Executive Summary"

PARTNERS_SHEET = "Partners"

CHANGES_SHEET = "Changes"

LOGS_SHEET = "Logs"

# =========================================================
# Hyperlink Labels
# =========================================================

WEBSITE_LINK_TEXT = "🌐 Website"

CISCO_LINK_TEXT = "🔗 Cisco Profile"

# =========================================================
# Report Colors
# =========================================================

COLOR_HEADER = "1F4E78"

COLOR_TITLE = "D9EAD3"

COLOR_NEW = "C6EFCE"

COLOR_UPDATED = "FFEB9C"

COLOR_REMOVED = "FFC7CE"

# =========================================================
# Email Configuration
# =========================================================

EMAIL_ENABLED = True

EMAIL_SUBJECT_PREFIX = (
    "Cisco Partner Intelligence Report"
)

EMAIL_RECIPIENTS = [
    # Example:
    # "sunil@company.com",
]

# =========================================================
# Cisco API Endpoints
# =========================================================

BASE_API = (
    "https://locatr.cloudapps.cisco.com/"
    "WWChannels/LOCATR/service/api/v1"
)

PARTNER_SEARCH_API = (
    f"{BASE_API}/getPfPartners"
)

PARTNER_OVERVIEW_API = (
    f"{BASE_API}/getPartnerOverview"
)

# =========================================================
# Cisco Search Filters
# =========================================================

PORTFOLIO_IDS = (
    "^137061562^,^137061563^,^137061564^,^137061565^,"
    "^137061566^,^137061567^,^137061568^,^137061569^,"
    "^137061570^,^137061571^,^137061572^,^137061573^"
)

COMPETENCY_IDS = (
    "^137061583^,^137061577^,^137061578^,^137063260^,"
    "^137061579^,^137064818^,^137061585^,^137061582^,"
    "^137061586^,^137061581^,^137061580^,^137061778^"
)

# =========================================================
# API Configuration
# =========================================================

LANGUAGE = "EN"

PAGE_SIZE = 12

REQUEST_TIMEOUT = 30

MAX_RETRIES = 3

RETRY_DELAY = 5

REQUEST_DELAY = 0.5

CHECKPOINT_INTERVAL = 50

USER_AGENT = (
    f"{PROJECT_NAME}/{VERSION}"
)

RANDOM_SEED = "123456789"

# =========================================================
# Date & Time Formats
# =========================================================

DATE_FORMAT = "%Y-%m-%d"

TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M:%S"

# =========================================================
# Application Settings
# =========================================================

SAVE_BACKUP = True

ENABLE_CHANGE_LOG = True

ENABLE_RUN_LOG = True

# =========================================================
# CSV Write Defaults
# =========================================================
# Applied to every CSV written across the project so output
# encoding and indexing stay consistent.

CSV_INDEX = False

FILE_ENCODING = "utf-8"

# =========================================================
# Excel Configuration
# =========================================================

FREEZE_HEADER_ROW = True

FREEZE_NAME_COLUMN = True

AUTO_FILTER = True

AUTO_FIT_COLUMNS = True

# =========================================================
# Dataset Schemas
# =========================================================
# Shared column contracts consumed across the pipeline.
# These are the single source of truth for CSV structure;
# do not redefine them inside individual scripts and do not
# reorder them without updating every downstream stage.

# Inventory schema produced by collect_partner_inventory.py.
INVENTORY_COLUMNS = [
    "partner_name",
    "be_geo_id",
    "location_id",
    "partner_id",
]

# Detail schema produced by get_partner_details.py.
OUTPUT_COLUMNS = [
    "partner_name",
    "site_id",
    "be_geo_id",
    "partner_since",
    "website",
    "phone",
    "address_line_1",
    "address_line_2",
    "address_line_3",
    "city",
    "state",
    "country",
    "zip_code",
    "coordinates",
    "description",
    "location",
    "portfolios",
    "competencies",
]

# Audit schema for partners that could not be collected.
FAILED_COLUMNS = [
    "partner_name",
    "site_id",
    "be_geo_id",
    "location_id",
    "attempts",
    "error_message",
    "timestamp",
]

# Inventory columns that must be present for detail collection.
REQUIRED_INVENTORY_COLUMNS = (
    "partner_name",
    "be_geo_id",
    "location_id",
)

# =========================================================
# Change Log Schema
# =========================================================
# Column contract for the change log produced by
# detect_changes.py, plus the mapping of snapshot fields that
# are compared between runs. Add a new entry to
# TRACKED_CHANGE_FIELDS to track an additional field; no script
# changes are required.

CHANGE_LOG_COLUMNS = [
    "timestamp",
    "partner_name",
    "change_type",
    "old_value",
    "new_value",
]

TRACKED_CHANGE_FIELDS = {
    "portfolios": "PORTFOLIO_CHANGED",
    "competencies": "COMPETENCY_CHANGED",
    "location": "LOCATION_CHANGED",
}

# Change types and value labels for partner additions and removals.
CHANGE_TYPE_NEW_PARTNER = "NEW_PARTNER"

CHANGE_TYPE_REMOVED_PARTNER = "REMOVED_PARTNER"

NEW_PARTNER_NEW_VALUE = "Added"

REMOVED_PARTNER_OLD_VALUE = "Present"

REMOVED_PARTNER_NEW_VALUE = "Removed"

# =========================================================
# Partner Classification Rules
# =========================================================
# Business rules used to classify partner portfolios in the
# final workbook. A portfolio containing PREFERRED_KEYWORD is
# treated as a preferred portfolio.

PREFERRED_KEYWORD = "Preferred"

PREFERRED_NETWORKING_PORTFOLIO = "Preferred Networking Partner"

PREFERRED_SECURITY_PORTFOLIO = "Preferred Security Partner"

PREFERRED_CLOUD_AI_PORTFOLIO = (
    "Preferred Cloud and AI Infrastructure Partner"
)

# =========================================================
# Partners Sheet Schema
# =========================================================
# Column order for the business-friendly Partners sheet in the
# final Excel workbook.

PARTNERS_SHEET_COLUMNS = [
    "Name",
    "Preferred Portfolios",
    "Non Preferred Portfolios",
    "Preferred Networking",
    "Preferred Security",
    "Preferred Cloud + AI",
    "Sec Net Spez",
    "Sec AI Infra Spez",
    "Competencies",
    "Location",
]

# =========================================================
# Workbook Layout
# =========================================================
# Order in which worksheets appear in the final workbook, plus
# the subtitle used on the metadata (Logs) sheet.

WORKSHEET_ORDER = [
    SUMMARY_SHEET,
    PARTNERS_SHEET,
    CHANGES_SHEET,
    LOGS_SHEET,
]

LOGS_SUBTITLE = "Execution Metadata"

# =========================================================
# Excel Formatting
# =========================================================
# Reusable presentation settings for the final workbook. All
# colors are 6-digit hex RGB strings.

COLOR_HEADER_FONT = "FFFFFF"

COLOR_ALTERNATE_ROW = "F7F9FC"

COLOR_BORDER = "D9D9D9"

HEADER_FONT_SIZE = 11

TITLE_FONT_SIZE = 14

SUBTITLE_FONT_SIZE = 11

ENABLE_ALTERNATE_ROWS = True

ENABLE_BORDERS = True

SHOW_GRIDLINES = False

MIN_COLUMN_WIDTH = 15

MAX_COLUMN_WIDTH = 60

COLUMN_WIDTH_PADDING = 3

# =========================================================
# Pipeline Execution
# =========================================================
# Scripts executed, in order, by run_all.py. Formatting is now
# handled inside build_final_excel.py, so there is no separate
# formatting stage.

PIPELINE_SCRIPTS = [
    "src/collect_partner_inventory.py",
    "src/get_partner_details.py",
    "src/detect_changes.py",
    "src/build_final_excel.py",
]

# =========================================================
# Run History
# =========================================================
# Layout and status colors for the run-history workbook.

RUN_HISTORY_SHEET_NAME = "Run History"

RUN_HISTORY_HEADERS = [
    "Timestamp",
    "Status",
    "Duration (mins)",
]

RUN_HISTORY_HEADER_COLOR = COLOR_HEADER

RUN_HISTORY_SUCCESS_COLOR = "008000"

RUN_HISTORY_FAILURE_COLOR = "FF0000"

# =========================================================
# Qualification Categories
# =========================================================
# Categories exposed by getPartnerOverview. Any category
# outside this set is ignored rather than treated as an error.

CATEGORY_PORTFOLIO = "PORTFOLIO"

CATEGORY_COMPETENCY = "COMPETENCY"

# =========================================================
# Data Quality
# =========================================================

# URL schemes that indicate a website value is already
# addressable and must be left untouched.
VALID_URL_SCHEMES = (
    "http://",
    "https://",
    "mailto:",
    "tel:",
    "javascript:",
)

# Scheme prepended to bare website values.
HTTPS_PREFIX = "https://"

# Separator used to serialize portfolio and competency lists.
# The surrounding spaces intentionally match the format already
# stored in partner_details.csv so change detection stays stable.
QUALIFICATION_SEPARATOR = " | "

# Separator used to compose "City, Country" location labels.
LOCATION_SEPARATOR = ", "

# =========================================================
# Run Thresholds
# =========================================================

# Maximum tolerated fraction of failed partners before a run
# is treated as unsuccessful.
FAILURE_THRESHOLD = 0.02

# =========================================================
# Reporting & Time
# =========================================================

# Banner rule used for startup and summary log blocks.
SUMMARY_BORDER = "=" * 60

# Time conversion helpers for human-readable durations.
SECONDS_PER_MINUTE = 60

SECONDS_PER_HOUR = 3600

# =========================================================
# HTTP Request Headers
# =========================================================

ACCEPT_JSON = "application/json"

CONTENT_TYPE_JSON = "application/json"

CONNECTION_KEEP_ALIVE = "keep-alive"

# Shared header set applied to every requests.Session.
REQUEST_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": ACCEPT_JSON,
    "Content-Type": CONTENT_TYPE_JSON,
    "Connection": CONNECTION_KEEP_ALIVE,
}
