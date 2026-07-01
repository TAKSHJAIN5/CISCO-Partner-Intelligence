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

All configurable values used across the project
should be maintained here.

===========================================================
"""

# =========================================================
# Project Information
# =========================================================

PROJECT_NAME = "Cisco Partner Intelligence Automation Platform"

VERSION = "1.0.0"

AUTHOR = "Taksh Jain"

COMPANY = "NetCom Learning"

RELEASE_DATE = "June 2026"

STATUS = "Production Ready"

REFRESH_FREQUENCY = "Bi-Weekly"

# =========================================================
# Data Files
# =========================================================

PARTNER_INVENTORY_FILE = "data/partner_inventory.csv"

PARTNER_DETAILS_FILE = "data/partner_details.csv"

PARTNER_DETAILS_PREVIOUS_FILE = "data/partner_details_previous.csv"

# =========================================================
# Log Files
# =========================================================

LOG_FOLDER = "data/logs"

FAILED_PARTNERS_FILE = "data/logs/failed_partners.csv"

CHANGE_LOG_FILE = "data/logs/changes_log.csv"

RUN_LOG_FILE = "data/logs/run_log.txt"

RUN_HISTORY_FILE = "data/logs/Run_History.xlsx"

# =========================================================
# Output
# =========================================================

OUTPUT_FOLDER = "output"

OUTPUT_EXCEL_FILE = "output/Cisco_Partner_Master.xlsx"

# =========================================================
# Cisco API Endpoints
# =========================================================

PARTNER_SEARCH_API = (
    "https://locatr.cloudapps.cisco.com/"
    "WWChannels/LOCATR/service/api/v1/getPfPartners"
)

PARTNER_OVERVIEW_API = (
    "https://locatr.cloudapps.cisco.com/"
    "WWChannels/LOCATR/service/api/v1/getPartnerOverview"
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

RANDOM_SEED = "123456789"

# =========================================================
# Application Settings
# =========================================================

SAVE_BACKUP = True

ENABLE_CHANGE_LOG = True

ENABLE_RUN_LOG = True

# =========================================================
# Excel Configuration
# =========================================================

FREEZE_HEADER_ROW = True

FREEZE_NAME_COLUMN = True

AUTO_FILTER = True

AUTO_FIT_COLUMNS = True