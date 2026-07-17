# Cisco Partner Intelligence Automation Platform

## Overview

The Cisco Partner Intelligence Automation Platform is an internal automation solution developed for NetCom Learning to streamline the collection, enrichment, change detection, and reporting of Cisco Partner Locator data.

The platform replaces a manual spreadsheet maintenance process with a fully automated ETL pipeline that retrieves Cisco partner information, enriches partner records, detects business-critical changes between refreshes, and generates a professionally formatted Excel report for business users.

---

# Business Objectives

- Eliminate manual Cisco partner data collection.
- Maintain a centralized and up-to-date partner intelligence repository.
- Detect meaningful changes between data refreshes.
- Improve visibility into partner portfolios, competencies, and geographic coverage.
- Deliver a business-ready Excel report for marketing and partner management teams.

---

# Key Features

- Automated Cisco Partner Locator data collection
- Partner detail enrichment
- Snapshot-based change detection
- Executive summary generation
- Business-friendly Excel reporting
- Run history tracking
- Execution logging
- Configuration-driven architecture
- Suitable for scheduled unattended execution

---

# Architecture

```text
run_all.py
      │
      ▼
collect_partner_inventory.py
      │
      ▼
get_partner_details.py
      │
      ▼
detect_changes.py
      │
      ▼
build_final_excel.py
      │
      ▼
Cisco Partner Intelligence Report.xlsx
```

---

# Project Structure

```text
CiscoPartnerAutomation/

│
├── run_all.py
├── requirements.txt
├── README.md
├── INSTALLATION.md
│
├── src/
│   ├── config.py
│   ├── logger.py
│   ├── collect_partner_inventory.py
│   ├── get_partner_details.py
│   ├── detect_changes.py
│   └── build_final_excel.py
│
├── data/
│   ├── partner_inventory.csv
│   ├── partner_details.csv
│   ├── partner_details_previous.csv
│   └── logs/
│       ├── changes_log.csv
│       ├── failed_partners.csv
│       ├── run_log.txt
│       └── Run_History.xlsx
│
└── output/
    └── Cisco_Partner_Intelligence_Report.xlsx
```

---

# Core Components

## collect_partner_inventory.py

Collects Cisco partner inventory matching the configured business filters.

---

## get_partner_details.py

Retrieves detailed information for every collected Cisco partner.

---

## detect_changes.py

Compares the latest snapshot with the previous snapshot and detects:

- New Partners
- Removed Partners
- Portfolio Changes
- Competency Changes
- Location Changes

---

## build_final_excel.py

Generates the final business report by:

- Transforming partner data
- Building an executive summary
- Including change history
- Applying professional workbook formatting
- Producing the final Excel deliverable

---

## config.py

Centralized configuration for the entire platform, including:

- File paths
- API settings
- Report settings
- Excel formatting
- Pipeline configuration
- Business rules

---

## logger.py

Provides consistent console and file logging across the entire platform.

---

# Generated Outputs

| File | Purpose |
|------|---------|
| partner_inventory.csv | Raw Cisco partner inventory |
| partner_details.csv | Latest partner snapshot |
| partner_details_previous.csv | Previous snapshot used for comparison |
| changes_log.csv | Change history between snapshots |
| failed_partners.csv | Partners that could not be processed |
| run_log.txt | Detailed execution log |
| Run_History.xlsx | Historical execution records |
| Cisco_Partner_Intelligence_Report.xlsx | Final business report |

---

# Technology Stack

- Python
- Requests
- Pandas
- OpenPyXL

---

# Refresh Frequency

Recommended: **Bi-Weekly**

The platform is designed for scheduled execution using Windows Task Scheduler, GitHub Actions, or any equivalent scheduling solution.

---

# Documentation

- INSTALLATION.md — Installation and setup guide
- README.md — Project overview and architecture

---

# Author

**Taksh Jain**

NetCom Learning
