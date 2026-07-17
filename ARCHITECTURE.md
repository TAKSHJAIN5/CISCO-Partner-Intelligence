# Architecture

## Cisco Partner Intelligence Automation Platform

---

# Overview

The Cisco Partner Intelligence Automation Platform is an automated ETL (Extract, Transform, Load) and reporting solution that collects Cisco Partner Locator data, enriches partner records, detects changes between executions, and produces a business-ready Excel report.

The platform is designed for unattended execution and follows a modular, configuration-driven architecture.

---

# High-Level Workflow

```text
                Cisco Partner Locator API
                          │
                          ▼
        collect_partner_inventory.py
                          │
                          ▼
          partner_inventory.csv
                          │
                          ▼
          get_partner_details.py
                          │
                          ▼
           partner_details.csv
                          │
                          ▼
           detect_changes.py
                          │
                          ▼
             changes_log.csv
                          │
                          ▼
          build_final_excel.py
                          │
                          ▼
Cisco_Partner_Intelligence_Report.xlsx
```

---

# Execution Flow

The platform is orchestrated by:

```text
run_all.py
```

Execution order:

1. Collect partner inventory
2. Retrieve detailed partner information
3. Detect changes from the previous snapshot
4. Generate the final business report

Every stage must complete successfully before the next stage begins.

If any stage fails, execution stops immediately.

---

# Project Structure

```text
CiscoPartnerAutomation/

│
├── run_all.py
├── requirements.txt
├── README.md
├── INSTALLATION.md
├── ARCHITECTURE.md
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

# Component Responsibilities

## run_all.py

Coordinates the complete execution pipeline.

Responsibilities:

- Execute scripts in order
- Stop execution on failure
- Record execution history
- Log execution progress

---

## collect_partner_inventory.py

Extracts Cisco partner inventory using the configured search filters.

Output:

```text
partner_inventory.csv
```

---

## get_partner_details.py

Retrieves detailed information for every discovered partner.

Responsibilities:

- API communication
- Validation
- Error handling
- Snapshot management

Output:

```text
partner_details.csv
```

---

## detect_changes.py

Compares the latest partner snapshot against the previous execution.

Tracks:

- New Partners
- Removed Partners
- Portfolio Changes
- Competency Changes
- Location Changes

Output:

```text
changes_log.csv
```

---

## build_final_excel.py

Presentation layer of the platform.

Responsibilities:

- Transform business data
- Build executive summary
- Generate partner report
- Include change history
- Apply workbook formatting
- Produce the final deliverable

Output:

```text
Cisco_Partner_Intelligence_Report.xlsx
```

---

## config.py

Centralized configuration for the platform.

Contains:

- File paths
- API configuration
- Business rules
- Report settings
- Excel formatting settings
- Pipeline configuration
- Logging configuration

No business constants should be duplicated elsewhere.

---

## logger.py

Provides consistent logging across the entire platform.

Responsibilities:

- Console logging
- Run log generation
- Timestamp formatting
- Consistent message formatting

---

# Data Flow

```text
Cisco API

↓

partner_inventory.csv

↓

partner_details.csv

↓

partner_details_previous.csv

↓

changes_log.csv

↓

Cisco_Partner_Intelligence_Report.xlsx
```

Every stage consumes the output of the previous stage.

No script performs unrelated responsibilities.

---

# Design Principles

The platform follows the following principles:

- Single Responsibility Principle
- Modular architecture
- Configuration-driven design
- Fail-fast execution
- Deterministic processing
- Reproducible outputs
- Clear separation of extraction, transformation, comparison, and presentation

---

# Error Handling

Each script:

- validates its inputs
- validates required files
- validates required columns
- logs meaningful errors
- terminates on unrecoverable failures

The pipeline stops immediately if a critical stage fails.

---

# Logging

Execution information is written to:

```text
data/logs/run_log.txt
```

Execution history is recorded in:

```text
data/logs/Run_History.xlsx
```

This provides an audit trail of all executions.

---

# Change Detection Strategy

The platform uses snapshot comparison.

Current execution:

```text
partner_details.csv
```

Previous execution:

```text
partner_details_previous.csv
```

The comparison identifies only business-relevant changes defined in the configuration.

---

# Configuration Philosophy

All reusable values are centralized in:

```text
src/config.py
```

This includes:

- Paths
- Report settings
- Business constants
- Excel formatting
- API settings
- Change tracking rules

Scripts should never hardcode reusable values.

---

# Final Deliverable

The platform produces a single business-ready Excel workbook.

```text
output/Cisco_Partner_Intelligence_Report.xlsx
```

The workbook includes:

- Executive Summary
- Partners
- Changes
- Logs

It is intended for marketing, partner management, and leadership teams.

---

# Future Enhancements

The modular architecture allows additional capabilities without changing the existing pipeline.

Examples include:

- Automated email distribution
- SharePoint or OneDrive publishing
- Microsoft Teams notifications
- Dashboard integration
- Scheduled cloud execution
- Additional business analytics
- Historical trend reporting

These enhancements can be added as separate modules while preserving the existing architecture.
