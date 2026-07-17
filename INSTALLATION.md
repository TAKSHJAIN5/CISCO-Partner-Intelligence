# Installation Guide

## Cisco Partner Intelligence Automation Platform

---

# Prerequisites

Before installing, ensure the following are available:

- Python 3.11 or later
- Internet connection
- Access to the Cisco Partner API
- Microsoft Excel (for viewing generated reports)

---

# Installation

## Step 1 — Extract the Project

Extract the project to a local directory.

Example:

```text
C:\CiscoPartnerAutomation
```

---

## Step 2 — Open PowerShell or Command Prompt

Navigate to the project directory.

```powershell
cd C:\CiscoPartnerAutomation
```

---

## Step 3 — Create a Virtual Environment

```powershell
python -m venv venv
```

---

## Step 4 — Activate the Virtual Environment

Windows PowerShell

```powershell
venv\Scripts\activate
```

---

## Step 5 — Install Dependencies

```powershell
pip install -r requirements.txt
```

---

# Configuration

All application settings are centralized in:

```text
src/config.py
```

This file contains:

- File paths
- API configuration
- Report settings
- Excel formatting options
- Pipeline configuration
- Logging settings

Modify configuration values here before running the platform if required.

---

# Running the Platform

From the project root directory:

```powershell
python run_all.py
```

The platform will automatically execute the complete pipeline.

```text
collect_partner_inventory.py

↓

get_partner_details.py

↓

detect_changes.py

↓

build_final_excel.py
```

---

# Generated Files

After a successful execution the following files are created or updated.

| File | Description |
|------|-------------|
| data/partner_inventory.csv | Cisco partner inventory |
| data/partner_details.csv | Latest detailed partner dataset |
| data/partner_details_previous.csv | Previous snapshot for change detection |
| data/logs/changes_log.csv | Partner changes between executions |
| data/logs/failed_partners.csv | Failed partner processing records |
| data/logs/run_log.txt | Detailed execution log |
| data/logs/Run_History.xlsx | Historical execution log |
| output/Cisco_Partner_Intelligence_Report.xlsx | Final business report |

---

# Scheduled Execution

The platform is designed for unattended execution.

It can be scheduled using:

- Windows Task Scheduler
- GitHub Actions
- Cron (Linux)
- Any enterprise scheduling platform

Each execution automatically:

- Refreshes partner inventory
- Updates partner details
- Detects business changes
- Generates a new Excel report
- Records execution history

---

# Troubleshooting

## Dependencies Missing

Reinstall project dependencies.

```powershell
pip install -r requirements.txt
```

---

## Virtual Environment Not Active

Activate the virtual environment before running the platform.

```powershell
venv\Scripts\activate
```

---

## Excel Workbook Is Open

If the generated workbook is open during execution, close it and rerun the platform.

```powershell
python run_all.py
```

---

## Cisco API Connectivity

Verify:

- Internet connectivity
- Cisco Partner API availability
- Configuration values in `src/config.py`

---

# Support

For enhancements, maintenance, or configuration changes, refer to:

- README.md
- Source code in the `src` directory
- Configuration in `src/config.py`

---

# Author

**Taksh Jain**

NetCom Learning
