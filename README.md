# Cisco Partner Automation

## Overview

Cisco Partner Automation is an internal NetCom Learning solution that automates the collection, enrichment, validation, change tracking, and reporting of Cisco Partner Locator data.

The solution replaces a manual spreadsheet maintenance process by automatically extracting Cisco partner information, applying business logic, tracking changes between refreshes, and generating a formatted Excel workbook for business users.

---

## Business Objective

* Eliminate manual Cisco partner data collection.
* Maintain an up-to-date partner intelligence repository.
* Reduce stale partner information.
* Improve visibility into partner portfolio and competency changes.
* Provide a business-ready Excel workbook for partner management and marketing teams.

---

## Refresh Frequency

Bi-Weekly

---

## Project Structure

```text
CiscoPartnerAutomation/

run_all.py

src/
│
├── collect_partner_inventory.py
├── get_partner_details.py
├── detect_changes.py
├── build_final_excel.py
├── format_excel.py
├── config.py
└── logger.py

data/
│
├── partner_inventory.csv
├── partner_details.csv
├── partner_details_previous.csv
│
└── logs/
    ├── changes_log.csv
    └── failed_partners.csv

output/
│
└── Cisco_Partner_Master.xlsx

requirements.txt
README.md
```

---

## First-Time Setup (One-Time Only)

### Step 1 – Open PowerShell

```powershell
cd C:\Users\Dell\Desktop\CiscoPartnerAutomation
```

### Step 2 – Create Virtual Environment

```powershell
python -m venv venv
```

### Step 3 – Activate Virtual Environment

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& .\venv\Scripts\Activate.ps1)
```

### Step 4 – Install Dependencies

```powershell
pip install -r requirements.txt
```

---

## Refresh Process (Every 2 Weeks)

### Step 1 – Open PowerShell

```powershell
cd C:\Users\Dell\Desktop\CiscoPartnerAutomation
```

### Step 2 – Activate Virtual Environment

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& .\venv\Scripts\Activate.ps1)
```

### Step 3 – Run Full Refresh

```powershell
python run_all.py
```

---

## What Happens During Refresh

### 1. Partner Inventory Collection

Collects all Cisco partners matching the configured portfolio and competency filters.

Output:

```text
data/partner_inventory.csv
```

---

### 2. Partner Detail Collection

Retrieves detailed information for each partner.

Output:

```text
data/partner_details.csv
```

Before creating a new file, the previous version is automatically backed up as:

```text
data/partner_details_previous.csv
```

---

### 3. Change Detection

Compares:

```text
partner_details_previous.csv
```

vs

```text
partner_details.csv
```

and identifies:

* New Partners
* Removed Partners
* Portfolio Changes
* Competency Changes
* Location Changes

Output:

```text
data/logs/changes_log.csv
```

---

### 4. Excel Generation

Transforms partner data into the business reporting structure.

Output:

```text
output/Cisco_Partner_Master.xlsx
```

---

### 5. Excel Formatting

Applies:

* Professional header formatting
* Frozen header row
* Frozen Name column
* Filters
* Alternate row shading
* Auto-sized columns
* Business-friendly layout

---

## Final Deliverable

```text
output/Cisco_Partner_Master.xlsx
```

Columns:

* Name
* Preferred Portfolios
* Non Preferred Portfolios
* Sec Net Spez
* Sec AI Infra Spez
* Preferred NW
* Preferred SEC
* Preferred Cloud + AI
* Competencies
* Location

---

## Dependencies

Install using:

```powershell
pip install -r requirements.txt
```

Packages:

```text
requests
pandas
openpyxl
```

---

## Troubleshooting

### Excel File Open Error

If the workbook is open while formatting runs, close:

```text
Cisco_Partner_Master.xlsx
```

and rerun:

```powershell
python run_all.py
```

---

### Missing Python Packages

```powershell
pip install -r requirements.txt
```

---

### Virtual Environment Not Active

```powershell
(Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned) ; (& .\venv\Scripts\Activate.ps1)
```

---

## Owner

Taksh Jain

NetCom Learning
