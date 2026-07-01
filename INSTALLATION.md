# Installation Guide

## Cisco Partner Intelligence Automation Platform

### Prerequisites

* Python 3.11 or later
* Internet connection
* Access to the Cisco Partner API
* Microsoft Excel (to view the generated workbook)

---

## Step 1 – Extract the Project

Extract the project folder to a local directory, for example:

```text
C:\CiscoPartnerAutomation
```

---

## Step 2 – Open Command Prompt or PowerShell

Navigate to the project folder:

```powershell
cd C:\CiscoPartnerAutomation
```

---

## Step 3 – Create a Virtual Environment

```powershell
python -m venv venv
```

---

## Step 4 – Activate the Virtual Environment

Windows:

```powershell
venv\Scripts\activate
```

---

## Step 5 – Install Dependencies

```powershell
pip install -r requirements.txt
```

---

## Step 6 – Run the Automation

If `run_all.py` is in the project root:

```powershell
python run_all.py
```

If `run_all.py` is inside the `src` folder:

```powershell
python src\run_all.py
```

---

## Generated Files

After a successful execution, the project automatically creates and updates:

* `data/partner_inventory.csv`
* `data/partner_details.csv`
* `data/partner_details_previous.csv`
* `data/logs/Run_History.xlsx`
* `data/logs/run_log.txt`
* `data/logs/changes_log.csv`
* `data/logs/failed_partners.csv`
* `output/Cisco_Partner_Master.xlsx`

---

## Scheduled Execution

This project is intended to be executed on a bi-weekly schedule using Windows Task Scheduler.

The generated Excel workbook can be stored in a SharePoint document library or a shared OneDrive folder so the latest version is available to all users.

---

## Support

For any enhancements or maintenance, refer to the project documentation and source code in the `src` folder.
