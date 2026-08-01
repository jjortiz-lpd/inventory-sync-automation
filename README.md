# Inventory Sync Automation

## Overview

A Python automation project that synchronizes inventory using Excel workbooks.

The application reads inventory and order data, updates stock automatically, generates an updated Excel report, creates a PDF summary, and logs the execution.

---

## Features

- Read inventory from Excel
- Process customer orders
- Update stock quantities
- Detect low inventory
- Generate updated Excel workbook
- Generate PDF report
- Logging
- Error handling

---

## Tech Stack

- Python
- Pandas
- OpenPyXL
- ReportLab

---

## Project Structure

```
Inventory.xlsx
Orders.xlsx
main.py
requirements.txt
README.md
UpdatedInventory.xlsx
InventoryReport.pdf
```

---

## Example Workflow

Inventory.xlsx
↓

Orders.xlsx

↓

Python Automation

↓

UpdatedInventory.xlsx

↓

InventoryReport.pdf

---

## Future Improvements

- SQL Database Integration
- REST API
- Email Notifications
- Dashboard
- Google Sheets Sync
