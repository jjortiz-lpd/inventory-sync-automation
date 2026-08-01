# Inventory Sync Automation

A Python automation project that processes sales orders, updates an Excel inventory, detects low-stock and out-of-stock products, and generates a PDF summary report.

## Business problem

Small businesses often maintain inventory and sales in separate spreadsheets. Updating stock manually is repetitive and prone to errors.

This project automates the complete workflow:

1. Reads the current inventory from `Inventory.xlsx`.
2. Reads completed sales from `Orders.xlsx`.
3. Groups sales by product.
4. Updates stock automatically.
5. Prevents negative inventory values.
6. Classifies each product as:
   - Available
   - Low Stock
   - Out of Stock
7. Generates `UpdatedInventory.xlsx`.
8. Generates `InventoryReport.pdf`.
9. Records execution details in `inventory_sync.log`.

## Technologies

- Python
- pandas
- openpyxl
- ReportLab
- Excel / XLSX
- Data validation and processing

## Installation

```bash
python -m venv .venv
```

### Windows

```bash
.venv\Scripts\activate
```

### macOS/Linux

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run the project

```bash
python main.py
```

## Expected output

After running the program, the following files are created:

- `UpdatedInventory.xlsx`
- `InventoryReport.pdf`
- `inventory_sync.log`

## Portfolio description

**Inventory Sync Automation**

Python automation that processes Excel sales records, updates product inventory, detects low-stock items, and generates formatted Excel and PDF reports. The solution includes input validation, error handling, logging, and reusable code.

## Suggested Upwork skills

- Python
- Automation
- Data Processing
- Microsoft Excel
- pandas
- openpyxl
- Report Generation
- Scripting
