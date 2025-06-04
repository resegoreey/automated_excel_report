# automated_excel_report
# Sales Report Generator

This Python script reads sales data from a CSV file and generates a styled Excel report (`.xlsx`) with the data.

## Features

- Reads sales data from a CSV file (`data/sales_data.csv`).
- Creates a new Excel workbook with one worksheet named **Sales Report**.
- Writes the CSV data into the Excel sheet with headers.
- Applies basic styling to the header row (bold font, background color, borders, and centered text).
- Automatically adjusts column widths to fit the content.
- Saves the report to the `reports` folder as `sales_reports.xlsx`.

## Requirements

- Python 3.6 or higher
- `pandas` library
- `openpyxl` library

Install the required libraries using:

```bash
pip install pandas openpyxl
