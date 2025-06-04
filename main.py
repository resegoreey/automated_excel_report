import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Side
from pathlib import Path

DATA_FILE = Path("data/sales_data.csv")
REPORT_File = Path("reports/sales_reports.xlsx")

#this reads from the CSV file
try:
    data = pd.read_csv(DATA_FILE)
except FileNotFoundError:
    print(f"Error: File not found at {DATA_FILE}")


#this will create the Workbook
wb = Workbook()
ws = wb.active
ws.title = "Sales Report"