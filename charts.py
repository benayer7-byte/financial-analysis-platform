import pandas as pd 

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, LineChart 

df = pd.read_csv('data/raw/apple.csv') 
wb = load_workbook("data/processed/apple_info.xlsx")


if "Apple Charts" in wb.sheetnames:
    ws = wb["Apple Charts"]
else:
    ws = wb.create_sheet("Apple Charts")

ws["A1"] = "Fiscal Year"
ws["B1"] = "Revenue"

for row, (_, data) in enumerate(df.iterrows(), start=2):
    ws[f"A{row}"] = data["fiscal_year"]
    ws[f"B{row}"] = data["revenue"] / 1000


wb.save("data/processed/apple_info.xlsx")
print("Charts created and saved to data/processed/apple_info.xlsx")