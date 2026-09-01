import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


df = pd.read_csv("data/raw/apple.csv")



df["Current Ratio"] = df["Current Assets"] / df["Current Liabilities"]
df["Long-Term Debt Ratio"] = df["Long-Term Debt"] / df["Total Assets"]
df["Profit Margin"] = df["Net Income"] / df["Revenue"]
df["ROA"] = df["Net Income"] / df["Total Assets"]
df["Debt Ratio"] = df["Total Liabilities"] / df["Total Assets"]

print(
    df[
        [
            "Year",
            "Current Ratio",
            "Long-Term Debt Ratio",
            "Profit Margin",
            "ROA",
            "Debt Ratio"
        ]
    ].to_string(
        index=False,
        formatters={
            "Current Ratio": "{:.2f}".format,
            "Long-Term Debt Ratio": "{:.2%}".format,
            "Profit Margin": "{:.2%}".format,
            "ROA": "{:.2%}".format,
            "Debt Ratio": "{:.2%}".format
        }
    )
)

ratio_table = df[
    [
        "Year",
        "Current Ratio",
        "Long-Term Debt Ratio",
        "Profit Margin",
        "ROA",
        "Debt Ratio"
    ]
]
ratio_table.to_excel(
    "data/processed/apple_analysis.xlsx",
    sheet_name="Ratios",
    index=False
)
file_path = "data/processed/apple_analysis.xlsx"

workbook = load_workbook(file_path)
worksheet = workbook["Ratios"]

worksheet.freeze_panes = "A2"

worksheet.column_dimensions["A"].width = 10
worksheet.column_dimensions["B"].width = 16
worksheet.column_dimensions["C"].width = 22
worksheet.column_dimensions["D"].width = 16
worksheet.column_dimensions["E"].width = 12
worksheet.column_dimensions["F"].width = 14

for cell in worksheet["B"][1:]:
    cell.number_format = "0.00"
for column in ["C", "D", "E", "F"]:
    for cell in worksheet[column][1:]:
        cell.number_format = "0.00%"


table = Table(displayName="RatiosTable", 
              ref=f"A1:F{worksheet.max_row}")


style = TableStyleInfo(name = "TableStyleMedium9", 
    showFirstColumn = False,
    showLastColumn = False, 
    showRowStripes = True, 
    showColumnStripes = False)


table.tableStyleInfo = style 
worksheet.add_table(table)
workbook.save(file_path)
