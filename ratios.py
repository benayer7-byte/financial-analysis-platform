import pandas as pd 
from openpyxl import load_workbook 
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

company_name = "Microsoft"
ticker = "MSFT"
csv_path = f"data/raw/{ticker.lower()}.csv"
excel_path = f"data/processed/{ticker.lower()}_info.xlsx"



df = pd.read_csv(csv_path) 

#All ratios are calculated using the following formulas:
df["current_Ratio"] = df['current_assets'] / df['current_liabilities']
df["net_profit_margin"] = df['net_income'] / df['revenue']
df["roa"]= df['net_income'] / df['total_assets']
df["roe"]= df['net_income'] / df['shareholders_equity']
df["debt_to_equity_ratio"] = df['total_liabilities'] / df['shareholders_equity']
df["operating_cash_flow_ratio"] = df['operating_cash_flow'] / df['current_liabilities']


#Ratio table 
ratio_table = df[["fiscal_year", 
                  "current_Ratio", 
                  "net_profit_margin", 
                  "roa", 
                  "roe", 
                  "debt_to_equity_ratio", 
                  "operating_cash_flow_ratio"]]


#flip the table so that fiscal_year is the column and ratios are the rows
ratio_table = ratio_table.set_index("fiscal_year").T

#rename index to be simpler and more readable
ratio_table.index = ["Current Ratio", 
                     "Net Profit Margin", 
                     "Return on Assets (ROA)", 
                     "Return on Equity (ROE)", 
                     "Debt to Equity Ratio", 
                     "Operating Cash Flow Ratio"]

#format numbers in table 
for ratio in ["Net Profit Margin", "Return on Assets (ROA)", "Return on Equity (ROE)"]:
    ratio_table.loc[ratio] = ratio_table.loc[ratio]*100
ratio_table = ratio_table.round(2)

ratio_table.to_excel(excel_path, sheet_name=f"{company_name} Ratios")

# Load the existing Excel file
wb = load_workbook(excel_path)
ws = wb.active

ws.title = f"{company_name} Ratios"
for column in ws.columns:

    max_length = 0
    column_letter = column[0].column_letter 
     # Get the column letter
    for cell in column: 
        if cell.value is not None: 
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column_letter].width = max_length + 2

#A1 adjustments and formatting 
ws["A1"] = f"{company_name} Financial Ratios"
ws["A1"].font = Font(size=14, bold=True, color="1F4E78")
ws["A1"].alignment = Alignment(horizontal="center")
ws["A1"].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

# Format year headers
for cell in ws[1][1:]:
    cell.font = Font(size=12, bold=True)

# Format ratio headers
for cell in ws["A"][1:]:
    cell.font = Font(size=12, bold=True)

# Define a thin border style
thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

# Apply border and alignment to all cells
for row in ws.iter_rows():
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

#Add excel table formatting 
for cell in ws[1]:
    cell.value = str(cell.value)  
    # Ensure the year headers are strings for the table


table = Table(displayName=f"{ticker}Ratios", 
              ref=f"A1:E{ws.max_row}") 
 # Adjust the range as needed

style = TableStyleInfo(name="TableStyleMedium9", 
                       showFirstColumn=False,
                       showLastColumn=False, 
                       showRowStripes=True, 
                       showColumnStripes=False)


table.tableStyleInfo = style 

if f"{ticker}Ratios" in ws.tables:
    del ws.tables[f"{ticker}Ratios"] 
    # Remove existing table if it exists to avoid duplication


ws.add_table(table) 

# Save the updated workbook

# Format percentage ratios
percentage_ratios = [
    "Net Profit Margin",
    "Return on Assets (ROA)",
    "Return on Equity (ROE)"
]

for row in ws.iter_rows(min_row=2):
    ratio_name = row[0].value

    for cell in row[1:]:
        if ratio_name in percentage_ratios:
            cell.number_format = '0.00"%"'
        else:
            cell.number_format = "0.00"

wb.save(excel_path)


print(ratio_table)

#Print all ratios if needed
#print(df[["fiscal_year", "current_Ratio", "net_profit_margin", "roa", "roe", "debt_to_equity_ratio", "operating_cash_flow_ratio"]])

