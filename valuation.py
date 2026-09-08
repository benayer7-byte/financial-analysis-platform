import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

company_name = "Apple"
ticker = "AAPL"
csv_path = f"data/raw/{ticker.lower()}.csv"
excel_path = f"data/processed/{ticker.lower()}_info.xlsx"

df = pd.read_csv(csv_path)

market_cap = df["stock_price"] * df["diluted_weighted_average_shares"]
free_cash_flow = df["operating_cash_flow"] - df["capital_expenditures"]

df["pe_ratio"] = df["stock_price"] / df["diluted_eps"]
df["ps_ratio"] = market_cap / df["revenue"]
df["price_to_book"] = market_cap / df["shareholders_equity"]
df["fcf_yield"] = free_cash_flow / market_cap

valuation_table = df[["fiscal_year", "pe_ratio", "ps_ratio", "price_to_book", "fcf_yield"]]
valuation_table = valuation_table.set_index("fiscal_year").T

valuation_table.index = ["P/E Ratio", "P/S Ratio", "Price-to-Book", "FCF Yield"]
valuation_table = valuation_table.round(2)
valuation_table.loc["FCF Yield"] = valuation_table.loc["FCF Yield"] * 100

sheet_name = f"{company_name} Valuation"

with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    valuation_table.to_excel(writer, sheet_name=sheet_name)

wb = load_workbook(excel_path)
ws = wb[sheet_name]

for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column[1:]:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column_letter].width = max_length + 2


ws["A1"] = f"Valuation for {company_name} Inc."
ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width, len(ws["A1"].value) + 2)
ws["A1"].font = Font(size=14, bold=True, color="1F4E78")
ws["A1"].alignment = Alignment(horizontal="center")
ws["A1"].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

for cell in ws[1][1:]:
    cell.font = Font(size=12, bold=True)
for cell in ws["A"][1:]:
    cell.font = Font(size=12, bold=True)


thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))

for row in ws.iter_rows():
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

for cell in ws[1]:
    cell.value = str(cell.value)

table = Table(displayName=f"{ticker}Valuation", ref=f"A1:E{ws.max_row}")
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
if f"{ticker}Valuation" in ws.tables:
    del ws.tables[f"{ticker}Valuation"]
ws.add_table(table)

for row in ws.iter_rows(min_row=2):
    ratio_name = row[0].value
    for cell in row[1:]:
        if ratio_name == "FCF Yield":
            cell.number_format = '0.00"%"'
        else:
            cell.number_format = "0.00"

wb.save(excel_path)
print("saved")
print(valuation_table)