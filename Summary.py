import pandas as pd
from openpyxl import load_workbook

company_name = "Microsoft"
ticker = "MSFT"
csv_path = f"data/raw/{ticker.lower()}.csv"
excel_path = f"data/processed/{ticker.lower()}_info.xlsx"

df = pd.read_csv(csv_path)

market_cap = df["stock_price"] * df["diluted_weighted_average_shares"]
free_cash_flow = df["operating_cash_flow"] - df["capital_expenditures"]

#All KPI s are calculated here
df["revenue_growth"] = df["revenue"].pct_change() * 100
df["net_margin"] = (df["net_income"] / df["revenue"]) * 100
df["roe"] = (df["net_income"] / df["shareholders_equity"]) * 100
df["fcf"] = free_cash_flow / 1000
df["debt_to_equity"] = df["total_liabilities"] / df["shareholders_equity"]
df["pe_ratio"] = df["stock_price"] / df["diluted_eps"]

kpi_table = df[["fiscal_year", "revenue_growth", "net_margin", "roe", "fcf", "debt_to_equity", "pe_ratio"]]
kpi_table = kpi_table.round(2)

kpi_table = kpi_table.set_index("fiscal_year").T

kpi_table.index = ["Revenue Growth", "Net Margin", "ROE", "Free Cash Flow ($B)", "Debt to Equity", "P/E Ratio"]

current_values = kpi_table[2025]
trend_direction = []

for metric in kpi_table.index:
    first_value = kpi_table.loc[metric].iloc[0]
    last_value = kpi_table.loc[metric].iloc[-1]
    if last_value > first_value:
        trend_direction.append("↑")
    elif last_value < first_value:
        trend_direction.append("↓")
    else:
        trend_direction.append("→")

kpi_table["Trend"] = trend_direction

for metric, value, direction in zip(kpi_table.index, current_values, trend_direction):
    print(f"{metric}: {value} {direction}")

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

sheet_name = f"{company_name} Summary"

with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    kpi_table.to_excel(writer, sheet_name=sheet_name)

wb = load_workbook(excel_path)
ws = wb[sheet_name]

for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column[1:]:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))


ws["A1"] = f"{company_name} Financial Summary"
ws.column_dimensions["A"].width = len(ws["A1"].value) + 6
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

table = Table(displayName=f"{ticker}Summary", ref=f"A1:F{ws.max_row}")
style = TableStyleInfo(name="TableStyleMedium9", 
                       showFirstColumn=False,
                       showLastColumn=False, 
                       showRowStripes=True, 
                       showColumnStripes=False)
table.tableStyleInfo = style
if f"{ticker}Summary" in ws.tables:
    del ws.tables[f"{ticker}Summary"]
ws.add_table(table)

percentage_rows = ["Revenue Growth", "Net Margin", "ROE"]

for row in ws.iter_rows(min_row=2):
    metric_name = row[0].value
    for cell in row[1:-1]:
        if metric_name in percentage_rows:
            cell.number_format = '0.00"%"'
        else:
            cell.number_format = "0.00"

#ALL KPIs are now calculated and saved to the excel file.
#Trend from year to year based on revenue is also saved to excel file
#Start of KPIs and company rating, start by pulling in current ratio 

df["current_ratio"] = df["current_assets"] / df["current_liabilities"]

current_ratio_table = df[["fiscal_year", "current_ratio"]].round(2)
print(current_ratio_table)
def score_percent(value, bands=(25, 20, 15, 10)):
    if value > bands[0]: return 10, f"> {bands[0]}%"
    if value > bands[1]: return 8, f"{bands[1]}-{bands[0]}%"
    if value > bands[2]: return 6, f"{bands[2]}-{bands[1]}%"
    if value > bands[3]: return 4, f"{bands[3]}-{bands[2]}%"
    return 2, f"< {bands[3]}%"

def score_ratio_low_is_good(value, bands=(0.5, 1, 2, 3)):
    if value < bands[0]: return 10, f"< {bands[0]}"
    if value < bands[1]: return 8, f"{bands[0]}-{bands[1]}"
    if value < bands[2]: return 6, f"{bands[1]}-{bands[2]}"
    if value < bands[3]: return 4, f"{bands[2]}-{bands[3]}"
    return 2, f"> {bands[3]}"

def score_ratio_high_is_good(value, bands=(2, 1.5, 1, 0.5)):
    if value > bands[0]: return 10, f"> {bands[0]}"
    if value > bands[1]: return 8, f"{bands[1]}-{bands[0]}"
    if value > bands[2]: return 6, f"{bands[2]}-{bands[1]}"
    if value > bands[3]: return 4, f"{bands[3]}-{bands[2]}"
    return 2, f"< {bands[3]}"

fcf_margin_2025 = (df["fcf"].iloc[-1] / (df["revenue"].iloc[-1] / 1000)) * 100

metrics_2025 = {
    "Profitability (Net Margin)": (df["net_margin"].iloc[-1], score_percent),
    "Profitability (ROE)": (df["roe"].iloc[-1], score_percent),
    "Cash Flow (FCF Margin)": (fcf_margin_2025, score_percent),
    "Growth (Revenue Growth)": (df["revenue_growth"].iloc[-1], lambda v: score_percent(v, bands=(10, 5, 0, -5))),
    "Financial Risk (Debt to Equity)": (df["debt_to_equity"].iloc[-1], score_ratio_low_is_good),
    "Valuation (P/E Ratio)": (df["pe_ratio"].iloc[-1], lambda v: score_ratio_low_is_good(v, bands=(15, 20, 30, 40))),
    "Liquidity (Current Ratio)": (df["current_ratio"].iloc[-1], score_ratio_high_is_good),
}

profitability_score = (metrics_2025["Profitability (Net Margin)"][1](metrics_2025["Profitability (Net Margin)"][0])[0] +
                        metrics_2025["Profitability (ROE)"][1](metrics_2025["Profitability (ROE)"][0])[0]) / 2

cash_flow_score = metrics_2025["Cash Flow (FCF Margin)"][1](metrics_2025["Cash Flow (FCF Margin)"][0])[0]
growth_score = metrics_2025["Growth (Revenue Growth)"][1](metrics_2025["Growth (Revenue Growth)"][0])[0]
risk_score = metrics_2025["Financial Risk (Debt to Equity)"][1](metrics_2025["Financial Risk (Debt to Equity)"][0])[0]
valuation_score = metrics_2025["Valuation (P/E Ratio)"][1](metrics_2025["Valuation (P/E Ratio)"][0])[0]
liquidity_score = metrics_2025["Liquidity (Current Ratio)"][1](metrics_2025["Liquidity (Current Ratio)"][0])[0]

overall_score = (
    profitability_score * 0.20 +
    cash_flow_score * 0.20 +
    growth_score * 0.20 +
    risk_score * 0.15 +
    valuation_score * 0.15 +
    liquidity_score * 0.10
)

print(f"Overall Score: {round(overall_score, 1)}/10")
scoring_criteria = {
    "Profitability (Net Margin)": "≥25%: 10 | 20-25%: 8 | 15-20%: 6 | 10-15%: 4 | <10%: 2",
    "Profitability (ROE)": "≥25%: 10 | 20-25%: 8 | 15-20%: 6 | 10-15%: 4 | <10%: 2",
    "Cash Flow (FCF Margin)": "≥25%: 10 | 20-25%: 8 | 15-20%: 6 | 10-15%: 4 | <10%: 2",
    "Growth (Revenue Growth)": "≥10%: 10 | 5-10%: 8 | 0-5%: 6 | -5-0%: 4 | <-5%: 2",
    "Financial Risk (Debt to Equity)": "<0.5: 10 | 0.5-1: 8 | 1-2: 6 | 2-3: 4 | >3: 2",
    "Valuation (P/E Ratio)": "<15: 10 | 15-20: 8 | 20-30: 6 | 30-40: 4 | >40: 2",
    "Liquidity (Current Ratio)": ">2: 10 | 1.5-2: 8 | 1-1.5: 6 | 0.5-1: 4 | <0.5: 2",
}

for label, (value, scorer) in metrics_2025.items():
    score, criteria = scorer(value)
    print(f"{label}: {round(value, 2)} → {score}/10")
    print(f"   Criteria: {scoring_criteria[label]}")

start_row = ws.max_row + 2

header_row = start_row
headers = ["Category", "2025 Value", "Score"]
for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=header_row, column=col, value=header)
    cell.font = Font(size=12, bold=True)
    cell.fill = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

row_num = header_row + 1
for label, (value, scorer) in metrics_2025.items():
    score, _ = scorer(value)

    ws.cell(row=row_num, column=1, value=label)
    ws.cell(row=row_num, column=2, value=round(value, 2))
    ws.cell(row=row_num, column=3, value=score)

    for col in range(1, 4):
        cell = ws.cell(row=row_num, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row_num += 1

ws.cell(row=row_num, column=1, value="Overall Score")
ws.cell(row=row_num, column=2, value="")
ws.cell(row=row_num, column=3, value=f"{round(overall_score, 1)}/10")

for col in range(1, 4):
    cell = ws.cell(row=row_num, column=col)
    cell.border = thin_border
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.font = Font(size=12, bold=True, color="1F4E78")

ws.column_dimensions["A"].width = 35
ws.column_dimensions["B"].width = 15


wb.save(excel_path)
print("saved")