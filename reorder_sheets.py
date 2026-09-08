from openpyxl import load_workbook

company_name = "Apple"
ticker = "AAPL"
excel_path = f"data/processed/{ticker.lower()}_info.xlsx"

wb = load_workbook(excel_path)

desired_order = [
    f"{company_name} Summary",
    f"{company_name} Charts",
    f"{company_name} Forecasting",
    f"{company_name} Valuation",
    f"{company_name} Ratios",
]

wb._sheets = [wb[name] for name in desired_order]

wb.save(excel_path)
print("sheets reordered")