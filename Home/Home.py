import subprocess

import os

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
codes_dir = os.path.join(project_root, "Codes")

companies = {
    "1": ("Apple", os.path.join(codes_dir, "apple_master.py")),
    "2": ("Microsoft", os.path.join(codes_dir, "microsoft_master.py")),
    "3":("Amazon", os.path.join(codes_dir, "amazon_master.py")), 
    "4": ("Walmart", os.path.join(codes_dir, "walmart_master.py")),
    "5": ("Nvidia", os.path.join(codes_dir, "nvidia_master.py")) 

}

print("=" * 60)
print("FINANCIAL ANALYSIS PLATFORM")
print("=" * 60)
print("""
This tool builds a full financial analysis workbook for a
chosen company, using its 2022-2025 annual financial data.

Each workbook includes:
  - Ratios: liquidity, profitability, and leverage ratios
  - Charts: Revenue, Net Income, EPS, and Stock Price trends
  - Valuation: P/E, P/S, Price-to-Book, and FCF Yield
  - Forecasting: 2026-2027 revenue projections (two methods)
  - Summary: a KPI dashboard plus an overall 0-10 score

The overall score is a weighted average across 6 categories:
  Profitability 20% | Cash Flow 20% | Growth 20%
  Financial Risk 15% | Valuation 15% | Liquidity 10%

Each category is scored 0-10 using fixed thresholds set by
the author (general finance heuristics, not industry-specific
benchmarks) - e.g. Net Margin above 25% scores a 10, below
10% scores a 2.

Feel Free to edit tables and charts as you please and adjust formatting in excel. 


"""

)
print("=" * 60)
print()

print("Additional tools (run separately from the Codes folder):")
print("  - add_company.py : add a new company's data to the database")
print("  - query.py       : write your own SQL queries to explore the data")
print()

print("Which company would you like to build?")
for key, (name, _) in companies.items():
    print(f"{key}. {name}")

choice = input("Enter a number: ")

if choice in companies:
    name, script = companies[choice]
    print(f"\nRunning {name}...\n")
    subprocess.run(["python", script])

    ticker_map = {"Apple": "aapl", "Microsoft": "msft", "Amazon":"amzn", "Walmart":"wmt", "Nvidia":"nvda"}
    ticker = ticker_map[name]
    excel_file = os.path.join(project_root, "data", "processed", f"{ticker}_info.xlsx")

    print(f"\nOpening {excel_file}...")
    os.startfile(excel_file)
else:
    print("Invalid choice. Please run again and enter a valid number.")