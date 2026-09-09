import pandas as pd
import sqlite3
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

db_path = os.path.join(project_root, "data", "financials.db")
raw_folder = os.path.join(project_root, "data", "raw")

tickers = ["aapl", "msft", "amzn", "nvda", "wmt"]

conn = sqlite3.connect(db_path)

for i, ticker in enumerate(tickers):
    csv_path = os.path.join(raw_folder, f"{ticker}.csv")
    df = pd.read_csv(csv_path)

    if i == 0:
        df.to_sql("financials", conn, if_exists="replace", index=False)
    else:
        df.to_sql("financials", conn, if_exists="append", index=False)

    print(f"Loaded {ticker}")

check = pd.read_sql("SELECT ticker, fiscal_year, revenue FROM financials", conn)
print(check)

conn.close()
print("\nDatabase built with all 5 companies.")