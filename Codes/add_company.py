import pandas as pd
import sqlite3
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

db_path = os.path.join(project_root, "data", "financials.db")

new_ticker = input("Enter the new company's ticker (e.g. TSLA), or press Enter to cancel: ").upper()

if new_ticker == "":
    print("Cancelled.")
    exit()

csv_path = input("Enter the full path to the new company's CSV: ")

import shutil

df = pd.read_csv(csv_path)

raw_folder = os.path.join(project_root, "data", "raw")
destination = os.path.join(raw_folder, f"{new_ticker.lower()}.csv")
shutil.copy(csv_path, destination)
print(f"Copied CSV to {destination}")

conn = sqlite3.connect(db_path)

existing = pd.read_sql(f"SELECT * FROM financials WHERE ticker = '{new_ticker}'", conn)
if len(existing) > 0:
    print(f"\n{new_ticker} already exists in the database. Nothing was added.")
else:
    df.to_sql("financials", conn, if_exists="append", index=False)
    print(f"\n{new_ticker} added successfully.")

conn.close()