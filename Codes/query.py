import pandas as pd
import sqlite3
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
db_path = os.path.join(project_root, "data", "financials.db")

print("=" * 60)
print("FINANCIALS DATABASE - QUERY TOOL")
print("=" * 60)
print("""
Table name: financials
Columns: ticker, company_name, fiscal_year, revenue, net_income,
         and the rest of your standard 16 fields.

Example queries:
  SELECT * FROM financials WHERE ticker = 'AAPL';
  SELECT ticker, revenue FROM financials WHERE fiscal_year = 2025;
  SELECT ticker, AVG(net_income) FROM financials GROUP BY ticker;

Type 'exit' to quit.
""")

conn = sqlite3.connect(db_path)

while True:
    query = input("SQL> ")
    if query.lower() == "exit":
        break

    try:
        result = pd.read_sql(query, conn)
        print(result)
    except Exception as e:
        print(f"Error: {e}")

    print()

conn.close()
print("Goodbye.")