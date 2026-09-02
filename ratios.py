import pandas as pd 
df = pd.read_csv('data/raw/apple.csv') 
df["current_Ratio"] = df['current_assets'] / df['current_liabilities']
df["net_profit_margin"] = df['net_income'] / df['revenue']
df["roa"]= df['net_income'] / df['total_assets']
df["roe"]= df['net_income'] / df['shareholders_equity']
print(df[["fiscal_year", "current_Ratio", "net_profit_margin", "roa", "roe"]])
