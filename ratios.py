import pandas as pd 
df = pd.read_csv('data/raw/apple.csv') 

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

ratio_table.to_excel("data/processed/apple_ratios.xlsx", index=True)


print(ratio_table)

#Print all ratios if needed
#print(df[["fiscal_year", "current_Ratio", "net_profit_margin", "roa", "roe", "debt_to_equity_ratio", "operating_cash_flow_ratio"]])

