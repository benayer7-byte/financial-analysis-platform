import pandas as pd

df = pd.read_csv("data/raw/apple.csv")

print(df.to_string(
    index = False, 
    formatters={ 
        "Revenue": "${:,.2f}".format,
        "Net Income": "${:,.2f}".format,
        "Total Assets": "${:,.2f}".format,
        "Total Liabilities": "${:,.2f}".format,
    }))

