from pathlib import Path
import pandas as pd


def load_financial_data(file_path):
    df = pd.read_csv(file_path)
    return df.sort_values("Year")


# Ensure the export folder exists.
Path("data/processed").mkdir(parents=True, exist_ok=True)

# Load and sort the CSV data.
apple_data = load_financial_data("data/raw/apple.csv")

# Display in the terminal.
print(
    apple_data.to_string(
        index=False,
        formatters={
            "Revenue": "${:,.0f}".format,
            "Net Income": "${:,.0f}".format,
            "Total Assets": "${:,.0f}".format,
            "Total Liabilities": "${:,.0f}".format,
        },
    )
)

# Write the data to Excel.
apple_data.to_excel(
    "data/processed/apple_processed.xlsx",
    sheet_name="Apple Financials",
    index=False,
)
from pathlib import Path
import pandas as pd


def load_financial_data(file_path):
    df = pd.read_csv(file_path)
    return df.sort_values("Year").reset_index(drop=True)


Path("data/processed").mkdir(parents=True, exist_ok=True)

apple_data = load_financial_data("data/raw/apple.csv")

output_file = "data/processed/apple_processed.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    apple_data.to_excel(
        writer,
        sheet_name="Apple Financials",
        index=False,
    )

    workbook = writer.book
    worksheet = writer.sheets["Apple Financials"]

    currency_format = workbook.add_format({
        "num_format": "$#,##0",
    })

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
    })

    for column_num, column_name in enumerate(apple_data.columns):
        worksheet.write(0, column_num, column_name, header_format)

    worksheet.set_column("A:A", 12)

    worksheet.set_column("B:E", 20, currency_format)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(apple_data), len(apple_data.columns) - 1)