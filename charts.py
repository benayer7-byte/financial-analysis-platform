import pandas as pd
from openpyxl import load_workbook
#from openpyxl.chart import LineChart, Reference

df = pd.read_csv("data/raw/apple.csv")

workbook = load_workbook("data/processed/apple_analysis.xlsx")

if "Charts" in workbook.sheetnames:
    worksheet = workbook["Charts"]
else:
    worksheet = workbook.create_sheet("Charts")

worksheet["N1"] = "Year"
worksheet["O1"] = "Revenue"
worksheet["P1"] = "Net Income"

for index, row in df.iterrows():
    worksheet.cell(row=index + 2, column=14, value=row["Year"])
    worksheet.cell(row=index + 2, column=15, value=row["Revenue"])
    worksheet.cell(row=index + 2, column=16, value=row["Net Income"])



#chart = LineChart()
#chart.title = "Revenue and Net Income over Time"
#chart.x_axis.title = "Year"
#chart.y_axis.title = "USD"

#chart.y_axis.majorGridlines = None


#chart.x_axis.majorTickMark = "out"
#chart.y_axis.majorTickMark = "out"

#years = Reference(
#    worksheet,
#    min_col=14,
#    min_row=2,
#    max_row=len(df) + 1
#)

#revenue_and_income = Reference(
#    worksheet,
#    min_col=15,
#    max_col=16,
#    min_row=2,
#    max_row=len(df) + 1
#)

#chart.add_data(revenue_and_income, titles_from_data=True)
#chart.set_categories(years)

#for series in chart.series:
#    series.marker.symbol = "circle"
#    series.marker.size = 7

#chart.title = "Revenue and Net Income over Time"
#chart.width = 20
#chart.height = 10

#worksheet.add_chart(chart, "A1")

print("Charts created and saved to 'data/processed/apple_analysis.xlsx'.")
workbook.save("data/processed/apple_analysis.xlsx")