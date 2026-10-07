"""
04_build_excel.py
------------------
Builds the Sales Performance Dashboard workbook: a Raw Data sheet plus four
analysis sheets that compute everything with live formulas (SUMIFS/COUNTIFS/
AVERAGEIFS) referencing Raw Data, so the workbook recalculates if the data
changes. Charts are native Excel chart objects bound to those formula cells.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

df = pd.read_csv("sales_project/excel/raw_data_for_excel.csv", parse_dates=["Order Date"])
N = len(df)
LAST_ROW = N + 1  # header is row 1

wb = Workbook()

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="2F6F6B")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1C2321")
LABEL_FONT = Font(name=FONT, bold=True, size=10, color="5B655F")
VALUE_FONT = Font(name=FONT, bold=True, size=16, color="2F6F6B")
BODY_FONT = Font(name=FONT, size=10)
THIN = Border(bottom=Side(style="thin", color="DDE2DC"))

# ============================================================
# Sheet 1: Raw Data
# ============================================================
ws = wb.active
ws.title = "Raw Data"

for c, col in enumerate(df.columns, start=1):
    cell = ws.cell(row=1, column=c, value=col)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL

for r, row in enumerate(df.itertuples(index=False), start=2):
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = BODY_FONT
        col_name = df.columns[c - 1]
        if col_name == "Order Date":
            cell.number_format = "yyyy-mm-dd"
        elif col_name in ("Sales", "MSRP", "Est. Unit Cost", "Est. Profit", "Price Each"):
            cell.number_format = "$#,##0.00"
        elif col_name == "Est. Profit Margin":
            cell.number_format = "0.0%"

for c, col in enumerate(df.columns, start=1):
    width = max(12, min(22, df[col].astype(str).map(len).max() + 2))
    ws.column_dimensions[get_column_letter(c)].width = width

tab = Table(displayName="SalesData", ref=f"A1:R{LAST_ROW}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tab)
ws.freeze_panes = "A2"

# Column letter references used by every formula sheet below
COL = {name: get_column_letter(i + 1) for i, name in enumerate(df.columns)}
RNG = {name: f"'Raw Data'!${COL[name]}$2:${COL[name]}${LAST_ROW}" for name in df.columns}


def style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL


def kpi_card(ws, col, label, formula, number_format, row=3):
    lcell = ws.cell(row=row, column=col, value=label)
    lcell.font = LABEL_FONT
    vcell = ws.cell(row=row + 1, column=col, value=formula)
    vcell.font = VALUE_FONT
    vcell.number_format = number_format


# ============================================================
# Sheet 2: KPI Summary
# ============================================================
ws2 = wb.create_sheet("KPI Summary")
ws2["A1"] = "Sales Performance — KPI Summary"
ws2["A1"].font = TITLE_FONT
ws2["A2"] = "All formulas reference the Raw Data sheet and recalculate automatically"
ws2["A2"].font = Font(name=FONT, italic=True, size=9, color="5B655F")

kpi_card(ws2, 1, "Total Orders", f'=SUMPRODUCT(1/COUNTIF({RNG["Order Number"]},{RNG["Order Number"]}))', "#,##0")
kpi_card(ws2, 3, "Total Sales", f'=SUM({RNG["Sales"]})', "$#,##0")
kpi_card(ws2, 5, "Total Est. Profit", f'=SUM({RNG["Est. Profit"]})', "$#,##0")
kpi_card(ws2, 7, "Overall Margin", f'=SUM({RNG["Est. Profit"]})/SUM({RNG["Sales"]})', "0.0%")

kpi_card(ws2, 1, "Avg Order Value",
         f'=SUM({RNG["Sales"]})/SUMPRODUCT(1/COUNTIF({RNG["Order Number"]},{RNG["Order Number"]}))',
         "$#,##0", row=6)
kpi_card(ws2, 3, "Total Customers",
         f'=SUMPRODUCT((COUNTIF({RNG["Customer Name"]},{RNG["Customer Name"]})>0)/COUNTIF({RNG["Customer Name"]},{RNG["Customer Name"]}))',
         "#,##0", row=6)
kpi_card(ws2, 5, "Territories Served",
         f'=SUMPRODUCT((COUNTIF({RNG["Territory"]},{RNG["Territory"]})>0)/COUNTIF({RNG["Territory"]},{RNG["Territory"]}))',
         "#,##0", row=6)
kpi_card(ws2, 7, "Product Lines",
         f'=SUMPRODUCT((COUNTIF({RNG["Product Line"]},{RNG["Product Line"]})>0)/COUNTIF({RNG["Product Line"]},{RNG["Product Line"]}))',
         "#,##0", row=6)

for col in range(1, 9):
    ws2.column_dimensions[get_column_letter(col)].width = 16

# ============================================================
# Sheet 3: Region Analysis (Territory)
# ============================================================
ws3 = wb.create_sheet("Region Analysis")
ws3["A1"] = "Region-Wise Performance (Territory)"
ws3["A1"].font = TITLE_FONT

headers = ["Territory", "Orders", "Total Sales", "Total Profit", "Profit Margin %", "% of Total Profit"]
for c, h in enumerate(headers, start=1):
    ws3.cell(row=3, column=c, value=h)
style_header(ws3, 3, len(headers))

territories = sorted(df["Territory"].unique())
start_row = 4
for i, terr in enumerate(territories):
    r = start_row + i
    ws3.cell(row=r, column=1, value=terr).font = BODY_FONT
    ws3.cell(row=r, column=2,
             value=f'=SUMPRODUCT(({RNG["Territory"]}=A{r})/COUNTIFS({RNG["Order Number"]},{RNG["Order Number"]},{RNG["Territory"]},{RNG["Territory"]}))'
             ).number_format = "#,##0"
    ws3.cell(row=r, column=3, value=f'=SUMIFS({RNG["Sales"]},{RNG["Territory"]},A{r})').number_format = "$#,##0"
    ws3.cell(row=r, column=4, value=f'=SUMIFS({RNG["Est. Profit"]},{RNG["Territory"]},A{r})').number_format = "$#,##0"
    ws3.cell(row=r, column=5, value=f"=D{r}/C{r}").number_format = "0.0%"
    ws3.cell(row=r, column=6, value=f'=D{r}/SUM({RNG["Est. Profit"]})').number_format = "0.0%"
end_row = start_row + len(territories) - 1

chart3 = BarChart()
chart3.title = "Total Profit by Territory"
chart3.y_axis.title = "Profit ($)"
data = Reference(ws3, min_col=4, min_row=3, max_row=end_row)
cats = Reference(ws3, min_col=1, min_row=start_row, max_row=end_row)
chart3.add_data(data, titles_from_data=True)
chart3.set_categories(cats)
chart3.width, chart3.height = 16, 9
ws3.add_chart(chart3, f"A{end_row + 3}")

for col, w in zip("ABCDEF", [16, 10, 14, 14, 16, 16]):
    ws3.column_dimensions[col].width = w

# ============================================================
# Sheet 4: Product Line Analysis
# ============================================================
ws4 = wb.create_sheet("Product Line Analysis")
ws4["A1"] = "Product Line Performance"
ws4["A1"].font = TITLE_FONT

headers = ["Product Line", "Total Sales", "Total Profit", "Profit Margin %", "vs. Company Avg Margin"]
for c, h in enumerate(headers, start=1):
    ws4.cell(row=3, column=c, value=h)
style_header(ws4, 3, len(headers))

lines = df.groupby("Product Line")["Sales"].sum().sort_values(ascending=False).index.tolist()
start_row = 4
for i, pl in enumerate(lines):
    r = start_row + i
    ws4.cell(row=r, column=1, value=pl).font = BODY_FONT
    ws4.cell(row=r, column=2, value=f'=SUMIFS({RNG["Sales"]},{RNG["Product Line"]},A{r})').number_format = "$#,##0"
    ws4.cell(row=r, column=3, value=f'=SUMIFS({RNG["Est. Profit"]},{RNG["Product Line"]},A{r})').number_format = "$#,##0"
    ws4.cell(row=r, column=4, value=f"=C{r}/B{r}").number_format = "0.0%"
    ws4.cell(row=r, column=5,
              value=f'=D{r}-(SUM({RNG["Est. Profit"]})/SUM({RNG["Sales"]}))').number_format = "+0.0%;-0.0%"
end_row4 = start_row + len(lines) - 1

chart4 = BarChart()
chart4.type = "col"
chart4.title = "Profit Margin % by Product Line"
data = Reference(ws4, min_col=4, min_row=3, max_row=end_row4)
cats = Reference(ws4, min_col=1, min_row=start_row, max_row=end_row4)
chart4.add_data(data, titles_from_data=True)
chart4.set_categories(cats)
chart4.width, chart4.height = 16, 9
ws4.add_chart(chart4, f"A{end_row4 + 3}")

for col, w in zip("ABCDE", [18, 14, 14, 16, 20]):
    ws4.column_dimensions[col].width = w

# ============================================================
# Sheet 5: Monthly Trend
# ============================================================
ws5 = wb.create_sheet("Monthly Trend")
ws5["A1"] = "Monthly Sales Trend"
ws5["A1"].font = TITLE_FONT

headers = ["Month", "Total Sales", "Total Profit"]
for c, h in enumerate(headers, start=1):
    ws5.cell(row=3, column=c, value=h)
style_header(ws5, 3, len(headers))

months = sorted(df["Order Month"].unique())
start_row = 4
for i, m in enumerate(months):
    r = start_row + i
    ws5.cell(row=r, column=1, value=m).font = BODY_FONT
    ws5.cell(row=r, column=2, value=f'=SUMIFS({RNG["Sales"]},{RNG["Order Month"]},A{r})').number_format = "$#,##0"
    ws5.cell(row=r, column=3, value=f'=SUMIFS({RNG["Est. Profit"]},{RNG["Order Month"]},A{r})').number_format = "$#,##0"
end_row5 = start_row + len(months) - 1

chart5 = LineChart()
chart5.title = "Sales & Profit Trend by Month"
data = Reference(ws5, min_col=2, max_col=3, min_row=3, max_row=end_row5)
cats = Reference(ws5, min_col=1, min_row=start_row, max_row=end_row5)
chart5.add_data(data, titles_from_data=True)
chart5.set_categories(cats)
chart5.width, chart5.height = 22, 10
ws5.add_chart(chart5, f"E3")

for col, w in zip("ABC", [12, 14, 14]):
    ws5.column_dimensions[col].width = w

# ============================================================
# Sheet 6: Top Customers (pre-sorted in Python, values computed
# by live formulas — sort order isn't a volatile Excel function
# this runtime can evaluate, so it's fixed at build time; totals
# still recalculate if Raw Data changes)
# ============================================================
ws6 = wb.create_sheet("Top Customers")
ws6["A1"] = "Top 10 Customers by Total Sales"
ws6["A1"].font = TITLE_FONT
ws6["A2"] = "Row order fixed at build time; $ values are live SUMIFS formulas"
ws6["A2"].font = Font(name=FONT, italic=True, size=9, color="5B655F")

headers = ["Rank", "Customer", "Territory", "Total Sales", "Orders", "% of Total Revenue"]
for c, h in enumerate(headers, start=1):
    ws6.cell(row=4, column=c, value=h)
style_header(ws6, 4, len(headers))

top10 = (df.groupby(["Customer Name", "Territory"])["Sales"].sum()
         .sort_values(ascending=False).head(10).reset_index())
start_row = 5
for i, row in top10.iterrows():
    r = start_row + i
    ws6.cell(row=r, column=1, value=i + 1).font = BODY_FONT
    ws6.cell(row=r, column=2, value=row["Customer Name"]).font = BODY_FONT
    ws6.cell(row=r, column=3, value=row["Territory"]).font = BODY_FONT
    ws6.cell(row=r, column=4,
             value=f'=SUMIFS({RNG["Sales"]},{RNG["Customer Name"]},B{r},{RNG["Territory"]},C{r})').number_format = "$#,##0"
    ws6.cell(row=r, column=5,
             value=(f'=SUMPRODUCT(({RNG["Customer Name"]}=B{r})*({RNG["Territory"]}=C{r})'
                    f'/COUNTIFS({RNG["Order Number"]},{RNG["Order Number"]},'
                    f'{RNG["Customer Name"]},{RNG["Customer Name"]},{RNG["Territory"]},{RNG["Territory"]}))')
             ).number_format = "#,##0"
    ws6.cell(row=r, column=6, value=f'=D{r}/SUM({RNG["Sales"]})').number_format = "0.0%"

for col, w in zip("ABCDEF", [8, 28, 14, 14, 10, 18]):
    ws6.column_dimensions[col].width = w

wb.save("sales_project/excel/Sales_Performance_Dashboard.xlsx")
print("Saved workbook")
