import csv
from pathlib import Path
import xlsxwriter


ROOT = Path(__file__).parent
SOURCE = ROOT / "data" / "customer_churn_cleaned.csv"
OUTPUT = ROOT / "customer_churn_kpi_dashboard.xlsx"


def tenure_band(months: int) -> str:
    if months <= 12:
        return "0-12 months"
    if months <= 24:
        return "13-24 months"
    return "25+ months"


with SOURCE.open(newline="", encoding="utf-8") as source_file:
    records = list(csv.DictReader(source_file))

for record in records:
    record["Age"] = int(record["Age"])
    record["TenureMonths"] = int(record["TenureMonths"])
    record["MonthlyCharges"] = float(record["MonthlyCharges"])
    record["TotalCharges"] = float(record["TotalCharges"])
    record["SupportTickets"] = int(record["SupportTickets"])
    record["TenureBand"] = tenure_band(record["TenureMonths"])

contracts = ["Month-to-Month", "One Year", "Two Year"]
tenure_bands = ["0-12 months", "13-24 months", "25+ months"]
payments = sorted({record["PaymentMethod"] for record in records})

workbook = xlsxwriter.Workbook(OUTPUT)
workbook.set_properties({
    "title": "Customer Churn KPI Dashboard",
    "subject": "Business KPI Dashboard – customer churn",
    "author": "Data analysis project",
    "comments": "Metrics calculated from the supplied 15-record cleaned customer churn dataset.",
})
workbook.set_calc_mode("auto")

# Palette: high contrast navy surface, teal positive, amber risk, and coral churn.
navy = "#13243A"
blue = "#1F5FAA"
teal = "#0F766E"
coral = "#C2414C"
amber = "#B45309"
ink = "#1F2937"
muted = "#52606D"
line = "#D8E0E8"
surface = "#F6F8FB"
white = "#FFFFFF"

title = workbook.add_format({"bold": True, "font_size": 22, "font_color": white, "bg_color": navy, "valign": "vcenter"})
subtitle = workbook.add_format({"font_size": 10, "font_color": "#D6E4F0", "bg_color": navy, "valign": "vcenter"})
section = workbook.add_format({"bold": True, "font_size": 11, "font_color": navy, "bottom": 1, "bottom_color": blue})
label = workbook.add_format({"font_size": 9, "font_color": muted, "bold": True})
note = workbook.add_format({"font_size": 9, "font_color": muted, "text_wrap": True, "valign": "top"})
filter_label = workbook.add_format({"bold": True, "font_size": 9, "font_color": navy, "bg_color": "#E8F0F8", "align": "center", "valign": "vcenter"})
filter_value = workbook.add_format({"font_size": 10, "font_color": ink, "bg_color": white, "border": 1, "border_color": "#9FB3C8", "align": "center", "valign": "vcenter"})
card_label = workbook.add_format({"bold": True, "font_size": 9, "font_color": muted, "bg_color": white, "align": "center", "valign": "vcenter"})
card_number = workbook.add_format({"bold": True, "font_size": 20, "font_color": navy, "bg_color": white, "align": "center", "valign": "vcenter", "num_format": "0"})
card_percent = workbook.add_format({"bold": True, "font_size": 20, "font_color": navy, "bg_color": white, "align": "center", "valign": "vcenter", "num_format": "0.0%"})
card_currency = workbook.add_format({"bold": True, "font_size": 20, "font_color": amber, "bg_color": white, "align": "center", "valign": "vcenter", "num_format": "$#,##0.00"})
card_decimal = workbook.add_format({"bold": True, "font_size": 20, "font_color": navy, "bg_color": white, "align": "center", "valign": "vcenter", "num_format": "0.0"})
definition_head = workbook.add_format({"bold": True, "font_color": white, "bg_color": navy, "border": 1, "border_color": navy, "align": "left"})
definition = workbook.add_format({"font_size": 9, "font_color": ink, "text_wrap": True, "valign": "top", "border": 1, "border_color": line})
data_header = workbook.add_format({"bold": True, "font_color": white, "bg_color": navy, "border": 1, "border_color": navy})
currency = workbook.add_format({"num_format": "$#,##0.00"})
integer = workbook.add_format({"num_format": "0"})
small_header = workbook.add_format({"bold": True, "font_size": 9, "font_color": white, "bg_color": blue, "align": "center", "border": 1, "border_color": white})
small_label = workbook.add_format({"font_size": 9, "font_color": ink, "border": 1, "border_color": line})
small_number = workbook.add_format({"font_size": 9, "font_color": ink, "align": "center", "border": 1, "border_color": line, "num_format": "0"})

dashboard = workbook.add_worksheet("Dashboard")
data = workbook.add_worksheet("Data")
definitions = workbook.add_worksheet("Metric Definitions")
charts = workbook.add_worksheet("Chart Data")
charts.hide()

dashboard.hide_gridlines(2)
dashboard.set_tab_color(navy)
dashboard.set_zoom(90)
dashboard.set_column("A:A", 2)
dashboard.set_column("B:N", 14)
dashboard.set_column("O:O", 2)
dashboard.set_row(0, 30)
dashboard.set_row(1, 22)
dashboard.merge_range("B1:N1", "Customer Churn | Business KPI Dashboard", title)
dashboard.merge_range("B2:N2", "Interactive summary based on the supplied cleaned dataset (15 records). Use the filters below; charts and KPI cards recalculate in Excel.", subtitle)

dashboard.write("B4", "FILTER", filter_label)
dashboard.write("C4", "SELECTION", filter_label)
dashboard.write("B5", "Contract type", filter_label)
dashboard.write("B6", "Tenure band", filter_label)
dashboard.write("B7", "Payment method", filter_label)
for cell in ("C5", "C6", "C7"):
    dashboard.write(cell, "All", filter_value)
dashboard.data_validation("C5", {"validate": "list", "source": ["All"] + contracts})
dashboard.data_validation("C6", {"validate": "list", "source": ["All"] + tenure_bands})
dashboard.data_validation("C7", {"validate": "list", "source": ["All"] + payments})
dashboard.merge_range("E4:N4", "KPI SUMMARY", section)

criteria = (
    '--(((Data!$H$2:$H$16=$C$5)+($C$5="All"))>0),'
    '--(((Data!$L$2:$L$16=$C$6)+($C$6="All"))>0),'
    '--(((Data!$J$2:$J$16=$C$7)+($C$7="All"))>0)'
)
records_formula = "=SUMPRODUCT(" + "".join(criteria) + ")"
churn_formula = "=IFERROR(SUMPRODUCT(--(Data!$K$2:$K$16=\"Yes\")," + "".join(criteria) + ")/E6,0)"
retention_formula = "=IFERROR(SUMPRODUCT(--(Data!$K$2:$K$16=\"No\")," + "".join(criteria) + ")/E6,0)"
risk_formula = "=SUMPRODUCT(Data!$F$2:$F$16,--(Data!$K$2:$K$16=\"Yes\")," + "".join(criteria) + ")"
tenure_formula = "=IFERROR(SUMPRODUCT(Data!$D$2:$D$16," + "".join(criteria) + ")/E6,0)"

for cell, text in [("E5", "Records"), ("G5", "Churn rate"), ("I5", "Retention rate"), ("K5", "Monthly revenue at risk"), ("M5", "Average tenure")]:
    dashboard.merge_range(f"{cell}:{chr(ord(cell[0])+1)}5", text, card_label)
dashboard.merge_range("E6:F7", "", card_number)
dashboard.merge_range("G6:H7", "", card_percent)
dashboard.merge_range("I6:J7", "", card_percent)
dashboard.merge_range("K6:L7", "", card_currency)
dashboard.merge_range("M6:N7", "", card_decimal)
dashboard.write_formula("E6", records_formula, card_number, len(records))
dashboard.write_formula("G6", churn_formula, card_percent, 7 / 15)
dashboard.write_formula("I6", retention_formula, card_percent, 8 / 15)
dashboard.write_formula("K6", risk_formula, card_currency, sum(r["MonthlyCharges"] for r in records if r["Churn"] == "Yes"))
dashboard.write_formula("M6", tenure_formula, card_decimal, sum(r["TenureMonths"] for r in records) / len(records))
dashboard.merge_range("B9:N9", "CHURN PATTERNS", section)

# Helper data uses the selected tenure band and payment method, but not the grouped chart dimension.
charts.write_row("A1", ["Contract type", "Churned", "Retained"], small_header)
for row, contract in enumerate(contracts, start=1):
    excel_row = row + 1
    charts.write(row, 0, contract, small_label)
    common = (
        f'--(Data!$H$2:$H$16=A{excel_row}),'
        '--(((Data!$L$2:$L$16=Dashboard!$C$6)+(Dashboard!$C$6="All"))>0),'
        '--(((Data!$J$2:$J$16=Dashboard!$C$7)+(Dashboard!$C$7="All"))>0)'
    )
    churned = '=SUMPRODUCT(--(Data!$K$2:$K$16="Yes"),' + ''.join(common) + ')'
    retained = '=SUMPRODUCT(--(Data!$K$2:$K$16="No"),' + ''.join(common) + ')'
    charts.write_formula(row, 1, churned, small_number, sum(1 for r in records if r["ContractType"] == contract and r["Churn"] == "Yes"))
    charts.write_formula(row, 2, retained, small_number, sum(1 for r in records if r["ContractType"] == contract and r["Churn"] == "No"))

charts.write_row("E1", ["Tenure band", "Churned", "Retained"], small_header)
for row, band in enumerate(tenure_bands, start=1):
    excel_row = row + 1
    charts.write(row, 4, band, small_label)
    common = (
        f'--(Data!$L$2:$L$16=E{excel_row}),'
        '--(((Data!$H$2:$H$16=Dashboard!$C$5)+(Dashboard!$C$5="All"))>0),'
        '--(((Data!$J$2:$J$16=Dashboard!$C$7)+(Dashboard!$C$7="All"))>0)'
    )
    churned = '=SUMPRODUCT(--(Data!$K$2:$K$16="Yes"),' + ''.join(common) + ')'
    retained = '=SUMPRODUCT(--(Data!$K$2:$K$16="No"),' + ''.join(common) + ')'
    charts.write_formula(row, 5, churned, small_number, sum(1 for r in records if r["TenureBand"] == band and r["Churn"] == "Yes"))
    charts.write_formula(row, 6, retained, small_number, sum(1 for r in records if r["TenureBand"] == band and r["Churn"] == "No"))

charts.write_row("I1", ["Payment method", "Churned", "Retained"], small_header)
for row, payment in enumerate(payments, start=1):
    excel_row = row + 1
    charts.write(row, 8, payment, small_label)
    common = (
        f'--(Data!$J$2:$J$16=I{excel_row}),'
        '--(((Data!$H$2:$H$16=Dashboard!$C$5)+(Dashboard!$C$5="All"))>0),'
        '--(((Data!$L$2:$L$16=Dashboard!$C$6)+(Dashboard!$C$6="All"))>0)'
    )
    churned = '=SUMPRODUCT(--(Data!$K$2:$K$16="Yes"),' + ''.join(common) + ')'
    retained = '=SUMPRODUCT(--(Data!$K$2:$K$16="No"),' + ''.join(common) + ')'
    charts.write_formula(row, 9, churned, small_number, sum(1 for r in records if r["PaymentMethod"] == payment and r["Churn"] == "Yes"))
    charts.write_formula(row, 10, retained, small_number, sum(1 for r in records if r["PaymentMethod"] == payment and r["Churn"] == "No"))

def make_chart(title_text, categories, churned_values, retained_values, width=500):
    chart = workbook.add_chart({"type": "column", "subtype": "stacked"})
    chart.add_series({"name": "Churned", "categories": categories, "values": churned_values, "fill": {"color": coral}, "border": {"none": True}})
    chart.add_series({"name": "Retained", "categories": categories, "values": retained_values, "fill": {"color": teal}, "border": {"none": True}})
    chart.set_title({"name": title_text, "name_font": {"color": navy, "size": 11, "bold": True}})
    chart.set_legend({"position": "bottom", "font": {"color": muted, "size": 9}})
    chart.set_y_axis({"name": "Customers", "major_gridlines": {"visible": False}, "num_font": {"color": muted, "size": 8}, "name_font": {"color": muted, "size": 8}, "min": 0})
    chart.set_x_axis({"num_font": {"color": muted, "size": 8}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": white}})
    chart.set_chartarea({"border": {"color": line}, "fill": {"color": white}})
    chart.set_size({"width": width, "height": 270})
    return chart

dashboard.insert_chart("B10", make_chart("Churn by contract", ["Chart Data", 1, 0, 3, 0], ["Chart Data", 1, 1, 3, 1], ["Chart Data", 1, 2, 3, 2]), {"x_scale": 1.0, "y_scale": 1.0})
dashboard.insert_chart("H10", make_chart("Churn by tenure band", ["Chart Data", 1, 4, 3, 4], ["Chart Data", 1, 5, 3, 5], ["Chart Data", 1, 6, 3, 6]), {"x_scale": 1.0, "y_scale": 1.0})
dashboard.insert_chart("B25", make_chart("Churn by payment method", ["Chart Data", 1, 8, len(payments), 8], ["Chart Data", 1, 9, len(payments), 9], ["Chart Data", 1, 10, len(payments), 10], width=620), {"x_scale": 1.0, "y_scale": 1.0})

dashboard.merge_range("H25:N25", "How to read this dashboard", section)
dashboard.merge_range("H26:N31", "Use the three dropdown filters to change the KPI cards and charts. “Monthly revenue at risk” is the sum of MonthlyCharges for customers marked as churned in the selected view. This dataset has 15 records, so differences are descriptive observations—not causal conclusions or population estimates.", note)
dashboard.merge_range("H33:N33", "Decision note", section)
dashboard.merge_range("H34:N37", "The unfiltered view indicates a 46.7% churn rate. Churn is concentrated among month-to-month customers in this small sample; investigate the pattern with a larger dataset before taking action.", note)
dashboard.freeze_panes(4, 1)

# Source data sheet.
headers = ["CustomerID", "Gender", "Age", "TenureMonths", "SubscriptionType", "MonthlyCharges", "TotalCharges", "ContractType", "SupportTickets", "PaymentMethod", "Churn", "TenureBand"]
data.hide_gridlines(2)
data.set_tab_color(blue)
data.set_column("A:A", 14)
data.set_column("B:B", 11)
data.set_column("C:D", 13)
data.set_column("E:E", 17)
data.set_column("F:G", 16)
data.set_column("H:H", 18)
data.set_column("I:I", 15)
data.set_column("J:J", 17)
data.set_column("K:L", 14)
for col, header in enumerate(headers):
    data.write(0, col, header, data_header)
for row, record in enumerate(records, start=1):
    values = [record[h] for h in headers]
    for col, value in enumerate(values):
        cell_format = None
        if headers[col] in {"Age", "TenureMonths", "SupportTickets"}:
            cell_format = integer
        elif headers[col] in {"MonthlyCharges", "TotalCharges"}:
            cell_format = currency
        data.write(row, col, value, cell_format)
data.add_table(0, 0, len(records), len(headers) - 1, {"name": "ChurnData", "style": "Table Style Medium 2", "columns": [{"header": h} for h in headers]})
data.freeze_panes(1, 0)

# Definitions sheet documents the KPI assumptions and source limitations.
definitions.hide_gridlines(2)
definitions.set_tab_color(teal)
definitions.set_column("A:A", 25)
definitions.set_column("B:B", 80)
definitions.merge_range("A1:B1", "KPI Definitions and Assumptions", title)
definitions.write_row("A3", ["Metric", "Definition used in this workbook"], definition_head)
definition_rows = [
    ("Churn rate", "Customers with Churn = Yes ÷ selected customers."),
    ("Retention rate", "Customers with Churn = No ÷ selected customers."),
    ("Monthly revenue at risk", "Sum of MonthlyCharges for customers with Churn = Yes in the selected view. It is a monthly recurring-charge proxy, not a forecast of realized revenue loss."),
    ("Average tenure", "Average TenureMonths for selected customers."),
    ("Tenure bands", "0-12 months, 13-24 months, and 25+ months. These are display groupings created for this dashboard."),
    ("Filters", "Contract type, tenure band, and payment method dropdowns on the Dashboard sheet. Metrics and charts recalculate when opened in Excel."),
    ("Scope limitation", "The source contains 15 records. Dashboard patterns are descriptive only and should not be generalized or treated as causal."),
]
for row, values in enumerate(definition_rows, start=3):
    definitions.write(row, 0, values[0], definition)
    definitions.write(row, 1, values[1], definition)
    definitions.set_row(row, 38)
definitions.write("A12", "Source", label)
definitions.write("B12", "data/customer_churn_cleaned.csv", note)
definitions.freeze_panes(2, 0)

workbook.close()
print(f"Created {OUTPUT}")
