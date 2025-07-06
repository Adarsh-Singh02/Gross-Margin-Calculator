import pandas as pd
import matplotlib.pyplot as plt
import os

# ───────────────────────────────────────────────
# 1.  FILE LOCATIONS  (🔄 change if needed)
# ───────────────────────────────────────────────
file_path         = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/New/FAAS Working File.xlsx"
output_excel_path = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/New/Gross_Margin_Output.xlsx"
output_chart_path = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/New/Gross_Margin_Dashboard.png"

# ───────────────────────────────────────────────
# 2.  READ & CLEAN SHEETS
# ───────────────────────────────────────────────
xls = pd.ExcelFile("D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/New/FAAS Working File.xlsx")

# Employee ─────────────────────────────────────
employee_df = xls.parse("Employee")
employee_df.columns = employee_df.columns.str.strip()
employee_df["Salary"]      = pd.to_numeric(employee_df["Salary"], errors="coerce")
employee_df["Involvement"] = pd.to_numeric(employee_df["Involvement"], errors="coerce")
employee_df["Cost"]        = employee_df["Salary"] * employee_df["Involvement"]

# Client (revenue) ─────────────────────────────
client_raw = xls.parse("Clinet Name ", header=None)
client_raw.columns = client_raw.iloc[0].str.strip()          # promote first row to header
client_df = client_raw[1:].copy()
client_df = client_df.rename(columns={"Client Name": "Project",
                                      "Amount":      "Revenue"})
client_df["Revenue"] = pd.to_numeric(client_df["Revenue"], errors="coerce")

# ➡️  GROUP revenue in case of duplicate/valid invoices
#client_df = client_df.groupby("Project", as_index=False)["Revenue"].sum()
client_df = client_df.groupby(["Project", "Ownership"], as_index=False)["Revenue"].sum()

# Direct Expense ───────────────────────────────
direct_df = xls.parse("Direct Expense")
direct_df.columns = direct_df.columns.str.strip()
direct_df = direct_df.rename(columns={"Client": "Project",
                                      "Amount": "Direct Expense"})
direct_df["Direct Expense"] = pd.to_numeric(direct_df["Direct Expense"], errors="coerce")

# ───────────────────────────────────────────────
# 3.  AGGREGATE COST & EXPENSE PER PROJECT
# ───────────────────────────────────────────────
project_costs    = employee_df.groupby("Project", as_index=False)["Cost"].sum()
project_expenses = direct_df.groupby("Project",  as_index=False)["Direct Expense"].sum()

# ───────────────────────────────────────────────
# 4.  MERGE EVERYTHING  (full outer keeps all projects)
# ───────────────────────────────────────────────
#gm_df = (
    #project_costs
    #.merge(client_df,  on="Project", how="outer")
    #.merge(project_expenses, on="Project", how="outer")
#)
gm_df = (
    client_df
    .merge(project_costs, on="Project", how="outer")
    .merge(project_expenses, on="Project", how="outer")
)

# Fill blanks with zero
for col in ["Cost", "Revenue", "Direct Expense"]:
    gm_df[col] = pd.to_numeric(gm_df[col], errors="coerce").fillna(0)

# ───────────────────────────────────────────────
# 5.  CALCULATE GROSS MARGIN
# ───────────────────────────────────────────────
gm_df["Gross Margin"]   = gm_df["Revenue"] - gm_df["Cost"] - gm_df["Direct Expense"]
gm_df["Gross Margin %"] = gm_df.apply(
    lambda r: (r["Gross Margin"] / r["Revenue"]) if r["Revenue"] else None,
    axis=1
).round(2)

# Optional preview
print("✅ Data preview:\n", gm_df.head())

# ───────────────────────────────────────────────
# 6.  EXPORT TO EXCEL  (main sheet + pivot + chart)
# ───────────────────────────────────────────────
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)

with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:
    # Main sheet
    gm_df.to_excel(writer, sheet_name="Gross Margin", index=False)

    wb  = writer.book
    ws  = writer.sheets["Gross Margin"]

    cur_fmt = wb.add_format({"num_format": "#,##0.00"})
    pct_fmt = wb.add_format({"num_format": "0.00%"})

    #ws.set_column("B:B", 14, cur_fmt)    # Cost
    #ws.set_column("C:D", 14, cur_fmt)    # Revenue & Direct Expense
    #ws.set_column("E:E", 16, cur_fmt)    # Gross Margin
    #ws.set_column("F:F", 16, pct_fmt)    # Gross Margin %
    ws.set_column("C:C", 14, cur_fmt)  # Cost
    ws.set_column("D:E", 14, cur_fmt)  # Revenue & Direct Expense
    ws.set_column("F:F", 16, cur_fmt)  # Gross Margin
    ws.set_column("G:G", 16, pct_fmt)  # Gross Margin %

    # Pivot: average GM % by project
    pivot_df = (
        gm_df.pivot_table(index="Project",
                          values="Gross Margin %",
                          aggfunc="mean")
        .reset_index()
        .sort_values("Gross Margin %", ascending=False)
    )
    pivot_df.to_excel(writer, sheet_name="GM % Pivot", index=False)

    ws_piv = writer.sheets["GM % Pivot"]
    ws_piv.set_column("A:A", 25)
    ws_piv.set_column("B:B", 16, pct_fmt)

    # Column chart for GM % in pivot sheet
    chart = wb.add_chart({'type': 'column'})
    max_row = len(pivot_df)
    chart.add_series({
        'name':       'Gross Margin %',
        'categories': ['GM % Pivot', 1, 0, max_row, 0],  # Project names
        'values':     ['GM % Pivot', 1, 1, max_row, 1],  # GM %
        'data_labels': {'value': False},
    })
    chart.set_title({'name': 'Gross Margin % by Project'})
    chart.set_x_axis({'name': 'Project'})
    chart.set_y_axis({'name': 'Gross Margin %'})
    chart.set_legend({'none': True})
    ws_piv.insert_chart('D2', chart,
                        {'x_scale': 2.0, 'y_scale': 1.4,
                         'x_offset': 20, 'y_offset': 10})

print(f"✅ Excel saved to {output_excel_path}")

# ───────────────────────────────────────────────
# 7.  PNG BAR CHART OF GROSS MARGIN ₹
# ───────────────────────────────────────────────
plt.figure(figsize=(10, 6))
plt.bar(gm_df["Project"], gm_df["Gross Margin"])
plt.xlabel("Project")
plt.ylabel("Gross Margin")
plt.title("Gross Margin per Project")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.savefig(output_chart_path, dpi=300)
plt.close()

print(f"✅ PNG chart saved to {output_chart_path}")

# Auto-open the Excel file (Windows only; ignore on Mac/Linux)
try:
    os.startfile(output_excel_path)
except AttributeError:
    pass