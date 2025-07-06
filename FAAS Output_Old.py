import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime
from pathlib import Path

# ───────────────────────────────────────────────
# 1.  FILE LOCATIONS  (🔄 change if needed)
# ───────────────────────────────────────────────
#file_path         = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/FAAS Working File.xlsx"
#output_excel_path = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/Gross_Margin_Output.xlsx"
#output_chart_path = r"D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/Gross_Margin_Dashboard.png"

month_folder = datetime.now().strftime("%Y-%m")
base_file_path = Path("D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/FAAS Working File.xlsx")
base_dir = base_file_path.parent

output_dir = base_dir / month_folder
output_dir.mkdir(parents=True, exist_ok=True)

file_path = base_file_path
output_excel_path = output_dir / "Gross_Margin_Output.xlsx"
output_chart_path = output_dir / "Gross_Margin_Dashboard.png"

# ───────────────────────────────────────────────
# 2.  READ & CLEAN SHEETS
# ───────────────────────────────────────────────
xls = pd.ExcelFile("D:/OneDrive - valueonshore.com/Desktop/Allocation Working/FAAS/FAAS Working File.xlsx")

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
# Sort and get Top 5 and Bottom 5 by Gross Margin %
#top5_df = gm_df.sort_values(by="Gross Margin %", ascending=False).head(5)
#bottom5_df = gm_df.sort_values(by="Gross Margin %", ascending=True).head(5)

with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:

    #top5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=1, startcol=7, header=True)
    #bottom5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=15, startcol=7, header=True)

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
    #chart = wb.add_chart({'type': 'column'})
    #max_row = len(pivot_df)
    #chart.add_series({
    #    'name':       'Gross Margin %',
    #    'categories': ['GM % Pivot', 1, 0, max_row, 0],  # Project names
    #    'values':     ['GM % Pivot', 1, 1, max_row, 1],  # GM %
    #    'data_labels': {'value': False},
    #})
    #chart.set_title({'name': 'Gross Margin % by Project'})
    #chart.set_x_axis({'name': 'Project'})
    #chart.set_y_axis({'name': 'Gross Margin %'})
    #chart.set_legend({'none': True})
    #ws_piv.insert_chart('D2', chart,
    #                    {'x_scale': 2.0, 'y_scale': 1.4,
    #                     'x_offset': 20, 'y_offset': 10})

    # Create a new Column Chart
    chart = wb.add_chart({'type': 'column'})

    # Get number of data rows
    max_row = len(pivot_df)
    
    # Add series to the chart
    chart.add_series({
        'name':       'Gross Margin %',
        'categories': ['GM % Pivot', 1, 0, max_row, 0],  # A2:A{n} – Project
        'values':     ['GM % Pivot', 1, 1, max_row, 1],  # B2:B{n} – GM %
        'data_labels': {
            'value': False,
            'num_format': '0.0%',
            'position': 'outside_end'
        },
    })

    # Set chart title & axis formatting
    chart.set_title({
        'name': 'Gross Margin % by Project',
        'name_font': {'bold': True, 'color': '#1F4E78', 'size': 14}
    })

    chart.set_x_axis({
        'name': 'Project',
        'name_font': {'bold': True},
        'num_font': {'rotation': -45, 'color': '#555555'},
        'label_position': 'low'
    })

    chart.set_y_axis({
        'name': 'Gross Margin %',
        'name_font': {'bold': True},
        'num_font': {'color': '#555555'},
        'major_gridlines': {'visible': False}
    })

    chart.set_legend({'none': True})

    # Add a nice chart style
    chart.set_style(10)  # Choose from 1–48, 10 is clean blue

    # Insert chart into worksheet with size and padding
    ws_piv.insert_chart('D2', chart, {
        'x_scale': 2.2,
        'y_scale': 1.6,
        'x_offset': 10,
        'y_offset': 20
    })

    '''pivot_start_row = 0
    pivot_num_rows = len(pivot_df)   # e.g., 20
    pivot_end_row = pivot_start_row + pivot_num_rows  # Row 20

    top5_start_row = pivot_end_row + 5     # Row 22
    bottom5_start_row = top5_start_row + len(top5_df) + 5  # Row 29

    start_col = 0  # Column A

    pivot_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=0)
    pivot_end_row = len(pivot_df) + 2  # Add 2 rows space
    
    pivot_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=pivot_start_row, startcol=start_col)
    top5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=top5_start_row, startcol=start_col)
    bottom5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=bottom5_start_row, startcol=start_col)

    # === Chart for Top 5 ===
    chart_top5 = wb.add_chart({'type': 'column'})
    chart_top5.add_series({
        'name': 'Top 5 GM %',
        'categories': ['GM % Pivot', top5_start_row + 1, 0, top5_start_row + 5, 0],   # Project
        'values':     ['GM % Pivot', top5_start_row + 1, 5, top5_start_row + 5, 5],   # GM %
        'data_labels': {'value': True, 'num_format': '0.0'}
    })
    chart_top5.set_title({'name': 'Top 5 Projects by GM %'})
    chart_top5.set_x_axis({'num_font': {'rotation': -45}})
    chart_top5.set_legend({'none': True})
    chart_top5.set_style(10)
    ws_piv.insert_chart(f'H{top5_start_row + 1}', chart_top5, {'x_scale': 2, 'y_scale': 1.5})

    # === Chart for Bottom 5 ===
    chart_bottom5 = wb.add_chart({'type': 'column'})
    chart_bottom5.add_series({
        'name': 'Bottom 5 GM %',
        'categories': ['GM % Pivot', bottom5_start_row + 1, 0, bottom5_start_row + 5, 0],   # Project
        'values':     ['GM % Pivot', bottom5_start_row + 1, 5, bottom5_start_row + 5, 5],   # GM %
        'data_labels': {'value': True, 'num_format': '0.0'}
    })
    chart_bottom5.set_title({'name': 'Bottom 5 Projects by GM %'})
    chart_bottom5.set_x_axis({'num_font': {'rotation': -45}})
    chart_bottom5.set_legend({'none': True})
    chart_bottom5.set_style(11)
    ws_piv.insert_chart(f'H{bottom5_start_row + 1}', chart_bottom5, {'x_scale': 2, 'y_scale': 1.5})'''

    
    # ─── Constants ───────────────────────────────────────────────
    proj_col = 0         # 'Project' column index
    gm_pct_col = 5       # 'Gross Margin %' column index
    table_to_chart_col = 9  # Place chart in column G (index 6)
    x_scale = 1.2
    y_scale = 1.0

    # ─── 1) Write Pivot Table ────────────────────────────────────
    pivot_start_row = 0
    pivot_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=pivot_start_row, startcol=0)
    ws_piv = writer.sheets["GM % Pivot"]
    pivot_rows = len(pivot_df)

    # ─── 2) Write Top 5 Table ─────────────────────────────────────
    #top5_start_row = pivot_start_row + pivot_rows + 3
    #top5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=top5_start_row, startcol=0)

    # ─── 3) Insert Top 5 Chart (beside table) ─────────────────────
    '''chart_top5 = wb.add_chart({'type': 'column'})
    chart_top5.add_series({
        'name': 'Top 5 Gross Margin %',
        'categories': ['GM % Pivot', top5_start_row + 1, proj_col, top5_start_row + len(top5_df), proj_col],
        'values':     ['GM % Pivot', top5_start_row + 1, gm_pct_col, top5_start_row + len(top5_df), gm_pct_col],
        'data_labels': {'value': True, 'num_format': '0.0'},
    })
    chart_top5.set_title({'name': 'Top 5 Projects by GM %'})
    chart_top5.set_x_axis({'num_font': {'rotation': -45}})
    chart_top5.set_legend({'none': True})
    chart_top5.set_style(10)

    # Insert next to top 5 table
    ws_piv.insert_chart(top5_start_row, table_to_chart_col, chart_top5,
                        {'x_scale': x_scale, 'y_scale': y_scale})

    # ─── 4) Write Bottom 5 Table ──────────────────────────────────
    bottom5_start_row = top5_start_row + len(top5_df) + 9
    bottom5_df.to_excel(writer, sheet_name="GM % Pivot", index=False, startrow=bottom5_start_row, startcol=0)

    # ─── 5) Insert Bottom 5 Chart (beside table) ──────────────────
    chart_bottom5 = wb.add_chart({'type': 'column'})
    chart_bottom5.add_series({
        'name': 'Bottom 5 Gross Margin %',
        'categories': ['GM % Pivot', bottom5_start_row + 1, proj_col, bottom5_start_row + len(bottom5_df), proj_col],
        'values':     ['GM % Pivot', bottom5_start_row + 1, gm_pct_col, bottom5_start_row + len(bottom5_df), gm_pct_col],
        'data_labels': {'value': True, 'num_format': '0.0'},
    })
    chart_bottom5.set_title({'name': 'Bottom 5 Projects by GM %'})
    chart_bottom5.set_x_axis({'num_font': {'rotation': -45}})
    chart_bottom5.set_legend({'none': True})
    chart_bottom5.set_style(11)

    # Insert next to bottom 5 table
    ws_piv.insert_chart(bottom5_start_row, table_to_chart_col, chart_bottom5,
                        {'x_scale': x_scale, 'y_scale': y_scale})'''

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