import pandas as pd

# Load the Excel file
file_path = 'sheet.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load each sheet into a DataFrame
baseline_sheet = pd.read_excel(file_path, sheet_name='Baseline Sheet ')
outcomes_sheet = pd.read_excel(file_path, sheet_name='Outcomes Sheet')
change_in_variables_sheet = pd.read_excel(file_path, sheet_name='Change in Variables')

# Data Cleaning: Set the first row as the header and drop it
baseline_sheet.columns = baseline_sheet.iloc[0]
baseline_sheet = baseline_sheet.drop(0).reset_index(drop=True)

outcomes_sheet.columns = outcomes_sheet.iloc[0]
outcomes_sheet = outcomes_sheet.drop(0).reset_index(drop=True)

change_in_variables_sheet.columns = change_in_variables_sheet.iloc[0]
change_in_variables_sheet = change_in_variables_sheet.drop(0).reset_index(drop=True)

# Descriptive Statistics for each sheet
baseline_stats = baseline_sheet.describe(include='all')
outcomes_stats = outcomes_sheet.describe(include='all')
change_stats = change_in_variables_sheet.describe(include='all')

# Convert summary tables to HTML
html_tables = {
    'Baseline Sheet Statistics': baseline_stats.to_html(),
    'Outcomes Sheet Statistics': outcomes_stats.to_html(),
    'Change in Variables Statistics': change_stats.to_html()
}

# Save HTML tables to a file
html_content = "<html><head><title>Summary Tables</title></head><body>"
for name, html_table in html_tables.items():
    html_content += f"<h2>{name}</h2>{html_table}<br><br>"
html_content += "</body></html>"

with open('summary_tables.html', 'w') as file:
    file.write(html_content)

print("Summary tables have been saved to summary_tables.html")
