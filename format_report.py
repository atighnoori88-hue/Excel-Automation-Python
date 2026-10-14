import openpyxl
from openpyxl.styles import Font, PatternFill

# 1. Load the generated Excel workbook
workbook = openpyxl.load_workbook("sales_report.xlsx")
sheet = workbook.active

# 2. Create styles for the summary rows (Font and Background color)
bold_font = Font(name="Arial", size=11, bold=True)
highlight_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft green background

# 3. Find and style the last two rows (GRAND TOTAL and AVERAGE PRICE)
max_row = sheet.max_row

# Style the GRAND TOTAL row
for col in range(1, sheet.max_column + 1):
    cell = sheet.cell(row=max_row - 1, column=col)
    cell.font = bold_font
    cell.fill = highlight_fill

# Style the AVERAGE PRICE row
for col in range(1, sheet.max_column + 1):
    cell = sheet.cell(row=max_row, column=col)
    cell.font = bold_font

# 4. Save the stylized report
workbook.save("stylized_sales_report.xlsx")
print("✨ Visual styles and clean formatting applied successfully to the Excel sheet!")
