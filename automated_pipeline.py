import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill

def process_and_calculate_data(input_data):
    """بخش پانداس: دریافت داده‌ها، ایجاد دیتافریم و محاسبات مالی"""
    df = pd.DataFrame(input_data)
    
    # محاسبه خودکار درآمد کل برای هر سطر
    df["Total_Income"] = df["Price"] * df["Quantity"]
    
    # محاسبه شاخص‌های کلی مالی
    grand_total = df["Total_Income"].sum()
    average_product_price = df["Price"].mean()
    
    # اضافه کردن سطر مجموع کل و میانگین قیمت با استفاده از .loc
    df.loc[len(df)] = {"Product": "GRAND TOTAL", "Price": "-", "Quantity": "-", "Customer": "-", "Total_Income": grand_total}
    df.loc[len(df)] = {"Product": "AVERAGE PRICE", "Price": average_product_price, "Quantity": "-", "Customer": "-", "Total_Income": "-"}
    
    return df

def apply_visual_styles(file_path):
    """بخش اوپن‌پای‌اکسل: بارگذاری فایل خروجی و اعمال استایل‌های گرافیکی حرفه‌ای"""
    if not os.path.exists(file_path):
        print(f"❌ Error: The file {file_path} was not found.")
        return
        
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    max_row = sheet.max_row
    
    # تعریف استایل‌ها (فونت ضخیم و پس‌زمینه سبز ملایم)
    bold_font = Font(name="Arial", size=11, bold=True)
    highlight_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    # استایل‌دهی به سطر GRAND TOTAL (یکی مانده به آخر)
    for col in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=max_row - 1, column=col)
        cell.font = bold_font
        cell.fill = highlight_fill
        
    # استایل‌دهی به سطر AVERAGE PRICE (سطر آخر)
    for col in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=max_row, column=col)
        cell.font = bold_font
        
    # ذخیره فایل نهایی استایل‌دهی شده
    final_output = "stylized_sales_report.xlsx"
    workbook.save(final_output)
    print(f"🎯 Visual styles and clean formatting applied successfully to: {final_output}")

# --- اجرای خط لوله اتوماسیون (Data Pipeline Execution) ---
if __name__ == "__main__":
    try:
        # داده‌های خام فروش شما
        raw_sales_data = [
            {"Product": "Laptop", "Price": 800, "Quantity": 3, "Customer": "John"},
            {"Product": "Mouse", "Price": 25, "Quantity": 10, "Customer": "Alice"},
            {"Product": "Keyboard", "Price": 45, "Quantity": 5, "Customer": "Bob"},
            {"Product": "Monitor", "Price": 150, "Quantity": 2, "Customer": "David"}
        ]
        
        print("🚀 Starting data automation pipeline...")
        
        # گام اول: پردازش پانداس
        processed_df = process_and_calculate_data(raw_sales_data)
        temp_file = "sales_report.xlsx"
        processed_df.to_excel(temp_file, index=False)
        print("✅ Step 1: Data processed and temporary excel file generated.")
        
        # گام دوم: زیباسازی اوپن‌پای‌اکسل
        apply_visual_styles(temp_file)
        
        print("✨ Project Success: The entire data pipeline completed successfully!")
        
    except Exception as e:
        print(f"❌ An error occurred during pipeline execution: {e}")