import pandas as pd

# 1. Raw sales data
raw_sales_data = [
    {"Product": "Laptop", "Price": 800, "Quantity": 3, "Customer": "John"},
    {"Product": "Mouse", "Price": 25, "Quantity": 10, "Customer": "Alice"},
    {"Product": "Keyboard", "Price": 45, "Quantity": 5, "Customer": "Bob"},
    {"Product": "Monitor", "Price": 150, "Quantity": 2, "Customer": "David"}
]

# Create initial DataFrame
df = pd.DataFrame(raw_sales_data)

# 2. Automatically calculate Total Income for each product row
df["Total_Income"] = df["Price"] * df["Quantity"]

# 3. Calculate financial metrics for the entire file
grand_total = df["Total_Income"].sum()
average_product_price = df["Price"].mean()

# 4. Append metrics to the bottom of the table using .loc
df.loc[len(df)] = {"Product": "GRAND TOTAL", "Price": "-", "Quantity": "-", "Customer": "-", "Total_Income": grand_total}
df.loc[len(df)] = {"Product": "AVERAGE PRICE", "Price": average_product_price, "Quantity": "-", "Customer": "-", "Total_Income": "-"}

# 5. Output the finalized data to an Excel file
output_file = "sales_report.xlsx"
df.to_excel(output_file, index=False)

print("🏆 Project Success: The professional sales report has been generated successfully!")

