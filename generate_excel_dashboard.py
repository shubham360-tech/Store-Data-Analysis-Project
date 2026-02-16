import openpyxl
import pandas as pd

# Sample data for different sheets
sales_data = {
    'Product': ['Widget A', 'Widget B', 'Widget C'],
    'Sales': [1000, 1500, 900],
    'Date': pd.to_datetime(['2026-02-15', '2026-02-15', '2026-02-15'])
}

customer_data = {
    'Customer ID': [1, 2, 3],
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [30, 25, 35],
    'Location': ['New York', 'Los Angeles', 'Chicago']
}

product_data = {
    'Product': ['Widget A', 'Widget B', 'Widget C'],
    'Performance': ['Good', 'Better', 'Best']
}

kpi_summary = {
    'KPI': ['Total Sales', 'Average Order Value'],
    'Value': [3400, 1133.33]
}

# Create a Pandas Excel writer using openpyxl as the engine
excel_file_path = 'Store_Data_Analysis.xlsx'
with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
    # Write each DataFrame to a different worksheet
    pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales Data', index=False)
    pd.DataFrame(customer_data).to_excel(writer, sheet_name='Customer Demographics', index=False)
    pd.DataFrame(product_data).to_excel(writer, sheet_name='Product Performance', index=False)
    pd.DataFrame(kpi_summary).to_excel(writer, sheet_name='KPI Summary', index=False)

    # Access the workbook and the sheets to format them
    workbook = writer.book
    for sheet in workbook.sheetnames:
        worksheet = workbook[sheet]
        # Set column widths for better visibility
        if sheet == 'Sales Data':
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 15
        elif sheet == 'Customer Demographics':
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 10
            worksheet.column_dimensions['D'].width = 20
        elif sheet == 'Product Performance':
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 20
        elif sheet == 'KPI Summary':
            worksheet.column_dimensions['A'].width = 30
            worksheet.column_dimensions['B'].width = 15

# Save the workbook
workbook.save(excel_file_path)
print('Excel dashboard has been generated successfully!')
