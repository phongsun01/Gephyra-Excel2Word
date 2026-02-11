"""
Script to create sample Excel file with sample data and Config sheet
Run this after installing dependencies: python examples/create_sample_excel.py
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Create sample data
text_data = {
    'ID': [1, 2, 3, 4, 5],
    'HoTen': ['Nguyễn Văn A', 'Trần Thị B', 'Lê Văn C', 'Phạm Thị D', 'Hoàng Văn E'],
   'SoHD': ['HD2024001', 'HD2024002', 'HD2024003', 'HD2024004', 'HD2024005'],
    'DiaChi': ['Hà Nội', 'Hồ Chí Minh', 'Đà Nẵng', 'Hải Phòng', 'Cần Thơ']
}

# Config sheet data (table mappings)
config_data = {
    'TableName': ['BangThietBi'],
    'SourceSheet': ['Table'],
    'Range': ['A1:D10'],
    'WordBookmark': ['TablePos'],
    'Mode': ['KeepFormat']
}

# Table data (sample table to paste into Word)
table_data = {
    'STT': list(range(1, 11)),
    'TenThietBi': ['Laptop Dell', 'Mouse Logitech', 'Keyboard', 'Monitor', 'Printer',
                   'Scanner', 'Webcam', 'Headset', 'USB Drive', 'External HDD'],
    'SoLuong': [2, 5, 3, 2, 1, 1, 3, 4, 10, 2],
    'DonGia': ['20.000.000', '200.000', '500.000', '5.000.000', '3.000.000',
               '1.500.000', '800.000', '1.200.000', '150.000', '2.500.000']
}

# Create Excel file with multiple sheets
with pd.ExcelWriter('Project_Sample/Data/input.xlsx', engine='openpyxl') as writer:
    # Write Text sheet
    df_text = pd.DataFrame(text_data)
    df_text.to_excel(writer, sheet_name='Text', index=False)
    
    # Write Config sheet
    df_config = pd.DataFrame(config_data)
    df_config.to_excel(writer, sheet_name='Config', index=False)
    
    # Write Table sheet
    df_table = pd.DataFrame(table_data)
    df_table.to_excel(writer, sheet_name='Table', index=False)

print("✅ Sample Excel file created: Project_Sample/Data/input.xlsx")
print("   - Text sheet: 5 rows")
print("   - Config sheet: 1 table mapping")
print("   - Table sheet: 10 rows sample data")
