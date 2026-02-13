
import pandas as pd
from pathlib import Path

def check_data_columns():
    excel_path = Path("examples/Project_Sample/Data/sample_data.xlsx")
    if not excel_path.exists():
        print(f"File not found: {excel_path}")
        return

    try:
        df = pd.read_excel(excel_path, dtype=str)
        print("Excel Columns found:")
        for col in df.columns:
            print(f" - '{col}'")
            
        print("\nFirst row data:")
        print(df.iloc[0].to_dict())
    except Exception as e:
        print(f"Error reading Excel: {e}")

if __name__ == "__main__":
    check_data_columns()
