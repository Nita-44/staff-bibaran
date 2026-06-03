import pandas as pd

# Load the Excel file to inspect its sheets and structure
file_path = "कर्मचारी.xlsx"
try:
    xls = pd.ExcelFile(file_path)
    print("Sheet names:", xls.sheet_names)
    
    # Read the first sheet to see what it looks like
    df = pd.read_excel(file_path, sheet_name=0)
    print("\nFirst 5 rows of the data:")
    print(df.head())
    print("\nData info:")
    print(df.info())
except Exception as e:
    print("Error:", e)