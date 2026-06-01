import pandas as pd
import os

base_path = '/Users/stoyantodorov/Downloads/Consegne Happy - Sporttime'
file1 = os.path.join(base_path, 'Happy Sport - Nike SP26 - Confirmation.xlsx')

print(f"--- Searching for headers in {os.path.basename(file1)} ---")
try:
    # Read first 20 rows without header
    df = pd.read_excel(file1, header=None, nrows=20)
    
    # Iterate through rows to find the one containing "EAN UPC Cd"
    for index, row in df.iterrows():
        row_values = [str(val).strip() for val in row.values]
        if "EAN UPC Cd" in row_values:
            print(f"Found header at row index: {index}")
            print("Row content:", row_values)
            break
    else:
        print("Header 'EAN UPC Cd' not found in the first 20 rows.")

except Exception as e:
    print("Error reading file:", e)
