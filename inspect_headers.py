import pandas as pd
import os

base_path = '/Users/stoyantodorov/Downloads/Consegne Happy - Sporttime'
file1 = os.path.join(base_path, 'Happy Sport - Nike SP26 - Confirmation.xlsx')
file2 = os.path.join(base_path, 'PLEVEN_delivery NIKE_050226.xlsx')

print(f"--- Inspecting {os.path.basename(file1)} ---")
try:
    df1 = pd.read_excel(file1, nrows=5)
    print("Columns:", list(df1.columns))
except Exception as e:
    print("Error reading file 1:", e)

print(f"\n--- Inspecting {os.path.basename(file2)} (Sheet: Dati_imp) ---")
try:
    df2 = pd.read_excel(file2, sheet_name='Dati_imp', nrows=5)
    print("Columns:", list(df2.columns))
except Exception as e:
    print("Error reading file 2:", e)
