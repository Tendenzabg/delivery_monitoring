import pandas as pd
import os

# Paths
DATA_DIR = '/Users/stoyantodorov/Downloads/Consegne Happy - Sporttime'
ORDER_FILE = os.path.join(DATA_DIR, 'Happy Sport - Nike SP26 - Confirmation.xlsx')
DELIV_FILE = os.path.join(DATA_DIR, 'PLEVEN_delivery NIKE_050226.xlsx')

COL_MAP_ORDER = {
    'EAN UPC Cd': 'Barcode',
    'Qty status NNT': 'Ordered_Qty'
}

print("--- 1. Loading Order File ---")
try:
    df_order = pd.read_excel(ORDER_FILE, header=1)
    df_order = df_order.rename(columns=COL_MAP_ORDER)
    df_order['Barcode'] = df_order['Barcode'].astype(str).str.strip()
    # Group by Barcode
    df_order = df_order.groupby('Barcode', as_index=False)['Ordered_Qty'].sum()
    print(f"Loaded {len(df_order)} unique order items.")
    print(f"Total Ordered Qty: {df_order['Ordered_Qty'].sum()}")
except Exception as e:
    print(f"Error loading order: {e}")
    exit()

print("\n--- 2. Loading Delivery File ---")
try:
    df_deliv = pd.read_excel(DELIV_FILE, sheet_name='Dati_imp')
    df_deliv['Barcode'] = df_deliv['Barcode'].astype(str).str.strip()
    delivery_map = df_deliv.groupby('Barcode')['Dlv.qty'].sum().to_dict()
    print(f"Loaded deliveries for {len(delivery_map)} unique items.")
    print(f"Total Delivered Qty: {sum(delivery_map.values())}")
except Exception as e:
    print(f"Error loading delivery: {e}")
    exit()

print("\n--- 3. Merging and Verifying ---")
df_order['Delivered'] = df_order['Barcode'].map(delivery_map).fillna(0)
df_order['Remaining'] = df_order['Ordered_Qty'] - df_order['Delivered']

total_ordered = df_order['Ordered_Qty'].sum()
total_delivered = df_order['Delivered'].sum()
total_remaining = df_order['Remaining'].sum()

print(f"Total Ordered: {total_ordered}")
print(f"Total Delivered: {total_delivered}")
print(f"Total Remaining: {total_remaining}")

fully_delivered_count = len(df_order[df_order['Remaining'] <= 0])
print(f"Items fully delivered (Green rows): {fully_delivered_count}")
print("Verification Succcessful" if total_remaining < total_ordered else "Verification Failed")
