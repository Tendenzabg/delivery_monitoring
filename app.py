import streamlit as st
import pandas as pd
import os
import json

# --- Configuration ---
# --- Configuration ---
# State is saved locally in the running directory
STATE_FILE = 'app_state.json'
LOCAL_ORDER_PATH = 'saved_order.xlsx'

# --- Header Mapping ---
# Order File Columns -> App Columns
COL_MAP_ORDER = {
    'EAN UPC Cd': 'Barcode',
    'CCM': 'Concatenate',
    'Article Name': 'Description',
    'EU Size': 'SizeConverted',
    'Qty status NNT': 'Ordered_Qty'
}

# Delivery File Columns
COL_BARCODE_DELIV = 'Barcode'
COL_QTY_DELIV = 'Dlv.qty'

# --- State Management ---
def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                state = json.load(f)
                # Ensure structure is correct (migration for new feature)
                if 'processed_files' not in state: state['processed_files'] = []
                if 'delivery_data' not in state: state['delivery_data'] = {}
                if 'delivery_meta' not in state: state['delivery_meta'] = {}
                return state
        except:
            pass
    return {'processed_files': [], 'delivery_data': {}, 'delivery_meta': {}} 

def save_state(state):
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f)

# --- Data Loading ---
@st.cache_data
def load_initial_order(file_path_or_buffer):
    # Header is at row index 1 (0-based) based on inspection
    try:
        df = pd.read_excel(file_path_or_buffer, header=1)
        # Verify columns exist
        missing_cols = [c for c in COL_MAP_ORDER.keys() if c not in df.columns]
        if missing_cols:
            st.error(f"Липсващи колони във файла с поръчката: {missing_cols}")
            return pd.DataFrame()
        
        # Select and Rename
        df_clean = df[list(COL_MAP_ORDER.keys())].rename(columns=COL_MAP_ORDER)
        
        # Ensure Barcode is string for matching
        df_clean['Barcode'] = df_clean['Barcode'].astype(str).str.strip()
        
        # Group by Barcode to handle duplicates in order file (if any)
        # User said EAN UPC Cd is Primary Key, but let's be safe
        df_clean = df_clean.groupby('Barcode', as_index=False).agg({
            'Concatenate': 'first',
            'Description': 'first',
            'SizeConverted': 'first',
            'Ordered_Qty': 'sum'
        })
        
        return df_clean
    except Exception as e:
        st.error(f"Грешка при зареждане на файла с поръчката: {e}")
        return pd.DataFrame()

def process_delivery_file(uploaded_file, current_state):
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Dati_imp')
        
        # Check columns
        if COL_BARCODE_DELIV not in df.columns or COL_QTY_DELIV not in df.columns:
            st.error(f"Файлът {uploaded_file.name} няма задължителни колони: {COL_BARCODE_DELIV}, {COL_QTY_DELIV}")
            return current_state
            
        # Extract and Aggregate
        df['Barcode'] = df[COL_BARCODE_DELIV].astype(str).str.strip()
        
        # 1. Sum Quantities
        delivery_sums = df.groupby('Barcode')[COL_QTY_DELIV].sum().to_dict()
        
        # 2. Extract Metadata (Concatenate, Description, SizeConverted)
        # We take the first occurrence for each barcode
        # Ensure these columns exist in delivery file before accessing
        meta_cols = ['Concatenate', 'Description', 'SizeConverted']
        available_meta_cols = [c for c in meta_cols if c in df.columns]
        
        if available_meta_cols:
            delivery_meta = df.groupby('Barcode')[available_meta_cols].first().to_dict('index')
        else:
            delivery_meta = {}

        # Update State
        for barcode, qty in delivery_sums.items():
            current_state['delivery_data'][barcode] = current_state['delivery_data'].get(barcode, 0) + qty
            
            # Store metadata if available
            if barcode in delivery_meta:
                # Merge existing meta with new (new wins or keep old? keep old usually better for stability, but let's overwrite to be fresh)
                # Actually, check if we already have it
                if barcode not in current_state['delivery_meta']:
                     current_state['delivery_meta'][barcode] = delivery_meta[barcode]

        current_state['processed_files'].append(uploaded_file.name)
        save_state(current_state)
        st.success(f"Обработен {uploaded_file.name}")
        return current_state
        
    except Exception as e:
        st.error(f"Грешка при обработка на {uploaded_file.name}: {e}")
        return current_state

# --- Main App ---
st.set_page_config(page_title="Контрол на Стоките", layout="wide")
st.title("📦 Контрол на Пристигащи Стоки")

# 1. Load Order Data
st.sidebar.header("📁 Стъпка 1: Първоначална Поръчка")
order_file = st.sidebar.file_uploader("Качете файл 'Потвърждение'", type=['xlsx'])

# Logic to handle persistence of the Order File
if order_file:
    # Save the uploaded file locally to persist it
    with open(LOCAL_ORDER_PATH, "wb") as f:
        f.write(order_file.getbuffer())
    st.sidebar.success("Файлът с поръчката е запазен успешно!")
    df_order = load_initial_order(order_file)
elif os.path.exists(LOCAL_ORDER_PATH):
    st.sidebar.info("Използва се предишно качен файл с поръчка.")
    df_order = load_initial_order(LOCAL_ORDER_PATH)
else:
    st.info("👈 Моля, качете файла 'Happy Sport - Nike SP26 - Confirmation.xlsx' в страничната лента, за да започнете.")
    st.image("https://placehold.co/600x400?text=Waiting+for+Order+File", caption="Upload Verification File")
    st.stop()

if df_order.empty:
    st.warning("Неуспешно зареждане на валидни данни за поръчка.")
    st.stop()

# 2. Load State
if 'app_state' not in st.session_state:
    st.session_state.app_state = load_state()

# 3. Sidebar - File Upload
st.sidebar.markdown("---")
st.sidebar.header("📥 Стъпка 2: Качване на Доставки")
uploaded_files = st.sidebar.file_uploader("Качете Excel файл(ове) с доставки", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name not in st.session_state.app_state['processed_files']:
            st.session_state.app_state = process_delivery_file(uploaded_file, st.session_state.app_state)
        else:
            st.sidebar.info(f"Пропуснат {uploaded_file.name} (вече е обработен)")

# Show Processed Files History
if st.session_state.app_state['processed_files']:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Обработени Файлове")
    for fname in st.session_state.app_state['processed_files']:
        st.sidebar.text(f"✅ {fname}")

# 4. Sidebar - Reset
if st.sidebar.button("⚠️ Нулирай Всичко"):
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
    st.session_state.app_state = {'processed_files': [], 'delivery_data': {}, 'delivery_meta': {}}
    st.rerun()

# 5. Merge Data
delivery_map = st.session_state.app_state['delivery_data']
delivery_meta = st.session_state.app_state.get('delivery_meta', {})

# Only map delivery quantities that correspond to ordered items here
df_order['Delivered'] = df_order['Barcode'].map(delivery_map).fillna(0)
df_order['Remaining'] = df_order['Ordered_Qty'] - df_order['Delivered']

# Handle Unordered Items (Items in delivery but not in order)
ordered_barcodes = set(df_order['Barcode'])
delivered_barcodes = set(delivery_map.keys())
unordered_barcodes = list(delivered_barcodes - ordered_barcodes)

if unordered_barcodes:
    data_unordered = []
    for bc in unordered_barcodes:
        qty = delivery_map[bc]
        meta = delivery_meta.get(bc, {})
        
        data_unordered.append({
            'Barcode': bc,
            'Concatenate': meta.get('Concatenate', 'ИЗВЪН ПОРЪЧКА'),
            'Description': meta.get('Description', 'Непоръчан артикул'),
            'SizeConverted': meta.get('SizeConverted', '-'),
            'Ordered_Qty': 0,
            'Delivered': qty,
            'Remaining': -qty # Logical: 0 - Delivered
        })
    
    df_unordered = pd.DataFrame(data_unordered)
else:
    df_unordered = pd.DataFrame()

# Handle Excess Items (Ordered > 0 but Delivered > Ordered)
# Remaining is Negative
mask_excess = (df_order['Ordered_Qty'] > 0) & (df_order['Remaining'] < 0)
df_excess = df_order[mask_excess].copy()
df_others = df_order[~mask_excess].copy()

# Combine: 1. Unordered, 2. Excess, 3. Others
df_total = pd.concat([df_unordered, df_excess, df_others], ignore_index=True)

# 6. Styling
def highlight_row_bg(row):
    ordered = row['Поръчано']
    remaining = row['Оставащо']
    
    # Red for Unordered (Ordered == 0)
    if ordered == 0:
        return ['background-color: #f8d7da; color: red; font-weight: bold'] * len(row) # Red Text
    # Purple/Blue for Excess (Ordered > 0 AND Remaining < 0)
    elif ordered > 0 and remaining < 0:
        return ['background-color: #cce5ff; color: #004085'] * len(row) # Blue
    # Green for Completed (Remaining == 0) - Note: Remaining <= 0 covered Excess before, so strict check needed
    elif remaining == 0:
        return ['background-color: #d4edda; color: #155724'] * len(row) # Green
    # Yellow for Pending (Remaining > 0)
    else:
        return ['background-color: #fff3cd; color: #856404'] * len(row) # Yellow/Orange

# 7. Filter
st.sidebar.markdown("---")
st.sidebar.header("🔍 Филтриране на Данни")

# Filter: Special Categories
show_unordered = st.sidebar.checkbox("⚠️ Покажи ИЗВЪН поръчка (Red)", value=False)
show_excess = st.sidebar.checkbox("📈 Покажи НАДВИШЕНИ количества (Blue)", value=False)
show_pending = st.sidebar.checkbox("📉 Покажи НЕДОСТАВЕНИ (Yellow)", value=False)

# Text Search
search_text = st.sidebar.text_input("Търсене по Concatenate (Текст)")

# Multiselect
all_concats = sorted(df_total['Concatenate'].dropna().unique().tolist())
selected_concats = st.sidebar.multiselect("Изберете конкретен Concatenate", options=all_concats)

# Apply Filters
df_visible = df_total.copy()

# Logic for combining checkbox filters (OR logic if multiple selected to show inclusive sets, or specific? 
# Usually users want to toggle visibility of specific groups. 
# Let's make them additive: If ANY checkbox is checked, show items from those categories.
# If NO checkbox is checked, show ALL (default behavior).
# This is more intuitive for "Show me X" buttons.

filter_mask = pd.Series([False] * len(df_visible))
any_checkbox_active = show_unordered or show_excess or show_pending

if any_checkbox_active:
    if show_unordered:
        filter_mask |= (df_visible['Ordered_Qty'] == 0)
    if show_excess:
        filter_mask |= ((df_visible['Ordered_Qty'] > 0) & (df_visible['Remaining'] < 0))
    if show_pending:
        filter_mask |= (df_visible['Remaining'] > 0)
    
    df_visible = df_visible[filter_mask]

if search_text:
    df_visible = df_visible[df_visible['Concatenate'].astype(str).str.contains(search_text, case=False, na=False)]

if selected_concats:
    df_visible = df_visible[df_visible['Concatenate'].isin(selected_concats)]

# 8. Display Stats
st.subheader("Общ Преглед")

# Calculate Valid Metrics
# Total Ordered: Sum of Ordered_Qty (Unaffected by delivery)
total_ordered = df_visible['Ordered_Qty'].sum()

# Total Delivered (Planned): Sum of Delivered WHERE Ordered_Qty > 0
# Logic: We usually want to know how much of the PLAN was fulfilled.
# If I ordered 10 and got 12, I "fulfilled" 10. The extra 2 are excess.
# But "Total Delivered" in simple terms often just sums the column.
# Let's keep "Доставено" as the simple sum of everything that MATCHED an order line (including excess).
total_delivered_matching = df_visible[df_visible['Ordered_Qty'] > 0]['Delivered'].sum()

# Total Unordered: Sum of Delivered WHERE Ordered_Qty == 0
total_unordered = df_visible[df_visible['Ordered_Qty'] == 0]['Delivered'].sum()

# Total Excess Qty: Sum of (Delivered - Ordered) WHERE Ordered > 0 AND Delivered > Ordered
# This is equivalent to ABS(Remaining) where Remaining < 0 AND Ordered > 0
excess_mask = (df_visible['Ordered_Qty'] > 0) & (df_visible['Remaining'] < 0)
total_excess = df_visible[excess_mask]['Remaining'].abs().sum()

# Pending: Sum of Remaining WHERE Remaining > 0
total_pending = df_visible[df_visible['Remaining'] > 0]['Remaining'].sum()

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Общо Поръчано", int(total_ordered))
col2.metric("Доставено (Вкл. надвишени)", int(total_delivered_matching))
col3.metric("⚠️ Извън поръчка", int(total_unordered))
col4.metric("📈 Превишени к-ва", int(total_excess)) # Excess Quantities
col5.metric("Оставащо", int(total_pending))

st.subheader("Детален Статус")

# Maps columns to Bulgarian for display
display_cols_map = {
    'Concatenate': 'Concatenate',
    'SizeConverted': 'Размер',
    'Description': 'Описание',
    'Barcode': 'Баркод',
    'Ordered_Qty': 'Поръчано',
    'Delivered': 'Доставено',
    'Remaining': 'Оставащо'
}

cols_to_show = list(display_cols_map.keys())
df_display = df_visible[cols_to_show].rename(columns=display_cols_map)

# Apply Styling (Logic still uses 'Remaining', but display uses 'Оставащо')
# We need to apply style BEFORE renaming if we use column logic, 
# OR adjust the highlight function to use the new name.
# Easier to apply style to the original df_visible's subset, then rename? 
# Streamlit dataframe properties are tricky. 
# Let's adjust the highlight function to handle the original DF, 
# but st.dataframe expects the styled object to match the display.
# Strategy: Rename first, then style using new column names.

def highlight_row_bg(row):
    remaining = row['Оставащо']
    if remaining <= 0:
        return ['background-color: #d4edda; color: #155724'] * len(row) # Green
    else:
        return ['background-color: #fff3cd; color: #856404'] * len(row) # Yellow/Orange

st.dataframe(
    df_display.style.apply(highlight_row_bg, axis=1),
    use_container_width=True,
    height=800
)
