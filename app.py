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
                return json.load(f)
        except:
            pass
    return {'processed_files': [], 'delivery_data': {}} # delivery_data: {barcode: qty}

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
        delivery_sums = df.groupby('Barcode')[COL_QTY_DELIV].sum().to_dict()
        
        # Update State
        for barcode, qty in delivery_sums.items():
            current_state['delivery_data'][barcode] = current_state['delivery_data'].get(barcode, 0) + qty
            
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
    st.session_state.app_state = {'processed_files': [], 'delivery_data': {}}
    st.rerun()

# 5. Merge Data
delivery_map = st.session_state.app_state['delivery_data']
df_order['Delivered'] = df_order['Barcode'].map(delivery_map).fillna(0)
df_order['Remaining'] = df_order['Ordered_Qty'] - df_order['Delivered']

# 6. Styling
def highlight_row(row):
    remaining = row['Remaining']
    if remaining <= 0:
        return ['background-color: #d4edda; color: #155724'] * len(row) # Green
    else:
        return ['background-color: #fff3cd; color: #856404'] * len(row) # Yellow/Orange

# 7. Filter
st.sidebar.markdown("---")
st.sidebar.header("🔍 Филтриране на Данни")

# Text Search
search_text = st.sidebar.text_input("Търсене по Concatenate (Текст)")

# Multiselect
all_concats = sorted(df_order['Concatenate'].dropna().unique().tolist())
selected_concats = st.sidebar.multiselect("Изберете конкретен Concatenate", options=all_concats)

# Apply Filters
df_visible = df_order.copy()

if search_text:
    df_visible = df_visible[df_visible['Concatenate'].astype(str).str.contains(search_text, case=False, na=False)]

if selected_concats:
    df_visible = df_visible[df_visible['Concatenate'].isin(selected_concats)]

# 8. Display Stats
st.subheader("Общ Преглед")
col1, col2, col3 = st.columns(3)
col1.metric("Общо Поръчано", int(df_visible['Ordered_Qty'].sum()))
col2.metric("Общо Доставено", int(df_visible['Delivered'].sum()))
col3.metric("Оставащо", int(df_visible['Remaining'].sum()))

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
