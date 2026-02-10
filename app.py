import streamlit as st
import pandas as pd
import os
import json

# --- Configuration ---
DATA_DIR = '/Users/stoyantodorov/Downloads/Consegne Happy - Sporttime'
ORDER_FILE = os.path.join(DATA_DIR, 'Happy Sport - Nike SP26 - Confirmation.xlsx')
STATE_FILE = os.path.join(DATA_DIR, 'app_state.json')

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
def load_initial_order(file_path):
    # Header is at row index 1 (0-based) based on inspection
    try:
        df = pd.read_excel(file_path, header=1)
        # Verify columns exist
        missing_cols = [c for c in COL_MAP_ORDER.keys() if c not in df.columns]
        if missing_cols:
            st.error(f"Missing columns in Order file: {missing_cols}")
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
        st.error(f"Error loading Order file: {e}")
        return pd.DataFrame()

def process_delivery_file(uploaded_file, current_state):
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Dati_imp')
        
        # Check columns
        if COL_BARCODE_DELIV not in df.columns or COL_QTY_DELIV not in df.columns:
            st.error(f"File {uploaded_file.name} missing required columns: {COL_BARCODE_DELIV}, {COL_QTY_DELIV}")
            return current_state
            
        # Extract and Aggregate
        df['Barcode'] = df[COL_BARCODE_DELIV].astype(str).str.strip()
        delivery_sums = df.groupby('Barcode')[COL_QTY_DELIV].sum().to_dict()
        
        # Update State
        for barcode, qty in delivery_sums.items():
            current_state['delivery_data'][barcode] = current_state['delivery_data'].get(barcode, 0) + qty
            
        current_state['processed_files'].append(uploaded_file.name)
        save_state(current_state)
        st.success(f"Processed {uploaded_file.name}")
        return current_state
        
    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {e}")
        return current_state

# --- Main App ---
st.set_page_config(page_title="Merchandise Control", layout="wide")
st.title("📦 Merchandise Arrival Control")

# 1. Load Order Data
df_order = load_initial_order(ORDER_FILE)

if df_order.empty:
    st.warning("Please ensure the 'Confirmation' file is in the correct directory.")
    st.stop()

# 2. Load State
if 'app_state' not in st.session_state:
    st.session_state.app_state = load_state()

# 3. Sidebar - File Upload
st.sidebar.header("📥 Upload Deliveries")
uploaded_files = st.sidebar.file_uploader("Upload Delivery Excel", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.name not in st.session_state.app_state['processed_files']:
            st.session_state.app_state = process_delivery_file(uploaded_file, st.session_state.app_state)
        else:
            st.sidebar.info(f"Skipped {uploaded_file.name} (already processed)")

# Show Processed Files History
if st.session_state.app_state['processed_files']:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Processed Files")
    for fname in st.session_state.app_state['processed_files']:
        st.sidebar.text(f"✅ {fname}")

# 4. Sidebar - Reset
if st.sidebar.button("⚠️ Reset All Progress"):
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
st.sidebar.header("🔍 Filter Data")

# Text Search
search_text = st.sidebar.text_input("Search Concatenate (Text Match)")

# Multiselect
all_concats = sorted(df_order['Concatenate'].dropna().unique().tolist())
selected_concats = st.sidebar.multiselect("Select Specific Concatenate", options=all_concats)

# Apply Filters
df_visible = df_order.copy()

if search_text:
    df_visible = df_visible[df_visible['Concatenate'].astype(str).str.contains(search_text, case=False, na=False)]

if selected_concats:
    df_visible = df_visible[df_visible['Concatenate'].isin(selected_concats)]

# 8. Display Stats
st.subheader("Overview")
col1, col2, col3 = st.columns(3)
col1.metric("Total Ordered", int(df_visible['Ordered_Qty'].sum()))
col2.metric("Total Delivered", int(df_visible['Delivered'].sum()))
col3.metric("Pending", int(df_visible['Remaining'].sum()))

st.subheader("Detailed Status")

# Column Order
cols_to_show = ['Concatenate', 'SizeConverted', 'Description', 'Barcode', 'Ordered_Qty', 'Delivered', 'Remaining']
df_display = df_visible[cols_to_show]

# Apply Styling
st.dataframe(
    df_display.style.apply(highlight_row, axis=1),
    use_container_width=True,
    height=800
)
