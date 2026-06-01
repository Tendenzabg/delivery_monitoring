import streamlit as st
import pandas as pd
import os
import json
import datetime

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

# --- Logging Helper ---
def scrivi_log(messaggio):
    """
    Scrive un messaggio nel file di log 'output/log.txt' con data e ora.
    Crea la cartella 'output' se non esiste.
    
    Parametri:
        messaggio (str): Il messaggio descrittivo dell'errore da registrare.
        
    Ritorna:
        None
    """
    try:
        # Nome della cartella di output
        cartella_output = 'output'
        # Se la cartella di output non esiste, la crea
        if not os.path.exists(cartella_output):
            os.makedirs(cartella_output)
        # Percorso completo del file di log
        percorso_log = os.path.join(cartella_output, 'log.txt')
        # Ottiene la data e l'ora corrente formattata come stringa
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Apre il file in modalità append con codifica UTF-8
        with open(percorso_log, 'a', encoding='utf-8') as file_log:
            # Scrive la riga di log con data, ora e il messaggio
            file_log.write(f"[{timestamp}] {messaggio}\n")
    except Exception as e_log:
        # Stampa l'errore a schermo solo in caso di fallimento della scrittura su file
        print(f"Impossibile scrivere il log: {e_log}")

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
    """
    Carica e pulisce il file dell'ordine iniziale ("Conferma").
    Gestisce sia il file grezzo del fornitore (header riga 1) sia quello già salvato (header riga 0).
    
    Parametri:
        file_path_or_buffer (str o file-like): Il percorso del file o il buffer del file caricato.
        
    Ritorna:
        pd.DataFrame: Il dataframe pulito e aggregato per codice a barre.
    """
    # Blocco try/except per catturare eventuali errori durante la lettura dei file
    try:
        # Legge il file con header alla riga 0 per verificare se si tratta del file pulito e già salvato in locale
        df_primo_tentativo = pd.read_excel(file_path_or_buffer, header=0)
        # Elenco delle colonne richieste nel file d'ordine pulito e standardizzato
        colonne_pulite = ['Barcode', 'Concatenate', 'Description', 'SizeConverted', 'Ordered_Qty']
        # Verifica se tutte le colonne pulite sono presenti nel dataframe letto al primo tentativo
        if all(col in df_primo_tentativo.columns for col in colonne_pulite):
            # Converte la colonna Barcode in stringa ed elimina gli spazi bianchi iniziali e finali
            df_primo_tentativo['Barcode'] = df_primo_tentativo['Barcode'].astype(str).str.strip()
            # Ritorna il dataframe già formattato
            return df_primo_tentativo
        
        # Se non è un file pulito salvato, legge il file originale del fornitore con header alla riga 1
        df = pd.read_excel(file_path_or_buffer, header=1)
        # Controlla se ci sono colonne obbligatorie mancanti rispetto a quelle mappate
        colonne_mancanti = [c for c in COL_MAP_ORDER.keys() if c not in df.columns]
        # Se ci sono colonne mancanti
        if colonne_mancanti:
            # Mostra un messaggio di errore all'utente Streamlit in lingua bulgara
            st.error(f"Липсващи колони във файла с поръчката: {colonne_mancanti}")
            # Registra l'errore nel file di log in italiano
            scrivi_log(f"Errore caricamento ordine: Colonne mancanti {colonne_mancanti}")
            # Ritorna un dataframe vuoto per interrompere l'elaborazione del file errato
            return pd.DataFrame()
        
        # Filtra solo le colonne necessarie e le rinomina usando la mappatura standard
        df_clean = df[list(COL_MAP_ORDER.keys())].rename(columns=COL_MAP_ORDER)
        
        # Forza la colonna Barcode a tipo stringa e rimuove gli spazi per garantire corrispondenza corretta
        df_clean['Barcode'] = df_clean['Barcode'].astype(str).str.strip()
        
        # Raggruppa per codice a barre (Barcode) per sommare le quantità ed evitare righe duplicate
        df_clean = df_clean.groupby('Barcode', as_index=False).agg({
            'Concatenate': 'first',
            'Description': 'first',
            'SizeConverted': 'first',
            'Ordered_Qty': 'sum'
        })
        
        # Ritorna il dataframe dell'ordine elaborato e pulito
        return df_clean
    # Gestione delle eccezioni generiche
    except Exception as e:
        # Mostra l'errore nell'interfaccia utente in lingua bulgara
        st.error(f"Грешка при зареждане на файла с поръчката: {e}")
        # Scrive l'eccezione riscontrata nel file log in lingua italiana
        scrivi_log(f"Eccezione durante il caricamento del file dell'ordine: {str(e)}")
        # Ritorna un dataframe vuoto
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
        # Registra l'errore riscontrato durante l'elaborazione del file di consegna nel log
        scrivi_log(f"Errore nell'elaborazione del file di consegna {uploaded_file.name}: {str(e)}")
        return current_state

# --- Main App ---
st.set_page_config(page_title="Контрол на Стоките", layout="wide")
st.title("📦 Контрол на Пристигащи Стоки")

# 1. Carica i dati dell'ordine (Step 1)
# Titolo della sezione dell'ordine nella barra laterale (bulgaro)
st.sidebar.header("📁 Стъпка 1: Първоначална Поръчка")
# Consente il caricamento di file multipli di conferma impostando accept_multiple_files=True
order_files = st.sidebar.file_uploader("Качете файл(ове) 'Потвърждение'", type=['xlsx'], accept_multiple_files=True)

# Inizializza il dataframe per i dati dell'ordine
df_order = pd.DataFrame()

# Se l'utente ha caricato uno o più file dell'ordine
if order_files:
    # Crea una lista temporanea per raccogliere i dataframe puliti di ciascun file
    lista_df_ordini = []
    # Iteriamo su ciascun file caricato
    for file_ordine in order_files:
        # Carica il singolo file dell'ordine pulendone le colonne
        df_singolo = load_initial_order(file_ordine)
        # Se il dataframe del file corrente non è vuoto
        if not df_singolo.empty:
            # Lo aggiunge alla lista dei dataframe caricati
            lista_df_ordini.append(df_singolo)
    
    # Se abbiamo caricato con successo almeno un file valido
    if lista_df_ordini:
        # Unisce tutti i dataframe degli ordini caricati in un unico dataframe cumulativo
        df_combinato = pd.concat(lista_df_ordini, ignore_index=True)
        # Raggruppa per codice a barre (Barcode) per sommare le quantità ed eliminare i doppioni tra file
        df_order = df_combinato.groupby('Barcode', as_index=False).agg({
            'Concatenate': 'first',
            'Description': 'first',
            'SizeConverted': 'first',
            'Ordered_Qty': 'sum'
        })
        # Blocco try/except per salvare il dataframe combinato in locale
        try:
            # Salva il dataframe dell'ordine combinato come file Excel locale in sovrascrittura
            df_order.to_excel(LOCAL_ORDER_PATH, index=False)
            # Mostra messaggio di successo all'utente Streamlit in lingua bulgara
            st.sidebar.success("Файловете с поръчки са обединени и запазени успешно!")
        # Gestione di errori nel salvataggio del file Excel
        except Exception as e_salva:
            # Registra l'errore di scrittura locale nel log in italiano
            scrivi_log(f"Impossibile salvare il file combinato dell'ordine locale: {e_salva}")
            # Mostra avviso in lingua bulgara all'utente
            st.sidebar.warning("Грешка при локално запазване на обединения файл.")

# Se non ci sono file appena caricati, controlla se esiste una versione unita salvata in locale
elif os.path.exists(LOCAL_ORDER_PATH):
    # Mostra messaggio informativo in lingua bulgara indicando l'uso del file salvato
    st.sidebar.info("Използва се предишно качен файл с поръчка.")
    # Carica i dati dell'ordine dal file locale (essendo pre-pulito verrà riconosciuto da load_initial_order)
    df_order = load_initial_order(LOCAL_ORDER_PATH)

# Se non ci sono file caricati e non c'è alcun file salvato locale su cui appoggiarsi
else:
    # Mostra messaggio informativo in bulgaro che invita a caricare i file per iniziare
    st.info("👈 Моля, качете файл(ове) 'Потвърждение' в страничната лента, за да започнете.")
    # Visualizza un'immagine segnaposto per indicare l'attesa del file
    st.image("https://placehold.co/600x400?text=Waiting+for+Order+File", caption="Upload Verification File")
    # Ferma l'esecuzione dello script Streamlit finché non viene caricato un file
    st.stop()

# Se il dataframe dell'ordine combinato è comunque vuoto (es. file non validi o errori di parsing)
if df_order.empty:
    # Mostra un messaggio di avviso in lingua bulgara
    st.warning("Неуспешно зареждане на валидни данни за поръчка.")
    # Ferma l'esecuzione di Streamlit
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

# 4. Sidebar - Pulsante per azzerare e resettare lo stato e i file salvati
if st.sidebar.button("⚠️ Нулирай Всичко"):
    # Se il file dello stato cumulativo delle consegne esiste su disco
    if os.path.exists(STATE_FILE):
        # Elimina il file dello stato
        os.remove(STATE_FILE)
    # Se il file Excel dell'ordine locale salvato esiste su disco
    if os.path.exists(LOCAL_ORDER_PATH):
        # Elimina il file dell'ordine
        os.remove(LOCAL_ORDER_PATH)
    # Ripristina lo stato iniziale dell'applicazione all'interno della sessione di Streamlit
    st.session_state.app_state = {'processed_files': [], 'delivery_data': {}, 'delivery_meta': {}}
    # Esegue il rerun dell'applicazione per caricare la pagina aggiornata
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
        return ['background-color: #d4edda; color: #155724'] * len(row) # Green (Excess)
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
show_excess = st.sidebar.checkbox("📈 Покажи НАДВИШЕНИ количества (Green)", value=False)
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
