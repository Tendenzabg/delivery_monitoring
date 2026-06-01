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
    """
    Carica lo stato dell'applicazione dal file JSON locale.
    Se il file non esiste o è corrotto, restituisce lo stato iniziale vuoto.
    
    Ritorna:
        dict: Lo stato dell'applicazione con tutte le chiavi necessarie.
    """
    try:
        # Se il file dello stato esiste su disco
        if os.path.exists(STATE_FILE):
            # Apriamo il file JSON in modalità lettura
            with open(STATE_FILE, 'r', encoding='utf-8') as file_stato:
                stato = json.load(file_stato)
                # Assicuriamo che tutte le liste e i dizionari per consegne e ordini siano presenti (migrazione e stabilità)
                if 'processed_files' not in stato: stato['processed_files'] = []
                if 'delivery_data' not in stato: stato['delivery_data'] = {}
                if 'delivery_meta' not in stato: stato['delivery_meta'] = {}
                if 'processed_order_files' not in stato: stato['processed_order_files'] = []
                if 'order_data' not in stato: stato['order_data'] = {}
                if 'order_meta' not in stato: stato['order_meta'] = {}
                # Ritorna lo stato caricato e validato
                return stato
    except Exception as e_carica_stato:
        # Logga l'errore in italiano
        scrivi_log(f"Errore nel caricamento dello stato dell'app: {e_carica_stato}")
    # In caso di errore o se il file non esiste, ritorna lo schema vuoto iniziale
    return {
        'processed_files': [],
        'delivery_data': {},
        'delivery_meta': {},
        'processed_order_files': [],
        'order_data': {},
        'order_meta': {}
    }

def save_state(state):
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f)

def process_order_file(uploaded_file, stato_corrente):
    """
    Elabora un singolo file Excel di conferma dell'ordine ("Потвърждение"),
    estrae e pulisce i dati relativi alle quantità ordinate e ai metadati
    e li somma/accumula nello stato dell'applicazione.
    
    Parametri:
        uploaded_file: Il file Excel caricato da Streamlit.
        stato_corrente (dict): Lo stato corrente dell'applicazione.
        
    Ritorna:
        dict: Lo stato dell'applicazione aggiornato con i nuovi dati dell'ordine.
    """
    try:
        # Legge il file dell'ordine con l'intestazione posizionata alla riga 1 (seconda riga)
        dataframe_grezzo = pd.read_excel(uploaded_file, header=1)
        
        # Controlla quali colonne richieste mancano nel file caricato
        colonne_mancanti = [c for c in COL_MAP_ORDER.keys() if c not in dataframe_grezzo.columns]
        # Se ci sono colonne obbligatorie mancanti
        if colonne_mancanti:
            # Mostra una notifica di errore in bulgaro all'utente
            st.error(f"Файлът {uploaded_file.name} няма задължителни колони: {colonne_mancanti}")
            # Salva l'evento nel file di log in italiano
            scrivi_log(f"Errore caricamento ordine {uploaded_file.name}: colonne mancanti {colonne_mancanti}")
            # Ritorna lo stato senza modifiche
            return stato_corrente
            
        # Seleziona solo le colonne di interesse e le rinomina con i nomi standard
        dataframe_pulito = dataframe_grezzo[list(COL_MAP_ORDER.keys())].rename(columns=COL_MAP_ORDER)
        
        # Converte la colonna dei codici a barre in stringa e rimuove eventuali spazi bianchi ai lati
        dataframe_pulito['Barcode'] = dataframe_pulito['Barcode'].astype(str).str.strip()
        
        # Somma le quantità ordinate raggruppando per codice a barre per questo specifico file
        somme_barcode = dataframe_pulito.groupby('Barcode')['Ordered_Qty'].sum().to_dict()
        
        # Estrae i metadati descrittivi associati a ciascun codice a barre (prima occorrenza per ogni barcode)
        colonne_metadati = ['Concatenate', 'Description', 'SizeConverted']
        metadati_barcode = dataframe_pulito.groupby('Barcode')[colonne_metadati].first().to_dict('index')
        
        # Cicla sulle quantità ordinate estratte dal file corrente
        for codice_a_barre, quantita in somme_barcode.items():
            # Aggiunge/somma la quantità a quella già eventualmente presente nello stato cumulativo
            stato_corrente['order_data'][codice_a_barre] = stato_corrente['order_data'].get(codice_a_barre, 0) + quantita
            
            # Se ci sono metadati associati a questo codice a barre nel file corrente
            if codice_a_barre in metadati_barcode:
                # Memorizza o aggiorna i metadati nello stato
                stato_corrente['order_meta'][codice_a_barre] = metadati_barcode[codice_a_barre]
                
        # Aggiunge il nome del file all'elenco dei file d'ordine già elaborati nello stato
        stato_corrente['processed_order_files'].append(uploaded_file.name)
        
        # Salva lo stato aggiornato su file JSON locale
        save_state(stato_corrente)
        # Mostra messaggio di successo all'utente Streamlit in lingua bulgara
        st.success(f"Обработен {uploaded_file.name} (Потвърждение)")
        # Ritorna lo stato aggiornato
        return stato_corrente
        
    except Exception as errore_lettura:
        # Mostra un messaggio di errore a schermo in bulgaro
        st.error(f"Грешка при обработка на {uploaded_file.name}: {errore_lettura}")
        # Registra il dettaglio dell'eccezione riscontrata nel file log in italiano
        scrivi_log(f"Errore nella lettura del file d'ordine {uploaded_file.name}: {str(errore_lettura)}")
        # Ritorna lo stato corrente inalterato
        return stato_corrente

def genera_df_ordine_da_stato(stato_corrente):
    """
    Costruisce e restituisce un DataFrame pandas aggregato a partire dalle
    informazioni sull'ordine memorizzate nello stato corrente dell'applicazione.
    
    Parametri:
        stato_corrente (dict): Lo stato corrente dell'applicazione.
        
    Ritorna:
        pd.DataFrame: Il DataFrame dell'ordine consolidato pronto per il confronto.
    """
    try:
        # Verifica se lo stato contiene dati dell'ordine
        if not stato_corrente.get('order_data'):
            # Ritorna un dataframe vuoto se non c'è nulla registrato
            return pd.DataFrame()
            
        lista_righe = []
        # Cicla su ciascun codice a barre registrato nello stato dell'ordine
        for codice_a_barre, quantita in stato_corrente['order_data'].items():
            # Recupera i metadati descrittivi associati a questo codice a barre
            metadati = stato_corrente['order_meta'].get(codice_a_barre, {})
            # Aggiunge un dizionario rappresentante la riga del dataframe alla lista temporanea
            lista_righe.append({
                'Barcode': codice_a_barre,
                'Concatenate': metadati.get('Concatenate', 'SCONOSCIUTO'),
                'Description': metadati.get('Description', 'Articolo d\'ordine cumulato'),
                'SizeConverted': metadati.get('SizeConverted', '-'),
                'Ordered_Qty': quantita
            })
            
        # Ritorna il dataframe creato a partire dall'elenco di righe
        return pd.DataFrame(lista_righe)
    except Exception as errore_generazione:
        # Scrive l'errore riscontrato nel log in italiano
        scrivi_log(f"Errore nella generazione del DataFrame dell'ordine dallo stato: {errore_generazione}")
        # Ritorna un dataframe vuoto
        return pd.DataFrame()

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

# Inizializziamo lo stato dell'applicazione se non è già presente nella sessione di Streamlit (caricato subito)
if 'app_state' not in st.session_state:
    st.session_state.app_state = load_state()

# 1. Carica i dati dell'ordine (Step 1)
# Titolo della sezione dell'ordine nella barra laterale (bulgaro)
st.sidebar.header("📁 Стъпка 1: Първоначална Поръчка")
# Consente il caricamento di file multipli di conferma impostando accept_multiple_files=True
order_files = st.sidebar.file_uploader("Качете файл(ове) 'Потвърждение'", type=['xlsx'], accept_multiple_files=True)

# Se l'utente ha inserito uno o più file dell'ordine
if order_files:
    # Flag per verificare se abbiamo processato nuovi file non ancora registrati
    nuovi_file_processati = False
    # Iteriamo su ciascun file caricato
    for file_ordine in order_files:
        # Se il file non è già presente nella lista dei file di conferma elaborati
        if file_ordine.name not in st.session_state.app_state['processed_order_files']:
            # Elabora il file dell'ordine e aggiorna lo stato cumulativo
            st.session_state.app_state = process_order_file(file_ordine, st.session_state.app_state)
            # Segnala che c'è stato almeno un nuovo inserimento
            nuovi_file_processati = True
        else:
            # Informa l'utente che il file è già stato elaborato in precedenza
            st.sidebar.info(f"Пропуснат {file_ordine.name} (вече е обработен)")
            
    # Se sono stati aggiunti nuovi dati all'ordine in questo ciclo
    if nuovi_file_processati:
        # Rigeneriamo il DataFrame cumulativo dallo stato aggiornato
        df_consolidato = genera_df_ordine_da_stato(st.session_state.app_state)
        # Se il DataFrame non è vuoto, lo salviamo in locale su file Excel
        if not df_consolidato.empty:
            try:
                # Salva il file consolidato aggiornato
                df_consolidato.to_excel(LOCAL_ORDER_PATH, index=False)
            except Exception as e_salva_excel:
                # Logga l'errore in italiano
                scrivi_log(f"Impossibile salvare il file consolidato Excel in locale: {e_salva_excel}")

# Se non ci sono file appena inseriti nel file uploader ma lo stato contiene dati dell'ordine
if st.session_state.app_state.get('order_data'):
    # Genera il DataFrame direttamente dallo stato per essere confrontato
    df_order = genera_df_ordine_da_stato(st.session_state.app_state)

# Se lo stato dell'ordine è vuoto ma esiste il file Excel locale (recupero/persistenza all'avvio)
elif os.path.exists(LOCAL_ORDER_PATH):
    # Mostra messaggio informativo in lingua bulgara indicando l'uso del file salvato
    st.sidebar.info("Използва се предишно качен файл с поръчка.")
    # Carica i dati dell'ordine dal file Excel locale consolidato
    df_recuperato = load_initial_order(LOCAL_ORDER_PATH)
    # Se il caricamento è andato a buon fine
    if not df_recuperato.empty:
        # Ripopoliamo lo stato con i dati letti dal file Excel locale
        for _, riga in df_recuperato.iterrows():
            bc = str(riga['Barcode']).strip()
            qta = riga['Ordered_Qty']
            st.session_state.app_state['order_data'][bc] = qta
            st.session_state.app_state['order_meta'][bc] = {
                'Concatenate': riga['Concatenate'],
                'Description': riga['Description'],
                'SizeConverted': riga['SizeConverted']
            }
        # Aggiungiamo un nome fittizio se non presente per indicare l'avvenuto recupero
        if "saved_order.xlsx" not in st.session_state.app_state['processed_order_files']:
            st.session_state.app_state['processed_order_files'].append("saved_order.xlsx")
        # Salviamo lo stato aggiornato su file JSON
        save_state(st.session_state.app_state)
    # Assegna il DataFrame recuperato a df_order
    df_order = df_recuperato

# Se lo stato dell'ordine è vuoto e non c'è nessun file Excel locale da recuperare
else:
    # Mostra messaggio informativo in bulgaro che invita a caricare i file per iniziare
    st.info("👈 Моля, качете файл(ове) 'Потвърждение' в страничната лента, за да започнете.")
    # Visualizza un'immagine segnaposto per indicare l'attesa del file
    st.image("https://placehold.co/600x400?text=Waiting+for+Order+File", caption="Upload Verification File")
    # Ferma l'esecuzione dello script Streamlit finché non viene caricato un file
    st.stop()

# Se il dataframe dell'ordine consolidato è comunque vuoto (es. file non validi o errori di parsing)
if df_order.empty:
    # Mostra un messaggio di avviso in lingua bulgara
    st.warning("Неуспешно зареждане на валидни данни за поръчка.")
    # Ferma l'esecuzione di Streamlit
    st.stop()

# 2. Lo stato dell'applicazione è già caricato all'inizio del file
pass

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

# Mostra la cronologia dei file d'ordine elaborati (Потвърждение)
if st.session_state.app_state.get('processed_order_files'):
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Обработени Поръчки")
    for nome_file in st.session_state.app_state['processed_order_files']:
        # Esclude la visualizzazione del file di ripristino locale fittizio per pulizia visiva
        if nome_file != "saved_order.xlsx":
            st.sidebar.text(f"✅ {nome_file}")

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
    # Ripristina lo stato iniziale completo dell'applicazione all'interno della sessione di Streamlit
    st.session_state.app_state = {
        'processed_files': [],
        'delivery_data': {},
        'delivery_meta': {},
        'processed_order_files': [],
        'order_data': {},
        'order_meta': {}
    }
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
