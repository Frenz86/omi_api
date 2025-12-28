import streamlit as st
import pandas as pd
import requests

# ============================================
# CONFIGURAZIONE PAGINA
# ============================================
st.set_page_config(
    page_title="Quotazioni OMI",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# STILE PERSONALIZZATO
# ============================================
st.markdown("""
<style>
    /* Stile generale */
    .main {
        padding: 1rem 2rem;
    }
    
    /* Header */
    .header-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    
    .header-title {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    
    .header-subtitle {
        font-size: 1.1rem;
        opacity: 0.9;
    }
    
    /* Cards */
    .info-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-left: 4px solid #667eea;
        margin-bottom: 1rem;
    }
    
    /* Tabella risultati */
    .result-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 1rem;
    }
    
    .result-table th {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 12px 15px;
        text-align: left;
        font-weight: 600;
    }
    
    .result-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #eee;
    }
    
    .result-table tr:hover {
        background-color: #f8f9ff;
    }
    
    /* Metriche */
    .metric-container {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        border-radius: 12px;
        padding: 1rem 1.5rem;
        text-align: center;
        margin-bottom: 1rem;
    }
    
    .metric-label {
        font-size: 0.85rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .metric-value {
        font-size: 1.3rem;
        font-weight: 700;
        color: #333;
    }
    
    /* Accordion personalizzato */
    .streamlit-expanderHeader {
        background-color: #f8f9ff;
        border-radius: 10px;
    }
    
    /* Bottoni */
    .stButton > button {
        width: 100%;
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
    }
    
    /* Selectbox */
    .stSelectbox > div > div {
        border-radius: 10px;
    }
    
    /* Nasconde footer Streamlit */
    footer {visibility: hidden;}
    
    /* Link mappa */
    .map-link {
        display: inline-block;
        padding: 12px 24px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        text-decoration: none;
        border-radius: 8px;
        font-weight: 600;
        margin: 10px 0;
        transition: all 0.3s ease;
    }
    
    .map-link:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        color: white;
    }
    
    /* Divider */
    .custom-divider {
        height: 3px;
        background: linear-gradient(90deg, #667eea, #764ba2, #667eea);
        border: none;
        border-radius: 2px;
        margin: 1.5rem 0;
    }
</style>
""", unsafe_allow_html=True)


# ============================================
# FUNZIONI UTILITY
# ============================================
def format_currency(value):
    """Formatta un valore come valuta euro"""
    if value is None:
        return "N/D"
    return f"€ {value:,.0f}".replace(",", ".")


@st.cache_data(ttl=3600)
def carica_dati_comuni():
    """Carica e prepara i dati dei comuni dal file Excel"""
    df_comuni = pd.read_excel("COD_CAT.xlsx")

    sigle_mancanti = {
        'Napoli': 'NA',
        'Casalnuovo di Napoli': 'NA',
        'Casola di Napoli': 'NA', 
        'Marano di Napoli': 'NA',
        'Melito di Napoli': 'NA',
        'Mugnano di Napoli': 'NA',
    }
    for comune, sigla in sigle_mancanti.items():
        df_comuni.loc[df_comuni['COMUNE'] == comune, 'Sigla automobilistica'] = sigla

    df_comuni = df_comuni.dropna(subset=['COMUNE', 'COD_CATASTO', 'Sigla automobilistica'])
    df_comuni['COMUNE_FULL'] = df_comuni['COMUNE'].astype(str) + " (" + df_comuni['Sigla automobilistica'].astype(str) + ")"
    comune_to_codice = dict(zip(df_comuni['COMUNE_FULL'], df_comuni['COD_CATASTO']))
    lista_comuni = sorted(df_comuni['COMUNE_FULL'].tolist())
    
    return comune_to_codice, lista_comuni


def geocode_indirizzo(indirizzo, comune=""):
    """Converte un indirizzo in coordinate GPS usando Nominatim"""
    url = "https://nominatim.openstreetmap.org/search"
    query = f"{indirizzo}, {comune}, Italia" if comune else f"{indirizzo}, Italia"
    
    params = {
        "q": query,
        "format": "json",
        "limit": 1,
        "addressdetails": 1
    }
    
    headers = {"User-Agent": "QuotazioniOMI/1.0"}
    
    try:
        response = requests.get(url, params=params, headers=headers, timeout=5)
        data = response.json()
        
        if data and len(data) > 0:
            lat = float(data[0]["lat"])
            lon = float(data[0]["lon"])
            return (lat, lon)
        return None
    except:
        return None


@st.cache_data(ttl=600)
def ottieni_zone_omi_comune(codice_comune, tipo_immobile="abitazioni_civili"):
    """Ottiene tutte le zone OMI disponibili per un comune"""
    url = "https://3eurotools.it/api-quotazioni-immobiliari-omi/ricerca"
    params = {
        "codice_comune": codice_comune,
        "tipo_immobile": tipo_immobile,
        "metri_quadri": 1
    }
    
    try:
        response = requests.get(url, params=params, timeout=10)
        data = response.json()
        zone = sorted(list(data.keys()))
        return zone
    except:
        return []


def cerca_tutti_valori_omi(codice_comune, zona_omi, mq):
    """Cerca le quotazioni OMI per TUTTI i tipi di immobile"""
    tipi_immobile = [
        'abitazioni_civili',
        'abitazioni_di_tipo_economico',
        'ville_e_villini',
        'negozi',
        'uffici',
        'capannoni_tipici',
        'box',
        'posti_auto_coperti'
    ]
    
    url = "https://3eurotools.it/api-quotazioni-immobiliari-omi/ricerca"
    dati_completi = {}
    
    for tipo in tipi_immobile:
        params = {
            "codice_comune": codice_comune,
            "metri_quadri": mq,
            "zona_omi": zona_omi.lower(),
            "tipo_immobile": tipo
        }
        
        try:
            response = requests.get(url, params=params, timeout=10)
            data = response.json()
            
            if tipo in data:
                dati_completi[tipo] = data[tipo]
        except:
            continue
    
    return dati_completi


# ============================================
# MAIN APP
# ============================================
def main():
    # Header
    st.markdown("""
    <div class="header-container">
        <div class="header-title">🏠 Quotazioni Immobiliari OMI</div>
        <div class="header-subtitle">Osservatorio del Mercato Immobiliare - Agenzia delle Entrate</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Carica dati comuni
    comune_to_codice, lista_comuni = carica_dati_comuni()
    
    # ============================================
    # SIDEBAR - PARAMETRI DI RICERCA
    # ============================================
    with st.sidebar:
        st.markdown("## 🔍 Parametri di Ricerca")
        st.markdown("---")
        
        # Selezione comune
        comune = st.selectbox(
            "📍 Comune",
            options=[""] + lista_comuni,
            index=0,
            help="Cerca e seleziona il comune di interesse"
        )
        
        # Carica zone OMI se comune selezionato
        zone_disponibili = []
        if comune:
            codice = comune_to_codice.get(comune)
            if codice:
                with st.spinner("Caricamento zone OMI..."):
                    zone_disponibili = ottieni_zone_omi_comune(codice)
        
        # Selezione zona OMI
        if zone_disponibili:
            zona_omi = st.selectbox(
                "🗺️ Zona OMI",
                options=[""] + zone_disponibili,
                index=0,
                help=f"{len(zone_disponibili)} zone disponibili per questo comune"
            )
        else:
            zona_omi = st.text_input(
                "🗺️ Zona OMI",
                placeholder="es. B1",
                help="Seleziona prima un comune per vedere le zone disponibili"
            )
        
        # Metri quadri
        mq = st.number_input(
            "📐 Superficie (mq)",
            min_value=1,
            max_value=10000,
            value=100,
            step=5,
            help="Inserisci la superficie dell'immobile in metri quadri"
        )
        
        st.markdown("---")
        
        # Bottone ricerca
        cerca = st.button("🔍 Cerca Quotazioni", type="primary", use_container_width=True)
        
        st.markdown("---")
        
        # Sezione geolocalizzazione
        with st.expander("📍 Trova zona da indirizzo"):
            indirizzo = st.text_input(
                "Indirizzo",
                placeholder="es. Via Rizzoli 1",
                help="Inserisci via e numero civico"
            )
            
            geocode_btn = st.button("🗺️ Geolocalizza", use_container_width=True)
            
            if geocode_btn and indirizzo and comune:
                nome_comune = comune.split(" (")[0] if " (" in comune else comune
                coordinate = geocode_indirizzo(indirizzo, nome_comune)
                
                if coordinate:
                    lat, lon = coordinate
                    st.success(f"✅ Trovato!")
                    st.code(f"Lat: {lat:.6f}\nLon: {lon:.6f}")
                    
                    
                    geopoi_url = "https://www1.agenziaentrate.gov.it/servizi/geopoi_omi/index.php"
                    st.markdown(f"""
                    <a href="{geopoi_url}" target="_blank" class="map-link">
                        🏛️ Apri GeoPOI Agenzia Entrate
                    </a>
                    """, unsafe_allow_html=True)
                    
                    st.info("📋 Coordinate per GeoPOI (clicca l'icona per copiare):")
                    st.code(f"{lon:.6f}, {lat:.6f}")
                else:
                    st.error("❌ Indirizzo non trovato")
            elif geocode_btn and not comune:
                st.warning("⚠️ Seleziona prima un comune")
    
    # ============================================
    # AREA PRINCIPALE - RISULTATI
    # ============================================
    
    # Info box iniziale
    if not cerca and not (comune and zona_omi):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div class="info-card">
                <h3>📍 1. Seleziona il Comune</h3>
                <p>Scegli il comune dalla barra laterale. Puoi cercare digitando il nome.</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="info-card">
                <h3>🗺️ 2. Scegli la Zona OMI</h3>
                <p>Le zone OMI disponibili verranno caricate automaticamente. Usa la geolocalizzazione se non conosci la zona.</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class="info-card">
                <h3>🔍 3. Cerca le Quotazioni</h3>
                <p>Visualizza i prezzi di acquisto e affitto per tutti i tipi di immobile nella zona selezionata.</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<div class='custom-divider'></div>", unsafe_allow_html=True)
        
        st.info("👈 Utilizza la barra laterale per iniziare la ricerca delle quotazioni OMI")
    
    # Esegui ricerca
    if cerca:
        if not comune:
            st.error("⚠️ Seleziona un comune")
        elif not zona_omi:
            st.error("⚠️ Inserisci la zona OMI")
        else:
            codice_comune = comune_to_codice.get(comune)
            
            if not codice_comune:
                st.error("❌ Codice comune non trovato")
            else:
                with st.spinner("🔍 Ricerca quotazioni in corso..."):
                    dati = cerca_tutti_valori_omi(codice_comune, zona_omi, mq)
                
                if not dati:
                    st.error(f"❌ Nessun dato trovato per la zona **{zona_omi.upper()}** in {comune}")
                else:
                    # Header risultati
                    st.markdown("<div class='custom-divider'></div>", unsafe_allow_html=True)
                    
                    # Info riepilogo
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">📍 Comune</div>
                            <div class="metric-value">{comune}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">🗺️ Zona OMI</div>
                            <div class="metric-value">{zona_omi.upper()}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">📐 Superficie</div>
                            <div class="metric-value">{mq} mq</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("<div class='custom-divider'></div>", unsafe_allow_html=True)
                    
                    # Tabs per Acquisto e Affitto
                    tab1, tab2, tab3 = st.tabs(["💰 Acquisto", "🔑 Affitto", "📊 Tutti i Dati"])
                    
                    # Prepara dati per tabelle
                    rows = []
                    for tipo, valori in dati.items():
                        nome = tipo.replace('_', ' ').title()
                        rows.append({
                            "Tipologia": nome,
                            "Acquisto Min": valori.get('prezzo_acquisto_min'),
                            "Acquisto Medio": valori.get('prezzo_acquisto_medio'),
                            "Acquisto Max": valori.get('prezzo_acquisto_max'),
                            "Affitto Min": valori.get('prezzo_affitto_min'),
                            "Affitto Medio": valori.get('prezzo_affitto_medio'),
                            "Affitto Max": valori.get('prezzo_affitto_max'),
                        })
                    
                    df = pd.DataFrame(rows)
                    
                    with tab1:
                        st.markdown("### 💰 Prezzi di Acquisto")
                        st.markdown("*Valori riferiti all'immobile completo*")
                        
                        df_acquisto = df[["Tipologia", "Acquisto Min", "Acquisto Medio", "Acquisto Max"]].copy()
                        df_acquisto.columns = ["Tipologia", "Prezzo Minimo", "Prezzo Medio", "Prezzo Massimo"]
                        
                        # Formatta i valori
                        for col in ["Prezzo Minimo", "Prezzo Medio", "Prezzo Massimo"]:
                            df_acquisto[col] = df_acquisto[col].apply(format_currency)
                        
                        st.dataframe(
                            df_acquisto,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Tipologia": st.column_config.TextColumn("🏠 Tipologia", width="large"),
                                "Prezzo Minimo": st.column_config.TextColumn("💶 Minimo", width="medium"),
                                "Prezzo Medio": st.column_config.TextColumn("💰 Medio", width="medium"),
                                "Prezzo Massimo": st.column_config.TextColumn("💎 Massimo", width="medium"),
                            }
                        )
                    
                    with tab2:
                        st.markdown("### 🔑 Canoni di Affitto")
                        st.markdown("*Valori riferiti al canone mensile*")
                        
                        df_affitto = df[["Tipologia", "Affitto Min", "Affitto Medio", "Affitto Max"]].copy()
                        df_affitto.columns = ["Tipologia", "Canone Minimo", "Canone Medio", "Canone Massimo"]
                        
                        # Formatta i valori
                        for col in ["Canone Minimo", "Canone Medio", "Canone Massimo"]:
                            df_affitto[col] = df_affitto[col].apply(format_currency)
                        
                        st.dataframe(
                            df_affitto,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Tipologia": st.column_config.TextColumn("🏠 Tipologia", width="large"),
                                "Canone Minimo": st.column_config.TextColumn("💶 Minimo", width="medium"),
                                "Canone Medio": st.column_config.TextColumn("💰 Medio", width="medium"),
                                "Canone Massimo": st.column_config.TextColumn("💎 Massimo", width="medium"),
                            }
                        )
                    
                    with tab3:
                        st.markdown("### 📊 Riepilogo Completo")
                        
                        df_completo = df.copy()
                        for col in df_completo.columns:
                            if col != "Tipologia":
                                df_completo[col] = df_completo[col].apply(format_currency)
                        
                        st.dataframe(
                            df_completo,
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        # Download buttons
                        col_dl1, col_dl2 = st.columns(2)
                        
                        with col_dl1:
                            # Download CSV
                            csv = df.to_csv(index=False)
                            st.download_button(
                                label="📥 Scarica CSV",
                                data=csv,
                                file_name=f"quotazioni_omi_{zona_omi}_{comune.split(' (')[0]}.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                        
                        with col_dl2:
                            # Download Excel
                            from io import BytesIO
                            buffer = BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                df.to_excel(writer, sheet_name='Quotazioni OMI', index=False)
                            excel_data = buffer.getvalue()
                            
                            st.download_button(
                                label="📥 Scarica Excel",
                                data=excel_data,
                                file_name=f"quotazioni_omi_{zona_omi}_{comune.split(' (')[0]}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    
                    # Note finali
                    st.markdown("<div class='custom-divider'></div>", unsafe_allow_html=True)
                    st.caption("📌 I dati sono forniti dall'Osservatorio del Mercato Immobiliare dell'Agenzia delle Entrate. I prezzi di acquisto si riferiscono all'immobile completo, i canoni di affitto sono mensili.")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #888; font-size: 0.85rem;'>"
        "🏠 Quotazioni OMI • Dati Osservatorio Mercato Immobiliare • Agenzia delle Entrate"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
