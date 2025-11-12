import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go

# Seiten-Konfiguration
st.set_page_config(
    page_title="Schulungsverwaltung",
    page_icon="📚",
    layout="wide"
)

# Titel
st.title("📚 Schulungsverwaltung - Matrix-Übersicht")

# Daten laden
@st.cache_data
def load_data():
    df_ausbildungen = pd.read_excel('Bericht_Ausbildungen_.xlsx')
    df_user = pd.read_excel('Bericht_User.xlsx')
    return df_ausbildungen, df_user

@st.cache_data
def process_data(df_ausbildungen, df_user):
    # Merge über Namen
    df_merged = df_ausbildungen.merge(
        df_user[['Name', 'Personalnummer']], 
        left_on='Teilnehmer', 
        right_on='Name', 
        how='left'
    )
    
    # Berechne Gültigkeitsdatum
    def calculate_gueltig_bis(row):
        if pd.notna(row['Gültig bis']):
            return row['Gültig bis']
        elif pd.notna(row['Datum der Durchführung']) and pd.notna(row['Intervall']) and row['Intervall'] > 0:
            return row['Datum der Durchführung'] + pd.DateOffset(months=int(row['Intervall']))
        return None
    
    df_merged['Gültig_bis_berechnet'] = df_merged.apply(calculate_gueltig_bis, axis=1)
    
    # Nur neueste Schulung pro User/Ausbildung behalten
    df_merged = df_merged.sort_values('Datum der Durchführung', ascending=False)
    df_latest = df_merged.groupby(['Personalnummer', 'Ausbildung (Bezeichnung)']).first().reset_index()
    
    # Status berechnen
    heute = pd.Timestamp.now()
    
    def get_status(gueltig_bis):
        if pd.isna(gueltig_bis):
            return pd.Series({
                'Status': 'Unklar',
                'StatusCode': 0,
                'Farbe': 'gray',
                'TageVerbleibend': None
            })
        
        tage = (gueltig_bis - heute).days
        
        if tage < 0:
            return pd.Series({
                'Status': 'Abgelaufen',
                'StatusCode': 3,
                'Farbe': '#ffc7ce',
                'TageVerbleibend': tage
            })
        elif tage <= 90:
            return pd.Series({
                'Status': 'Bald ablaufend',
                'StatusCode': 2,
                'Farbe': '#ffeb9c',
                'TageVerbleibend': tage
            })
        else:
            return pd.Series({
                'Status': 'Gültig',
                'StatusCode': 1,
                'Farbe': '#c6efce',
                'TageVerbleibend': tage
            })
    
    status_df = df_latest['Gültig_bis_berechnet'].apply(get_status)
    df_latest = pd.concat([df_latest, status_df], axis=1)
    
    return df_latest, df_user

try:
    df_ausbildungen, df_user = load_data()
    df_processed, df_user = process_data(df_ausbildungen, df_user)
    
    # Sidebar - Filter
    st.sidebar.header("🔍 Filter")
    
    # Status-Filter
    status_options = ['Alle'] + list(df_processed['Status'].unique())
    selected_status = st.sidebar.multiselect(
        "Status",
        options=status_options,
        default=['Alle']
    )
    
    # User-Filter
    user_options = ['Alle'] + sorted(df_user['Name'].dropna().unique().tolist())
    selected_user = st.sidebar.selectbox("Mitarbeiter", user_options)
    
    # Ausbildungs-Filter
    ausb_options = ['Alle'] + sorted(df_processed['Ausbildung (Bezeichnung)'].dropna().unique().tolist())
    selected_ausb = st.sidebar.selectbox("Ausbildung", ausb_options)
    
    # Daten filtern
    df_filtered = df_processed.copy()
    
    if 'Alle' not in selected_status:
        df_filtered = df_filtered[df_filtered['Status'].isin(selected_status)]
    
    if selected_user != 'Alle':
        pers_nr = df_user[df_user['Name'] == selected_user]['Personalnummer'].iloc[0]
        df_filtered = df_filtered[df_filtered['Personalnummer'] == pers_nr]
    
    if selected_ausb != 'Alle':
        df_filtered = df_filtered[df_filtered['Ausbildung (Bezeichnung)'] == selected_ausb]
    
    # Statistiken
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("👥 Mitarbeiter", len(df_user))
    
    with col2:
        st.metric("📚 Ausbildungen", df_processed['Ausbildung (Bezeichnung)'].nunique())
    
    with col3:
        abgelaufen = len(df_processed[df_processed['Status'] == 'Abgelaufen'])
        st.metric("🔴 Abgelaufen", abgelaufen)
    
    with col4:
        bald = len(df_processed[df_processed['Status'] == 'Bald ablaufend'])
        st.metric("🟡 Bald ablaufend", bald)
    
    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Matrix", "📋 Liste", "📈 Statistiken", "⚙️ Daten-Export"])
    
    with tab1:
        st.subheader("Schulungs-Matrix")
        
        # Erstelle Pivot-Tabelle
        pivot = df_processed.pivot_table(
            index='Teilnehmer',
            columns='Ausbildung (Bezeichnung)',
            values='StatusCode',
            aggfunc='first'
        )
        
        # Farb-Mapping
        def color_status(val):
            if pd.isna(val):
                return 'background-color: #f5f5f5'
            elif val == 1:
                return 'background-color: #c6efce; color: #006100; font-weight: bold'
            elif val == 2:
                return 'background-color: #ffeb9c; color: #9c5700; font-weight: bold'
            elif val == 3:
                return 'background-color: #ffc7ce; color: #9c0006; font-weight: bold'
            else:
                return 'background-color: #dcdcdc; color: #666666'
        
        # Symbol-Mapping
        def format_status(val):
            if pd.isna(val):
                return ''
            else:
                return '✓'
        
        # Zeige Matrix
        styled_pivot = pivot.style.applymap(color_status).format(format_status)
        
        st.dataframe(
            styled_pivot,
            use_container_width=True,
            height=600
        )
        
        # Legende
        st.markdown("---")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown("🟢 **Gültig** (>90 Tage)")
        with col2:
            st.markdown("🟡 **Bald ablaufend** (≤90 Tage)")
        with col3:
            st.markdown("🔴 **Abgelaufen**")
        with col4:
            st.markdown("⚪ **Unklar** (kein Datum)")
        with col5:
            st.markdown("◻️ **Nicht absolviert**")
    
    with tab2:
        st.subheader("Detaillierte Liste")
        
        # Zeige gefilterte Daten
        display_cols = [
            'Teilnehmer', 
            'Ausbildung (Bezeichnung)', 
            'Datum der Durchführung',
            'Gültig_bis_berechnet',
            'Status',
            'TageVerbleibend'
        ]
        
        df_display = df_filtered[display_cols].copy()
        df_display.columns = ['Mitarbeiter', 'Ausbildung', 'Durchführung', 'Gültig bis', 'Status', 'Tage verbleibend']
        
        # Formatiere Datumspalten
        df_display['Durchführung'] = pd.to_datetime(df_display['Durchführung']).dt.strftime('%d.%m.%Y')
        df_display['Gültig bis'] = pd.to_datetime(df_display['Gültig bis']).dt.strftime('%d.%m.%Y')
        
        # Färbe Zeilen nach Status
        def highlight_row(row):
            if row['Status'] == 'Gültig':
                return ['background-color: #c6efce'] * len(row)
            elif row['Status'] == 'Bald ablaufend':
                return ['background-color: #ffeb9c'] * len(row)
            elif row['Status'] == 'Abgelaufen':
                return ['background-color: #ffc7ce'] * len(row)
            else:
                return ['background-color: #dcdcdc'] * len(row)
        
        styled_df = df_display.style.apply(highlight_row, axis=1)
        
        st.dataframe(styled_df, use_container_width=True, height=600)
        
        st.markdown(f"**Anzeige: {len(df_display)} Einträge**")
    
    with tab3:
        st.subheader("Statistiken & Auswertungen")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Status-Verteilung
            status_counts = df_processed['Status'].value_counts()
            
            fig1 = px.pie(
                values=status_counts.values,
                names=status_counts.index,
                title="Status-Verteilung",
                color=status_counts.index,
                color_discrete_map={
                    'Gültig': '#c6efce',
                    'Bald ablaufend': '#ffeb9c',
                    'Abgelaufen': '#ffc7ce',
                    'Unklar': '#dcdcdc'
                }
            )
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # Top Ausbildungen nach Abgelaufen
            abgelaufen_per_ausb = df_processed[df_processed['Status'] == 'Abgelaufen'].groupby('Ausbildung (Bezeichnung)').size().sort_values(ascending=False).head(10)
            
            fig2 = px.bar(
                x=abgelaufen_per_ausb.values,
                y=abgelaufen_per_ausb.index,
                orientation='h',
                title="Top 10 Abgelaufene Schulungen",
                labels={'x': 'Anzahl', 'y': 'Ausbildung'},
                color_discrete_sequence=['#ffc7ce']
            )
            fig2.update_layout(height=400)
            st.plotly_chart(fig2, use_container_width=True)
        
        # Zeitstrahl - Ablaufende Schulungen
        st.subheader("Ablaufende Schulungen (nächste 180 Tage)")
        
        upcoming = df_processed[
            (df_processed['TageVerbleibend'] >= 0) & 
            (df_processed['TageVerbleibend'] <= 180)
        ].sort_values('Gültig_bis_berechnet')
        
        if len(upcoming) > 0:
            fig3 = px.scatter(
                upcoming,
                x='Gültig_bis_berechnet',
                y='Ausbildung (Bezeichnung)',
                color='Status',
                hover_data=['Teilnehmer', 'TageVerbleibend'],
                title="Zeitstrahl der ablaufenden Schulungen",
                color_discrete_map={
                    'Gültig': '#c6efce',
                    'Bald ablaufend': '#ffeb9c',
                    'Abgelaufen': '#ffc7ce'
                }
            )
            fig3.update_layout(height=600)
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("Keine ablaufenden Schulungen in den nächsten 180 Tagen")
    
    with tab4:
        st.subheader("Daten-Export")
        
        # Excel-Export
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📊 Matrix als Excel")
            
            @st.cache_data
            def convert_to_excel(df):
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pivot = df.pivot_table(
                        index='Teilnehmer',
                        columns='Ausbildung (Bezeichnung)',
                        values='Status',
                        aggfunc='first'
                    )
                    pivot.to_excel(writer, sheet_name='Matrix')
                    df.to_excel(writer, sheet_name='Details', index=False)
                return output.getvalue()
            
            excel_data = convert_to_excel(df_processed)
            
            st.download_button(
                label="📥 Excel herunterladen",
                data=excel_data,
                file_name=f"schulungsmatrix_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.markdown("### 📋 Liste als CSV")
            
            csv_data = df_processed.to_csv(index=False, encoding='utf-8-sig')
            
            st.download_button(
                label="📥 CSV herunterladen",
                data=csv_data,
                file_name=f"schulungen_liste_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        # JSON-Export der Dictionaries zum Debuggen
        st.markdown("### 🔍 Debug: Daten als JSON")
        
        if st.checkbox("JSON-Daten anzeigen"):
            import json
            
            # Erstelle Dictionary-Struktur
            user_dict = {}
            for _, row in df_processed.iterrows():
                pers_nr = str(row['Personalnummer'])
                if pers_nr not in user_dict:
                    user_dict[pers_nr] = {
                        'Name': row['Teilnehmer'],
                        'Schulungen': {}
                    }
                
                user_dict[pers_nr]['Schulungen'][row['Ausbildung (Bezeichnung)']] = {
                    'Status': row['Status'],
                    'StatusCode': int(row['StatusCode']) if pd.notna(row['StatusCode']) else None,
                    'Durchführung': row['Datum der Durchführung'].strftime('%Y-%m-%d') if pd.notna(row['Datum der Durchführung']) else None,
                    'GueltigBis': row['Gültig_bis_berechnet'].strftime('%Y-%m-%d') if pd.notna(row['Gültig_bis_berechnet']) else None,
                    'TageVerbleibend': int(row['TageVerbleibend']) if pd.notna(row['TageVerbleibend']) else None
                }
            
            st.json(user_dict)
            
            # Download als JSON
            json_str = json.dumps(user_dict, indent=2, ensure_ascii=False)
            st.download_button(
                label="📥 JSON herunterladen",
                data=json_str,
                file_name=f"schulungen_dict_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json"
            )

except FileNotFoundError:
    st.error("⚠️ Excel-Dateien nicht gefunden!")
    st.info("Bitte legen Sie die Dateien 'Bericht_Ausbildungen_.xlsx' und 'Bericht_User.xlsx' im gleichen Verzeichnis ab.")
except Exception as e:
    st.error(f"❌ Fehler beim Laden der Daten: {str(e)}")
    st.exception(e)
