# 📚 Schulungsverwaltung - Streamlit App

Eine interaktive Webanwendung zur Verwaltung und Überwachung von Mitarbeiter-Schulungen.

## ✨ Features

- 📊 **Matrix-Ansicht** - Übersichtliche Darstellung aller Schulungen mit Farbcodierung
- 📋 **Detail-Listen** - Filterbarer Tabellenansicht mit allen Informationen
- 📈 **Statistiken** - Interaktive Grafiken und Auswertungen
- 🔍 **Filter** - Nach Status, Mitarbeiter oder Ausbildung filtern
- 📥 **Export** - Als Excel, CSV oder JSON herunterladen

## 🎨 Status-System

- 🟢 **Grün** = Gültig (>90 Tage verbleibend)
- 🟡 **Gelb** = Bald ablaufend (≤90 Tage)
- 🔴 **Rot** = Abgelaufen
- ⚪ **Grau** = Unklar (kein Gültigkeitsdatum)
- ◻️ **Leer** = Nicht absolviert

## 🚀 Verwendung

### Online (Streamlit Cloud)
Die App läuft direkt im Browser - keine Installation nötig!

### Lokal
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## 📊 Datenstruktur

Die App benötigt zwei Excel-Dateien:

1. **Bericht_User.xlsx** - Mitarbeiterliste
   - Personalnummer
   - Name
   - Weitere Mitarbeiter-Details

2. **Bericht_Ausbildungen_.xlsx** - Schulungsdaten
   - Teilnehmer
   - Ausbildung (Bezeichnung)
   - Datum der Durchführung
   - Intervall
   - Gültig bis

## 🔧 Technologie

- Python 3.8+
- Streamlit
- Pandas
- Plotly
- OpenPyXL

## 📝 Lizenz

MIT License

## 👤 Autor

Erstellt mit Claude 🤖
