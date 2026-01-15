# -*- coding: utf-8 -*-
"""
Created on Fri Dec 12 15:17:28 2025
@author: e1012121
Hauptseite für interne Automatisierungs-Tools
"""
import streamlit as st

st.set_page_config(
    page_title="Dashboard",
    page_icon="🏠",
    layout="wide"
)

st.title("Dashboard")

st.markdown("""

### Zu beachten
- Während ein Skript ausgeführt wird, bitte nicht die Seite wechseln, sonst bricht es ab
- Bei der Bildverarbeitung darf im Ordner unter 1_Abbildungen -> 1_Originale keine tif-Dateien liegen! Dashboard stürzt sonst ab. 
Bitte zuerst den TIF zu JPG Konverter benutzen 
- ** abc

---

### 🚀 Über Streamlit
Dieses Dashboard wurde mit **Streamlit** erstellt - einem Python-Framework für die schnelle Erstellung von Web-Anwendungen.

**Vorteile:**
- **🎯 Einfache Bedienung**: Intuitive Benutzeroberfläche 
- **🗂️ Zentrale Anlaufstelle**: Alle Automatisierungs-Tools an einem Ort statt verstreuter Python-Skripte
- **🔀 Einfacher Wechsel**: Schnelles Umschalten zwischen verschiedenen Tools über die Sidebar-Navigation
- **📦 Keine Installation nötig**: Zugriff über Browser - keine Python-Umgebung auf jedem Arbeitsplatz erforderlich
- **🔍 Übersichtlichkeit**: Klare Struktur statt Ordner voller .py-Dateien

Weitere Informationen: [streamlit.io](https://streamlit.io)
""")

# Sidebar Info
with st.sidebar:
    st.info("""
    *TEXT**
    
    abcdefg
    """)
    
    st.divider()
    

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <small>Skript-Dashboard | Erstellt mit Streamlit</small>
</div>
""", unsafe_allow_html=True)