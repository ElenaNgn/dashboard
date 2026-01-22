# -*- coding: utf-8 -*-
"""
Streamlit App für CAD-Datei Konvertierung
Konvertiert CAD-Dateien automatisch in einzelne ZIP-Archive

Streamlit Multi-Page Version
"""

import streamlit as st
from pathlib import Path
import os
import shutil
import traceback
import tempfile
import zipfile
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing
import time

# ============================================================
# KEIN st.set_page_config() hier - wird in app.py gesetzt!
# ============================================================

# Titel und Beschreibung
st.title("📦 CAD zu ZIP Konvertierung")
st.markdown("""
Dieses Tool konvertiert CAD-Dateien automatisch in einzelne ZIP-Archive:
- **Drag & Drop**: Einfaches Hochladen mehrerer CAD-Dateien
- **Automatische Konvertierung**: Jede Datei wird in ein eigenes ZIP gepackt
- **Batch-Verarbeitung**: Parallele Verarbeitung mehrerer Dateien
- **Flexible Ausgabe**: Download als Einzel-ZIPs oder Gesamt-ZIP
""")

# Unterstützte CAD-Formate
SUPPORTED_CAD_FORMATS = [
    'dwg', 'dxf', 'dwf', 'dgn',  # AutoCAD & MicroStation
    'step', 'stp', 'iges', 'igs',  # STEP & IGES
    'sat', 'sab',  # ACIS
    'prt', 'asm',  # Creo/Pro-E
    'ipt', 'iam',  # Inventor
    'catpart', 'catproduct',  # CATIA
    'sldprt', 'sldasm', 'slddrw',  # SolidWorks
    'x_t', 'x_b',  # Parasolid
    'stl', 'obj', '3dm',  # 3D Mesh
    'rvt', 'rfa',  # Revit
]


st.divider()

# Session State initialisieren
if 'cad_processing_complete' not in st.session_state:
    st.session_state.cad_processing_complete = False
if 'cad_zip_files' not in st.session_state:
    st.session_state.cad_zip_files = {}
if 'cad_run_id' not in st.session_state:
    st.session_state.cad_run_id = str(int(time.time()))


def create_single_zip(cad_file_path, output_dir):
    """
    Erstellt ein ZIP-Archiv für eine einzelne CAD-Datei
    
    Args:
        cad_file_path: Path-Objekt zur CAD-Datei
        output_dir: Path-Objekt zum Ausgabe-Verzeichnis
        
    Returns:
        Path zum erstellten ZIP oder None bei Fehler
    """
    try:
        # ZIP-Dateiname = CAD-Dateiname ohne Extension + .zip
        zip_filename = cad_file_path.stem + '.zip'
        zip_path = output_dir / zip_filename
        
        # Erstelle ZIP
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Füge CAD-Datei mit ursprünglichem Namen hinzu
            zipf.write(cad_file_path, arcname=cad_file_path.name)
        
        return zip_path
        
    except Exception as e:
        st.error(f"❌ Fehler bei {cad_file_path.name}: {e}")
        return None


def process_cad_files_streamlit(uploaded_files, use_network_paths=False, 
                                 folder_path=None, compression_level=zipfile.ZIP_DEFLATED):
    """
    Verarbeitet hochgeladene CAD-Dateien oder Dateien aus Netzwerkpfad
    
    Args:
        uploaded_files: Liste von hochgeladenen Dateien (bei Upload-Modus)
        use_network_paths: Boolean, ob Netzwerkpfad verwendet werden soll
        folder_path: Vollständiger Pfad zum Quellordner
        compression_level: ZIP-Kompressionslevel
        
    Returns:
        Dictionary mit Ergebnissen oder None bei Fehler
    """
    try:
        # Bei Netzwerkpfad: Direkt im Projektordner arbeiten
        # Bei Upload: Temporäres Verzeichnis verwenden
        if use_network_paths and folder_path:
            # Netzwerkpfad-Modus
            source_path = Path(folder_path)
            
            # Prüfe ob Quellordner existiert
            if not source_path.exists():
                st.error("❌ Netzwerkpfad nicht gefunden: {source_path}")
                st.info("💡 **Prüfe folgendes:**")
                st.markdown(f"""
                1. Existiert der Ordner? `{source_path}`
                2. Ist das Netzlaufwerk verbunden?
                3. Hast du Zugriffsrechte?
                """)
                return None
            
            # Erstelle Ausgabe-Verzeichnis
            output_dir = source_path / "ZIP_Output"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Finde CAD-Dateien
            cad_files = []
            for ext in SUPPORTED_CAD_FORMATS:
                cad_files.extend(source_path.glob(f'*.{ext}'))
                cad_files.extend(source_path.glob(f'*.{ext.upper()}'))
            
            if not cad_files:
                st.error("❌ Keine CAD-Dateien gefunden in: {source_path}")
                st.info("Der Ordner existiert, enthält aber keine unterstützten CAD-Formate.")
                return None
            
            st.success(f"✅ {len(cad_files)} CAD-Dateien gefunden")
            
        else:
            # Upload-Modus: Temporäres Verzeichnis erstellen
            temp_dir = tempfile.mkdtemp()
            temp_path = Path(temp_dir)
            
            source_dir = temp_path / "cad_source"
            output_dir = temp_path / "zip_output"
            
            source_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Von Upload laden
            if not uploaded_files:
                st.warning("⚠️ Keine CAD-Dateien hochgeladen")
                return None
            
            cad_files = []
            for uploaded_file in uploaded_files:
                file_path = source_dir / uploaded_file.name
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                cad_files.append(file_path)
        
        if not cad_files:
            st.warning("⚠️ Keine gültigen CAD-Dateien gefunden")
            return None
        
        total_files = len(cad_files)
        
        # Progress Bars
        st.subheader("🔄 Konvertierung läuft...")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        st.write("**ZIP-Erstellung läuft...**")
        zip_progress = st.progress(0)
        
        # Parallel processing mit ThreadPoolExecutor
        max_workers = min(4, multiprocessing.cpu_count())
        
        completed_count = 0
        created_zips = []
        failed_files = []
        
        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit alle Dateien
                future_to_file = {
                    executor.submit(
                        create_single_zip,
                        cad_file,
                        output_dir
                    ): cad_file 
                    for cad_file in cad_files
                }
                
                # Sammle Ergebnisse
                for future in as_completed(future_to_file):
                    cad_file = future_to_file[future]
                    try:
                        zip_path = future.result(timeout=60)
                        if zip_path and zip_path.exists():
                            created_zips.append(zip_path)
                            completed_count += 1
                            status_text.text(f"✅ ZIP erstellt {completed_count}/{total_files}: {cad_file.name}")
                        else:
                            failed_files.append(cad_file.name)
                            st.warning(f"⚠️ Fehler bei: {cad_file.name}")
                    except Exception as e:
                        failed_files.append(cad_file.name)
                        st.warning(f"⚠️ Fehler bei {cad_file.name}: {e}")
                    
                    zip_progress.progress(len(created_zips + failed_files) / total_files)
        
        except Exception as e:
            st.error(f"❌ Kritischer Fehler bei paralleler Verarbeitung: {e}")
            st.warning("⚠️ Fallback: Verwende sequentielle Verarbeitung...")
            
            # Fallback: Sequentiell
            created_zips = []
            failed_files = []
            
            for idx, cad_file in enumerate(cad_files):
                try:
                    status_text.text(f"Erstelle ZIP {idx+1}/{total_files}: {cad_file.name}")
                    zip_path = create_single_zip(cad_file, output_dir)
                    
                    if zip_path and zip_path.exists():
                        created_zips.append(zip_path)
                    else:
                        failed_files.append(cad_file.name)
                    
                    zip_progress.progress((idx + 1) / total_files)
                except Exception as e:
                    failed_files.append(cad_file.name)
                    st.warning(f"⚠️ Fehler bei {cad_file.name}: {e}")
        
        progress_bar.progress(1.0)
        status_text.text("✅ Konvertierung abgeschlossen!")
        
        # Zeige Zusammenfassung
        if created_zips:
            st.success(f"🎉 {len(created_zips)} ZIP-Archive erfolgreich erstellt!")
        
        if failed_files:
            st.warning(f"⚠️ {len(failed_files)} Dateien konnten nicht konvertiert werden:")
            with st.expander("Fehlgeschlagene Dateien"):
                for filename in failed_files:
                    st.write(f"- {filename}")
        
        # Bei Netzwerkpfad: Zeige wo gespeichert wurde
        if use_network_paths and folder_path:
            st.info("📁 **ZIP-Dateien wurden gespeichert in:**")
            st.code(str(output_dir))
            
        else:
            # Bei Upload-Modus: Kopiere für Download in permanentes Verzeichnis
            perm_dir = Path(tempfile.gettempdir()) / "streamlit_cad_results" / st.session_state.cad_run_id
            perm_dir.mkdir(parents=True, exist_ok=True)
            
            perm_zips = []
            for zip_file in created_zips:
                perm_zip = perm_dir / zip_file.name
                shutil.copy(zip_file, perm_zip)
                perm_zips.append(perm_zip)
            
            created_zips = perm_zips
        
        # Ergebnisse speichern
        results = {
            'zip_files': created_zips,
            'output_dir': output_dir,
            'total_files': total_files,
            'successful': len(created_zips),
            'failed': len(failed_files),
            'failed_files': failed_files,
            'is_network_mode': use_network_paths
        }
        
        return results
        
    except Exception as e:
        st.error(f"❌ Fehler während der Verarbeitung: {e}")
        st.error(traceback.format_exc())
        return None


def create_master_zip(zip_files):
    """Erstellt ein Master-ZIP mit allen einzelnen ZIPs"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as master_zip:
        for zip_file in zip_files:
            master_zip.write(zip_file, arcname=zip_file.name)
    
    zip_buffer.seek(0)
    return zip_buffer


# ============================================================================
# STREAMLIT UI
# ============================================================================

# Sidebar
with st.sidebar:
    st.header("⚙️ Einstellungen")
    
    processing_mode = st.radio(
        "Verarbeitungsmodus:",
        ["📤 Dateien hochladen (Drag & Drop)", "📁 Netzwerkpfad verwenden"],
        help="Wähle Upload oder Netzwerkpfad"
    )
    
    st.divider()
    
    # Kompression-Level
    compression_option = st.selectbox(
        "Kompression:",
        ["Standard (empfohlen)", "Keine (schneller)", "Maximum"],
        help="Höhere Kompression = kleinere Dateien aber langsamer"
    )
    
    compression_map = {
        "Standard (empfohlen)": zipfile.ZIP_DEFLATED,
        "Keine (schneller)": zipfile.ZIP_STORED,
        "Maximum": zipfile.ZIP_DEFLATED  # ZIP_BZIP2 würde hier mehr bringen
    }
    compression_level = compression_map[compression_option]
    
    st.divider()
    
    folder_path_input = None
    
    if processing_mode == "📁 Netzwerkpfad verwenden":
        folder_path_input = st.text_input(
            "CAD-Ordner-Pfad:",
            placeholder=r"N:\CAD\Projekte\2024",
            help="Vollständiger Pfad zum Ordner mit CAD-Dateien"
        )
        
        if folder_path_input:
            st.caption(f"📍 Suche in: `{folder_path_input}`")


# Tabs
tab1, tab2 = st.tabs(["📦 CAD zu ZIP", "📖 Anleitung"])

with tab1:
    if processing_mode == "📤 Dateien hochladen (Drag & Drop)":
        st.subheader("📤 CAD-Dateien hochladen")
        
        # File Uploader mit allen CAD-Formaten
        uploaded_files = st.file_uploader(
            "Ziehe CAD-Dateien hierher oder klicke zum Auswählen",
            type=SUPPORTED_CAD_FORMATS,
            accept_multiple_files=True,
            help=f"Unterstützte Formate: {', '.join(SUPPORTED_CAD_FORMATS[:10])} ..."
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} Dateien hochgeladen")
            
            # Zeige Dateiliste
            with st.expander("📋 Hochgeladene Dateien"):
                for idx, file in enumerate(uploaded_files, 1):
                    file_size = len(file.getvalue()) / (1024 * 1024)  # MB
                    st.write(f"{idx}. **{file.name}** ({file_size:.2f} MB)")
        
        if st.button("🚀 ZIP-Konvertierung starten", type="primary", disabled=not uploaded_files, key="btn_start_upload"):
            results = process_cad_files_streamlit(
                uploaded_files, 
                compression_level=compression_level
            )
            
            if results:
                st.session_state.cad_processing_complete = True
                st.session_state.cad_zip_files = results
    
    else:  # Netzwerkpfad
        st.subheader("📁 Netzwerkpfad-Verarbeitung")
        
        if folder_path_input:
            full_path = Path(folder_path_input)
            st.info(f"📍 Pfad: `{full_path}`")
            
            if full_path.exists():
                st.success("✅ Pfad gefunden!")
                try:
                    # Zähle CAD-Dateien
                    files = []
                    for ext in SUPPORTED_CAD_FORMATS:
                        files.extend(full_path.glob(f'*.{ext}'))
                        files.extend(full_path.glob(f'*.{ext.upper()}'))
                    
                    if files:
                        st.info(f"📦 {len(files)} CAD-Dateien gefunden")
                        
                        # Zeige Vorschau
                        with st.expander("📋 Gefundene Dateien (Vorschau)"):
                            for file in files[:20]:
                                st.write(f"- {file.name}")
                            if len(files) > 20:
                                st.write(f"... und {len(files) - 20} weitere")
                    else:
                        st.warning("⚠️ Keine CAD-Dateien gefunden")
                except Exception as e:
                    st.warning(f"⚠️ Lesefehler: {e}")
            else:
                st.warning("⚠️ Pfad nicht erreichbar")
        else:
            st.warning("⚠️ Bitte Pfad eingeben")
        
        if st.button("🚀 ZIP-Konvertierung starten", type="primary", disabled=not folder_path_input, key="btn_start_network"):
            results = process_cad_files_streamlit(
                None, 
                use_network_paths=True, 
                folder_path=folder_path_input,
                compression_level=compression_level
            )
            
            if results:
                st.session_state.cad_processing_complete = True
                st.session_state.cad_zip_files = results
    
    # Ergebnisse anzeigen
    if st.session_state.cad_processing_complete and st.session_state.cad_zip_files:
        st.divider()
        st.subheader("📦 Ergebnisse")
        
        results = st.session_state.cad_zip_files
        
        # Metriken
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Gesamt", results['total_files'])
        with col2:
            st.metric("Erfolgreich", results['successful'], 
                     delta=None if results['failed'] == 0 else f"-{results['failed']}")
        with col3:
            if results.get('is_network_mode'):
                st.success("✅ Im Netzwerk gespeichert")
            else:
                st.info("📥 Downloads verfügbar")
        
        # Netzwerk-Modus: Zeige Speicherort
        if results.get('is_network_mode'):
            st.markdown("### 📁 Gespeicherte ZIP-Dateien")
            st.code(str(results['output_dir']))
            
            st.markdown(f"**{len(results['zip_files'])} ZIP-Archive erstellt:**")
            with st.expander("📋 Liste der erstellten ZIPs"):
                for zip_file in results['zip_files']:
                    file_size = zip_file.stat().st_size / (1024 * 1024)  # MB
                    st.write(f"- {zip_file.name} ({file_size:.2f} MB)")
            
            # Optional: Download auch im Netzwerk-Modus
            if st.checkbox("💾 ZIP-Dateien auch herunterladen", key="checkbox_download_network"):
                with st.spinner("Erstelle Master-ZIP..."):
                    master_zip = create_master_zip(results['zip_files'])
                
                st.download_button(
                    "📦 Alle ZIPs als Gesamt-ZIP herunterladen",
                    data=master_zip,
                    file_name=f"CAD_ZIPs_{st.session_state.cad_run_id}.zip",
                    mime="application/zip",
                    key="download_master_zip_network"
                )
        
        # Upload-Modus: Downloads
        else:
            st.markdown("### 💾 Downloads")
            
            # Einzelne ZIPs
            if len(results['zip_files']) <= 10:
                st.markdown("**Einzelne ZIP-Dateien:**")
                for idx, zip_file in enumerate(results['zip_files']):
                    with open(zip_file, 'rb') as f:
                        st.download_button(
                            f"📥 {zip_file.name}",
                            data=f,
                            file_name=zip_file.name,
                            mime="application/zip",
                            key=f"download_single_zip_{idx}"
                        )
            else:
                with st.expander(f"📋 {len(results['zip_files'])} einzelne ZIP-Dateien"):
                    for idx, zip_file in enumerate(results['zip_files']):
                        with open(zip_file, 'rb') as f:
                            st.download_button(
                                f"📥 {zip_file.name}",
                                data=f,
                                file_name=zip_file.name,
                                mime="application/zip",
                                key=f"download_zip_expanded_{idx}"
                            )
            
            st.divider()
            
            # Master-ZIP
            st.markdown("**Oder alle zusammen:**")
            with st.spinner("Erstelle Gesamt-ZIP..."):
                master_zip = create_master_zip(results['zip_files'])
            
            st.download_button(
                "📦 Alle ZIPs als Gesamt-ZIP herunterladen",
                data=master_zip,
                file_name=f"CAD_ZIPs_Komplett_{st.session_state.cad_run_id}.zip",
                mime="application/zip",
                key="download_master_zip_upload"
            )

with tab2:
    st.header("📖 Anleitung")
    
    st.markdown("""
    ### 🎯 Verwendungszweck
    
    Dieses Tool konvertiert CAD-Dateien automatisch in einzelne ZIP-Archive. 
    Ideal für:
    - Archivierung von CAD-Projekten
    - E-Mail-Versand (reduzierte Dateigrösse)
    - Batch-Konvertierung vieler Dateien
    
    ---
    
    ### 📤 Upload-Modus (empfohlen)
    
    **So geht's:**
    1. **Dateien hochladen**: Ziehe CAD-Dateien per Drag & Drop in den Upload-Bereich
    2. **Konvertierung starten**: Klicke auf "🚀 ZIP-Konvertierung starten"
    3. **Downloads**: Lade einzelne ZIPs oder Gesamt-ZIP herunter
    
    **Vorteile:**
    - ✅ Einfache Bedienung
    - ✅ Funktioniert überall
    - ✅ Keine Netzwerkverbindung nötig
    
    ---
    
    ### 📁 Netzwerkpfad-Modus
    
    **So geht's:**
    1. **Pfad eingeben**: Gib den vollständigen Pfad zum CAD-Ordner ein
    2. **Konvertierung starten**: Klicke auf "🚀 ZIP-Konvertierung starten"
    3. **Ergebnis**: ZIPs werden im Unterordner "ZIP_Output" gespeichert
    
    **Beispiel-Pfad:**
    ```
    N:\\CAD\\Projekte\\2024\\Projekt_XYZ
    ```
    
    **Vorteile:**
    - ✅ Direkt im Netzwerk arbeiten
    - ✅ Keine Upload-Wartezeit
    - ✅ Gut für große Dateien
    
    ---
    
    ### 🔧 Einstellungen
    
    **Kompression:**
    - **Standard**: Guter Kompromiss (empfohlen)
    - **Keine**: Schnellste Verarbeitung, größere Dateien
    - **Maximum**: Kleinste Dateien, längere Verarbeitung
    
    ---
    
    ### 📋 Unterstützte Formate
    
    **CAD-Software:**
    - **AutoCAD**: .dwg, .dxf, .dwf
    - **SolidWorks**: .sldprt, .sldasm, .slddrw
    - **Inventor**: .ipt, .iam
    - **CATIA**: .catpart, .catproduct
    - **Creo/Pro-E**: .prt, .asm
    - **Revit**: .rvt, .rfa
    
    **Austausch-Formate:**
    - **STEP**: .step, .stp
    - **IGES**: .iges, .igs
    - **STL**: .stl
    - **Parasolid**: .x_t, .x_b
    - **ACIS**: .sat, .sab
    
    **3D-Modelle:**
    - .obj, .3dm
    
    ---
    
    ### 💡 Tipps
    
    - ⚡ **Performance**: Bei vielen Dateien wird parallele Verarbeitung genutzt
    - 📦 **Organisation**: Jede CAD-Datei wird in ein eigenes ZIP gepackt
    - 🔐 **Dateinamen**: Original-Dateinamen bleiben erhalten
    - 💾 **Speicher**: Gesamt-ZIP enthält alle Einzel-ZIPs
    
    ---
    
    ### ❓ Häufige Fragen
    
    **Wie groß dürfen die Dateien sein?**
    - Upload-Modus: Abhängig von Streamlit-Einstellungen (Standard: 200 MB)
    - Netzwerk-Modus: Keine Beschränkung
    
    **Werden die Originaldateien verändert?**
    - Nein, Originaldateien bleiben unberührt
    
    **Wo werden die ZIPs gespeichert?**
    - Upload-Modus: Temporär, nur zum Download
    - Netzwerk-Modus: Im Unterordner "ZIP_Output"
    
    **Was passiert bei Fehlern?**
    - Fehlerhafte Dateien werden übersprungen
    - Erfolgreiche Konvertierungen werden trotzdem gespeichert
    """)

st.divider()
st.caption("CAD-zu-ZIP Konvertierungs-Tool | Version 1.0")