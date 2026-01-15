# -*- coding: utf-8 -*-
"""
Created on Mon Jan 12 09:59:20 2026

@author: e1012121
"""

import streamlit as st
import pandas as pd
import requests
import os
from pathlib import Path
import re
from urllib.parse import urlparse
import time

def sanitize_artikel_nr(artikel_nr):
    """
    Entfernt Leerzeichen und ersetzt Punkte durch Unterstriche
    Beispiel: '1313 414.105.000' -> '1313414_105_000'
    """
    # Leerzeichen entfernen
    cleaned = artikel_nr.replace(" ", "")
    # Punkte durch Unterstriche ersetzen
    cleaned = cleaned.replace(".", "_")
    return cleaned

def get_abbildung_name(artikel_nr):
    """
    Erstellt den Namen für Abbildungen
    Beispiel: '1313 414.105.000' -> '01313414_105_000'
    """
    cleaned = sanitize_artikel_nr(artikel_nr)
    # 0 am Anfang anhängen
    return f"0{cleaned}"

def get_ambiente_name(artikel_nr, index):
    """
    Erstellt den Namen für Ambientebilder
    Beginnt mit _2, dann _3, _4, etc.
    Beispiel: '1313 414.105.000' -> '01313414_105_000_2'
    """
    base_name = get_abbildung_name(artikel_nr)
    # Index beginnt bei 0, aber wir wollen bei _2 starten
    return f"{base_name}_{index + 2}"

def get_masszeichnung_name(artikel_nr):
    """
    Nimmt nur die ersten 7 Zahlen der Artikel-Nr
    Beispiel: '1313 414.105.000' -> '1313414'
    """
    # Alle nicht-numerischen Zeichen entfernen
    numbers_only = re.sub(r'[^0-9]', '', artikel_nr)
    # Nur die ersten 7 Zahlen nehmen
    return numbers_only[:7]

def get_file_extension(url):
    """Extrahiert die Dateiendung aus der URL"""
    path = urlparse(url).path
    return os.path.splitext(path)[1]

def download_file(url, output_path, progress_callback=None):
    """
    Lädt eine Datei von der URL herunter und speichert sie
    """
    try:
        response = requests.get(url, timeout=30, stream=True)
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        
        with open(output_path, 'wb') as f:
            if total_size == 0:
                f.write(response.content)
            else:
                downloaded = 0
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback:
                            progress_callback(downloaded / total_size)
        
        return True, None
    except Exception as e:
        return False, str(e)

def process_excel(df, output_folder, progress_bar, status_text):
    """
    Verarbeitet die Excel-Datei und lädt alle Mediendaten herunter
    """
    # Ausgabeordner erstellen, falls nicht vorhanden
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Log-Datei für Fehler
    log_file = output_path / "download_log.txt"
    errors = []
    success_count = 0
    total_files = 0
    
    # Zähle Gesamtanzahl der Dateien
    for _, row in df.iterrows():
        artikel_nr = row['Artikel-Nr']
        
        # Abbildungen
        if pd.notna(row['Abbildungen']) and row['Abbildungen'].strip():
            total_files += 1
        
        # Ambientebilder (können mehrere sein, getrennt durch ;)
        if pd.notna(row['Ambientebilder ']):
            urls = str(row['Ambientebilder ']).split(';')
            total_files += len([u for u in urls if u.strip()])
        
        # Masszeichnungen
        if pd.notna(row['Masszeichnungen']) and row['Masszeichnungen'].strip():
            total_files += 1
    
    current_file = 0
    
    # Durch jede Zeile iterieren
    for idx, row in df.iterrows():
        artikel_nr = row['Artikel-Nr']
        status_text.text(f"Verarbeite Zeile {idx + 1}/{len(df)}: {artikel_nr}")
        
        # 1. Abbildungen verarbeiten
        if pd.notna(row['Abbildungen']) and row['Abbildungen'].strip():
            url = row['Abbildungen'].strip()
            extension = get_file_extension(url)
            filename = f"{get_abbildung_name(artikel_nr)}{extension}"
            output_file = output_path / filename
            
            current_file += 1
            progress_bar.progress(current_file / total_files)
            
            success, error = download_file(url, output_file)
            if success:
                success_count += 1
            else:
                errors.append(f"Zeile {idx + 1} - Abbildung: {url} - Fehler: {error}")
        
        # 2. Ambientebilder verarbeiten (mehrere URLs möglich, getrennt durch ;)
        if pd.notna(row['Ambientebilder ']):
            urls = str(row['Ambientebilder ']).split(';')
            for ambiente_idx, url in enumerate(urls):
                url = url.strip()
                if url:
                    extension = get_file_extension(url)
                    filename = f"{get_ambiente_name(artikel_nr, ambiente_idx)}{extension}"
                    output_file = output_path / filename
                    
                    current_file += 1
                    progress_bar.progress(current_file / total_files)
                    
                    success, error = download_file(url, output_file)
                    if success:
                        success_count += 1
                    else:
                        errors.append(f"Zeile {idx + 1} - Ambiente {ambiente_idx + 2}: {url} - Fehler: {error}")
        
        # 3. Masszeichnungen verarbeiten
        if pd.notna(row['Masszeichnungen']) and row['Masszeichnungen'].strip():
            url = row['Masszeichnungen'].strip()
            extension = get_file_extension(url)
            filename = f"{get_masszeichnung_name(artikel_nr)}{extension}"
            output_file = output_path / filename
            
            current_file += 1
            progress_bar.progress(current_file / total_files)
            
            success, error = download_file(url, output_file)
            if success:
                success_count += 1
            else:
                errors.append(f"Zeile {idx + 1} - Masszeichnung: {url} - Fehler: {error}")
    
    # Fehler in Log-Datei schreiben
    if errors:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f"Download-Log vom {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Erfolgreich: {success_count}/{total_files}\n")
            f.write(f"Fehler: {len(errors)}\n\n")
            for error in errors:
                f.write(f"{error}\n")
    
    return success_count, total_files, errors

# Streamlit App
def main():
    st.set_page_config(
        page_title="Mediendaten Downloader",
        page_icon="📥",
        layout="wide"
    )
    
    st.title("📥 Mediendaten Downloader")
    st.markdown("---")
    
    # Anleitung
    with st.expander("ℹ️ Anleitung"):
        st.markdown("""
        
        1. **Excel-Datei hochladen**: Datei muss folgende Spalten enthalten:
           - Artikel-Nr
           - Abbildungen (URL zum Download)
           - Ambientebilder (URL zum Download) 
           - Masszeichnungen (URL zum Download)
        
        2. **Ausgabeordner angeben**: Pfad, wo die heruntergeladenen und umbenannten Dateien gespeichert werden sollen
        
        3. **Download starten**: Das Tool lädt alle Dateien herunter und benennt sie um
        
        ### Benennungsregeln:
        - **Abbildungen**: `0` + Artikel-Nr (ohne Leerzeichen, `.` → `_`)
          - Beispiel: `1313 414.105.000` → `01313414_105_000.png`
        
        - **Ambientebilder**: Wie Abbildungen, aber mit `_2`, `_3`, `_4` usw. am Ende
          - Beispiel: `01313414_105_000_2.jpg`, `01313414_105_000_3.jpg`
        
        - **Masszeichnungen**: Nur erste 7 Zahlen der Artikel-Nr
          - Beispiel: `1313 414.105.000` → `1313414.jpg`
        """)
    
    # Zwei Spalten für Upload und Pfad
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Excel-Datei hochladen")
        uploaded_file = st.file_uploader(
            "Wähle die Excel-Datei aus",
            type=['xlsx', 'xls'],
            help="Die Excel-Datei muss die Spalten 'Artikel-Nr', 'Abbildungen', 'Ambientebilder ' und 'Masszeichnungen' enthalten"
        )
    
    with col2:
        st.subheader("2. Ausgabeordner festlegen")
        output_folder = st.text_input(
            "Pfad zum Ausgabeordner",
            value=os.path.expanduser("~/Downloads/Mediendaten"),
            help="Gib den vollständigen Pfad an, wo die Dateien gespeichert werden sollen"
        )
    
    # Vorschau der Excel-Datei
    if uploaded_file is not None:
        st.markdown("---")
        st.subheader("📋 Vorschau der Excel-Datei")
        
        try:
            df = pd.read_excel(uploaded_file)
            
            # Prüfe, ob alle erforderlichen Spalten vorhanden sind
            required_columns = ['Artikel-Nr', 'Abbildungen', 'Ambientebilder ', 'Masszeichnungen']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Fehlende Spalten: {', '.join(missing_columns)}")
                st.info("Die Excel-Datei muss folgende Spalten enthalten: " + ", ".join(required_columns))
                return
            
            st.success(f"✅ Excel-Datei geladen: {len(df)} Zeilen gefunden")
            
            # Zeige erste Zeilen
            st.dataframe(df.head(10), use_container_width=True)
            
            # Statistiken
            col1, col2, col3 = st.columns(3)
            with col1:
                abbildungen_count = df['Abbildungen'].notna().sum()
                st.metric("Abbildungen", abbildungen_count)
            with col2:
                # Zähle Ambientebilder (können mehrere pro Zeile sein)
                ambiente_count = 0
                for val in df['Ambientebilder ']:
                    if pd.notna(val):
                        ambiente_count += len([u for u in str(val).split(';') if u.strip()])
                st.metric("Ambientebilder", ambiente_count)
            with col3:
                mass_count = df['Masszeichnungen'].notna().sum()
                st.metric("Masszeichnungen", mass_count)
            
            # Beispiel-Benennungen anzeigen
            st.markdown("---")
            st.subheader("🏷️ Beispiel-Benennungen")
            
            if len(df) > 0:
                example_row = df.iloc[0]
                artikel_nr = example_row['Artikel-Nr']
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Original Artikel-Nr:**")
                    st.code(artikel_nr)
                
                with col2:
                    st.markdown("**Neue Dateinamen:**")
                    st.code(f"Abbildung: {get_abbildung_name(artikel_nr)}.xxx")
                    st.code(f"Ambiente 1: {get_ambiente_name(artikel_nr, 0)}.xxx")
                    st.code(f"Ambiente 2: {get_ambiente_name(artikel_nr, 1)}.xxx")
                    st.code(f"Masszeichnung: {get_masszeichnung_name(artikel_nr)}.xxx")
            
            # Download-Button
            st.markdown("---")
            if st.button("🚀 Download starten", type="primary", use_container_width=True):
                
                # Prüfe ob Ausgabeordner angegeben wurde
                if not output_folder or output_folder.strip() == "":
                    st.error("❌ Bitte gib einen Ausgabeordner an!")
                    return
                
                # Progress-Bar und Status
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Starte Download
                with st.spinner("Lade Dateien herunter..."):
                    success_count, total_files, errors = process_excel(
                        df, 
                        output_folder, 
                        progress_bar, 
                        status_text
                    )
                
                # Ergebnis anzeigen
                st.markdown("---")
                if errors:
                    st.warning(f"⚠️ Download abgeschlossen mit {len(errors)} Fehlern")
                    st.metric("Erfolgreich heruntergeladen", f"{success_count}/{total_files}")
                    
                    with st.expander("❌ Fehler anzeigen"):
                        for error in errors:
                            st.text(error)
                    
                    st.info(f"📄 Vollständiges Fehler-Log wurde gespeichert unter:\n`{os.path.join(output_folder, 'download_log.txt')}`")
                else:
                    st.success(f"✅ Alle {success_count} Dateien erfolgreich heruntergeladen!")
                
                st.balloons()
                st.info(f"📁 Dateien wurden gespeichert unter:\n`{output_folder}`")
                
        except Exception as e:
            st.error(f"❌ Fehler beim Laden der Excel-Datei: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()