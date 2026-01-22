# -*- coding: utf-8 -*-
"""
Streamlit App für Bildverarbeitung
Konvertiert und verarbeitet Bilder für Katalog-Import

Streamlit Multi-Page Version
"""

import streamlit as st
import pandas as pd
from PIL import Image, ImageChops, ImageFile
from pathlib import Path
import os
import shutil
import traceback
import tempfile
import zipfile
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing

# Import Cache-Funktionen
from utils.caching import get_file_list_cached

# ============================================================
# KEIN st.set_page_config() hier - wird in app.py gesetzt!
# ============================================================

# Titel und Beschreibung
st.title("🖼️ Bildverarbeitung")
st.markdown("""
Dieses Skript verarbeitet Produktbilder folgendermassen:
- **Artikelbild max**: Beschnittene TIFF-Bilder ohne weissen Rand
- **Katalogbilder**: Skalierte Graustufenbilder im JPEG-Format
- **Excel-Import**: Automatische Generierung der Import-Datei
- **S-Laufwerk**: Kopiert die Bilder am Schluss in die korrekten Ordner im S-Laufwerk 
""")
st.info('Originalbilder dürfen momentan nicht im TIF-Format sein, sonst stürzt das Skript ab!', icon="ℹ️")

st.divider()

# Session State initialisieren
if 'img_processing_complete' not in st.session_state:
    st.session_state.img_processing_complete = False
if 'img_excel_path' not in st.session_state:
    st.session_state.img_excel_path = None
if 'img_processed_images' not in st.session_state:
    st.session_state.img_processed_images = {}

# Haupt-Funktion für Bildverarbeitung
def process_images_streamlit(uploaded_files, use_network_paths=False, 
                              folder_path=None, copy_to_s_drive=True):
    """
    Verarbeitet hochgeladene Bilder oder Bilder aus Netzwerkpfad
    
    Args:
        uploaded_files: Liste von hochgeladenen Dateien (bei Upload-Modus)
        use_network_paths: Boolean, ob Netzwerkpfad verwendet werden soll
        folder_path: Vollständiger Pfad zum Projektordner
        copy_to_s_drive: Boolean, ob Dateien ins S-Laufwerk kopiert werden sollen
    """
    try:
        # Variable für S-Laufwerk-Status
        s_drive_copied = False
        
        # Bei Netzwerkpfad: Direkt in Projektordner arbeiten
        # Bei Upload: Temporäres Verzeichnis verwenden
        if use_network_paths and folder_path:
            # Netzwerkpfad-Modus: Verwende Projektordner-Struktur
            project_path = Path(folder_path)
            
            # Definiere Verzeichnisse
            dir_originalbilder = project_path / "1_Abbildungen" / "1_Originale"
            dir_artikelbild_max = project_path / "1_Abbildungen" / "2_Bad_Artikelbild_max"
            dir_katalog = project_path / "1_Abbildungen" / "3_Katalog"
            dir_importfiles = project_path / "8_Importfiles_Media-Datenpfade"
            
            # Prüfe ob Originalbilder-Ordner existiert
            if not dir_originalbilder.exists():
                st.error(f"❌ Netzwerkpfad nicht gefunden: {dir_originalbilder}")
                st.info("💡 **Prüfe folgendes:**")
                st.markdown(f"""
                1. Existiert der Projektordner? `{folder_path}`
                2. Existiert der Unterordner? `{dir_originalbilder}`
                3. Ist das Netzlaufwerk verbunden?
                4. Hast du Zugriffsrechte?
                """)
                return None
            
            # Erstelle Ausgabe-Verzeichnisse falls nicht vorhanden
            dir_artikelbild_max.mkdir(parents=True, exist_ok=True)
            dir_katalog.mkdir(parents=True, exist_ok=True)
            dir_importfiles.mkdir(parents=True, exist_ok=True)
            
            # ============================================================
            # OPTIMIERT: Lade Bilddateien mit Cache
            # ============================================================
            image_files_str = get_file_list_cached(
                str(dir_originalbilder),
                ['.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp']
            )
            
            # Konvertiere String-Pfade zurück zu Path-Objekten
            image_files = [Path(f) for f in image_files_str]
            
            if not image_files:
                st.error(f"❌ Keine Bilddateien gefunden in: {dir_originalbilder}")
                st.info("Der Ordner existiert, ist aber leer oder enthält keine unterstützten Bildformate.")
                return None
            
            st.success(f"✅ {len(image_files)} Bilddateien gefunden (🚀 Cached)")
            
        else:
            # Upload-Modus: Temporäres Verzeichnis erstellen
            temp_dir = tempfile.mkdtemp()
            temp_path = Path(temp_dir)
            
            dir_originalbilder = temp_path / "1_Originale"
            dir_artikelbild_max = temp_path / "2_Artikelbild_max"
            dir_katalog = temp_path / "3_Katalog"
            dir_importfiles = temp_path / "8_Importfiles"
            
            dir_originalbilder.mkdir(parents=True, exist_ok=True)
            dir_artikelbild_max.mkdir(parents=True, exist_ok=True)
            dir_katalog.mkdir(parents=True, exist_ok=True)
            dir_importfiles.mkdir(parents=True, exist_ok=True)
            
            # Von Upload laden
            if not uploaded_files:
                st.warning("⚠️ Keine Bilder hochgeladen")
                return None
            
            for uploaded_file in uploaded_files:
                file_path = dir_originalbilder / uploaded_file.name
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            
            # Im Upload-Modus: Kein Cache nötig
            image_files = list(dir_originalbilder.glob('*'))
            image_files = [f for f in image_files 
                          if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp']]
        
        if not image_files:
            st.warning("⚠️ Keine gültigen Bilddateien gefunden")
            return None
        
        total_files = len(image_files)
        
        # Bildverarbeitung
        ImageFile.LOAD_TRUNCATED_IMAGES = True
        
        # Progress Bars
        st.subheader("🔄 Verarbeitung läuft...")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # SCHRITT 1: Artikelbild max erstellen (PARALLEL)
        st.write("**Schritt 1/3:** Erstelle Artikelbilder max (TIFF)...")
        step1_progress = st.progress(0)
        
        # Parallel processing mit ThreadPoolExecutor
        if use_network_paths:
            max_workers = min(2, multiprocessing.cpu_count())
            st.info("ℹ️ Netzwerk-Modus: Verwende reduzierte Thread-Anzahl für Stabilität")
        else:
            max_workers = min(8, multiprocessing.cpu_count())
        
        completed_count = 0
        
        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {
                    executor.submit(
                        crop_white_or_transparent_border,
                        file_path,
                        dir_artikelbild_max / file_path.name
                    ): file_path 
                    for file_path in image_files
                }
                
                for future in as_completed(future_to_file):
                    file_path = future_to_file[future]
                    try:
                        future.result(timeout=30)
                        completed_count += 1
                        status_text.text(f"Verarbeite Artikelbild {completed_count}/{total_files}: {file_path.name}")
                        step1_progress.progress(completed_count / total_files)
                    except Exception as e:
                        st.warning(f"⚠️ Fehler bei {file_path.name}: {e}")
        
        except Exception as e:
            st.error(f"❌ Kritischer Fehler bei paralleler Verarbeitung: {e}")
            st.warning("⚠️ Fallback: Verwende sequentielle Verarbeitung...")
            
            completed_count = 0
            for idx, file_path in enumerate(image_files):
                try:
                    status_text.text(f"Verarbeite Artikelbild {idx+1}/{total_files}: {file_path.name}")
                    crop_white_or_transparent_border(
                        file_path,
                        dir_artikelbild_max / file_path.name
                    )
                    completed_count += 1
                    step1_progress.progress((idx + 1) / total_files)
                except Exception as e:
                    st.warning(f"⚠️ Fehler bei {file_path.name}: {e}")
        
        progress_bar.progress(0.33)
        st.success(f"✅ {completed_count} Artikelbilder max erstellt")
        
        # SCHRITT 2: Katalogbilder erstellen (PARALLEL)
        st.write("**Schritt 2/3:** Erstelle Katalogbilder (JPEG)...")
        step2_progress = st.progress(0)
        
        completed_count = 0
        
        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {
                    executor.submit(
                        process_image_for_catalog,
                        dir_artikelbild_max / file_path.with_suffix('.tif').name,
                        dir_katalog / file_path.with_suffix('.jpg').name
                    ): file_path 
                    for file_path in image_files
                }
                
                for future in as_completed(future_to_file):
                    file_path = future_to_file[future]
                    try:
                        future.result(timeout=30)
                        completed_count += 1
                        status_text.text(f"Erstelle Katalogbild {completed_count}/{total_files}: {file_path.name}")
                        step2_progress.progress(completed_count / total_files)
                    except Exception as e:
                        st.warning(f"⚠️ Fehler bei {file_path.name}: {e}")
        
        except Exception as e:
            st.error(f"❌ Kritischer Fehler: {e}")
            st.warning("⚠️ Fallback: Sequentielle Verarbeitung...")
            
            completed_count = 0
            for idx, file_path in enumerate(image_files):
                try:
                    status_text.text(f"Erstelle Katalogbild {idx+1}/{total_files}: {file_path.name}")
                    tif_path = dir_artikelbild_max / file_path.with_suffix('.tif').name
                    jpg_path = dir_katalog / file_path.with_suffix('.jpg').name
                    process_image_for_catalog(tif_path, jpg_path)
                    completed_count += 1
                    step2_progress.progress((idx + 1) / total_files)
                except Exception as e:
                    st.warning(f"⚠️ Fehler bei {file_path.name}: {e}")
        
        progress_bar.progress(0.66)
        st.success(f"✅ {completed_count} Katalogbilder erstellt")
        
        # SCHRITT 3: Kopiere zu S-Laufwerk (wenn aktiviert)
        s_drive_copied = False
        if copy_to_s_drive:
            st.write("**Schritt 3/4:** Kopiere Dateien zu S-Laufwerk...")
            copy_progress = st.progress(0)
            copy_status = st.empty()
            
            target_originalbilder = Path(r"S:/Multimedia/Originale")
            target_artikelbild_max = Path(r"S:/Multimedia/Print/BAD_Artikelbild_maximal")
            target_katalog = Path(r"S:/Multimedia/Print/HAWAKatalog")
            
            try:
                if not target_originalbilder.parent.exists():
                    st.warning("⚠️ S-Laufwerk nicht erreichbar - Dateien werden nur im Projektordner gespeichert")
                    s_drive_copied = False
                else:
                    target_originalbilder.mkdir(parents=True, exist_ok=True)
                    target_artikelbild_max.mkdir(parents=True, exist_ok=True)
                    target_katalog.mkdir(parents=True, exist_ok=True)
                    s_drive_copied = True
            except Exception as e:
                st.warning(f"⚠️ S-Laufwerk nicht erreichbar: {e}")
                s_drive_copied = False
            
            if s_drive_copied:
                total_copy_files = len(image_files) * 3
                completed_copy = 0
                
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = []
                    
                    for img_file in image_files:
                        futures.append(
                            executor.submit(shutil.copy, img_file, target_originalbilder / img_file.name)
                        )
                    
                    for img_file in dir_artikelbild_max.glob('*.tif'):
                        futures.append(
                            executor.submit(shutil.copy, img_file, target_artikelbild_max / img_file.name)
                        )
                    
                    for img_file in dir_katalog.glob('*.jpg'):
                        futures.append(
                            executor.submit(shutil.copy, img_file, target_katalog / img_file.name)
                        )
                    
                    for future in as_completed(futures):
                        try:
                            future.result()
                            completed_copy += 1
                            copy_status.text(f"Kopiere Dateien {completed_copy}/{total_copy_files}...")
                            copy_progress.progress(completed_copy / total_copy_files)
                        except Exception as e:
                            st.warning(f"⚠️ Fehler beim Kopieren: {e}")
                
                copy_status.text("✅ Dateien zu S-Laufwerk kopiert!")
                st.success(f"✅ {completed_copy} Dateien ins S-Laufwerk kopiert")
                progress_bar.progress(0.80)
        else:
            st.info("ℹ️ Schritt 3: S-Laufwerk-Kopie übersprungen (deaktiviert)")
        
        # SCHRITT 4 (bzw. 3 ohne S-Laufwerk): Excel erstellen
        step_num = "4/4" if copy_to_s_drive else "3/3"
        st.write(f"**Schritt {step_num}:** Erstelle Import-Excel...")
        status_text.text("Generiere Excel-Datei...")
        
        excel_path = dir_importfiles / "Import_alle_Bilder_Status.xlsx"
        create_import_excel(dir_originalbilder, excel_path)
        
        progress_bar.progress(1.0)
        status_text.text("✅ Verarbeitung abgeschlossen!")
        st.success("🎉 Alle Bilder erfolgreich verarbeitet!")
        
        # Bei Netzwerkpfad: Zeige wo gespeichert wurde
        if use_network_paths and folder_path:
            st.info("📁 **Dateien wurden gespeichert in:**")
            
            st.markdown("**Projektordner:**")
            st.markdown(f"""
            - **Artikelbild max (TIFF):** `{dir_artikelbild_max}`
            - **Katalogbilder (JPEG):** `{dir_katalog}`
            - **Excel-Import:** `{excel_path}`
            """)
            
            st.markdown("**S-Laufwerk (Multimedia):**")
            if s_drive_copied:
                st.markdown("""
                - **Originale:** `S:/Multimedia/Originale`
                - **Artikelbild max:** `S:/Multimedia/Print/BAD_Artikelbild_maximal`
                - **Katalogbilder:** `S:/Multimedia/Print/HAWAKatalog`
                """)
                st.success("✅ Alle Dateien wurden in Projektordner UND S-Laufwerk gespeichert!")
            else:
                if copy_to_s_drive:
                    st.warning("⚠️ S-Laufwerk nicht erreichbar - nur im Projektordner gespeichert")
                else:
                    st.info("ℹ️ S-Laufwerk-Kopie deaktiviert - nur im Projektordner gespeichert")
        
        # Bei Upload-Modus: Kopiere für Download
        if not use_network_paths:
            perm_dir = Path(tempfile.gettempdir()) / "streamlit_image_results" / str(st.session_state.get('run_id', 'default'))
            perm_dir.mkdir(parents=True, exist_ok=True)
            
            perm_excel = perm_dir / "Import_alle_Bilder_Status.xlsx"
            perm_artikelbild = perm_dir / "artikelbild_max"
            perm_katalog = perm_dir / "katalog"
            
            shutil.copy(excel_path, perm_excel)
            shutil.copytree(dir_artikelbild_max, perm_artikelbild, dirs_exist_ok=True)
            shutil.copytree(dir_katalog, perm_katalog, dirs_exist_ok=True)
            
            excel_path = perm_excel
            dir_artikelbild_max = perm_artikelbild
            dir_katalog = perm_katalog
        
        # Ergebnisse speichern
        results = {
            'excel_path': excel_path,
            'artikelbild_dir': dir_artikelbild_max,
            'katalog_dir': dir_katalog,
            'total_files': total_files,
            'is_network_mode': use_network_paths,
            'copy_to_s_drive': s_drive_copied
        }
        
        return results
        
    except Exception as e:
        st.error(f"❌ Fehler während der Verarbeitung: {e}")
        st.error(traceback.format_exc())
        return None


def crop_white_or_transparent_border(image_path, output_path):
    """Entfernt weissen oder transparenten Rand von Bildern"""
    with Image.open(image_path) as img:
        if img.mode in ("RGBA", "LA"):
            img = img.convert("RGBA")
        else:
            img = img.convert("RGB")
        
        bg = Image.new(img.mode, img.size, 
                      (255, 255, 255, 0) if img.mode == "RGBA" else (255, 255, 255))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        
        if bbox:
            cropped_img = img.crop(bbox)
        else:
            cropped_img = img
        
        output_tiff_path = output_path.with_suffix(".tif")
        output_tiff_path.parent.mkdir(parents=True, exist_ok=True)
        cropped_img.save(output_tiff_path, format='TIFF', save_all=True)


def process_image_for_catalog(image_path, output_path):
    """Erstellt Katalogbild mit spezifischen Dimensionen"""
    target_size = int(3.49 / 2.54 * 300)
    smaller_width = int(3.2 / 2.54 * 300)
    
    with Image.open(image_path) as img:
        img = img.convert("RGBA")
        background = Image.new("RGBA", img.size, "WHITE")
        img = Image.alpha_composite(background, img).convert("L")
        
        width, height = img.size
        if height > width:
            new_height = target_size
            new_width = int((new_height / height) * width)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            new_img = Image.new("L", (target_size, target_size), "white")
            new_img.paste(img, ((target_size - new_width) // 2, 0))
        else:
            new_width = smaller_width
            new_height = int((new_width / width) * height)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            new_img = Image.new("L", (target_size, new_height), "white")
            new_img.paste(img, ((target_size - new_width) // 2, 0))
        
        new_img.save(output_path, 'JPEG')


def create_import_excel(original_dir, output_path):
    """Erstellt Import-Excel-Datei"""
    def process_filenames(folder_path):
        processed_names = []
        for file_name in os.listdir(folder_path):
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp')):
                try:
                    name = file_name.lstrip('0').rsplit('.', 1)[0]
                    if len(name) >= 4:
                        reformatted = name[:4] + ' ' + name[4:].replace('_', '.')
                    else:
                        reformatted = name
                    processed_names.append(reformatted)
                except Exception as e:
                    st.warning(f"⚠️ Fehler: '{file_name}': {e}")
        return processed_names
    
    def get_image_info(folder_path):
        file_names_without_extension = []
        combined_paths = []
        valid_extensions = {'.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp'}
        
        for file_name in os.listdir(folder_path):
            stem, ext = os.path.splitext(file_name)
            if ext.lower() in valid_extensions:
                file_names_without_extension.append(stem)
                combined_paths.append(stem + ext.lower())
        return file_names_without_extension, combined_paths
    
    processed_file_names = process_filenames(original_dir)
    names_without_ext, combined = get_image_info(original_dir)
    
    if not processed_file_names:
        raise ValueError("Keine gültigen Artikelnamen extrahiert")
    
    original = "\\Originale\\"
    bild_max = "\\Print\\BAD_Artikelbild_maximal\\"
    katalog = "\\Print\\HAWAKatalog\\"
    
    df = pd.DataFrame({
        "Reihenfolge": list(range(1, len(names_without_ext) + 1)),
        "Artikel-Nr": processed_file_names,
        "Orginalbild": [original + x for x in combined],
        "BAD Artikelbild maximal": [bild_max + x + ".tif" for x in names_without_ext],
        "BAD Hauptbild für Katalog": [katalog + x + ".jpg" for x in names_without_ext],
        "Status": ["allg. Mutation"] * len(names_without_ext)
    })
    
    df.to_excel(output_path, index=False)
    return output_path


def create_zip_download(results):
    """Erstellt ZIP mit allen Dateien"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.write(results['excel_path'], arcname="Import_alle_Bilder_Status.xlsx")
        
        for file in Path(results['artikelbild_dir']).glob('*.tif'):
            zip_file.write(file, arcname=f"Artikelbild_max/{file.name}")
        
        for file in Path(results['katalog_dir']).glob('*.jpg'):
            zip_file.write(file, arcname=f"Katalog/{file.name}")
    
    zip_buffer.seek(0)
    return zip_buffer


# ============================================================================
# STREAMLIT UI
# ============================================================================

if 'run_id' not in st.session_state:
    import time
    st.session_state.run_id = str(int(time.time()))

# Sidebar
with st.sidebar:
    st.header("⚙️ Einstellungen")
    
    processing_mode = st.radio(
        "Verarbeitungsmodus:",
        ["📁 Netzwerkpfad verwenden", "📤 Bilder hochladen"],
        help="Wähle Netzwerkpfad oder Upload"
    )
    
    st.divider()
    
    folder_path_input = None
    
    if processing_mode == "📁 Netzwerkpfad verwenden":
        folder_path_input = st.text_input(
            "Projektordner-Pfad:",
            placeholder=r"N:\Katalog\__Jira_Tasks_Media_Daten\_____Vorlage PDM-Jira-Nummer_Lieferant_Produktlinie",
            help="Vollständiger Pfad zum Projektordner"
        )
        
        if folder_path_input:
            preview_path = Path(folder_path_input) / "1_Abbildungen" / "1_Originale"
            st.caption(f"📍 Suche in: `{preview_path}`")
    
    st.divider()
    
    # Kopieren-Option - immer sichtbar
    copy_to_s_drive = st.checkbox(
        "Dateien ins S-Laufwerk kopieren",
        value=True,
        key="copy_to_s_drive_checkbox",
        help="Kopiert verarbeitete Bilder automatisch ins S-Laufwerk (wenn verfügbar)"
    )
    

# Tabs
tab1, tab2 = st.tabs(["🖼️ Bildverarbeitung", "📖 Anleitung"])

with tab1:
    if processing_mode == "📤 Bilder hochladen":
        st.subheader("📤 Bilder hochladen")
        uploaded_files = st.file_uploader(
            "Wähle Bilder",
            type=['png', 'jpg', 'jpeg', 'tiff', 'tif', 'bmp'],
            accept_multiple_files=True
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} Dateien")
            
            with st.expander("🔍 Vorschau"):
                cols = st.columns(4)
                for idx, file in enumerate(uploaded_files[:8]):
                    if not file.name.lower().endswith(('.tif', '.tiff')):
                        with cols[idx % 4]:
                            try:
                                st.image(file, caption=file.name, use_container_width=True)
                            except:
                                pass
        
        if st.button("🚀 Verarbeitung starten", type="primary", disabled=not uploaded_files):
            results = process_images_streamlit(uploaded_files, copy_to_s_drive=copy_to_s_drive)
            
            if results:
                st.session_state.img_processing_complete = True
                st.session_state.img_processed_images = results
    
    else:  # Netzwerkpfad
        st.subheader("📁 Netzwerkpfad-Verarbeitung")
        
        if folder_path_input:
            full_path = Path(folder_path_input) / "1_Abbildungen" / "1_Originale"
            st.info(f"📍 Pfad: `{full_path}`")
            
            if full_path.exists():
                st.success("✅ Pfad gefunden!")
                try:
                    files = [f for f in full_path.iterdir() 
                            if f.suffix.lower() in ['.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp']]
                    if files:
                        st.info(f"🖼️ {len(files)} Bilder gefunden")
                    else:
                        st.warning("⚠️ Keine Bilder gefunden")
                except Exception as e:
                    st.warning(f"⚠️ Lesefehler: {e}")
            else:
                st.warning("⚠️ Pfad nicht erreichbar")
        else:
            st.warning("⚠️ Bitte Pfad eingeben")
        
        if st.button("🚀 Verarbeitung starten", type="primary", disabled=not folder_path_input):
            results = process_images_streamlit(None, use_network_paths=True, folder_path=folder_path_input, copy_to_s_drive=copy_to_s_drive)
            
            if results:
                st.session_state.img_processing_complete = True
                st.session_state.img_processed_images = results
    
    # Ergebnisse
    if st.session_state.img_processing_complete and st.session_state.img_processed_images:
        st.divider()
        st.subheader("📦 Ergebnisse")
        
        results = st.session_state.img_processed_images
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Verarbeitete Bilder", results['total_files'])
        with col2:
            if results.get('is_network_mode'):
                st.success("✅ Im Netzwerk gespeichert")
            else:
                st.info("📥 Downloads verfügbar")
        
        if results.get('is_network_mode'):
            st.markdown("### 📁 Gespeicherte Dateien")
            st.markdown(f"""
            **Projektordner:**
            - Artikelbild max: `{results['artikelbild_dir']}`
            - Katalog: `{results['katalog_dir']}`
            - Excel: `{results['excel_path']}`
            """)
            
            if results.get('copy_to_s_drive'):
                st.success("✅ Auch auf S-Laufwerk gespeichert!")
            else:
                # Prüfe ob Option aktiviert war
                if copy_to_s_drive:
                    st.warning("⚠️ S-Laufwerk nicht erreichbar")
                else:
                    st.info("ℹ️ S-Laufwerk-Kopie deaktiviert")
            
            with st.expander("💾 Optional: Download"):
                with open(results['excel_path'], 'rb') as f:
                    st.download_button(
                        "📥 Excel herunterladen",
                        data=f,
                        file_name="Import_alle_Bilder_Status.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            with open(results['excel_path'], 'rb') as f:
                st.download_button(
                    "📥 Excel herunterladen",
                    data=f,
                    file_name="Import_alle_Bilder_Status.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with st.spinner("Erstelle ZIP..."):
                zip_data = create_zip_download(results)
            
            st.download_button(
                "📦 ZIP herunterladen",
                data=zip_data,
                file_name="Bildverarbeitung_Komplett.zip",
                mime="application/zip"
            )
        
        with st.expander("📊 Excel-Vorschau"):
            df = pd.read_excel(results['excel_path'])
            st.dataframe(df, use_container_width=True)

with tab2:
    st.header("📖 Anleitung")
    st.markdown("""
    ### Verwendung
    
    **Netzwerk-Modus (Standard):**
    1. Projektpfad eingeben
    2. Verarbeitung starten
    3. Dateien im Netzwerk gespeichert
    
    **Upload-Modus:**
    1. Bilder hochladen
    2. Verarbeitung starten
    3. Downloads nutzen
    
    ### Ausgabe
    - TIFF (Artikelbild max)
    - JPEG (Katalog)
    - XLSX (Import-Datei)
    """)

st.divider()
st.caption("Bildverarbeitungs-Tool | Version 2.0")