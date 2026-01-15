# -*- coding: utf-8 -*-
"""

Verarbeitet Masszeichnungen - 
erstellt JPG mit aktuellem Datum und generiert das PDF. 
Man kann wählen, ob halbseitig oder ganzseitig dargestellt werden soll 

ÄNDERUNG v2.1: Datum 5mm nach unten verschoben
ÄNDERUNG v2.2: Bild im PDF 10mm nach oben verschoben
ÄNDERUNG v2.3: Bild im PDF 10mm nach unten (nur halbseitig)

"""

from pathlib import Path
from datetime import datetime
import shutil
import traceback
import json
import threading
import queue
from typing import Optional
import time

import pandas as pd
from PIL import Image, ImageDraw, ImageFont, EpsImagePlugin
import streamlit as st
from utils.caching import get_file_list_cached


# ==========================================================
# KONFIGURATION
# ==========================================================

class Config:
    def __init__(self, project_path: Path):
        """
        Initialisiert Config mit Projekt-Pfad.
        Alle anderen Pfade werden automatisch daraus abgeleitet.
        """
        self.project_path = project_path
        
        # Ordnerstruktur gemäß Vorgabe
        self.dir_masszeichnungen = project_path / "2_Masszeichnungen" / "1_Originale"
        self.dir_mzbearbeitet = project_path / "2_Masszeichnungen" / "2_bearbeitet"
        self.dir_webjpeg = project_path / "2_Masszeichnungen" / "2b_WebJPEG"
        self.dir_importfiles = project_path / "8_Importfiles_Media-Datenpfade"
        
        # Input/Output-Pfade (für interne Verwendung)
        self.eps_data = self.dir_masszeichnungen
        self.ziel_ordner = self.dir_webjpeg
        
        # SAP-Pfade (optional, falls benötigt)
        self.target_jpg_sap = Path(r"S:\Multimedia\SAP\YM1")
        self.target_pdf_sap = Path(r"S:\Multimedia\SAP\YM2")
        
        # Excel wird im Importfiles-Ordner gespeichert
        self.excel_output = None  # Wird später gesetzt
        
        # Fester Pfad zur Word-Vorlage (falls verfügbar)
        self.vorlage_docx = Path(r"N:\Katalog\_ARCHIV_enguyen\ST_Masszeichnungen1.docx")
        
        # Ghostscript-Pfad - versuche automatisch zu finden
        self.ghostscript = self._find_ghostscript()
        
        # Bildverarbeitungs-Parameter
        self.dpi = 300
        self.width_cm = 17.3
        self.height_cm = 21.7
        self.jpg_quality = 97
        self.date_margin = 10
        
        # Padding/Margins für sicheren Bereich (in cm)
        self.margin_top_cm = 0.5
        self.margin_bottom_cm = 1.0  # Mehr Platz unten für Datum
        self.margin_left_cm = 0.3
        self.margin_right_cm = 0.3
        
        self.font_path = "arial.ttf"
        self.font_size = 30
    
    def _find_ghostscript(self):
        """Versucht Ghostscript automatisch zu finden"""
        import sys
        import os
        
        # Mögliche Pfade
        possible_paths = [
            # Conda-Installation
            Path(sys.prefix) / "Library" / "bin" / "gswin64c.exe",
            Path(sys.prefix) / "Library" / "bin" / "gs.exe",
            # Netzwerk-Installation
            Path(r"N:\Katalog\_ARCHIV_enguyen\gs900w64\gs901w64.exe"),
            # Standard Windows-Installation
            Path(r"C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe"),
            Path(r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe"),
        ]
        
        # Prüfe auch PATH
        if "PATH" in os.environ:
            for path_dir in os.environ["PATH"].split(os.pathsep):
                possible_paths.append(Path(path_dir) / "gswin64c.exe")
                possible_paths.append(Path(path_dir) / "gs.exe")
        
        # Finde ersten existierenden Pfad
        for gs_path in possible_paths:
            if gs_path.exists():
                return str(gs_path)
        
        # Fallback
        return r"N:\Katalog\_ARCHIV_enguyen\gs900w64\gs901w64.exe"
    
    def ensure_directories(self):
        """Erstellt alle benötigten Ordner, falls sie nicht existieren"""
        self.dir_masszeichnungen.mkdir(parents=True, exist_ok=True)
        self.dir_mzbearbeitet.mkdir(parents=True, exist_ok=True)
        self.dir_webjpeg.mkdir(parents=True, exist_ok=True)
        self.dir_importfiles.mkdir(parents=True, exist_ok=True)
    
    def get_project_info(self) -> dict:
        """Gibt Informationen über das Projekt zurück"""
        return {
            "project_path": str(self.project_path),
            "input_folder": str(self.dir_masszeichnungen),
            "output_folder": str(self.dir_webjpeg),
            "import_folder": str(self.dir_importfiles),
            "input_exists": self.dir_masszeichnungen.exists(),
            "output_exists": self.dir_webjpeg.exists(),
        }


# ==========================================================
# HILFSFUNKTIONEN
# ==========================================================

def cm_to_px(cm: float, dpi: int) -> int:
    return int(cm / 2.54 * dpi)


def load_font(cfg: Config):
    try:
        return ImageFont.truetype(cfg.font_path, cfg.font_size)
    except Exception:
        return ImageFont.load_default()


def load_image(path: Path) -> Image.Image:
    """Lädt ein Bild (EPS oder JPG) mit verbesserter Fehlerbehandlung"""
    import tempfile
    import subprocess
    
    # Konvertiere zu absolutem Pfad
    abs_path = path.resolve()
    
    # Prüfe ob Datei existiert
    if not abs_path.exists():
        raise FileNotFoundError(f"Datei existiert nicht: {abs_path}")
    
    # Prüfe ob Datei lesbar ist
    if not abs_path.is_file():
        raise ValueError(f"Pfad ist keine Datei: {abs_path}")
    
    try:
        if path.suffix.lower() == ".eps":
            # WORKAROUND: Kopiere EPS IMMER temporär lokal
            temp_dir = Path(tempfile.gettempdir()) / "eps_temp"
            temp_dir.mkdir(exist_ok=True)
            temp_eps = temp_dir / path.name
            temp_png = temp_dir / f"{path.stem}.png"
            
            # Kopiere Datei zu temp
            shutil.copy2(abs_path, temp_eps)
            
            try:
                # Methode 1: Versuche mit PIL/Ghostscript (kann Admin-Rechte benötigen)
                try:
                    with Image.open(str(temp_eps)) as img:
                        img.load(scale=12)
                        result = img.convert("RGB")
                    return result
                except (OSError, PermissionError) as e:
                    # Fehler 740 = Admin-Rechte erforderlich
                    if "740" in str(e) or "erhöhte Rechte" in str(e):
                        # Methode 2: Versuche Ghostscript direkt aufzurufen (ohne Admin)
                        raise Exception("Ghostscript benötigt Admin-Rechte - verwende Fallback")
                    else:
                        raise
                        
            except Exception as gs_error:
                # Fallback: Versuche mit ImageMagick (falls installiert)
                try:
                    # Versuche ImageMagick convert
                    result = subprocess.run(
                        ["magick", "convert", str(temp_eps), str(temp_png)],
                        capture_output=True,
                        timeout=30
                    )
                    if result.returncode == 0 and temp_png.exists():
                        with Image.open(temp_png) as img:
                            return img.convert("RGB")
                except:
                    pass
                
                # Letzter Fallback: Fehlermeldung mit Hinweis
                raise Exception(
                    f"EPS-Verarbeitung fehlgeschlagen: {gs_error}. "
                    f"TIPP: Starte Streamlit als Administrator oder verwende JPG-Dateien statt EPS."
                )
            finally:
                # Cleanup
                try:
                    if temp_eps.exists():
                        temp_eps.unlink()
                    if temp_png.exists():
                        temp_png.unlink()
                except:
                    pass
        else:
            # JPG/JPEG-Verarbeitung
            with Image.open(str(abs_path)) as img:
                return img.convert("RGB")
    except Exception as e:
        # Detaillierte Fehlermeldung
        if "Datei existiert nicht" in str(e) or "ist keine Datei" in str(e):
            raise  # Diese Fehler schon oben geworfen
        raise Exception(f"PIL-Fehler: {type(e).__name__}: {e}")


# ==========================================================
# BILDVERARBEITUNG - VERBESSERT MIT ZENTRIERUNG
# ==========================================================

def create_layout_image(img: Image.Image, name: str, cfg: Config, font) -> Path:
    """
    Erstellt ein Layout-Bild mit automatischer Zentrierung und Skalierung.
    Das Bild wird so skaliert, dass es vollständig in den verfügbaren Bereich passt.
    
    """
    # Canvas-Grösse in Pixel
    canvas_w = cm_to_px(cfg.width_cm, cfg.dpi)
    canvas_h = cm_to_px(cfg.height_cm, cfg.dpi)
    
    # Berechne verfügbaren Bereich (abzüglich Ränder)
    margin_top = cm_to_px(cfg.margin_top_cm, cfg.dpi)
    margin_bottom = cm_to_px(cfg.margin_bottom_cm, cfg.dpi)
    margin_left = cm_to_px(cfg.margin_left_cm, cfg.dpi)
    margin_right = cm_to_px(cfg.margin_right_cm, cfg.dpi)
    
    available_w = canvas_w - margin_left - margin_right
    available_h = canvas_h - margin_top - margin_bottom
    
    # Berechne Skalierungsfaktor (fit to available area)
    scale_w = available_w / img.width
    scale_h = available_h / img.height
    
    # Verwende kleineren Skalierungsfaktor, damit Bild vollständig passt
    scale = min(scale_w, scale_h)
    
    # Neue Bildgrösse berechnen
    new_w = int(img.width * scale)
    new_h = int(img.height * scale)
    
    # Bild skalieren
    img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    
    # Weisser Canvas erstellen
    canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
    
    # Berechne Position für Zentrierung im verfügbaren Bereich
    x_offset = margin_left + (available_w - new_w) // 2
    y_offset = margin_top + (available_h - new_h) // 2
    
    # Bild auf Canvas einfügen (zentriert)
    canvas.paste(img_resized, (x_offset, y_offset))
    
    # Datum hinzufügen
    draw = ImageDraw.Draw(canvas)
    date = datetime.now().strftime("%d.%m.%Y")
    
    # Textgrösse berechnen
    bbox = draw.textbbox((0, 0), date, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    
    # ============================================================
    #  Position für Datum (wie in Beispiel-Zeichnung)
    # - Datum wird rechts unten IM Bildbereich platziert
    # - ÄNDERUNG: 5mm nach unten verschoben
    # ============================================================
    is_halfpage = cfg.height_cm < 15
    
    # Datum rechts unten im Bildbereich (nicht auf weißem Rand)
    # 5mm = ca. 19 Pixel bei 300 DPI (5mm / 25.4mm * 300 = 59.06 Pixel, aber vom Rand aus gemessen)
    padding_right = 30  # Abstand vom rechten Bildrand in Pixeln
    padding_bottom = -49  # GEÄNDERT: 5mm nach unten = 10 - 59 = -49 Pixel (negativ = weiter nach unten)
    text_x = x_offset + new_w - text_w - padding_right
    text_y = y_offset + new_h - text_h - padding_bottom
    
    draw.text((text_x, text_y), date, fill="black", font=font)
    
    # Speichern im WebJPEG-Ordner
    out = cfg.dir_webjpeg / f"{name}.jpg"
    canvas.save(out, "JPEG", dpi=(cfg.dpi, cfg.dpi), quality=cfg.jpg_quality)
    
    return out


# ==========================================================
# PDF
# ==========================================================

def create_pdf(image: Path, artikel: str, cfg: Config):
    """
    Erstellt PDF mit Firmen-Template - OHNE Word zu öffnen

    """
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    from PyPDF2 import PdfReader, PdfWriter
    from io import BytesIO
    
    # PDF wird im gleichen Ordner wie JPG gespeichert (WebJPEG)
    pdf_path = image.with_suffix(".pdf")
    
    # Template-PDF Pfad (muss gleicher Name wie .docx sein, nur .pdf)
    template_pdf = cfg.vorlage_docx.with_suffix('.pdf')
    
    # Prüfe ob Template existiert
    if not template_pdf.exists():
        raise FileNotFoundError(
            f"❌ Template-PDF nicht gefunden: {template_pdf}\n"
            f"Bitte konvertiere die Word-Vorlage einmalig zu PDF:\n"
            f"1. Öffne: {cfg.vorlage_docx}\n"
            f"2. Speichern als PDF mit gleichem Namen"
        )
    
    # ============================================================
    # SCHRITT 1: Erstelle Overlay mit Bild und Artikelnummer
    # ============================================================
    overlay_buffer = BytesIO()
    c = canvas.Canvas(overlay_buffer, pagesize=A4)
    width, height = A4
    
    # Artikelnummer hinzufügen
    c.setFont("Helvetica-Bold", 14)
    artikel_text = f"Artikelnummer {artikel[:4]} {artikel[4:]}"
    # Position: 3cm von links, 3cm von oben
    c.drawString(3*cm, height - 3*cm, artikel_text)
    
    # Bild hinzufügen
    if image.exists():
        img_reader = ImageReader(str(image))
        
        # ============================================================
        # OPTIMIERT: Position und Grösse des Bildes
        # - Halbseitig (< 15cm): Vergrössert, näher am Rand, 10mm nach UNTEN
        # - Ganzseitig: Standard zentriert
        # v2.3: Bild bei halbseitig 10mm nach UNTEN verschoben (4cm statt 3cm)
        # ============================================================
        is_halfpage = cfg.height_cm < 15
        
        if is_halfpage:
            # MAXIMALE Grösse für halbseitiges Format
            img_width = (cfg.width_cm + 1.5) * cm  # +1.5cm breiter (18.8cm statt 17.3cm)
            img_height = (cfg.height_cm + 1.0) * cm  # +1.0cm höher (10.6cm statt 9.6cm)
            
            # Näher am linken Rand (2cm statt 3cm)
            x = 2*cm
            # v2.3: 10mm nach UNTEN verschoben: 3cm + 1cm = 4cm Abstand von oben
            y = height - img_height - 4*cm
        else:
            # Standard für ganzseitiges Format
            img_width = cfg.width_cm * cm
            img_height = cfg.height_cm * cm
            x = (width - img_width) / 2  # Zentriert
            # 10mm nach oben verschoben: 5cm - 1cm = 4cm Abstand von oben
            y = height - img_height - 4*cm
        
        c.drawImage(img_reader, x, y, width=img_width, height=img_height)
    
    c.save()
    overlay_buffer.seek(0)
    
    # ============================================================
    # SCHRITT 2: Kombiniere Template + Overlay
    # ============================================================
    
    # Lade Template-PDF
    template_reader = PdfReader(str(template_pdf))
    overlay_reader = PdfReader(overlay_buffer)
    
    writer = PdfWriter()
    
    # Hole erste Seite des Templates
    template_page = template_reader.pages[0]
    
    # Hole Overlay (Bild + Text)
    overlay_page = overlay_reader.pages[0]
    
    # Merge: Template (mit Logo) + Overlay (mit Bild)
    template_page.merge_page(overlay_page)
    
    # Füge zur Writer hinzu
    writer.add_page(template_page)
    
    # Speichere finales PDF
    with open(pdf_path, 'wb') as output_file:
        writer.write(output_file)
    
    return pdf_path


# ==========================================================
# BACKGROUND WORKER
# ==========================================================

class ProcessingJob:
    """Repräsentiert einen Verarbeitungsjob"""
    def __init__(self, job_id: str, cfg: Config):
        self.job_id = job_id
        self.cfg = cfg
        self.status = "queued"  # queued, processing, completed, error
        self.progress = 0.0
        self.current_file = ""
        self.total_files = 0
        self.processed_files = 0
        self.errors = []
        self.result_df: Optional[pd.DataFrame] = None
        self.start_time = datetime.now()
        self.end_time: Optional[datetime] = None
        self.copied_to_sap = False
        self.found_files = []  # Liste der gefundenen Dateien für Debugging


def process_job_worker(job: ProcessingJob, status_queue: queue.Queue):
    """Worker-Funktion die im Hintergrund läuft"""
    try:
        job.status = "processing"
        status_queue.put(("status", job))
        
        # Ordner erstellen
        job.cfg.ensure_directories()
        
        # Ghostscript konfigurieren und prüfen
        gs_available = False
        gs_error_details = ""
        
        try:
            gs_path = Path(job.cfg.ghostscript)
            
            # Detaillierte Prüfung
            if not gs_path.exists():
                gs_error_details = f"Datei existiert nicht: {gs_path}"
            elif not gs_path.is_file():
                gs_error_details = f"Pfad ist keine Datei: {gs_path}"
            else:
                # Versuche Ghostscript zu setzen
                EpsImagePlugin.gs_windows_binary = str(gs_path)
                gs_available = True
                status_queue.put(("info", f"✓ Ghostscript konfiguriert: {gs_path}"))
        except Exception as e:
            gs_error_details = f"Fehler bei Konfiguration: {type(e).__name__}: {e}"
        
        if not gs_available:
            status_queue.put(("warning", f"⚠️ Ghostscript-Problem: {gs_error_details}"))
            status_queue.put(("warning", "EPS-Dateien können nicht verarbeitet werden. Verwende JPG-Alternativen falls vorhanden."))
        
        Image.MAX_IMAGE_PIXELS = None
        
        font = load_font(job.cfg)
        files = {}
        
        # Dateien sammeln (MIT CACHE)
        job.current_file = "Scanne Dateien..."
        status_queue.put(("status", job))

        # ============================================================
        # OPTIMIERT: Verwende Cache für File-Listing
        # ============================================================
        try:
            # Hole Liste der Dateien (cached)
            file_paths_str = get_file_list_cached(
                str(job.cfg.eps_data),
                ['.eps', '.jpg', '.jpeg']
            )
    
            # Konvertiere zu Path-Objekten und gruppiere nach Dateinamen
            for file_path_str in file_paths_str:
                p = Path(file_path_str)
                files.setdefault(p.stem, []).append(p)
                job.found_files.append(str(p))
    
            status_queue.put(("info", f"✓ {len(file_paths_str)} Dateien gefunden (cached)"))
    
        except Exception as e:
            # Fallback: Wenn Cache fehlschlägt, verwende Original-Methode
            status_queue.put(("warning", f"⚠️ Cache-Fehler, verwende Standard-Suche: {e}"))
    
            for p in job.cfg.eps_data.rglob("*"):
                if p.suffix.lower() in (".eps", ".jpg", ".jpeg"):
                    files.setdefault(p.stem, []).append(p)
                    job.found_files.append(str(p))
        
        job.total_files = len(files)
        status_queue.put(("status", job))
        
        if job.total_files == 0:
            raise ValueError(f"Keine EPS/JPG-Dateien gefunden in {job.cfg.eps_data}")
        
        # Warnung wenn nur EPS-Dateien aber kein Ghostscript
        if not gs_available:
            eps_count = sum(1 for variants in files.values() 
                          if any(p.suffix.lower() == ".eps" for p in variants))
            if eps_count > 0:
                status_queue.put(("warning", f"⚠️ {eps_count} EPS-Dateien gefunden, aber Ghostscript nicht verfügbar!"))
        else:
            # Info über EPS-Verarbeitung
            eps_count = sum(1 for variants in files.values() 
                          if any(p.suffix.lower() == ".eps" for p in variants))
            if eps_count > 0:
                status_queue.put(("info", f"ℹ️ {eps_count} EPS-Dateien werden über lokale Temp-Kopie verarbeitet (Ghostscript-Workaround für Netzwerkpfade)"))
        
        # Verarbeitung
        for idx, (name, variants) in enumerate(files.items()):
            src = None  # Initialisiere für Fehlerbehandlung
            try:
                job.current_file = name
                job.processed_files = idx
                job.progress = idx / job.total_files if job.total_files > 0 else 0
                status_queue.put(("status", job))
                
                # Bevorzuge EPS, sonst JPG
                src = next((p for p in variants if p.suffix.lower() == ".eps"), variants[0])
                
                # Prüfe ob Quelldatei existiert
                if not src.exists():
                    raise FileNotFoundError(f"Quelldatei nicht gefunden: {src}")
                
                # Lade Bild - mit Fallback auf JPG wenn EPS fehlschlägt
                img = None
                eps_failed = False
                try:
                    img = load_image(src)
                except Exception as e:
                    error_str = str(e)
                    # Wenn EPS fehlschlägt wegen Admin-Rechten
                    if src.suffix.lower() == ".eps" and ("740" in error_str or "erhöhte Rechte" in error_str or "Admin" in error_str):
                        eps_failed = True
                        # Suche nach JPG-Alternative
                        jpg_alternatives = [p for p in variants if p.suffix.lower() in (".jpg", ".jpeg")]
                        if jpg_alternatives:
                            job.errors.append(f"Warnung bei {name}: EPS benötigt Admin-Rechte, verwende JPG")
                            src = jpg_alternatives[0]
                            try:
                                img = load_image(src)
                            except Exception as jpg_error:
                                raise Exception(f"Auch JPG fehlgeschlagen: {jpg_error}")
                        else:
                            raise Exception("EPS-Verarbeitung benötigt Admin-Rechte und keine JPG-Alternative vorhanden. Bitte EPS zu JPG konvertieren.")
                    else:
                        raise Exception(f"Fehler beim Laden des Bildes von '{src}': {e}")
                
                if img is None:
                    raise Exception("Bild konnte nicht geladen werden")
                
                # Erstelle Layout-JPG (NEUE ZENTRIERUNG + OPTIMIERTES DATUM - 5MM NACH UNTEN)
                try:
                    jpg = create_layout_image(img, name, job.cfg, font)
                except Exception as e:
                    raise Exception(f"Fehler beim Erstellen des Layout-Bildes: {e}")
                
                # Prüfe ob JPG erstellt wurde
                if not jpg.exists():
                    raise FileNotFoundError(f"Layout-JPG wurde nicht erstellt: {jpg}")
                
                # Erstelle PDF (auch im WebJPEG-Ordner, OPTIMIERT FÜR HALBSEITIG, BILD 10MM NACH UNTEN)
                try:
                    create_pdf(jpg, name, job.cfg)
                except Exception as e:
                    raise Exception(f"Fehler beim Erstellen des PDFs: {e}")
                
            except Exception as e:
                error_msg = f"Fehler bei {name}"
                if src:
                    error_msg += f" (Quelle: {src})"
                error_msg += f": {str(e)}"
                job.errors.append(error_msg)
                status_queue.put(("error", error_msg))
        
        # Kopieren nach SAP (optional)
        job.current_file = "Kopiere nach SAP..."
        status_queue.put(("status", job))
        
        try:
            job.cfg.target_jpg_sap.mkdir(parents=True, exist_ok=True)
            job.cfg.target_pdf_sap.mkdir(parents=True, exist_ok=True)
            
            jpgs = list(job.cfg.dir_webjpeg.glob("*.jpg"))
            pdfs = list(job.cfg.dir_webjpeg.glob("*.pdf"))
            
            for f in jpgs:
                shutil.copy2(f, job.cfg.target_jpg_sap / f.name)
            for f in pdfs:
                shutil.copy2(f, job.cfg.target_pdf_sap / f.name)
            
            job.copied_to_sap = True
        except Exception as e:
            job.errors.append(f"Warnung: Kopieren nach SAP fehlgeschlagen: {e}")
            job.copied_to_sap = False
            status_queue.put(("warning", f"SAP-Kopieren fehlgeschlagen: {e}"))
        
        # Excel erstellen (im Importfiles-Ordner)
        job.current_file = "Erstelle Excel..."
        status_queue.put(("status", job))
        
        jpgs = list(job.cfg.dir_webjpeg.glob("*.jpg"))
        
        # Pfade für Excel - Format: SAP\YM1\datei.jpg
        jpg_paths = []
        pdf_paths = []
        
        for f in jpgs:
            jpg_paths.append(f"\\SAP\\YM1\\{f.name}")
            pdf_paths.append(f"\\SAP\\YM2\\{f.stem}.pdf")
        
        df = pd.DataFrame({
            "Reihenfolge": range(1, len(jpgs) + 1),
            "Artikel-Nr": [f.stem[:4] + " " + f.stem[4:] for f in jpgs],
            "Masszeichnung JPG (YM1)": jpg_paths,
            "Masszeichnung PDF (YM2)": pdf_paths,
            "Status": ["allg. Mutation"] * len(jpgs)
        })
        
        # Excel Output-Pfad im Importfiles-Ordner
        job.cfg.excel_output = job.cfg.dir_importfiles / f"Import_MZ_{job.job_id}.xlsx"
        
        df.to_excel(job.cfg.excel_output, index=False)
        job.result_df = df
        
        # Fertig!
        job.status = "completed"
        job.progress = 1.0
        job.end_time = datetime.now()
        job.current_file = "Abgeschlossen"
        status_queue.put(("status", job))
        status_queue.put(("complete", job))
        
    except Exception as e:
        job.status = "error"
        job.errors.append(f"Fataler Fehler: {str(e)}\n{traceback.format_exc()}")
        job.end_time = datetime.now()
        status_queue.put(("status", job))
        status_queue.put(("fatal_error", str(e)))


# ==========================================================
# STREAMLIT UI
# ==========================================================

st.set_page_config(
    page_title="Masszeichnungen Verarbeitung",
    page_icon="📐",
    layout="wide"
)

st.title("📐 Masszeichnungen Verarbeiten")
st.markdown("""
Dieses Skript verarbeitet Masszeichnungen nach der Projektordner-Struktur:
- **Input**: `2_Masszeichnungen/1_Originale` (EPS/JPG)
- **Output**: `2_Masszeichnungen/2b_WebJPEG` (JPG + PDF mit Datum)
- **Import**: `8_Importfiles_Media-Datenpfade` (Excel-File)
- **Hintergrund-Verarbeitung**: Jobs laufen weiter, auch bei Seitenwechsel
- **✨ ÄNDERUNGEN v2.3**: Halbseitig: Bild im PDF 10mm nach unten verschoben
""")

st.divider()

# Session State initialisieren
if 'mz_jobs' not in st.session_state:
    st.session_state.mz_jobs = {}
if 'mz_status_queues' not in st.session_state:
    st.session_state.mz_status_queues = {}
if 'mz_job_counter' not in st.session_state:
    st.session_state.mz_job_counter = 0
if 'mz_project_path' not in st.session_state:
    st.session_state.mz_project_path = None

# ==========================================================
# SIDEBAR - KONFIGURATION
# ==========================================================

with st.sidebar:
    
    st.subheader("⚙️ Optionen")
    
    # Format-Auswahl IMMER sichtbar (vor Projektpfad)
    halbseitig = st.checkbox(
        "Halbseitiges Format (9.6 cm statt 21.7 cm)",
        value=False,
        key="format_checkbox"
    )
    
    if halbseitig:
        st.success("✨ v2.3: Bild im PDF 10mm nach unten verschoben!")
    
    # Session State für Format-Einstellungen
    if 'mz_margin_top' not in st.session_state:
        st.session_state.mz_margin_top = 0.5
    if 'mz_margin_bottom' not in st.session_state:
        st.session_state.mz_margin_bottom = 1.0
    if 'mz_margin_left' not in st.session_state:
        st.session_state.mz_margin_left = 0.3
    if 'mz_margin_right' not in st.session_state:
        st.session_state.mz_margin_right = 0.3
    
    # Layout-Einstellungen IMMER sichtbar
    with st.expander("📏 Layout-Einstellungen"):
        st.markdown("**Seitenränder (in cm)**")
        st.session_state.mz_margin_top = st.number_input(
            "Oben:", 
            min_value=0.0, 
            max_value=2.0, 
            value=st.session_state.mz_margin_top, 
            step=0.1,
            key="margin_top_input"
        )
        st.session_state.mz_margin_bottom = st.number_input(
            "Unten:", 
            min_value=0.0, 
            max_value=3.0, 
            value=st.session_state.mz_margin_bottom, 
            step=0.1,
            key="margin_bottom_input"
        )
        st.session_state.mz_margin_left = st.number_input(
            "Links:", 
            min_value=0.0, 
            max_value=2.0, 
            value=st.session_state.mz_margin_left, 
            step=0.1,
            key="margin_left_input"
        )
        st.session_state.mz_margin_right = st.number_input(
            "Rechts:", 
            min_value=0.0, 
            max_value=2.0, 
            value=st.session_state.mz_margin_right, 
            step=0.1,
            key="margin_right_input"
        )
        
        st.info("💡 Die Bilder werden automatisch zentriert und skaliert, um in den verfügbaren Bereich zu passen.")
    
    st.divider()
    st.subheader("📁 Projektpfad")
    
    # Projekt-Pfad eingeben
    project_path_input = st.text_input(
        "Projektordner:",
        value=str(st.session_state.mz_project_path) if st.session_state.mz_project_path else "",
        help="Hauptordner des Projekts (enthält Unterordner 2_Masszeichnungen, 8_Importfiles_Media-Datenpfade, etc.)",
        placeholder="Z:\\Projekte\\MeinProjekt"
    )
    
    # Config erstellen wenn Pfad gesetzt
    cfg = None
    if project_path_input:
        project_path = Path(project_path_input)
        st.session_state.mz_project_path = project_path
        cfg = Config(project_path)
        
        # Wende Format-Einstellungen auf Config an
        cfg.height_cm = 9.6 if halbseitig else 21.7
        cfg.margin_top_cm = st.session_state.mz_margin_top
        cfg.margin_bottom_cm = st.session_state.mz_margin_bottom
        cfg.margin_left_cm = st.session_state.mz_margin_left
        cfg.margin_right_cm = st.session_state.mz_margin_right
        
        st.divider()
        st.markdown("### 📂 Ordnerstruktur")
        
        # Input-Ordner
        if cfg.dir_masszeichnungen.exists():
            st.success(f"✅ **1_Originale** gefunden")
            try:
                eps_count = len(list(cfg.dir_masszeichnungen.rglob("*.eps")))
                jpg_count = len(list(cfg.dir_masszeichnungen.rglob("*.jpg")))
                st.info(f"📊 {eps_count} EPS, {jpg_count} JPG")
            except:
                pass
        else:
            st.error(f"❌ **1_Originale** nicht gefunden")
        
        st.code(cfg.dir_masszeichnungen, language=None)
        
        # Output-Ordner
        if cfg.dir_webjpeg.exists():
            st.success(f"✅ **2b_WebJPEG** gefunden")
        else:
            st.warning(f"⚠️ **2b_WebJPEG** wird erstellt")
        
        st.code(cfg.dir_webjpeg, language=None)
        
        # Import-Ordner
        st.info("📄 **Importfiles**")
        st.code(cfg.dir_importfiles, language=None)
        
        # SAP-Pfade und erweiterte Einstellungen (optional änderbar)
        with st.expander("🔧 Erweiterte Einstellungen"):
            cfg.ghostscript = st.text_input(
                "Ghostscript-Pfad:",
                value=str(cfg.ghostscript),
                help="Pfad zur Ghostscript-EXE für EPS-Verarbeitung"
            )
            
            cfg.target_jpg_sap = Path(st.text_input(
                "SAP JPG-Ziel (YM1):",
                value=str(cfg.target_jpg_sap)
            ))
            cfg.target_pdf_sap = Path(st.text_input(
                "SAP PDF-Ziel (YM2):",
                value=str(cfg.target_pdf_sap)
            ))
            
            # Word-Vorlage (optional änderbar)
            cfg.vorlage_docx = Path(st.text_input(
                "Word-Vorlage:",
                value=str(cfg.vorlage_docx),
                help="Optional: DOCX-Vorlage für PDF-Erstellung"
            ))
    
    else:
        st.info("👆 Bitte Projektpfad eingeben")
    
    st.divider()
    st.markdown("### 📋 Info")
    st.markdown("""
    **Ordnerstruktur:**
    ```
    Projektordner/
    ├── 2_Masszeichnungen/
    │   ├── 1_Originale/      ← Input (EPS/JPG)
    │   ├── 2_bearbeitet/     
    │   └── 2b_WebJPEG/       ← Output (JPG+PDF)
    └── 8_Importfiles_Media-Datenpfade/
        └── Import_MZ_*.xlsx  ← Excel
    ```
    
    **Ausgabe:**
    - JPG (mit Datum, zentriert)
    - PDF (mit Artikelnummer)
    - Excel-Importfile
    
    **✨ Änderungen v2.3:**
    - Halbseitig: Bild im PDF 10mm nach unten verschoben
    """)


# ==========================================================
# HAUPTBEREICH - TABS
# ==========================================================

tab1, tab2, tab3 = st.tabs(["🚀 Verarbeitung", "📋 Jobs", "📖 Information"])

# TAB 1: VERARBEITUNG STARTEN
with tab1:
    st.subheader("Neue Verarbeitung starten")
    
    if cfg:
        # Zeige Konfiguration
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.info("📂 **Input (1_Originale)**")
            if cfg.dir_masszeichnungen.exists():
                try:
                    eps_count = len(list(cfg.dir_masszeichnungen.rglob("*.eps")))
                    jpg_count = len(list(cfg.dir_masszeichnungen.rglob("*.jpg")))
                    st.success(f"✅ {eps_count} EPS, {jpg_count} JPG")
                except Exception as e:
                    st.warning(f"⚠️ Konnte Dateien nicht zählen: {e}")
            else:
                st.error("❌ Ordner existiert nicht")
        
        with col2:
            st.info("📂 **Output (2b_WebJPEG)**")
            format_text = "Halbseitig (9.6 cm)" if halbseitig else "Ganzseitig (21.7 cm)"
            st.success("✅ Bereit")
            st.caption(f"Format: {format_text}")
        
        with col3:
            st.info("📄 **Import (Importfiles)**")
            st.success("✅ Bereit")
        
        st.divider()
        
        # Validierung
        can_start = True
        if not cfg.dir_masszeichnungen.exists():
            st.error("❌ Input-Ordner (1_Originale) existiert nicht")
            can_start = False
        
        if not cfg.vorlage_docx.exists():
            st.warning("⚠️ Word-Vorlage nicht gefunden - PDF-Erstellung evtl. eingeschränkt")
        
        # Start-Button
        if st.button("🚀 Verarbeitung starten", type="primary", disabled=not can_start):
            job_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            job = ProcessingJob(job_id, cfg)
            status_queue = queue.Queue()
            
            # Job speichern
            st.session_state.mz_jobs[job_id] = job
            st.session_state.mz_status_queues[job_id] = status_queue
            
            # Thread starten (läuft im Hintergrund!)
            thread = threading.Thread(
                target=process_job_worker,
                args=(job, status_queue),
                daemon=True
            )
            thread.start()
            
            st.success(f"✅ Job {job_id} gestartet!")
            st.info("💡 **Der Job läuft im Hintergrund!** Sie können zu anderen Seiten navigieren.")
            time.sleep(1)
            st.rerun()
    
    else:
        st.warning("⚠️ Bitte Projektpfad in der Sidebar eingeben")
        st.info("👈 Geben Sie den Hauptordner Ihres Projekts an (z.B. `Z:\\Projekte\\MeinProjekt`)")
        
        # Zeige gewähltes Format auch ohne Config
        format_text = "Halbseitig (9.6 cm)" if st.session_state.get("format_checkbox", False) else "Ganzseitig (21.7 cm)"
        st.info(f"📏 **Gewähltes Format:** {format_text}")
        st.caption("Format kann in der Sidebar angepasst werden")

# TAB 2: JOBS ÜBERSICHT
with tab2:
    st.subheader("📋 Aktive und abgeschlossene Jobs")
    
    # Status-Updates aus Queue holen
    for job_id, status_queue in list(st.session_state.mz_status_queues.items()):
        try:
            while True:
                msg_type, msg_data = status_queue.get_nowait()
                if msg_type in ["status", "complete"]:
                    st.session_state.mz_jobs[job_id] = msg_data
                elif msg_type == "error":
                    # Error wird in job.errors gespeichert
                    pass
                elif msg_type == "warning":
                    # Warnings auch in errors speichern
                    if job_id in st.session_state.mz_jobs:
                        st.session_state.mz_jobs[job_id].errors.append(str(msg_data))
        except queue.Empty:
            pass
    
    # Jobs anzeigen
    if st.session_state.mz_jobs:
        # Sortiere: Aktive zuerst, dann neueste zuerst
        sorted_jobs = sorted(
            st.session_state.mz_jobs.items(),
            key=lambda x: (x[1].status != "processing", x[0]),
            reverse=True
        )
        
        for job_id, job in sorted_jobs:
            # Status-Icon
            status_icons = {
                "queued": "⏳",
                "processing": "🔄",
                "completed": "✅",
                "error": "❌"
            }
            icon = status_icons.get(job.status, "❓")
            
            with st.expander(
                f"{icon} Job {job_id} - {job.status.upper()}", 
                expanded=(job.status == "processing")
            ):
                # Metriken
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Status", job.status)
                with col2:
                    st.metric("Dateien", f"{job.processed_files}/{job.total_files}")
                with col3:
                    if job.end_time:
                        duration = (job.end_time - job.start_time).total_seconds()
                        st.metric("Dauer", f"{duration:.1f}s")
                    else:
                        elapsed = (datetime.now() - job.start_time).seconds
                        st.metric("Läuft seit", f"{elapsed}s")
                with col4:
                    sap_status = "✅ Ja" if job.copied_to_sap else "⚠️ Nein"
                    st.metric("SAP-Kopie", sap_status)
                
                # Progress Bar
                if job.status == "processing":
                    st.progress(job.progress)
                    st.info(f"🔄 Aktuell: {job.current_file}")
                
                # Fehler
                if job.errors:
                    st.markdown(f"### ⚠️ {len(job.errors)} Fehler/Warnungen")
                    for error in job.errors:
                        st.warning(error)
                
                # Debug: Gefundene Dateien anzeigen
                if job.found_files and st.checkbox(f"🔍 Zeige {len(job.found_files)} gefundene Dateien", key=f"debug_{job_id}"):
                    st.markdown("**Gefundene Quelldateien:**")
                    for f in job.found_files[:20]:  # Nur erste 20 anzeigen
                        st.code(f, language=None)
                    if len(job.found_files) > 20:
                        st.info(f"... und {len(job.found_files) - 20} weitere")
                
                # Ergebnisse
                if job.status == "completed" and job.result_df is not None:
                    st.success("✅ Erfolgreich abgeschlossen!")
                    
                    # Speicherorte
                    st.markdown("### 📁 Gespeicherte Dateien")
                    
                    st.markdown("**Projektordner:**")
                    st.code(f"""
JPG & PDF: {job.cfg.dir_webjpeg}
Excel: {job.cfg.excel_output}
                    """)
                    
                    if job.copied_to_sap:
                        st.markdown("**SAP-Laufwerk:**")
                        st.code(f"""
JPG (YM1): {job.cfg.target_jpg_sap}
PDF (YM2): {job.cfg.target_pdf_sap}
                        """)
                    
                    # Excel-Vorschau
                    st.markdown("### 📊 Excel-Vorschau")
                    st.dataframe(job.result_df, use_container_width=True)
                    
                    # Download-Option
                    st.markdown("### 💾 Optional: Excel herunterladen")
                    if job.cfg.excel_output.exists():
                        with open(job.cfg.excel_output, 'rb') as f:
                            st.download_button(
                                label="📥 Excel-Import herunterladen",
                                data=f,
                                file_name=job.cfg.excel_output.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{job_id}"
                            )
                    
                    # Remove-Button
                    st.divider()
                    if st.button(f"🗑️ Job {job_id} aus Liste entfernen", key=f"remove_{job_id}"):
                        del st.session_state.mz_jobs[job_id]
                        if job_id in st.session_state.mz_status_queues:
                            del st.session_state.mz_status_queues[job_id]
                        st.rerun()
    
    else:
        st.info("ℹ️ Noch keine Jobs gestartet. Starte einen Job im Tab 'Verarbeitung'.")

# TAB 3: Information
with tab3:
    
    st.markdown("""
    
    ### Ordnerstruktur im Detail
    
    ```
    Projektordner/
    │
    ├── 2_Masszeichnungen/
    │   ├── 1_Originale/              
    │   │   ├── 12345_Artikel1.eps
    │   │   ├── 12346_Artikel2.jpg
    │   │   └── Unterordner/
    │   │       └── 12347_Artikel3.eps
    │   │
    │   ├── 2_bearbeitet/              (Reserviert für manuelle Bearbeitungen)
    │   │
    │   └── 2b_WebJPEG/                ← HIER: Verarbeitete Dateien
    │       ├── 12345_Artikel1.jpg    (mit Datum, zentriert/optimiert)
    │       ├── 12345_Artikel1.pdf    (mit Artikelnummer, optimiert)
    │       ├── 12346_Artikel2.jpg
    │       └── 12346_Artikel2.pdf
    │
    └── 8_Importfiles_Media-Datenpfade/
        └── Import_MZ_20240109_153045.xlsx  ← HIER: Excel-File
    ```
    
    **✨ Änderungen v2.3:**
    - Halbseitiges Format: Bild im PDF 10mm nach unten verschoben (4cm statt 3cm von oben)
    - Ganzseitiges Format: Bleibt unverändert bei 4cm von oben
    
    """)
    

# Auto-Refresh für laufende Jobs (nur wenn mindestens ein Job läuft)
if any(job.status == "processing" for job in st.session_state.mz_jobs.values()):
    time.sleep(2)
    st.rerun()

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <small>Masszeichnungs-Verarbeitung v2.3 | Halbseitig: Bild 10mm nach unten ✨</small>
</div>
""", unsafe_allow_html=True)