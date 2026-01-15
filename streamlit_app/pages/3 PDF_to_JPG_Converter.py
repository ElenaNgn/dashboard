# -*- coding: utf-8 -*-
"""
PDF to JPG Converter - Streamlit Dashboard
"""
import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import os
import io
import zipfile
from pathlib import Path

# Streamlit Page Configuration
st.set_page_config(
    page_title="PDF zu JPG Konverter",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF zu JPG Konverter")
st.markdown("---")

def convert_pdf_to_jpg(pdf_file, zoom_factor=2.0):
    """
    Converts a PDF file to JPG images
    Returns a list of tuples (image, filename)
    """
    images = []
    
    # Read PDF from uploaded file
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Define zoom matrix
    zoom_matrix = fitz.Matrix(zoom_factor, zoom_factor)
    
    # Iterate through all pages
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        
        # Convert page to image
        pix = page.get_pixmap(matrix=zoom_matrix)
        
        # Create PIL image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Generate filename
        base_filename = os.path.splitext(pdf_file.name)[0]
        if pdf_document.page_count > 1:
            output_filename = f"{base_filename}_seite{page_num + 1}.jpg"
        else:
            output_filename = f"{base_filename}.jpg"
        
        images.append((img, output_filename))
    
    pdf_document.close()
    return images

def create_zip_file(images):
    """
    Creates a ZIP file containing all converted images
    Returns the ZIP file as bytes
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for img, filename in images:
            # Convert image to bytes
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG', quality=95)
            img_bytes = img_byte_arr.getvalue()
            
            # Add to ZIP
            zip_file.writestr(filename, img_bytes)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# Sidebar for settings
with st.sidebar:
    st.header("⚙️ Einstellungen")
    zoom_factor = st.slider(
        "Zoom-Faktor (Bildqualität)",
        min_value=1.0,
        max_value=5.0,
        value=3.0,
        step=0.5,
        help="Höhere Werte = bessere Qualität, aber grössere Dateien"
    )
    
    st.markdown("---")
    st.info("**Hinweis:** Höhere Zoom-Faktoren erzeugen schärfere Bilder, benötigen aber mehr Speicher.")

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📤 PDF-Dateien hochladen")
    uploaded_files = st.file_uploader(
        "Wähle eine oder mehrere PDF-Dateien",
        type=['pdf'],
        accept_multiple_files=True,
        help="Du kannst mehrere PDFs gleichzeitig auswählen"
    )

with col2:
    if uploaded_files:
        st.metric("Hochgeladene Dateien", len(uploaded_files))

# Process uploaded files
if uploaded_files:
    st.markdown("---")
    
    if st.button("🔄 Konvertierung starten", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_images = []
        
        for idx, pdf_file in enumerate(uploaded_files):
            status_text.text(f"Konvertiere: {pdf_file.name}...")
            
            try:
                images = convert_pdf_to_jpg(pdf_file, zoom_factor=zoom_factor)
                all_images.extend(images)
                
                st.success(f"✅ {pdf_file.name} erfolgreich konvertiert ({len(images)} Seite(n))")
                
            except Exception as e:
                st.error(f"❌ Fehler bei {pdf_file.name}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.text("Konvertierung abgeschlossen!")
        
        # Display results
        if all_images:
            st.markdown("---")
            st.subheader("📥 Download-Bereich")
            
            # Main download button for all images as ZIP
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                zip_data = create_zip_file(all_images)
                st.download_button(
                    label=f"📦 Alle Bilder herunterladen ({len(all_images)} Dateien als ZIP)",
                    data=zip_data,
                    file_name="konvertierte_bilder.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )
            
            st.markdown("---")
            st.subheader("📄 Einzelne Dateien")
            
            # Create columns for individual download buttons
            cols = st.columns(min(3, len(all_images)))
            
            for idx, (img, filename) in enumerate(all_images):
                with cols[idx % 3]:
                    # Convert image to bytes for download
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG', quality=95)
                    img_byte_arr = img_byte_arr.getvalue()
                    
                    st.download_button(
                        label=f"⬇️ {filename}",
                        data=img_byte_arr,
                        file_name=filename,
                        mime="image/jpeg",
                        use_container_width=True
                    )
            
            # Preview section
            st.markdown("---")
            st.subheader("👁️ Vorschau")
            
            preview_cols = st.columns(min(3, len(all_images)))
            
            for idx, (img, filename) in enumerate(all_images):
                with preview_cols[idx % 3]:
                    st.image(img, caption=filename, use_container_width=True)
                    st.caption(f"Größe: {img.width}x{img.height} px")

else:
    st.info("Lade eine oder mehrere PDF-Dateien hoch, um zu beginnen.")
    
    # Show example/instructions
    with st.expander("ℹ️ Anleitung"):
        st.markdown("""
        ### So funktioniert's:
        
        1. **PDF-Dateien hochladen**: Klicke auf "Browse files" oder ziehe Dateien in das Upload-Feld
        2. **Zoom-Faktor anpassen**: Stelle in der Sidebar die gewünschte Bildqualität ein
        3. **Konvertierung starten**: Klicke auf den Button "Konvertierung starten"
        4. **Download**: 
           - Lade alle Bilder auf einmal als ZIP-Datei herunter, oder
           - Lade einzelne JPG-Dateien herunter
        
        ### Tipps:
        - Mehrseitige PDFs werden automatisch in separate JPG-Dateien aufgeteilt
        - Höhere Zoom-Faktoren (3.0-5.0) eignen sich für technische Zeichnungen
        - Niedrigere Zoom-Faktoren (1.0-2.0) für einfache Dokumente
        - Bei mehreren Dateien empfiehlt sich der ZIP-Download
        """)

# Footer
st.markdown("---")
st.caption("PDF zu JPG Konverter | Powered by Streamlit & PyMuPDF")