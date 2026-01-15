# -*- coding: utf-8 -*-
"""
Image Format Converter - Streamlit Dashboard
TIF/BMP/PNG to JPG Converter
"""
import streamlit as st
from PIL import Image
import io
import zipfile

# Streamlit Page Configuration
st.set_page_config(
    page_title="TIF to JPG",
    page_icon="🖼️",
    layout="wide"
)

st.title("🖼️ TIF to JPG")
st.markdown("**TIF/BMP/PNG zu JPG**")
st.markdown("---")

def convert_image_to_jpg(image_file):
    """
    Convert image to JPG
    """
    # Open image
    im = Image.open(image_file)
    
    # Get new filename
    base_name = image_file.name.rsplit('.', 1)[0]
    new_filename = f"{base_name}.jpg"
    
    # Convert to RGB
    out = im.convert("RGB")
    
    # Save to bytes
    img_byte_arr = io.BytesIO()
    out.save(img_byte_arr, format='JPEG', quality=95)
    img_byte_arr.seek(0)
    
    return img_byte_arr.getvalue(), new_filename, out

# Main content
st.subheader("📤 Bilder hochladen")

uploaded_files = st.file_uploader(
    "Wähle TIF/BMP/PNG Dateien",
    type=['tif', 'tiff', 'bmp', 'png'],
    accept_multiple_files=True
)

if uploaded_files:
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Hochgeladene Dateien", len(uploaded_files))
    
    with col2:
        total_size = sum(f.size for f in uploaded_files) / (1024 * 1024)
        st.metric("Gesamtgröße", f"{total_size:.2f} MB")
    
    st.markdown("---")
    
    # Convert button
    if st.button("🔄 Zu JPG konvertieren", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        converted_files = []
        errors = []
        
        for idx, file in enumerate(uploaded_files):
            status_text.text(f"Konvertiere: {file.name}...")
            
            try:
                img_data, new_filename, converted_img = convert_image_to_jpg(file)
                
                converted_files.append({
                    'original_name': file.name,
                    'new_name': new_filename,
                    'data': img_data,
                    'image': converted_img,
                    'size': len(img_data)
                })
                
                file.seek(0)
                
            except Exception as e:
                errors.append(f"❌ Fehler bei {file.name}: {str(e)}")
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.text("Konvertierung abgeschlossen!")
        
        if errors:
            for error in errors:
                st.error(error)
        
        if converted_files:
            st.success(f"✅ {len(converted_files)} Dateien erfolgreich konvertiert!")
            st.session_state['converted_files'] = converted_files
            st.rerun()

# Download section
if 'converted_files' in st.session_state and st.session_state['converted_files']:
    converted_files = st.session_state['converted_files']
    
    st.markdown("---")
    st.subheader("📥 Download")
    
    # Download as ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in converted_files:
            zip_file.writestr(file_info['new_name'], file_info['data'])
    
    zip_buffer.seek(0)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="⬇️ Alle JPG-Dateien als ZIP herunterladen",
            data=zip_buffer.getvalue(),
            file_name="konvertierte_bilder.zip",
            mime="application/zip",
            use_container_width=True
        )
    
    st.markdown("---")
    
    # Preview
    st.subheader("👁️ Vorschau")
    
    cols = st.columns(min(4, len(converted_files)))
    for idx, file_info in enumerate(converted_files):
        with cols[idx % 4]:
            st.image(file_info['image'], caption=file_info['new_name'], use_container_width=True)
            st.caption(f"{file_info['size'] / 1024:.1f} KB")
            
            st.download_button(
                label=f"⬇️ Download",
                data=file_info['data'],
                file_name=file_info['new_name'],
                mime="image/jpeg",
                use_container_width=True,
                key=f"download_{idx}"
            )
    
    # Clear button
    st.markdown("---")
    if st.button("🗑️ Neu starten", use_container_width=True):
        del st.session_state['converted_files']
        st.rerun()

else:
    if not uploaded_files:
        st.info("👆 Bitte lade TIF/BMP/PNG Dateien hoch, um zu beginnen.")

# Footer
st.markdown("---")
st.caption("Bildformat-Konverter | Powered by Streamlit")