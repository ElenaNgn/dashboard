# -*- coding: utf-8 -*-
"""
Created on Mon Jan  5 13:40:54 2026

@author: e1012121
"""

import streamlit as st
from pathlib import Path
from typing import Optional, List

def create_file_upload_section(
    label: str,
    accepted_types: List[str],
    multiple: bool = True,
    help_text: Optional[str] = None
) -> Optional[List]:
    """Wiederverwendbare File-Upload-Komponente"""
    st.subheader(label)
    
    if help_text:
        st.info(help_text)
    
    uploaded_files = st.file_uploader(
        "Dateien auswählen",
        type=accepted_types,
        accept_multiple_files=multiple,
        key=f"upload_{label}"
    )
    
    if uploaded_files:
        if isinstance(uploaded_files, list):
            st.success(f"✅ {len(uploaded_files)} Dateien hochgeladen")
        else:
            st.success("✅ Datei hochgeladen")
    
    return uploaded_files


def create_path_input_section(
    label: str,
    default_path: str = "",
    path_type: str = "folder",  # "folder" oder "file"
    help_text: Optional[str] = None
) -> Optional[Path]:
    """Wiederverwendbare Pfad-Eingabe mit Validierung"""
    st.subheader(label)
    
    if help_text:
        st.info(help_text)
    
    path_input = st.text_input(
        f"{path_type.capitalize()}-Pfad",
        value=default_path,
        help=f"Vollständiger Pfad zum {path_type}",
        key=f"path_{label}"
    )
    
    if path_input:
        path = Path(path_input)
        
        # Validierung
        if path_type == "folder":
            if path.exists() and path.is_dir():
                st.success(f"✅ Ordner gefunden: {path}")
                return path
            else:
                st.error(f"❌ Ordner nicht gefunden: {path}")
        else:
            if path.exists() and path.is_file():
                st.success(f"✅ Datei gefunden: {path}")
                return path
            else:
                st.error(f"❌ Datei nicht gefunden: {path}")
    
    return None


def create_progress_section(title: str, steps: List[str]) -> dict:
    """Erstellt Progress-Tracking mit mehreren Schritten"""
    st.subheader(title)
    
    progress_bars = {}
    status_texts = {}
    
    for idx, step in enumerate(steps, 1):
        st.write(f"**Schritt {idx}/{len(steps)}:** {step}")
        progress_bars[step] = st.progress(0)
        status_texts[step] = st.empty()
    
    return {
        "progress_bars": progress_bars,
        "status_texts": status_texts
    }


def show_job_status(job, show_details: bool = True):
    """Einheitliche Job-Status-Anzeige"""
    status_icons = {
        "queued": "⏳",
        "processing": "🔄",
        "completed": "✅",
        "error": "❌"
    }
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            "Status",
            f"{status_icons.get(job.status, '❓')} {job.status.upper()}"
        )
    
    with col2:
        st.metric("Dateien", f"{job.processed_files}/{job.total_files}")
    
    with col3:
        if job.end_time:
            duration = (job.end_time - job.start_time).total_seconds()
            st.metric("Dauer", f"{duration:.1f}s")
        else:
            from datetime import datetime
            runtime = (datetime.now() - job.start_time).seconds
            st.metric("Läuft seit", f"{runtime}s")
    
    if show_details and job.status == "processing":
        st.progress(job.progress)
        st.text(f"Aktuell: {job.current_file}")
    
    if job.errors:
        with st.expander(f"⚠️ {len(job.errors)} Fehler"):
            for error in job.errors:
                st.warning(error)