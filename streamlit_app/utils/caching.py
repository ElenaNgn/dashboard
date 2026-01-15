# -*- coding: utf-8 -*-
"""
Created on Mon Jan  5 13:52:20 2026

@author: e1012121
"""

import streamlit as st
from pathlib import Path
from typing import List

@st.cache_data(ttl=300)
def get_file_list_cached(directory: str, extensions: List[str]) -> List[str]:
    """Cached Directory-Listing"""
    path = Path(directory)
    if not path.exists():
        return []
    
    files = []
    try:
        for file in path.iterdir():
            if file.is_file() and file.suffix.lower() in extensions:
                files.append(str(file))
    except PermissionError:
        st.warning(f"⚠️ Keine Berechtigung: {directory}")
        return []
    
    return files