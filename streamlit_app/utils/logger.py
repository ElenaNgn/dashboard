# -*- coding: utf-8 -*-
"""
Created on Mon Jan  5 13:42:23 2026

@author: e1012121
"""

import logging
from pathlib import Path
from datetime import datetime
import streamlit as st
from core.config import get_config

class StreamlitLogger:
    """Logger mit Streamlit-Integration"""
    
    def __init__(self, name: str):
        self.name = name
        self.config = get_config()
        
        # Setup Logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        
        # File Handler
        log_file = self.config.logs_dir / f"{name}_{datetime.now():%Y%m%d}.log"
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        
        # Session State für UI-Logs
        if 'ui_logs' not in st.session_state:
            st.session_state.ui_logs = []
    
    def debug(self, msg: str):
        """Debug-Level Logging"""
        self.logger.debug(msg)
    
    def info(self, msg: str, show_in_ui: bool = False):
        """Info-Level Logging"""
        self.logger.info(msg)
        if show_in_ui:
            st.info(msg)
            st.session_state.ui_logs.append(("info", msg, datetime.now()))
    
    def warning(self, msg: str, show_in_ui: bool = True):
        """Warning-Level Logging"""
        self.logger.warning(msg)
        if show_in_ui:
            st.warning(msg)
            st.session_state.ui_logs.append(("warning", msg, datetime.now()))
    
    def error(self, msg: str, show_in_ui: bool = True, exception: Exception = None):
        """Error-Level Logging"""
        if exception:
            self.logger.error(msg, exc_info=True)
        else:
            self.logger.error(msg)
        
        if show_in_ui:
            st.error(msg)
            st.session_state.ui_logs.append(("error", msg, datetime.now()))
    
    def success(self, msg: str, show_in_ui: bool = True):
        """Success message"""
        self.logger.info(f"SUCCESS: {msg}")
        if show_in_ui:
            st.success(msg)
            st.session_state.ui_logs.append(("success", msg, datetime.now()))
    
    def show_log_viewer(self):
        """Zeigt Log-Viewer in Streamlit"""
        if st.session_state.ui_logs:
            with st.expander("📋 Log-Verlauf anzeigen", expanded=False):
                for log_type, msg, timestamp in reversed(st.session_state.ui_logs[-50:]):
                    st.text(f"[{timestamp:%H:%M:%S}] {log_type.upper()}: {msg}")


def get_logger(name: str) -> StreamlitLogger:
    """Factory-Funktion für Logger"""
    return StreamlitLogger(name)