import io
from datetime import datetime
import os
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG VÀ SESSION STATE
# ==========================================
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thư mục mặc định ban đầu
DEFAULT_EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(DEFAULT_EXCEL_DIR):
  DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

if "is_admin" not in st.session_state:
  st.session_state.is_admin = False
if "excel_dir" not in st.session_state:
  st.session_state.excel_dir = DEFAULT_EXCEL_DIR
if "lan_hieuchinh" not in st.session_state:
  st.session_state.lan_hieuchinh = "001"
if "main_menu" not in st.session_state:
  st.session_state.main_menu = "1 🌐 DS WEBsites_CV"

# ==========================================
# GIAO DIỆN CSS iPHONE SLIDE TO UNLOCK
# ==========================================
css_style = """
<style>
/* Khoảng cách chính */
.block-container { 
    padding-top: 3.5rem !important; 
    padding-bottom: 1rem !important; 
}

/* Nền Dark Mode */
.stApp { 
    background-color: #0E1117 !important; 
    color: #E0E6ED !important; 
}

/* ---------------------------------------------------- */
/* 🔥 BIẾN NÚT SIDEBAR THU GỌN THÀNH IPHONE SLIDE BAR   */
/* ---------------------------------------------------- */
[data-testid="stSidebarCollapsedControl"] {
    position: fixed !important;
    top: 15px !important;
    left: 15px !important;
    z-index: 99999 !important;
    display: flex !important;
    align-items: center !important;
    background: rgba(22, 27, 34, 0.75) !important;
    backdrop-filter: blur(12px) !important;
    -webkit-backdrop-filter: blur(12px) !important;
    border: 1px solid rgba(66, 165, 245, 0.3) !important;
    border-radius: 30px !important;
    padding: 4px 16px 4px 6px !important;
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.4), 0 0 12px rgba(30, 136, 229, 0.2) !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    cursor: pointer !important;
}

/* Hiệu ứng Hover rực sáng cho Slide Bar */
[data-testid="stSidebarCollapsedControl"]:hover {
    border-color: #42A5F5 !important;
    background: rgba(30, 136, 229, 0.25) !important;
    box-shadow: 0 8px 25px rgba(30, 136, 229, 0.5), 0 0 15px rgba(66, 165, 245, 0.4) !important;
    transform: translateY(-1px) scale(1.02) !important;
}

/* Định dạng Nút icon tròn bên trong (Knob Slide) */
[data-testid="stSidebarCollapsedControl"] button {
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important;
    color: #FFFFFF !important;
    border-radius: 50% !important;
    width: 32px !important;
    height: 32px !important;
    border: none !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3) !important;
    transition: transform 0.3s ease !important;
}

[data-testid="stSidebarCollapsedControl"]:hover button {
    transform: translateX(4px) !important;
}

/* Tạo dòng chữ hiệu ứng vệt sáng "TRƯỢT ĐỂ MỞ ❯❯" dạng iOS */
[data-testid="stSidebarCollapsedControl"]::after {
    content: "TRƯỢT ĐỂ MỞ ❯❯" !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    letter-spacing: 1.5px !important;
    margin-left: 10px !important;
    background: linear-gradient(90deg, rgba(255,255,255,0.2) 0%, rgba(255,255,255,0.95) 50%, rgba(255,255,255,0.2) 100%) !important;
    background-size: 200% auto !important;
    color: transparent !important;
    -webkit-background-clip: text !important;
    background-clip: text !important;
    animation: iphoneSlideShimmer 2.5s infinite linear !important;
    white-space: nowrap !important;
}

@keyframes iphoneSlideShimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

/* ---------------------------------------------------- */
/* CÁC THIẾT KẾ SIDEBAR KHI MỞ                           */
/* ---------------------------------------------------- */
[data-testid="stSidebar"] {
    background-color: #12161F !important;
    border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
}

[data-testid="stSidebar"] div.stButton > button {
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 12px 16px !important;
    font-weight: 500 !important;
    font-size: 15px !important;
    border-radius: 10px !important;
    margin-bottom: 6px !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    background: rgba(255, 255, 255, 0.03) !important;
    color: #B0BEC5 !important;
}

[data-testid="stSidebar"] div.stButton > button:hover {
    background: rgba(30, 136, 229, 0.12) !important;
    color: #FFFFFF !important;
    border-color: rgba(30, 136, 229, 0.4) !important;
    transform: translateX(4px) !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
}

[data-testid="stSidebar"] div.stButton > button[kind="primary"],
[data-testid="stSidebar"] div.stButton > button[data-testid="baseButton-primary"] {
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #42A5F5 !important;
    box-shadow: 0 4px 15px rgba(30, 136, 229, 0.4) !important;
    font-weight: 700 !important;
}
</style>
"""
st.markdown(css_style, unsafe_allow_html=True)
