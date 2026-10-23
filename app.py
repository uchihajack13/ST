import os
import io
import re
import lzma
import json
import base64
import calendar
import datetime
from collections import Counter
import streamlit as st
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# KONFIGURASI HALAMAN (MOBILE-FIRST & RESPONSIVE)
# ==========================================
st.set_page_config(
    page_title="Sistem Otomasi ST & Perjadin — KPPBC TMC Malang",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# HELPER WAJIB: CLEAN HTML (100% ANTI BOCOR TAG HTML/SVG DI MARKDOWN)
# ==========================================
def clean_html(raw_html: str) -> str:
    """
    Menghapus spasi indentasi di awal setiap baris pada multi-line HTML/SVG string
    agar parser Markdown Streamlit tidak pernah menganggapnya sebagai blok kode <pre><code>.
    """
    if not raw_html:
        return ""
    return "\n".join(line.lstrip() for line in str(raw_html).splitlines() if line.strip())

# ==========================================
# INISIALISASI TEMA & 4 MENU ALUR KERJA UTAMA (SIDEBAR + QUICK BAR)
# ==========================================
MENU_OPTIONS = [
    "📊 1. Dashboard & Analitik",
    "📝 2. Buat Surat Tugas Baru",
    "⚡ 3. Tindak Lanjut ST & Perjadin",
    "🗂️ 4. Riwayat, Cetak Ulang & Edit"
]
THEME_OPTIONS = [
    "☀️ Light Mode",
    "🌙 Dark Mode"
]
PERIOD_OPTIONS_2026 = [
    "Semua Bulan (2026)",
    "Oktober 2026",
    "September 2026",
    "Agustus 2026",
    "Juli 2026",
    "November 2026",
    "Desember 2026"
]

if "active_main_menu" not in st.session_state or st.session_state["active_main_menu"] not in MENU_OPTIONS:
    st.session_state["active_main_menu"] = MENU_OPTIONS[0]
if "nav_sidebar_radio" not in st.session_state or st.session_state["nav_sidebar_radio"] not in MENU_OPTIONS:
    st.session_state["nav_sidebar_radio"] = st.session_state["active_main_menu"]
if "nav_top_select" not in st.session_state or st.session_state["nav_top_select"] not in MENU_OPTIONS:
    st.session_state["nav_top_select"] = st.session_state["active_main_menu"]

if "app_theme_mode" not in st.session_state:
    st.session_state["app_theme_mode"] = THEME_OPTIONS[0]
if "sidebar_theme_radio" not in st.session_state:
    st.session_state["sidebar_theme_radio"] = st.session_state["app_theme_mode"]

if "active_period_2026" not in st.session_state:
    st.session_state["active_period_2026"] = PERIOD_OPTIONS_2026[0]
if "sb_period_select" not in st.session_state:
    st.session_state["sb_period_select"] = st.session_state["active_period_2026"]
if "fin_period_select" not in st.session_state:
    st.session_state["fin_period_select"] = st.session_state["active_period_2026"]

# State untuk Deep-Link Action Center -> Menu 3 (Tindak Lanjut ST & Perjadin)
if "tl_focus_mode" not in st.session_state:
    st.session_state["tl_focus_mode"] = "A"  # "A" = Nomor ST & Rekap, "B" = Update ID Perjadin
if "nost_work_mode" not in st.session_state:
    st.session_state["nost_work_mode"] = "📱 Mode Form Satuan"
if "perj_work_mode" not in st.session_state:
    st.session_state["perj_work_mode"] = "📱 Mode Form Satuan"
if "jump_target_nost_urut" not in st.session_state:
    st.session_state["jump_target_nost_urut"] = None
if "jump_target_perj_row" not in st.session_state:
    st.session_state["jump_target_perj_row"] = None

def set_main_menu(target_menu: str):
    if target_menu in MENU_OPTIONS:
        st.session_state["active_main_menu"] = target_menu
        st.session_state["nav_sidebar_radio"] = target_menu
        st.session_state["nav_top_select"] = target_menu

def on_sidebar_nav_change():
    pil = st.session_state.get("nav_sidebar_radio", MENU_OPTIONS[0])
    set_main_menu(pil)

def on_top_nav_change():
    pil = st.session_state.get("nav_top_select", MENU_OPTIONS[0])
    set_main_menu(pil)

def on_theme_change():
    st.session_state["app_theme_mode"] = st.session_state.get("sidebar_theme_radio", THEME_OPTIONS[0])

def on_sb_period_change():
    val = st.session_state.get("sb_period_select", PERIOD_OPTIONS_2026[0])
    st.session_state["active_period_2026"] = val
    st.session_state["fin_period_select"] = val

def on_fin_period_change():
    val = st.session_state.get("fin_period_select", PERIOD_OPTIONS_2026[0])
    st.session_state["active_period_2026"] = val
    st.session_state["sb_period_select"] = val

def jump_to_nost_action(no_urut_str: str):
    """Callback 1-Klik dari Action Center Dashboard langsung ke Form Input Nomor ST tertentu."""
    set_main_menu(MENU_OPTIONS[2])
    st.session_state["tl_focus_mode"] = "A"
    st.session_state["nost_work_mode"] = "📱 Mode Form Satuan"
    st.session_state["jump_target_nost_urut"] = str(no_urut_str).strip()

def jump_to_perj_id_action(row_num: int):
    """Callback 1-Klik dari Action Center Dashboard langsung ke Form Update ID Perjadin tanggal tertentu."""
    set_main_menu(MENU_OPTIONS[2])
    st.session_state["tl_focus_mode"] = "B"
    st.session_state["perj_work_mode"] = "📱 Mode Form Satuan"
    st.session_state["jump_target_perj_row"] = int(row_num)

# ==========================================
# ARSITEKTUR UI/UX: DUAL-THEME BENTO GRID + ANTI-DARK-MODE BLEED (100% KONTRAS TAJAM)
# ==========================================
def inject_custom_css(theme_mode: str):
    is_dark = ("Dark" in str(theme_mode))
    if is_dark:
        pal = {
            "scheme": "dark",
            "bg_main": "#0B1120",
            "bg_header": "rgba(11, 17, 32, 0.88)",
            "bg_card": "#1E293B",
            "bg_card_alt": "#0F172A",
            "bg_input": "#0F172A",
            "bg_pill_group": "#0F172A",
            "border_card": "#334155",
            "border_input": "#475569",
            "text_main": "#F8FAFC",
            "text_sub": "#94A3B8",
            "text_muted": "#64748B",
            "accent_primary": "#3B82F6",
            "accent_secondary": "#2563EB",
            "accent_gold": "#FBBF24",
            "shadow_card": "0 10px 28px -6px rgba(0, 0, 0, 0.45)",
            "badge_bg": "linear-gradient(90deg, rgba(59, 130, 246, 0.18) 0%, rgba(30, 41, 59, 0.95) 100%)",
            "badge_border": "#3B82F6",
            "info_bg": "#0F172A",
            "info_border": "#334155",
            "info_text": "#E2E8F0",
            "warn_bg": "rgba(245, 158, 11, 0.14)",
            "warn_border": "#F59E0B",
            "warn_text": "#FDE68A",
            "ok_bg": "rgba(16, 185, 129, 0.14)",
            "ok_border": "#10B981",
            "ok_text": "#A7F3D0",
            "tag_bg": "#1E3A8A",
            "tag_border": "#3B82F6",
            "tag_text": "#F8FAFC",
            "hover_bg": "#1E3A8A",
            "chart_grid": "#334155",
            "cal_cell_bg": "#0F172A",
            "cal_cell_text": "#CBD5E1",
            "btn_sec_bg": "linear-gradient(180deg, #1E293B 0%, #0F172A 100%)",
        }
    else:
        pal = {
            "scheme": "light",
            "bg_main": "#F4F7FB",
            "bg_header": "rgba(244, 247, 251, 0.88)",
            "bg_card": "#FFFFFF",
            "bg_card_alt": "#F8FAFC",
            "bg_input": "#FFFFFF",
            "bg_pill_group": "#E2E8F0",
            "border_card": "#E2E8F0",
            "border_input": "#94A3B8",
            "text_main": "#0F172A",
            "text_sub": "#475569",
            "text_muted": "#64748B",
            "accent_primary": "#1E3A8A",
            "accent_secondary": "#2563EB",
            "accent_gold": "#F59E0B",
            "shadow_card": "0 8px 24px -4px rgba(15, 23, 42, 0.06)",
            "badge_bg": "linear-gradient(90deg, #EFF6FF 0%, #F8FAFC 100%)",
            "badge_border": "#DBEAFE",
            "info_bg": "#F8FAFC",
            "info_border": "#CBD5E1",
            "info_text": "#1E293B",
            "warn_bg": "#FFFBEB",
            "warn_border": "#FCD34D",
            "warn_text": "#92400E",
            "ok_bg": "#ECFDF5",
            "ok_border": "#6EE7B7",
            "ok_text": "#065F46",
            "tag_bg": "#DBEAFE",
            "tag_border": "#60A5FA",
            "tag_text": "#0A192F",
            "hover_bg": "#EFF6FF",
            "chart_grid": "#E2E8F0",
            "cal_cell_bg": "#F1F5F9",
            "cal_cell_text": "#334155",
            "btn_sec_bg": "linear-gradient(180deg, #FFFFFF 0%, #F1F5F9 100%)",
        }

    css_html = f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

        :root, html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
            color-scheme: {pal['scheme']} !important;
            --primary-color: {pal['accent_primary']} !important;
            --background-color: {pal['bg_main']} !important;
            --secondary-background-color: {pal['bg_card']} !important;
            --text-color: {pal['text_main']} !important;
            --bg-card: {pal['bg_card']};
            --bg-card-alt: {pal['bg_card_alt']};
            --text-main: {pal['text_main']};
            --text-sub: {pal['text_sub']};
            --border-card: {pal['border_card']};
            --accent-primary: {pal['accent_primary']};
            --accent-secondary: {pal['accent_secondary']};
            --accent-gold: {pal['accent_gold']};
            --chart-grid: {pal['chart_grid']};
            --cal-cell-bg: {pal['cal_cell_bg']};
            --cal-cell-text: {pal['cal_cell_text']};
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
            background-color: {pal['bg_main']} !important;
            color: {pal['text_main']} !important;
        }}

        header[data-testid="stHeader"] {{
            background: {pal['bg_header']} !important;
            backdrop-filter: blur(8px) !important;
            -webkit-backdrop-filter: blur(8px) !important;
            border-bottom: 1px solid {pal['border_card']} !important;
        }}

        .block-container {{
            padding-top: 2.8rem !important;
            padding-bottom: 3.2rem !important;
            max-width: 1280px !important;
        }}

        label,
        .stApp label,
        div[data-testid="stWidgetLabel"],
        div[data-testid="stWidgetLabel"] p,
        div[data-testid="stWidgetLabel"] span,
        div[data-testid="stMarkdownContainer"] p,
        div[data-testid="stMarkdownContainer"] li,
        div[data-testid="stMarkdownContainer"] span,
        div[data-testid="stMarkdownContainer"] strong,
        div[data-testid="stMarkdownContainer"] b,
        div[data-testid="stMarkdownContainer"] h1,
        div[data-testid="stMarkdownContainer"] h2,
        div[data-testid="stMarkdownContainer"] h3,
        div[data-testid="stMarkdownContainer"] h4,
        div[role="radiogroup"] label,
        div[role="radiogroup"] label p,
        div[role="radiogroup"] label span,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stCheckbox"] label p,
        div[data-testid="stCheckbox"] label span,
        div[data-testid="stExpander"] summary p,
        div[data-testid="stExpander"] summary span {{
            color: {pal['text_main']} !important;
        }}

        div[data-testid="stWidgetLabel"] p {{
            font-weight: 700 !important;
            font-size: 0.86rem !important;
            color: {pal['text_main']} !important;
            margin-bottom: 2px !important;
        }}

        div[data-testid="stCaptionContainer"] p,
        div[data-testid="stCaptionContainer"] span {{
            color: {pal['text_sub']} !important;
            font-size: 0.8rem !important;
        }}

        /* Sidebar Executive Dark Navy */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #0B1120 0%, #0F172A 100%) !important;
            border-right: 1px solid #1E293B !important;
        }}
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] div[data-testid="stWidgetLabel"] p,
        section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
        section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] span,
        section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] strong,
        section[data-testid="stSidebar"] div[data-testid="stCaptionContainer"] p,
        section[data-testid="stSidebar"] div[data-testid="stCaptionContainer"] span {{
            color: #F8FAFC !important;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] {{
            background-color: #1E293B !important;
            border: 1px solid #334155 !important;
            padding: 6px !important;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label {{
            background-color: #0F172A !important;
            border: 1px solid #334155 !important;
            width: 100% !important;
            padding: 9px 12px !important;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label p,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label span {{
            color: #CBD5E1 !important;
            -webkit-text-fill-color: #CBD5E1 !important;
            font-weight: 700 !important;
            font-size: 0.86rem !important;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {{
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
            border-color: #60A5FA !important;
            border-left: 4px solid #FBBF24 !important;
            box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
        }}
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p,
        section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) span {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 800 !important;
        }}
        section[data-testid="stSidebar"] div.stButton > button {{
            background: #1E293B !important;
            border: 1px solid #334155 !important;
            color: #F8FAFC !important;
            min-height: 42px !important;
        }}
        section[data-testid="stSidebar"] div.stButton > button p,
        section[data-testid="stSidebar"] div.stButton > button span {{
            color: #F8FAFC !important;
            -webkit-text-fill-color: #F8FAFC !important;
        }}
        section[data-testid="stSidebar"] div.stButton > button:hover {{
            background: #2563EB !important;
            border-color: #60A5FA !important;
        }}

        /* Bento Grid Container */
        div[data-testid="stVerticalBlockBorderWrapper"],
        div[data-testid="stLayoutWrapper"] > div[style*="border"] {{
            background-color: {pal['bg_card']} !important;
            border-radius: 18px !important;
            border: 1.5px solid {pal['border_card']} !important;
            box-shadow: {pal['shadow_card']} !important;
            padding: 16px 18px !important;
            margin-bottom: 0.8rem !important;
        }}

        /* Ultra-Clean Top Header Bar (1 Baris Ramping & Berwibawa, Bebas Duplikasi) */
        .bc-ultraclean-header {{
            background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 65%, #1D4ED8 100%) !important;
            color: #FFFFFF !important;
            padding: 0.85rem 1.3rem !important;
            border-radius: 16px !important;
            border-left: 5px solid #F59E0B !important;
            box-shadow: 0 8px 22px -6px rgba(10, 25, 47, 0.25) !important;
            margin-bottom: 0.55rem !important;
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            flex-wrap: wrap !important;
            gap: 10px !important;
        }}
        .bc-ultraclean-title {{
            font-size: 1.12rem !important;
            font-weight: 800 !important;
            margin: 0 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            line-height: 1.25 !important;
        }}
        .bc-ultraclean-sub {{
            font-size: 0.77rem !important;
            color: #DBEAFE !important;
            -webkit-text-fill-color: #DBEAFE !important;
            margin: 2px 0 0 0 !important;
            font-weight: 500 !important;
        }}
        .bc-online-badge {{
            display: inline-flex !important;
            align-items: center !important;
            gap: 7px !important;
            background: rgba(16, 185, 129, 0.18) !important;
            border: 1px solid rgba(52, 211, 153, 0.55) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            padding: 6px 13px !important;
            border-radius: 999px !important;
            font-size: 0.78rem !important;
            font-weight: 700 !important;
            white-space: nowrap !important;
        }}
        .bc-online-dot {{
            width: 8px !important;
            height: 8px !important;
            border-radius: 50% !important;
            background: #10B981 !important;
            box-shadow: 0 0 8px #10B981 !important;
            display: inline-block !important;
        }}

        /* Section Badges & Info Pills */
        .section-badge {{
            display: inline-flex !important;
            align-items: center !important;
            gap: 8px !important;
            background: {pal['badge_bg']} !important;
            color: {pal['text_main']} !important;
            -webkit-text-fill-color: {pal['text_main']} !important;
            padding: 6px 13px !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
            font-size: 0.9rem !important;
            border: 1px solid {pal['badge_border']} !important;
            border-left: 4px solid {pal['accent_primary']} !important;
            margin-bottom: 0.65rem !important;
        }}
        .info-pill {{
            background-color: {pal['info_bg']} !important;
            border: 1px solid {pal['info_border']} !important;
            border-left: 4px solid {pal['accent_primary']} !important;
            padding: 9px 13px !important;
            border-radius: 10px !important;
            font-size: 0.83rem !important;
            color: {pal['info_text']} !important;
            margin-top: 4px !important;
            margin-bottom: 10px !important;
            line-height: 1.5 !important;
        }}
        .info-pill-warn {{
            background-color: {pal['warn_bg']} !important;
            border: 1px solid {pal['warn_border']} !important;
            border-left: 4px solid #F59E0B !important;
            padding: 9px 13px !important;
            border-radius: 10px !important;
            font-size: 0.83rem !important;
            color: {pal['warn_text']} !important;
            margin-top: 4px !important;
            margin-bottom: 10px !important;
            line-height: 1.5 !important;
        }}
        .info-pill-ok {{
            background-color: {pal['ok_bg']} !important;
            border: 1px solid {pal['ok_border']} !important;
            border-left: 4px solid #10B981 !important;
            padding: 9px 13px !important;
            border-radius: 10px !important;
            font-size: 0.83rem !important;
            color: {pal['ok_text']} !important;
            margin-top: 4px !important;
            margin-bottom: 10px !important;
            line-height: 1.5 !important;
        }}

        /* Petugas Bento Mini-Cards */
        .petugas-bento-grid {{
            display: grid !important;
            grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)) !important;
            gap: 10px !important;
            margin-top: 8px !important;
            margin-bottom: 6px !important;
        }}
        .petugas-bento-card {{
            background: {pal['bg_card_alt']} !important;
            border: 1px solid {pal['border_card']} !important;
            border-left: 4px solid #F59E0B !important;
            border-radius: 12px !important;
            padding: 10px 13px !important;
            font-size: 0.83rem !important;
            line-height: 1.45 !important;
        }}
        .petugas-bento-card .p-title {{
            font-weight: 800 !important;
            color: {pal['text_main']} !important;
            font-size: 0.88rem !important;
            margin-bottom: 3px !important;
        }}
        .petugas-bento-card .p-sub {{
            color: {pal['text_sub']} !important;
            font-size: 0.79rem !important;
        }}

        /* Input, Selectbox, DateInput, Textarea & Dropdown */
        input,
        textarea,
        div[data-baseweb="input"],
        div[data-baseweb="base-input"],
        div[data-baseweb="textarea"],
        div[data-baseweb="select"] {{
            font-size: 16px !important;
            border-radius: 12px !important;
        }}

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="base-input"],
        div[data-baseweb="textarea"] {{
            min-height: 44px !important;
            background-color: {pal['bg_input']} !important;
            color: {pal['text_main']} !important;
            border: 1.5px solid {pal['border_input']} !important;
            border-radius: 12px !important;
        }}

        div[data-baseweb="select"] > div:focus-within,
        div[data-baseweb="input"] > div:focus-within,
        div[data-baseweb="textarea"]:focus-within {{
            background-color: {pal['bg_input']} !important;
            border-color: {pal['accent_primary']} !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.2) !important;
        }}

        input,
        textarea,
        div[data-baseweb="select"] span,
        div[data-baseweb="select"] div {{
            color: {pal['text_main']} !important;
            -webkit-text-fill-color: {pal['text_main']} !important;
            background-color: transparent !important;
        }}

        section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
            background-color: #1E293B !important;
            border-color: #334155 !important;
        }}
        section[data-testid="stSidebar"] div[data-baseweb="select"] span,
        section[data-testid="stSidebar"] div[data-baseweb="select"] div {{
            color: #F8FAFC !important;
            -webkit-text-fill-color: #F8FAFC !important;
        }}

        div[data-baseweb="select"] svg {{
            fill: {pal['text_main']} !important;
            color: {pal['text_main']} !important;
        }}

        input::placeholder,
        textarea::placeholder {{
            color: {pal['text_muted']} !important;
            -webkit-text-fill-color: {pal['text_muted']} !important;
            opacity: 1 !important;
        }}

        span[data-baseweb="tag"] {{
            background-color: {pal['tag_bg']} !important;
            border: 1px solid {pal['tag_border']} !important;
            border-radius: 8px !important;
            padding: 3px 8px !important;
        }}
        span[data-baseweb="tag"] span {{
            color: {pal['tag_text']} !important;
            -webkit-text-fill-color: {pal['tag_text']} !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
        }}

        div[data-baseweb="popover"],
        ul[data-baseweb="menu"],
        div[role="listbox"] {{
            background-color: {pal['bg_card']} !important;
            color: {pal['text_main']} !important;
            border: 1.5px solid {pal['border_card']} !important;
            border-radius: 12px !important;
            box-shadow: 0 12px 28px rgba(0, 0, 0, 0.22) !important;
        }}
        li[role="option"],
        div[role="option"] {{
            background-color: {pal['bg_card']} !important;
            color: {pal['text_main']} !important;
            -webkit-text-fill-color: {pal['text_main']} !important;
            font-size: 0.91rem !important;
            padding: 10px 14px !important;
        }}
        li[role="option"]:hover,
        li[role="option"][aria-selected="true"],
        div[role="option"]:hover,
        div[role="option"][aria-selected="true"] {{
            background-color: {pal['hover_bg']} !important;
            color: {pal['accent_primary']} !important;
            -webkit-text-fill-color: {pal['accent_primary']} !important;
            font-weight: 700 !important;
        }}

        /* Segmented Pill Control untuk Tab & Radio */
        div[data-baseweb="tab-list"] {{
            display: flex !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            white-space: nowrap !important;
            scrollbar-width: none !important;
            gap: 8px !important;
            background-color: {pal['bg_pill_group']} !important;
            padding: 6px !important;
            border-radius: 14px !important;
            border: 1px solid {pal['border_card']} !important;
            margin-bottom: 1rem !important;
        }}
        div[data-baseweb="tab-list"]::-webkit-scrollbar {{
            display: none !important;
        }}
        button[data-baseweb="tab"] {{
            font-weight: 700 !important;
            font-size: 0.87rem !important;
            padding: 9px 15px !important;
            border-radius: 10px !important;
            border: 1px solid {pal['border_card']} !important;
            background-color: {pal['bg_card']} !important;
            color: {pal['text_main']} !important;
            transition: all 0.18s ease !important;
            min-height: 42px !important;
            flex-shrink: 0 !important;
        }}
        button[data-baseweb="tab"] p,
        button[data-baseweb="tab"] span,
        button[data-baseweb="tab"] div {{
            color: {pal['text_main']} !important;
            -webkit-text-fill-color: {pal['text_main']} !important;
            font-weight: 700 !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%) !important;
            color: #FFFFFF !important;
            border: 1px solid #1E3A8A !important;
            border-bottom: 3px solid #F59E0B !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.28) !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] p,
        button[data-baseweb="tab"][aria-selected="true"] span,
        button[data-baseweb="tab"][aria-selected="true"] div {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 800 !important;
        }}
        div[data-baseweb="tab-highlight"],
        div[data-baseweb="tab-border"] {{
            display: none !important;
        }}

        /* Tombol Aksi & Radio Pill */
        div.stButton > button,
        div.stDownloadButton > button {{
            min-height: 44px !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            font-size: 0.86rem !important;
            background: {pal['btn_sec_bg']} !important;
            color: {pal['text_main']} !important;
            border: 1.5px solid {pal['border_input']} !important;
            transition: all 0.18s ease-in-out !important;
        }}
        div.stButton > button p,
        div.stButton > button span,
        div.stButton > button div {{
            color: {pal['text_main']} !important;
            -webkit-text-fill-color: {pal['text_main']} !important;
            font-weight: 700 !important;
        }}
        div.stButton > button:hover {{
            border-color: {pal['accent_primary']} !important;
        }}

        div.stButton > button[kind="primary"],
        div[data-testid="stFormSubmitButton"] > button,
        div[data-testid="stFormSubmitButton"] > button[kind="primary"] {{
            min-height: 52px !important;
            background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%) !important;
            color: #FFFFFF !important;
            border: 1px solid #1D4ED8 !important;
            border-bottom: 3.5px solid #F59E0B !important;
            box-shadow: 0 8px 20px -4px rgba(37, 99, 235, 0.38) !important;
            font-size: 0.93rem !important;
        }}
        div.stButton > button[kind="primary"] p,
        div.stButton > button[kind="primary"] span,
        div.stButton > button[kind="primary"] div,
        div[data-testid="stFormSubmitButton"] > button p,
        div[data-testid="stFormSubmitButton"] > button span,
        div[data-testid="stFormSubmitButton"] > button div {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 800 !important;
        }}

        div[role="radiogroup"] {{
            display: flex !important;
            flex-wrap: wrap !important;
            gap: 8px !important;
            background-color: {pal['bg_pill_group']} !important;
            padding: 6px !important;
            border-radius: 14px !important;
            border: 1.5px solid {pal['border_card']} !important;
            margin-top: 2px !important;
            margin-bottom: 6px !important;
        }}
        div[role="radiogroup"] > label {{
            background-color: {pal['bg_card']} !important;
            border: 1.5px solid {pal['border_card']} !important;
            border-radius: 10px !important;
            padding: 7px 13px !important;
            min-height: 40px !important;
            display: inline-flex !important;
            align-items: center !important;
            gap: 6px !important;
            cursor: pointer !important;
            transition: all 0.16s ease !important;
            margin: 0 !important;
            flex: 1 1 auto !important;
        }}
        div[role="radiogroup"] > label:has(input:checked) {{
            background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%) !important;
            border-color: #1D4ED8 !important;
            border-bottom: 3px solid #F59E0B !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
        }}
        div[role="radiogroup"] > label:has(input:checked) p,
        div[role="radiogroup"] > label:has(input:checked) span {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 800 !important;
        }}

        div[data-testid="stCheckbox"] {{
            background-color: {pal['bg_card_alt']} !important;
            border: 1.5px solid {pal['border_card']} !important;
            border-radius: 12px !important;
            padding: 8px 13px !important;
            min-height: 44px !important;
            display: flex !important;
            align-items: center !important;
            margin-bottom: 6px !important;
        }}
        div[data-testid="stCheckbox"]:has(input:checked) {{
            border-color: {pal['accent_primary']} !important;
            border-left: 4px solid #F59E0B !important;
        }}

        /* Dashboard KPI Cards */
        .dash-kpi-card {{
            background: var(--bg-card);
            border: 1.5px solid var(--border-card);
            border-radius: 18px;
            padding: 15px 17px;
            box-shadow: {pal['shadow_card']};
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            height: 100%;
            min-height: 104px;
            transition: transform 0.16s ease;
        }}
        .dash-kpi-card:hover {{
            transform: translateY(-2px);
        }}
        .dash-kpi-card.accent-navy {{
            background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%) !important;
            border-color: #1E3A8A !important;
            border-left: 5px solid #F59E0B !important;
        }}
        .dash-kpi-title {{
            font-size: 0.74rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-sub);
            margin-bottom: 4px;
        }}
        .dash-kpi-card.accent-navy .dash-kpi-title {{
            color: #DBEAFE !important;
        }}
        .dash-kpi-value {{
            font-size: 1.5rem;
            font-weight: 800;
            color: var(--text-main);
            line-height: 1.15;
        }}
        .dash-kpi-card.accent-navy .dash-kpi-value {{
            color: #FBBF24 !important;
        }}
        .dash-kpi-sub {{
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--text-sub);
            margin-top: 4px;
        }}
        .dash-kpi-card.accent-navy .dash-kpi-sub {{
            color: #E2E8F0 !important;
        }}
        .dash-kpi-icon {{
            width: 44px;
            height: 44px;
            border-radius: 13px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.3rem;
            flex-shrink: 0;
        }}

        /* Kartu Biru Gradien Finansial */
        .fin-gradient-card {{
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 55%, #1E3A8A 100%) !important;
            border-radius: 18px !important;
            padding: 18px 18px 16px 18px !important;
            color: #FFFFFF !important;
            box-shadow: 0 14px 30px -6px rgba(37, 99, 235, 0.42) !important;
            border: 1px solid rgba(255, 255, 255, 0.18) !important;
            margin-bottom: 10px !important;
            position: relative;
            overflow: hidden;
        }}
        .fin-gradient-card * {{
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }}
        .fin-bars-wrap {{
            display: flex;
            align-items: flex-end;
            gap: 5px;
            height: 48px;
            margin: 12px 0 10px 0;
            padding-bottom: 4px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.2);
        }}
        .fin-bar-col {{
            flex: 1;
            background: rgba(255, 255, 255, 0.28);
            border-radius: 999px;
            transition: all 0.2s ease;
        }}
        .fin-bar-col.active {{
            background: #FBBF24 !important;
            box-shadow: 0 0 8px rgba(251, 191, 36, 0.6);
        }}

        /* Active Users / Top Petugas Progress List */
        .user-row-item {{
            margin-bottom: 12px;
            padding-bottom: 9px;
            border-bottom: 1px solid var(--border-card);
        }}
        .user-row-item:last-child {{
            margin-bottom: 0;
            padding-bottom: 0;
            border-bottom: none;
        }}
        .user-avatar-circle {{
            width: 34px;
            height: 34px;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-weight: 800;
            font-size: 0.8rem;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            flex-shrink: 0;
        }}
        .prog-track {{
            width: 100%;
            height: 7px;
            background: var(--cal-cell-bg);
            border-radius: 999px;
            overflow: hidden;
            margin: 6px 0 4px 0;
        }}
        .prog-fill {{
            height: 100%;
            border-radius: 999px;
        }}

        /* Kalender Visual Bulanan */
        .cal-grid-header {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 5px;
            text-align: center;
            font-size: 0.72rem;
            font-weight: 800;
            color: var(--text-sub);
            margin-bottom: 6px;
        }}
        .cal-grid-days {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 5px;
        }}
        .cal-day-box {{
            aspect-ratio: 1;
            min-height: 30px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.77rem;
            font-weight: 700;
            background: var(--cal-cell-bg);
            color: var(--cal-cell-text);
            position: relative;
        }}
        .cal-day-box.empty {{
            background: transparent;
        }}
        .cal-day-box.luar-kota {{
            background: #F59E0B !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            box-shadow: 0 3px 8px rgba(245, 158, 11, 0.35);
        }}
        .cal-day-box.dalam-kota {{
            background: #10B981 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            box-shadow: 0 3px 8px rgba(16, 185, 129, 0.35);
        }}
        .cal-day-box.belum-id {{
            background: #8B5CF6 !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            box-shadow: 0 3px 8px rgba(139, 92, 246, 0.4);
            border: 1.5px solid #FBBF24 !important;
        }}

        /* Action Center */
        .topic-item {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 10px;
            border-radius: 12px;
            background: var(--bg-card-alt);
            border: 1px solid var(--border-card);
            margin-bottom: 4px;
        }}
        .topic-icon-box {{
            width: 34px;
            height: 34px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.98rem;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            flex-shrink: 0;
        }}

        /* Sidebar Queue Summary Card */
        .sb-queue-box {{
            background: rgba(30, 41, 59, 0.85);
            border: 1px solid #334155;
            border-left: 4px solid #FBBF24;
            border-radius: 12px;
            padding: 10px 12px;
            margin-top: 8px;
            margin-bottom: 8px;
            font-size: 0.78rem;
            color: #F8FAFC;
        }}

        @media (max-width: 640px) {{
            .block-container {{
                padding-left: 0.65rem !important;
                padding-right: 0.65rem !important;
                padding-top: 2.5rem !important;
            }}
            .bc-ultraclean-header {{
                padding: 0.75rem 1rem !important;
                border-radius: 14px !important;
            }}
            .bc-ultraclean-title {{
                font-size: 0.98rem !important;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"] {{
                padding: 13px !important;
                border-radius: 15px !important;
            }}
        }}
    </style>
    """
    st.markdown(clean_html(css_html), unsafe_allow_html=True)

inject_custom_css(st.session_state["app_theme_mode"])

# Kamus Nama Hari & Bulan Bahasa Indonesia
HARI_INDONESIA = {
    "Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu",
    "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"
}

BULAN_INDONESIA = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

BULAN_MAP_REV = {v.lower(): k for k, v in BULAN_INDONESIA.items()}
BULAN_MAP_REV.update({
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "jun": 6,
    "jul": 7, "agu": 8, "ags": 8, "sep": 9, "sept": 9, "okt": 10, "nov": 11, "des": 12,
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "aug": 8, "september": 9, "october": 10, "oct": 10,
    "november": 11, "december": 12, "dec": 12
})

PRESET_PERMOHONAN_OPTIONS = [
    "✍️ Ketik Manual / Kosongkan",
    "Permohonan Pelayanan Pelekatan Tanda Pengaman",
    "Permohonan Pelayanan Pembukaan Tanda Pengaman",
    "Permohonan Pemeriksaan Fisik Barang Kena Cukai Hasil Tembakau",
    "Permohonan Pelayanan Pengawasan Pengeluaran Barang Kena Cukai"
]

def format_hari_indo(d: datetime.date) -> str:
    if not d:
        return ""
    en_day = d.strftime("%A")
    return HARI_INDONESIA.get(en_day, en_day)

def format_tanggal_indo(d: datetime.date) -> str:
    if not d:
        return ""
    return f"{d.day} {BULAN_INDONESIA[d.month]} {d.year}"

def format_hari_tanggal_pelaksanaan(d_mulai: datetime.date, d_akhir: datetime.date = None):
    """
    Menghasilkan tuple (hari_str, tanggal_str, durasi_hari):
    - Jika 1 hari (d_akhir is None atau d_akhir <= d_mulai):
      -> ('Rabu', '21 Oktober 2026', 1)
    - Jika > 1 hari (d_akhir > d_mulai):
      -> ('Rabu s.d. Jumat', '21 Oktober 2026 s.d. 23 Oktober 2026', 3)
    """
    if not d_mulai:
        return "", "", 0
    if not d_akhir or d_akhir <= d_mulai:
        return format_hari_indo(d_mulai), format_tanggal_indo(d_mulai), 1
    hari_mulai = format_hari_indo(d_mulai)
    hari_akhir = format_hari_indo(d_akhir)
    tgl_mulai_str = format_tanggal_indo(d_mulai)
    tgl_akhir_str = format_tanggal_indo(d_akhir)
    durasi = (d_akhir - d_mulai).days + 1
    return f"{hari_mulai} s.d. {hari_akhir}", f"{tgl_mulai_str} s.d. {tgl_akhir_str}", durasi

def parse_indo_date(val) -> datetime.date:
    """
    Mengubah berbagai format string tanggal Indonesia / Inggris / ISO / Custom GSheet menjadi datetime.date.
    Mendukung penuh format Kolom A DATA PERJADIN ('Selasa, Oktober 20, 2026' / 'Rabu, Juli 01, 2026')
    maupun format Sheet1 & REKAP BAGIAN UMUM ('20 Oktober 2026').
    """
    if not val:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)) and 30000 <= val <= 80000:
        try:
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
        except Exception:
            pass
    s = str(val).strip()
    if not s:
        return None
    # 1. Coba format rentang ringkas "21 s.d. 23 Oktober 2026" (ambil tanggal awal)
    m_short = re.search(r'(\d{1,2})\s*(?:s\.?d\.?|s/d|-|–)\s*(\d{1,2})\s+([A-Za-z]+)[,\s]+(\d{4})', s, flags=re.IGNORECASE)
    if m_short:
        day1 = int(m_short.group(1))
        mon_str = m_short.group(3).lower()
        yr = int(m_short.group(4))
        if mon_str in BULAN_MAP_REV:
            try:
                return datetime.date(yr, BULAN_MAP_REV[mon_str], day1)
            except Exception:
                pass
    # 2. Coba format "15 Oktober 2026" atau "Selasa, 15 Oktober 2026" (DD Bulan YYYY)
    m = re.search(r'(\d{1,2})\s+([A-Za-z]+)[,\s]+(\d{4})', s)
    if m:
        day = int(m.group(1))
        mon_str = m.group(2).lower()
        yr = int(m.group(3))
        if mon_str in BULAN_MAP_REV:
            try:
                return datetime.date(yr, BULAN_MAP_REV[mon_str], day)
            except Exception:
                pass
    # 3. Coba format Kolom A DATA PERJADIN: "Rabu, Juli 01, 2026" / "Selasa, Oktober 20, 2026" (Bulan DD, YYYY)
    m_mdy = re.search(r'([A-Za-z]+)\s+(\d{1,2})[,\s]+(\d{4})', s)
    if m_mdy:
        mon_str = m_mdy.group(1).lower()
        day = int(m_mdy.group(2))
        yr = int(m_mdy.group(3))
        if mon_str in BULAN_MAP_REV:
            try:
                return datetime.date(yr, BULAN_MAP_REV[mon_str], day)
            except Exception:
                pass
    # 4. Coba format ISO "2026-10-15"
    m2 = re.match(r'^(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s)
    if m2:
        try:
            return datetime.date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except Exception:
            pass
    # 5. Coba format DD/MM/YYYY atau MM/DD/YYYY
    m3 = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m3:
        p1, p2, yr = int(m3.group(1)), int(m3.group(2)), int(m3.group(3))
        try:
            if p2 <= 12:
                return datetime.date(yr, p2, p1)
            return datetime.date(yr, p1, p2)
        except Exception:
            pass
    return None

def parse_indo_date_range(val) -> list:
    """
    Mengekstrak seluruh daftar datetime.date harian dari string tanggal tunggal maupun rentang tanggal.
    Contoh:
    - '21 Oktober 2026' -> [date(2026, 10, 21)]
    - 'Selasa, Oktober 20, 2026' -> [date(2026, 10, 20)]
    - '21 Oktober 2026 s.d. 23 Oktober 2026' -> [date(2026, 10, 21), date(2026, 10, 22), date(2026, 10, 23)]
    - '21 s.d. 23 Oktober 2026' -> [date(2026, 10, 21), date(2026, 10, 22), date(2026, 10, 23)]
    """
    if not val:
        return []
    if isinstance(val, (datetime.date, datetime.datetime)):
        d_single = parse_indo_date(val)
        return [d_single] if d_single else []
    s = str(val).strip()
    if not s:
        return []

    # 1. Cek jika ada >= 2 tanggal lengkap DD Bulan YYYY (misal: "21 Oktober 2026 s.d. 23 Oktober 2026")
    all_full = re.findall(r'(\d{1,2})\s+([A-Za-z]+)[,\s]+(\d{4})', s)
    if len(all_full) >= 2:
        try:
            d1_day, d1_mon, d1_yr = int(all_full[0][0]), all_full[0][1].lower(), int(all_full[0][2])
            d2_day, d2_mon, d2_yr = int(all_full[-1][0]), all_full[-1][1].lower(), int(all_full[-1][2])
            if d1_mon in BULAN_MAP_REV and d2_mon in BULAN_MAP_REV:
                dt1 = datetime.date(d1_yr, BULAN_MAP_REV[d1_mon], d1_day)
                dt2 = datetime.date(d2_yr, BULAN_MAP_REV[d2_mon], d2_day)
                if dt2 >= dt1 and (dt2 - dt1).days <= 60:
                    return [dt1 + datetime.timedelta(days=i) for i in range((dt2 - dt1).days + 1)]
        except Exception:
            pass

    # 2. Cek jika format ringkas dalam bulan yang sama (misal: "21 s.d. 23 Oktober 2026" atau "21 - 23 Oktober 2026")
    m_short = re.search(r'(\d{1,2})\s*(?:s\.?d\.?|s/d|-|–)\s*(\d{1,2})\s+([A-Za-z]+)[,\s]+(\d{4})', s, flags=re.IGNORECASE)
    if m_short:
        try:
            day1 = int(m_short.group(1))
            day2 = int(m_short.group(2))
            mon_str = m_short.group(3).lower()
            yr = int(m_short.group(4))
            if mon_str in BULAN_MAP_REV:
                dt1 = datetime.date(yr, BULAN_MAP_REV[mon_str], day1)
                dt2 = datetime.date(yr, BULAN_MAP_REV[mon_str], day2)
                if dt2 >= dt1 and (dt2 - dt1).days <= 60:
                    return [dt1 + datetime.timedelta(days=i) for i in range((dt2 - dt1).days + 1)]
        except Exception:
            pass

    # 3. Fallback tanggal tunggal (termasuk format "Selasa, Oktober 20, 2026")
    d_one = parse_indo_date(s)
    return [d_one] if d_one else []

def merge_nomor_st_cell(cur_no_st: str, new_no_st: str) -> str:
    """
    Menggabungkan atau memperbarui Nomor ST pada sel Kolom I (NO ST) di DATA PERJADIN:
    - Jika nomor urut ST sama (misal mengganti format ST-2041/KBC.1201/KBC.120101/2026 menjadi ST-2041/KBC.1201/2026),
      otomatis mengganti (replace) nomor tersebut tanpa duplikasi.
    - Jika di hari yang sama ada ST lain dengan nomor berbeda, menggabungkannya dengan koma.
    """
    new_clean = (new_no_st or "").strip()
    cur_clean = (cur_no_st or "").strip()
    if not new_clean:
        return cur_clean
    if not cur_clean:
        return new_clean

    m_new_num = re.match(r'^ST[-\s]*(\d+)/', new_clean, flags=re.IGNORECASE)
    new_num = m_new_num.group(1) if m_new_num else None

    parts = [p.strip() for p in cur_clean.split(",") if p.strip()]
    updated_parts = []
    replaced = False
    for p in parts:
        if p.upper() == new_clean.upper():
            updated_parts.append(new_clean)
            replaced = True
            continue
        m_p_num = re.match(r'^ST[-\s]*(\d+)/', p, flags=re.IGNORECASE)
        if new_num and m_p_num and m_p_num.group(1) == new_num:
            updated_parts.append(new_clean)
            replaced = True
        else:
            updated_parts.append(p)
    if not replaced:
        updated_parts.append(new_clean)
    return ", ".join(updated_parts)

def strip_permohonan(text: str) -> str:
    if not text:
        return ""
    clean = re.sub(r'^(permohonan\s+|permohonan:\s*)', '', text.strip(), flags=re.IGNORECASE).strip()
    return clean

def detect_tipe_ttd_from_nost(nomor_st: str) -> str:
    """
    Mendeteksi jenis penandatangan dari format Nomor ST yang sudah tersimpan:
    - Jika mengandung '/KBC.1201/20' (tanpa KBC.120101) -> 'Kepala Kantor Langsung'
    - Selain itu default -> 'a.n. Kepala Kantor (Kepala Subbagian Umum)'
    """
    s = (nomor_st or "").strip().upper()
    if re.search(r'/KBC\.1201/\d{4}$', s) and "KBC.120101" not in s:
        return "Kepala Kantor Langsung"
    return "a.n. Kepala Kantor (Kepala Subbagian Umum)"

def format_nomor_st_resmi(
    raw_st: str,
    tahun: int = 2026,
    tipe_ttd: str = "a.n. Kepala Kantor (Kepala Subbagian Umum)",
    auto_adjust_unit: bool = False
) -> str:
    """
    Memformat Nomor ST secara dinamis sesuai pejabat penandatangan:
    - Jika ditandatangani a.n. Kepala Kantor: ST-xx/KBC.1201/KBC.120101/2026
    - Jika ditandatangani Kepala Kantor langsung: ST-xx/KBC.1201/2026
    """
    if not raw_st or not str(raw_st).strip():
        return ""
    s = str(raw_st).strip()
    is_kk_langsung = ("kepala kantor langsung" in str(tipe_ttd).lower())
    unit_code = "KBC.1201" if is_kk_langsung else "KBC.1201/KBC.120101"

    # Jika user mengetik angka saja (misal '2013' atau 'ST-2013')
    m = re.match(r'^(?:ST[-\s]*)?(\d+)$', s, flags=re.IGNORECASE)
    if m:
        angka = m.group(1)
        return f"ST-{angka}/{unit_code}/{tahun}"

    # Jika sudah berformat ST-xx/KBC.1201.../tahun dan user mengganti opsi penandatangan (auto_adjust_unit=True)
    if auto_adjust_unit:
        m_full = re.match(r'^ST[-\s]*(\d+)/KBC\.1201(?:/KBC\.120101)?/(\d{4})$', s, flags=re.IGNORECASE)
        if m_full:
            return f"ST-{m_full.group(1)}/{unit_code}/{m_full.group(2)}"

    return s

def extract_lokasi_kecamatan(alamat: str, pt_name: str = "", perusahaan_dict: dict = None):
    """
    Mengekstrak LOKASI ('Kabupaten Malang' / 'Kota Malang' / 'Kota Batu')
    dan Kecamatan ('Singosari', 'Bululawang', 'Pakisaji', dll.) secara otomatis dari alamat pabrik.
    """
    alm = (alamat or "").strip()
    if not alm and pt_name and perusahaan_dict and pt_name in perusahaan_dict:
        for idx in range(1, 13):
            cand = perusahaan_dict[pt_name].get(idx, "").strip()
            if cand:
                alm = cand
                break

    if not alm:
        return "Kabupaten Malang", "-"

    # Tentukan LOKASI (Kabupaten / Kota)
    if re.search(r'Kota\s+Malang', alm, re.IGNORECASE):
        lokasi = "Kota Malang"
    elif re.search(r'Kota\s+Batu', alm, re.IGNORECASE):
        lokasi = "Kota Batu"
    elif re.search(r'Kab(?:upaten|\.?)\s+Malang', alm, re.IGNORECASE):
        lokasi = "Kabupaten Malang"
    else:
        lokasi = "Kabupaten Malang"

    # Ekstrak Kecamatan
    m_kec = re.search(
        r'(?:Kecamatan|Kec\.?)\s+([A-Za-z\s]+?)(?=\s*(?:,|\s+Kabupaten|\s+Kab\.?|\s+Kota|$))',
        alm,
        re.IGNORECASE
    )
    if m_kec:
        kec = m_kec.group(1).strip().title()
        if kec.lower() == "kendungkandang":
            kec = "Kedungkandang"
    else:
        kec = "-"

    return lokasi, kec

def render_direct_docx_download(docx_bytes: bytes, filename: str, button_label: str):
    """
    Merender tombol Download Word (.docx) menggunakan Base64 Data URI Murni + Blob JS.
    100% kebal terhadap error Chrome 'File wasn't available on site' di balik iframe Streamlit Cloud.
    """
    b64_data = base64.b64encode(docx_bytes).decode()
    safe_fn = filename.replace('"', '').replace("'", "")
    html_btn = f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{
            margin: 0;
            padding: 0;
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: transparent;
        }}
        .dl-btn {{
            display: flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            min-height: 48px;
            box-sizing: border-box;
            background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%);
            color: #ffffff !important;
            text-decoration: none;
            font-weight: 700;
            font-size: 14.5px;
            border-radius: 12px;
            border: 1px solid #0A192F;
            border-left: 5px solid #F59E0B;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(10, 25, 47, 0.2);
            transition: all 0.15s ease;
            padding: 8px 14px;
            text-align: center;
        }}
        .dl-btn:hover {{
            background: linear-gradient(135deg, #112240 0%, #2563EB 100%);
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3);
        }}
    </style>
    <script>
    function triggerDocxDownload(e) {{
        e.preventDefault();
        try {{
            var b64 = "{b64_data}";
            var binStr = atob(b64);
            var len = binStr.length;
            var arr = new Uint8Array(len);
            for (var i = 0; i < len; i++) {{
                arr[i] = binStr.charCodeAt(i);
            }}
            var blob = new Blob([arr], {{type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}});
            var url = URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = "{safe_fn}";
            document.body.appendChild(a);
            a.click();
            setTimeout(function() {{
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            }}, 300);
        }} catch (err) {{
            var aFallback = document.getElementById("direct-a");
            aFallback.click();
        }}
    }}
    </script>
    </head>
    <body>
        <a id="direct-a" class="dl-btn"
           href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64_data}"
           download="{safe_fn}"
           onclick="triggerDocxDownload(event)">
           {button_label}
        </a>
    </body>
    </html>
    """
    components.html(html_btn, height=56, scrolling=False)

# ==========================================
# INISIALISASI SESSION STATE & CALLBACKS
# ==========================================
if "in_kegiatan_1" not in st.session_state:
    st.session_state["in_kegiatan_1"] = ""
if "prev_permohonan_1" not in st.session_state:
    st.session_state["prev_permohonan_1"] = ""

if "in_kegiatan_2" not in st.session_state:
    st.session_state["in_kegiatan_2"] = ""
if "prev_permohonan_2" not in st.session_state:
    st.session_state["prev_permohonan_2"] = ""

def on_preset_permohonan_1_change():
    pil = st.session_state.get("preset_permohonan_1", PRESET_PERMOHONAN_OPTIONS[0])
    if pil and pil != PRESET_PERMOHONAN_OPTIONS[0]:
        st.session_state["in_permohonan_1"] = pil
        st.session_state["in_kegiatan_1"] = strip_permohonan(pil)
        st.session_state["prev_permohonan_1"] = pil
    elif pil == PRESET_PERMOHONAN_OPTIONS[0]:
        st.session_state["in_permohonan_1"] = ""
        st.session_state["in_kegiatan_1"] = ""
        st.session_state["prev_permohonan_1"] = ""

def on_preset_permohonan_2_change():
    pil2 = st.session_state.get("preset_permohonan_2", PRESET_PERMOHONAN_OPTIONS[0])
    if pil2 and pil2 != PRESET_PERMOHONAN_OPTIONS[0]:
        st.session_state["in_permohonan_2"] = pil2
        st.session_state["in_kegiatan_2"] = strip_permohonan(pil2)
        st.session_state["prev_permohonan_2"] = pil2
    elif pil2 == PRESET_PERMOHONAN_OPTIONS[0]:
        st.session_state["in_permohonan_2"] = ""
        st.session_state["in_kegiatan_2"] = ""
        st.session_state["prev_permohonan_2"] = ""

def on_permohonan_1_change():
    cur = st.session_state.get("in_permohonan_1", "")
    if cur != st.session_state.get("prev_permohonan_1", ""):
        st.session_state["in_kegiatan_1"] = strip_permohonan(cur)
        st.session_state["prev_permohonan_1"] = cur

def on_permohonan_2_change():
    cur2 = st.session_state.get("in_permohonan_2", "")
    if cur2 != st.session_state.get("prev_permohonan_2", ""):
        st.session_state["in_kegiatan_2"] = strip_permohonan(cur2)
        st.session_state["prev_permohonan_2"] = cur2

def on_tgl_mulai_change():
    """
    UX Pintar (Auto-Follow Date):
    Saat user memilih/mengubah Tanggal Mulai, otomatis sesuaikan Tanggal Selesai mengikuti Tanggal Mulai
    agar untuk ST 1 hari user cukup memilih 1 kali saja (dan tidak pernah terjadi Tanggal Selesai < Tanggal Mulai).
    """
    d_mulai = st.session_state.get("in_tgl_pelaks_mulai")
    if d_mulai:
        st.session_state["in_tgl_pelaks_akhir"] = d_mulai

def get_default_petugas(pegawai_list):
    """Mencari 'Derry Kurniawan' di daftar pegawai sebagai default otomatis."""
    for p in pegawai_list:
        if p.strip().lower() == "derry kurniawan":
            return [p]
    for p in pegawai_list:
        if "derry kurniawan" in p.strip().lower():
            return [p]
    return []

def reset_form_input(pegawai_list=None):
    st.session_state["in_permohonan_1"] = ""
    st.session_state["in_kegiatan_1"] = ""
    st.session_state["prev_permohonan_1"] = ""
    st.session_state["in_permohonan_2"] = ""
    st.session_state["in_kegiatan_2"] = ""
    st.session_state["prev_permohonan_2"] = ""
    st.session_state["preset_permohonan_1"] = PRESET_PERMOHONAN_OPTIONS[0]
    st.session_state["preset_permohonan_2"] = PRESET_PERMOHONAN_OPTIONS[0]
    for k in (
        "sel_pt1", "sel_pt2", "in_nosurat1", "in_nosurat2",
        "chk_pt2", "in_nost_opsional", "ms_petugas_st", "sel_lok1", "sel_lok2",
        "in_tgl_pelaks_mulai", "in_tgl_pelaks_akhir", "in_tipe_ttd"
    ):
        if k in st.session_state:
            del st.session_state[k]
    if pegawai_list is not None:
        st.session_state["ms_petugas_st"] = get_default_petugas(pegawai_list)
    if "last_submitted_data" in st.session_state:
        del st.session_state["last_submitted_data"]

# ==========================================
# KONEKSI & CACHING GOOGLE SHEETS (ANTI-TERTUKAR DATABASE)
# ==========================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ID Resmi Spreadsheet DATABASE SURAT TUGAS
SURAT_TUGAS_SHEET_ID = "1lMZmVmMXZO1SRq2JuAWLLN3EhoNHLuYWvd265WySVww"
REKAP_SHEET_NAME = "REKAP BAGIAN UMUM"
REKAP_HEADERS = [
    "NO", "TANGGAL", "PERIHAL", "Nama Pabrik",
    "LOKASI", "Kecamatan", "PETUGAS", "TANGGAL PELAKSANAAN", "NOMOR ST"
]

@st.cache_resource
def init_gspread_client():
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        credentials = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(credentials)

def get_spreadsheet():
    client = init_gspread_client()
    configured_id = st.secrets.get("ST_SPREADSHEET_ID", "")
    if configured_id:
        return client.open_by_key(configured_id)
    try:
        return client.open_by_key(SURAT_TUGAS_SHEET_ID)
    except Exception:
        fallback_id = st.secrets.get("SPREADSHEET_ID", SURAT_TUGAS_SHEET_ID)
        return client.open_by_key(fallback_id)

def find_worksheet_flexible(ss, candidate_names):
    """Mencari worksheet secara fleksibel (case-insensitive & tahan terhadap nama alternatif)."""
    all_ws = ss.worksheets()
    for cand in candidate_names:
        cand_clean = cand.strip().upper()
        for ws in all_ws:
            if ws.title.strip().upper() == cand_clean:
                return ws
    for cand in candidate_names:
        cand_clean = cand.strip().upper()
        for ws in all_ws:
            if cand_clean in ws.title.strip().upper():
                return ws
    return ss.worksheet(candidate_names[0])

def get_or_create_rekap_worksheet(ss):
    """
    Mendapatkan worksheet 'REKAP BAGIAN UMUM'.
    Jika belum ada di Google Sheets, otomatis membuatnya lengkap dengan header 9 kolom.
    """
    all_ws = ss.worksheets()
    for ws in all_ws:
        if ws.title.strip().upper() in ("REKAP BAGIAN UMUM", "LAPORAN BAGIAN UMUM", "REKAP ST"):
            return ws
    ws_new = ss.add_worksheet(title=REKAP_SHEET_NAME, rows=500, cols=10)
    ws_new.update("A1:I1", [REKAP_HEADERS], value_input_option="USER_ENTERED")
    try:
        ws_new.format("A1:I1", {
            "backgroundColor": {"red": 0.06, "green": 0.16, "blue": 0.26},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
            "horizontalAlignment": "CENTER"
        })
    except Exception:
        pass
    return ws_new

def find_perjadin_row_by_date(ws_perjadin, target_date: datetime.date, cached_col_a=None):
    """
    Mencari nomor baris (1-indexed) pada sheet DATA PERJADIN yang cocok dengan target_date.
    Menggunakan perhitungan kalender instan + verifikasi pembacaan kolom A.
    """
    if not target_date:
        return None
    col_a = cached_col_a if cached_col_a is not None else ws_perjadin.col_values(1)
    start_cal = datetime.date(2026, 7, 1)
    approx_row = 2 + (target_date - start_cal).days
    if 2 <= approx_row <= len(col_a):
        parsed_approx = parse_indo_date(col_a[approx_row - 1])
        if parsed_approx == target_date:
            return approx_row
    for idx, val in enumerate(col_a):
        if idx == 0:
            continue
        if parse_indo_date(val) == target_date:
            return idx + 1
    return None

def sync_st_to_rekap_and_perjadin(
    ss,
    no_urut: int,
    tgl_st_str: str,
    perihal_str: str,
    nama_pabrik_gabung: str,
    lokasi_kab_kota: str,
    kecamatan_str: str,
    petugas_gabung_str: str,
    tgl_pelaksanaan_str: str,
    nomor_st_resmi: str = ""
):
    """
    Mesin Sinkronisasi Otomatis 1-Pintu:
    1. Memasukkan / mengupdate baris di sheet 'REKAP BAGIAN UMUM' (9 Kolom).
    2. Mengupdate baris tanggal yang sesuai di sheet 'DATA PERJADIN'
       (Centang LUAR KOTA / DALAM KOTA = TRUE, TIDAK ADA ST = FALSE, dan isi NO ST jika ada)
       TANPA menimpa rumus NOMINAL (Kolom G) maupun rumus KETERANGAN 'BELUM ADA PERJADIN' (Kolom K).
    """
    hasil_log = {"rekap_ok": False, "perjadin_ok": False, "perjadin_row": None}

    # --- A. SINKRONISASI KE SHEET 'REKAP BAGIAN UMUM' ---
    try:
        ws_rekap = get_or_create_rekap_worksheet(ss)
        col_no_rekap = ws_rekap.col_values(1)
        target_rekap_row = None
        for r_idx, val in enumerate(col_no_rekap):
            if r_idx == 0:
                continue
            if str(val).strip() == str(no_urut).strip():
                target_rekap_row = r_idx + 1
                break

        tgl_rekap_sama = tgl_pelaksanaan_str.strip() if (tgl_pelaksanaan_str and tgl_pelaksanaan_str.strip()) else tgl_st_str
        rekap_payload = [
            no_urut,
            tgl_rekap_sama,
            perihal_str,
            nama_pabrik_gabung,
            lokasi_kab_kota,
            kecamatan_str,
            petugas_gabung_str,
            tgl_pelaksanaan_str,
            nomor_st_resmi
        ]

        if target_rekap_row:
            ws_rekap.update(f"A{target_rekap_row}:I{target_rekap_row}", [rekap_payload], value_input_option="USER_ENTERED")
        else:
            ws_rekap.append_row(rekap_payload, value_input_option="USER_ENTERED", table_range="A1")
        hasil_log["rekap_ok"] = True
    except Exception as e_rek:
        hasil_log["rekap_err"] = str(e_rek)

    # --- B. SINKRONISASI KE SHEET 'DATA PERJADIN' (MENDUKUNG 1 HARI MAUPUN MULTI-HARI s.d.) ---
    try:
        ws_perjadin = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
        daftar_tgl_pelaks = parse_indo_date_range(tgl_pelaksanaan_str)
        if daftar_tgl_pelaks:
            all_perj_vals = ws_perjadin.get_all_values()
            col_a_cached = [r[0] if r else "" for r in all_perj_vals]
            batch_payload_perj = []
            updated_rows = []
            updated_dates_str = []

            is_dalam_kota = ("kota malang" in lokasi_kab_kota.lower())

            for d_pelaks in daftar_tgl_pelaks:
                p_row = find_perjadin_row_by_date(ws_perjadin, d_pelaks, cached_col_a=col_a_cached)
                if p_row and 1 <= p_row <= len(all_perj_vals):
                    existing_row_vals = list(all_perj_vals[p_row - 1])
                    existing_row_vals += [""] * (11 - len(existing_row_vals))

                    cur_luar = str(existing_row_vals[1]).strip().upper() == "TRUE"
                    cur_dalam = str(existing_row_vals[2]).strip().upper() == "TRUE"
                    cur_no_st = str(existing_row_vals[8]).strip()

                    # Prioritas Luar Kota di atas Dalam Kota pada hari yang sama (agar COUNTIF tidak ganda)
                    new_luar = True if (not is_dalam_kota or cur_luar) else False
                    new_dalam = False if new_luar else True

                    # 1. Update B, C, D (LUAR KOTA, DALAM KOTA, TIDAK ADA ST = FALSE)
                    batch_payload_perj.append({
                        "range": f"B{p_row}:D{p_row}",
                        "values": [[new_luar, new_dalam, False]]
                    })

                    # 2. Jika Sabtu/Minggu dan status masih default "PREI BOS! MANDEK!", perbarui rumus Kolom E
                    if d_pelaks.weekday() >= 5:
                        cur_status_val = str(existing_row_vals[4]).strip()
                        if cur_status_val in ("", "PREI BOS! MANDEK!", "NGANTOR BOS!"):
                            formula_e = f'=IF(A{p_row}="", "", IF(D{p_row}=TRUE, IF(WEEKDAY(A{p_row}, 2)>=6, "PREI BOS! MANDEK!", "NGANTOR BOS!"), ""))'
                            batch_payload_perj.append({
                                "range": f"E{p_row}",
                                "values": [[formula_e]]
                            })

                    # 3. Update Kolom I (NO ST) jika nomor_st_resmi diisi
                    if nomor_st_resmi:
                        combined_st = merge_nomor_st_cell(cur_no_st, nomor_st_resmi)
                        batch_payload_perj.append({
                            "range": f"I{p_row}",
                            "values": [[combined_st]]
                        })

                    updated_rows.append(p_row)
                    updated_dates_str.append(format_tanggal_indo(d_pelaks))

            if batch_payload_perj:
                ws_perjadin.batch_update(batch_payload_perj, value_input_option="USER_ENTERED")
                hasil_log["perjadin_ok"] = True
                hasil_log["perjadin_row"] = updated_rows[0] if len(updated_rows) == 1 else f"{updated_rows[0]} s.d. {updated_rows[-1]}"
                hasil_log["perjadin_rows"] = updated_rows
                hasil_log["perjadin_dates"] = updated_dates_str
    except Exception as e_perj:
        hasil_log["perjadin_err"] = str(e_perj)

    return hasil_log

@st.cache_data(ttl=600)
def load_master_data():
    ss = get_spreadsheet()

    # 1. Master Pegawai (DATA PEGAWAI: B4:F120)
    ws_pegawai = find_worksheet_flexible(ss, ["DATA PEGAWAI", "PEGAWAI"])
    raw_pegawai = ws_pegawai.get("B4:F120")

    pegawai_dict = {}
    for row in raw_pegawai:
        if row and len(row) > 0 and row[0].strip():
            nama = row[0].strip()
            nip = row[2].strip() if len(row) > 2 else ""
            pangkat = row[3].strip() if len(row) > 3 else ""
            jabatan = row[4].strip() if len(row) > 4 else ""
            pegawai_dict[nama] = {
                "nip": nip,
                "pangkat": pangkat,
                "jabatan": jabatan
            }

    # 2. Master Perusahaan (DATA NPPBKC: B2:AK190)
    ws_nppbkc = find_worksheet_flexible(ss, ["DATA NPPBKC", "NPPBKC"])
    raw_nppbkc = ws_nppbkc.get("B2:AK190")

    perusahaan_dict = {}
    for row in raw_nppbkc:
        if row and len(row) > 0 and row[0].strip():
            nama_pt = row[0].strip()
            lokasi_list = {}
            for loc_idx in range(1, 13):
                offset = 24 + (loc_idx - 1)
                alamat_val = row[offset].strip() if len(row) > offset else ""
                lokasi_list[loc_idx] = alamat_val
            perusahaan_dict[nama_pt] = lokasi_list

    return pegawai_dict, perusahaan_dict

def repair_shifted_rows_in_sheet1(ws=None):
    """
    Memperbaiki otomatis baris ST yang sempat bergeser ke Kolom AG (indeks 32) akibat bug
    table_range Google Sheets API (contoh: ST No. 145 di baris 157 dan sel #ERROR! di AG156),
    memindahkannya kembali ke baris kosong tepat di bawah data terakhir Kolom A:AK,
    serta membersihkan sel yang nyasar di kolom AG:BQ.
    """
    try:
        if ws is None:
            ss = get_spreadsheet()
            ws = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
        all_rows = ws.get_all_values()
        if len(all_rows) <= 1:
            return False

        last_normal_row = 1
        shifted_entries = []
        error_cells_to_clear = []

        for idx, r_raw in enumerate(all_rows[1:], start=2):
            col_a = r_raw[0].strip() if len(r_raw) > 0 else ""
            if col_a.isdigit():
                last_normal_row = idx
                continue
            if not col_a:
                # Cek apakah ada data ST yang bergeser ke kolom AG (indeks 32) atau sekitarnya
                for shift_idx in range(1, len(r_raw)):
                    val_s = r_raw[shift_idx].strip()
                    if not val_s:
                        continue
                    if val_s.upper().startswith("#ERROR") or val_s.upper().startswith("#REF") or val_s.upper().startswith("#VALUE"):
                        error_cells_to_clear.append((idx, shift_idx))
                    elif val_s.isdigit() and len(r_raw) > shift_idx + 1 and r_raw[shift_idx + 1].strip():
                        extracted_37 = r_raw[shift_idx:shift_idx + 37]
                        extracted_37 += [""] * (37 - len(extracted_37))
                        shifted_entries.append((idx, shift_idx, extracted_37))
                        break

        if not shifted_entries and not error_cells_to_clear:
            return False

        batch_fix = []
        # 1. Bersihkan sel #ERROR! yang membuat append_row nyasar
        for err_row, _ in error_cells_to_clear:
            batch_fix.append({
                "range": f"A{err_row}:BQ{err_row}",
                "values": [[""] * 69]
            })

        # 2. Pindahkan baris yang bergeser (seperti ST #145) ke baris A:AK tepat setelah last_normal_row
        next_target_row = last_normal_row + 1
        for orig_row, shift_idx, data_37 in shifted_entries:
            # Bersihkan baris asal yang bergeser di kolom AG:BQ
            batch_fix.append({
                "range": f"A{orig_row}:BQ{orig_row}",
                "values": [[""] * 69]
            })
            # Tulis ulang dengan rapi di A{next_target_row}:AK{next_target_row}
            batch_fix.append({
                "range": f"A{next_target_row}:AK{next_target_row}",
                "values": [data_37]
            })
            next_target_row += 1

        if batch_fix:
            ws.batch_update(batch_fix, value_input_option="USER_ENTERED")
            return True
    except Exception:
        pass
    return False

@st.cache_data(ttl=15)
def load_riwayat_data():
    """
    Fungsi Murni READ (Tanpa Operasi Write API) untuk memuat riwayat Surat Tugas dari Sheet1.
    Dilengkapi pembacaan cerdas (in-memory slice) jika ada baris yang bergeser ke kolom AG,
    sehingga UI tetap memuat cepat tanpa risiko timeout atau silent crash.
    """
    try:
        ss = get_spreadsheet()
        ws = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
        all_rows = ws.get_all_values()
        if len(all_rows) <= 1:
            return []

        data_rows = all_rows[1:]
        riwayat = []
        for orig_idx, r_raw in enumerate(data_rows):
            if not r_raw:
                continue
            sheet_row_num = orig_idx + 2
            col_a_val = r_raw[0].strip() if len(r_raw) > 0 else ""

            # Fallback in-memory aman: jika ada baris yang sempat bergeser ke kolom AG, tetap baca penuh
            if not col_a_val:
                shifted_slice = None
                for s_idx in range(1, len(r_raw) - 1):
                    if r_raw[s_idx].strip().isdigit() and r_raw[s_idx + 1].strip():
                        shifted_slice = r_raw[s_idx:s_idx + 37]
                        break
                if not shifted_slice:
                    continue
                r = shifted_slice + [""] * (37 - len(shifted_slice))
            else:
                r = r_raw[:37] + [""] * max(0, 37 - len(r_raw[:37]))

            daftar_petugas = []
            for i in range(6):
                base_idx = 9 + (i * 3)
                nama_p = r[base_idx].strip()
                pangkat_p = r[base_idx + 1].strip()
                jabatan_p = r[base_idx + 2].strip()
                if nama_p:
                    daftar_petugas.append({
                        "no": len(daftar_petugas) + 1,
                        "nama": nama_p,
                        "nip": "",
                        "pangkat": pangkat_p,
                        "jabatan": jabatan_p
                    })

            no_st = r[0].strip()
            pt_1 = r[1].strip()
            tgl_pelaks = r[30].strip()
            p1_nama = daftar_petugas[0]["nama"] if daftar_petugas else "-"
            label_display = f"No. {no_st} | {pt_1} | {tgl_pelaks} (Petugas: {p1_nama})"

            riwayat.append({
                "label": label_display,
                "row_num": sheet_row_num,
                "no_urut": no_st,
                "pt_1": pt_1,
                "no_surat_1": r[2].strip(),
                "tgl_surat_1": r[3].strip(),
                "pt_2": r[4].strip(),
                "no_surat_2": r[5].strip(),
                "tgl_surat_2": r[6].strip(),
                "permohonan_1": r[7].strip(),
                "permohonan_2": r[8].strip(),
                "kegiatan_1": r[27].strip(),
                "kegiatan_2": r[28].strip(),
                "hari": r[29].strip(),
                "tgl_pelaksanaan": tgl_pelaks,
                "lokasi_idx_1": r[31].strip() or "1",
                "alamat_pabrik_1": r[32].strip(),
                "lokasi_idx_2": r[33].strip() or "1",
                "alamat_pabrik_2": r[34].strip(),
                "nomor_st_resmi": r[35].strip(),
                "tgl_st_resmi": r[36].strip(),
                "daftar_petugas": daftar_petugas
            })
        return list(reversed(riwayat))
    except Exception:
        return []

@st.cache_data(ttl=15)
def load_perjadin_and_rekap_data():
    """
    Fungsi Murni READ (Tanpa Operasi Write / batch_update di Dalam Cache) untuk memuat data
    DATA PERJADIN dan REKAP BAGIAN UMUM secara cepat, stabil, dan bebas timeout.
    Aturan Deteksi Peringatan:
    Jika (b_luar or c_dalam or i_nost) bernilai True DAN not j_idkeg (ID Kegiatan kosong)
    DAN not k_ket (Keterangan belum diisi manual), maka wajib k_ket = "BELUM ADA PERJADIN"
    tanpa memedulikan isi kolom Status (e_status).
    """
    ss = get_spreadsheet()
    perjadin_rows = []
    rekap_rows = []

    # 1. Baca DATA PERJADIN (Murni READ)
    try:
        ws_p = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
        raw_p = ws_p.get_all_values()
        for idx, r in enumerate(raw_p[1:], start=2):
            r_pad = r + [""] * (11 - len(r))
            tgl_str = r_pad[0].strip()
            if not tgl_str or tgl_str.upper() == "TOTAL":
                continue
            d_obj = parse_indo_date(tgl_str)
            b_luar = r_pad[1].strip().upper() == "TRUE"
            c_dalam = r_pad[2].strip().upper() == "TRUE"
            d_no_st_flag = r_pad[3].strip().upper() == "TRUE"
            e_status = r_pad[4].strip()
            i_nost = r_pad[8].strip()
            j_idkeg = r_pad[9].strip()
            k_ket = r_pad[10].strip()

            # Solusi Bug #1: Tanpa syarat `not e_status`!
            # Selama (b_luar or c_dalam or i_nost) True DAN ID Kegiatan kosong DAN Keterangan kosong,
            # wajib muncul "BELUM ADA PERJADIN" meskipun kolom Status berisi rumus/teks bawaan.
            if (b_luar or c_dalam or i_nost) and not j_idkeg and not k_ket:
                k_ket = "BELUM ADA PERJADIN"

            perjadin_rows.append({
                "row_num": idx,
                "tanggal_str": tgl_str,
                "date_obj": d_obj,
                "luar_kota": b_luar,
                "dalam_kota": c_dalam,
                "tidak_ada_st": d_no_st_flag,
                "status": e_status,
                "sudah_masuk": r_pad[5].strip().upper() == "TRUE",
                "nominal": r_pad[6].strip(),
                "tgl_cair": r_pad[7].strip(),
                "no_st": i_nost,
                "id_kegiatan": j_idkeg,
                "keterangan": k_ket
            })
    except Exception:
        pass

    # 2. Baca REKAP BAGIAN UMUM (Murni READ)
    try:
        ws_r = find_worksheet_flexible(ss, ["REKAP BAGIAN UMUM", "LAPORAN BAGIAN UMUM", "REKAP ST"])
        raw_r = ws_r.get_all_values()
        for idx, r in enumerate(raw_r[1:], start=2):
            r_pad = r + [""] * (9 - len(r))
            if not r_pad[0].strip() and not r_pad[3].strip():
                continue
            rekap_rows.append({
                "row_num": idx,
                "no": r_pad[0].strip(),
                "tanggal": r_pad[7].strip() or r_pad[1].strip(),
                "perihal": r_pad[2].strip(),
                "nama_pabrik": r_pad[3].strip(),
                "lokasi": r_pad[4].strip(),
                "kecamatan": r_pad[5].strip(),
                "petugas": r_pad[6].strip(),
                "tgl_pelaksanaan": r_pad[7].strip(),
                "nomor_st": r_pad[8].strip()
            })
    except Exception:
        pass

    # 3. Cross-Reference In-Memory (100% di RAM Python, 0% API Write / 0% Risiko Timeout):
    #    Jika ada penugasan di REKAP BAGIAN UMUM pada tanggal tertentu yang di DATA PERJADIN
    #    belum tercentang, lengkapi status tampilan di memori agar langsung muncul di Tab 3!
    if perjadin_rows and rekap_rows:
        perj_by_date = {p["date_obj"]: p for p in perjadin_rows if p.get("date_obj")}
        for r_item in rekap_rows:
            d_list = parse_indo_date_range(r_item.get("tgl_pelaksanaan", ""))
            is_dk = ("kota malang" in r_item.get("lokasi", "").lower())
            r_nost = r_item.get("nomor_st", "").strip()
            for d_pel in d_list:
                if d_pel and d_pel in perj_by_date:
                    p_target = perj_by_date[d_pel]
                    if not p_target["luar_kota"] and not p_target["dalam_kota"]:
                        p_target["luar_kota"] = not is_dk
                        p_target["dalam_kota"] = is_dk
                        p_target["tidak_ada_st"] = False
                    if r_nost:
                        p_target["no_st"] = merge_nomor_st_cell(p_target["no_st"], r_nost)
                    if not p_target["nominal"]:
                        if p_target["luar_kota"]:
                            p_target["nominal"] = "100.000" if d_pel.month == 10 else ("250.000" if d_pel.month <= 8 else "410.000")
                        elif p_target["dalam_kota"]:
                            p_target["nominal"] = "50.000" if d_pel.month == 10 else "100.000"
                    if (p_target["luar_kota"] or p_target["dalam_kota"] or p_target["no_st"]) and not p_target["id_kegiatan"] and not p_target["keterangan"]:
                        p_target["keterangan"] = "BELUM ADA PERJADIN"

    return perjadin_rows, rekap_rows

try:
    pegawai_data, perusahaan_data = load_master_data()
except Exception as e:
    st.error(f"Gagal memuat master data dari Google Sheets: {e}")
    pegawai_data, perusahaan_data = {}, {}

list_perusahaan = sorted(list(perusahaan_data.keys())) if perusahaan_data else []
list_pegawai = sorted(list(pegawai_data.keys())) if pegawai_data else []
riwayat_list_global = load_riwayat_data()
perjadin_list_global, rekap_list_global = load_perjadin_and_rekap_data()

# Hitung Nomor Urut ST Berikutnya & Indikator Perjadin
existing_nums_global = [int(item["no_urut"]) for item in riwayat_list_global if str(item["no_urut"]).isdigit()]
next_no_preview = max(existing_nums_global) + 1 if existing_nums_global else 1
jml_hari_penugasan_aktif = sum(1 for p in perjadin_list_global if p["luar_kota"] or p["dalam_kota"] or p["no_st"] or p["keterangan"] == "BELUM ADA PERJADIN")
jml_belum_perjadin = sum(1 for p in perjadin_list_global if p["keterangan"] == "BELUM ADA PERJADIN")
jml_sudah_id_perjadin = sum(1 for p in perjadin_list_global if (p["luar_kota"] or p["dalam_kota"] or p["no_st"]) and p["id_kegiatan"])
jml_st_belum_nomor = sum(1 for r in rekap_list_global if not r["nomor_st"])
jml_belum_cair = sum(1 for p in perjadin_list_global if (p["luar_kota"] or p["dalam_kota"] or p["no_st"]) and not p["sudah_masuk"])

# ==========================================
# FUNGSI HELPER: GENERATE WORD (.DOCX) - 100% LOCKED & PRESISI
# ==========================================
EMBEDDED_TEMPLATE_B85 = '{Wp48S^xk9=GL@E0stWa8~^|S5YJf5;^c8;sa*gl61-q2m(^%iG$G=wU{W3%F2l)KCm}k^G<#iSvtXjC*Guuax&3}IUS(%77`*Izey#Sd4K$=gb@!eZ6Q29h+jBK7du4)h!|)YSV(GHLxZnpwMc~+CPe>Ta^Wj|R4`MH7`>c`JLsm`Ggn>G~int2XaC@l-_lp+G(VwLhv}Rv6QeP3DqXHm=Zg#)V!5Y&fi}NyomD%;3iANsI;xHw<j<clq5piT?5_q*1#rlUC--bh@cWP@(L!)2tVoUG8lcSEEckJUtO)p2F<2l&gC0HTmcRFiez_BAD=!}2vc|nEvd8w)EK8^$7Lo<NRA@Wco#_u!51N8S{>H}7c{WBEnHvE#n%Ds{W50wPtFI?Zs7-b=imkIO7nBcFIq6}n9Bx-Mhxr~uawT^gPmaCotB&vz3a$eN`z_(xw;Pu)_N(OR{JbvXk9HN9Y-2V8knt_r#_9hOkH&!dlbwMEig*_VphwgYvgmF0q{BXyM94(ne0nKMGvlcU`8Q=$Ao|W@}!o+r-A|bZ`j>q``YIuP=rXA`;6T8?%7yO~G-rw$W%&DU~9j}$Q@-b%hU<4TGPag6VfRDO#b2k7DEE^VmXyu28Gq_a>ZWw%(q>i7o)BYQ|oODq5yzahG2vW+}5Fwq-qi<e^dD;eV<M#`y?od5QCs{Ho_P+fZxZ$Dyyr($-?%J?kUbnr){b*DX*esYei;0s_Osl+^w}pNa-=pWpc!@ob46aVi(e4nnUlCcxsYpYD<uZ*S4ndh)CfgL08>6w**=~Qkfcxl@wttZ9NXtX-E_&l2E_s{3B8QV^KT+@)%ykr23(vqtb%WxETV+V@1&~ue-Ojhc<-Y&Q|0+&ZHb>dLH+J4(4KaY9Jf|6eAMG%#qkq!qr~Kbq6eI99YyTneQot5j33<=Qkje6#WC=DERgv^m9Z(WILen{$x%4~PbC10LnZUWXJDH~N4cGXnz_zfzP#lBu>VkJS{HK<)SI<lhtL{BF*nR_8_D;7&dBz2mq`Y37l6syVq%k4`><2{U4F?+)CqP289S{)iW7s&Z(dBy}s(`QfS1hAgfTu|1ckLBG2ktxgfvksCk2#+OkIK=eTN!UB-KcK%UXvE5sq#BNE%v~{R=^utQ;#4CvRzb-QBll_sAa9pd&GO%bNxAZ*|c(!sOp&4i>`WZmbS{&@C{|FNFvfimF#tk;KF22p<SqlPK?pJFbEf<R^l^Ci9DVAA+8;|4GEC$8Lr*lfbUFDlho+09o7MXk#WEqm(LTbp?V2hl5hg!EiOccUQpmi!s|&1FVu|{VdWhHXOw+DoDhm4`{;OFrR`Gk{_DD4f~%~}VK$X<D`(-$l5E1)Xd~a3R)lM0Z49tEvyPJNW@k3=Xi+iQCA<M^v{Ew&CJ8g$Sc^vSNRJwejOfEGJEv>IbGBx%2vIlVk%MZAuP?Og`8)Fz8Z&PMdxjLt%Ie&l)xQ8JD-A`mNZl)mOB)kxg<g2Wvub<{lfY+>*+BD|1;2LB)DQu4ENV}oz$SmIeb(BM-XPmMLfrYFIuH`JAiSxck~Lv`>;LpAmO~}<bxYISstgjkz7@vLOyZfY&}ks563o$r#2ur9dVCf$VD)0djA(uJ1ugmUqj?K37FR=(iT`@wp9&QnXx@NNNU)GA?uF8OlFBbCwAdaA2v?@P`fV;KN1^HlnGH)3!T+5o>eGx_Aw!vvHnJt3{JwDnmt&?UTZAfsK=llTU}2^~`HWcfXi^nYYsblmGd037AAE5_erpe4Zc`Fj?FICIn!LGAvIF=HtZuF^X&e6(sgW?+>I((YO)cSUPv9GJG8C!YQ6U*m3>&jZ@oxn)>@*=zLrwc)SF=wFu})3W_fj2<Qh;LtTrbyVh8dJN)Di2FcpP3fZi)&!Ow+2*cy(5m0jYyGf_cL1B=V#)VnZB;AkZSa7e0?ShROna<4+k2+t&oPD1{d~Pl7R9^eE*fxk3zD>C<v40jGS1TP<BwY|2IK-!|5G;6n2funp=Kqq6``06nk6@jkH)(j9|u-7bA8FYigPAj#Yv$9rOBuy%)|h_rj$nlGGyAM6`__APr!;}ZiE;TT0&Q)qkP1@lDT=V=~y8gyWMtaA;F8t;BWm-%9!I}|eTal{9xw_y`IL+OKhEi}J3lEZ8iu1;WVWh#IG_3XZUDxp&&xQs*qhA_&`h$@hUf&i?ydr{1VX4EH<3rKAM@@}8q$glw^?|DzaX2Cyw!M$0U970-9+Vu1uz56u$e`+35%e;%Li7@Vj>o`J0QWg7ac|A<*+Us+e_2y}u?MxGhG_D&dAG!Z>5SO(@K|F3E?dspFk+Ra*OZttS%w0V9hy%DUfxPmBqD16)i(122W@BX*Gwo5IPNZ%Lo`z;W#&Q&!!_4^Q^gY)Y1~f-{(Q8?Z8!6sT699U0i5nEHgzCc6+Gn6CAm@n_=p<zGpsFOgyYY8ERFUMq=7Q`|B1F*>Cnv*#DyjXPP6;cPNuNv_!o@{&HgtX4C`}^~Pu@hPQOI$8%_pT#00&Xo1r!c1um|B_BsNqc!AndM>!Z16CC@ou2`)~?TOCw!3wAC0>!yKIEbe*B#md3|RkHHplNlPlQBB7)n5%TS${Bt>EdU>1cA^PXs0Bz^l|AWz$N@@83kn4Z4#N3=ppVUOGE>^#Eyy=k(y$+DMY93}jQWY;FL;g{<8ImSSL;4i_t@A&JfU_~cZ~A!f(=#gc?obQ%wkVAG|WVK&TZT#IN9M2PtDmdY!Mfcg|-{?*nlpc(VU_PcO0cri^@Arz}lLw_a>>+`TJ53;-6jysI$7Y4UN|%rBw@pTcIr#CVSBB$SqH)mE$pqd{^1;7lsG^p8_@s);#IB-R%s<1b^U>j15gS%xTQjn@l$IS5e+JDih^kQ6KZvLx}eB7EP43I8#fNU1q7!*E_6hB@QDf8bp`lGq=VMWO&*)Xjgf375>?^5Zjbd!+g5FL9o8e?TyK2T1yy^C-q0*H6a{UWb|1sV46qbpeMB>^K#VbQ&8%3I9GM%g3p6B6?M=TskvXW_Fq#tDkJz+E#e}Joq?^s^r_<zb`<|fn#8y;H14U)rYg+)EYB^S_N{O0i<7%4jP{K2;2UjAStc{kjV3j`HmA^?knUl^F)*ye=$V>U+1<qUre+YXp9bhJ$)ZTTYXf|$?2CepxzJ>nE@eb>>j#>Y%lNI3maC=;@x1sfU&3%4ZCW&Bkb`pocxMk@*jLrm?NYy(D&lAy-f@Wa2u`naK^zF=ZPJR?9bT}F_a_w6A1!W0*wJ324yV*NLS3!<m&+r#hdeZ^OW~5i`9HnxTl<KOYU7zrPd3vQ0$$5$_bvpEVV?jl{EE9M9`^?^PRv5DJ$%ZJuCmY>r-!bCXI=cu*l&tU{sy_4{37Pa#>h-sAw1{wc3IC##b|{K`})%ue_)yK_a~<ln?1Qm@`K80judkCO2L-U$FwJnmR6ZnXxb6H46i@!gH!7+?e;Xo0-Q-<YXK!PIr0@x9~+?x=da-d4;zZYBdL>)8E57rAkB<adb3w_dSyoHq<6?QRVP!Rf%uajHf<%NP~f;p5kU|-#*I2owy3KFt|HlV*dcC0+FGv|Qg+u8A`1A&Z!vp2N@?YC=_<C6ANG}kFBn5rck*WhoXJqw0&ridaB!4>@LlpG0tDu*%G~YZ+ky$A9+QMinpjU}O}9ab@+zE3GosJD%@2keCP9I7_>`V)$_!ZqD~!Z=C?HYS;SYGqppn-a!%VFRa`!T_!;{Cw{lGA-9ZN!~vNKk*!HcY{+ah4`>{ksm&q5O$37{$3)92u!zNH)Vo%J_37%Vf-=S}y12PERz8Y1?=3F-@dqp#QsKMWc*C0IGoC{aIkmolmVs~`@XSAWb-^($6ER+8j=_jzZrrnJ0q&A8hBOM)-t<BFUN%C^E^FV^NLnw^G^(^I>*R;r95RB&J8Aa$Mb9KB9lNBm)Suk*m=Z(!K#gt}T|l;hrILhSw25nf^@cEpkgvl4%Dz|s4YMB!_W5#AEeuwJLl8$GuCp$F+lX4;A~({_Xid#b;#{JNZ-_XoHRmD=2QUT+I6*CqTn;)5(c3j1bnC!??UqC(cLwtql?*zIF6J*<Rj#02ga{W3d)V&k_1VWRSm%?GANL4kIQye^SDCTy6mNwgFS>4)1KnA<B27Rq5d#vy9~&g@-hnuIRB4cQ^v06|k_K6-h?b6qxi)m8#dA&*?zgGKsq2v1Ub0O5FSr<VExeHa}bOSU@D_pQGXCq)Daf8i!8R-AyJ(rLvO(U4wLW9(Q5300PN%uujSTu_}XJ=zvn$iDT=M2QIxYHeCMWNhznKM{=sEar1ktkG~9Ek`N%D!nf{#GFsKuubbdVIS}>klALffiII=9(uAL`9_Tclv*SwsUCCnAHu2!0qh&`fZpneX^|Fbh?8}LckH=gv<v+L70h_tjS<6}J?)z4{__gp5jFP9Ks>ET&+63+5ILKjF(vt-Ryu$B=Z(E@uyEG`+bd9*hEJot9a;_c251#U&5=TF58r75Zf>jgpqM#Cp4(tD9Vlkh8Xvc!$H5NXHs|S4{8lb_(aP<3gmo2=0y1pisUrqx;S4=X#f~ZDW|7UhMYdlXYUSb1gnC@X$}av3U%V1gF?t|hYC@Ka*n6UZ{^N^f_S^vfV1`%uOfLdtZ8ns?Y}@Wrzw)GPYYK}*kc1e_`x}C$0N3mKR{7XPKKz4lKUJEv?v>+4<ZRVT#6n@xAErx%Qo1E!0d;ZW@!GTf&t!{RLy!8KS>peEf4iD3f3)Y0TN(<mAulwktJ2r^&$IdX4BW?sA0AhMeKYismOu#9QUONiArKlwEE+~U0H769R`sr`J8($CWCv2~mK!7IXh+okA@}oUpG)KM7vv^}b!wMp@fISzuN}7T+8@{c{;=c8k}{wR32=zsRH#ERY8PrUqFyMlvU_^4(s!UV_Q(b{$df2#i(DNQQh4+yN%>ik{?zo$mssHgw<fcD7Rr2;5z|Fhq4fW%h7Wc|UrI8q@N`lDVXDO%Bd@eSE*){4p_`kFT@#M4t%Oa$6CP<(sRoogwaw8UaCjrjP7!Pt5`UNxAq%R3U~A-Nrv}Nzw5qX}H#5}GDeLVPdNXQA@m=f=CjLQQ9}Itn;cNAA0hy|Rog$py+v?3Z)9VbOREJ1p#TFcD?z8ARuhmXmUoD3pGCb#U6T$;}4x58l&K?}P(H*k``x9%E)5hpSUvTc<1+CU^@Tb}o6=e*{s$Du5k@DL2dSC5@22J@dEz@skS^5kF1ktie7{YV+jPVhYSLRohp^O(y%1T=7qQBY|W(LX6-5jjI|Msl9vI>~mDUo<$e9@At*Tq&-9WnY6?_!;D2COsyzsadFnPMU*sU5ox@n2M_#`6BUPsXwL@-!jhqKOr#T6?Tdvq^sl+QzYtsa#d+x05Se)+btyR5z##uoEG>3A45u<CDA?I0(!bN)=SP;ZHd_86+yzPuLEZb1eU^simhn9Q*85Ik~tC0a-jA2}D0ah0U#!wUoP;Hw*ocGSfS9M*=YJv;F8OlS~}D{$BH&l0kNZcZ;?Y`_-Dy8^lOP8CMD<o5m?qU#Ogs{oriB;R&6^G0^wt>0dYcfu+`DbIaLIjaMZouKywA%DfOi^Z#w~05Ml3j)0GXgsbnskw6NAziuJ;129I8K87wkY^%r`h>N@DZ11m^UuOV*&!x7CJ#-Obh9hUA4F<ktR<9wD$1fXnZz~NML`RBs11V4HH6D(f)N{G3d(|PxaRHd9smKY{_99lGk>=-H<>!`wgcsp*(I954a9Z@rs=pOZlYE53*^2?ro9ibiig|^j?Z|lDY|Z?_&z7QB@N?_1AM}iS%lW?HcgJqppHLsKw`jDCrl92bf0Iwm>;a2H2Y^?-1rlpcH(g9vtHJSRYl{BxV5fYG_t2rH5I<+g3!L+ui>&3D^sqNOKaNk}m{|m&t}!XOxKf$auov@u6I{0(TG~w6$%+W+stQzsKILCbu;!qoRTofDVIh(~4*{t=+QC8`^KL<CRFa%?$Z35ed)ZSB#U7=CM?`=C-nwFtK9V>;>k?JvF@hBt#PHMaZdc~0R0Iv@b>OBjEFkW0kPVsc%@Uxa1NwepNV?bs{lmiTx=piIsDgudG`)G6ery<3C4BO*{RQ50H{|zQJJ!_UxuiSq-E#r>WTr`+ISFuZm`X`VtwuACjTLm0LZ@On(I0U2Tg&3=#^Q(iID@|v^dw`gN1z8a|Mn|QdPmflGcvEvwRsFsgJFJPE?ltHR&<gAS1<+iugFe7!Z}O0DCk8_ln@ngPbG_>Wb_+zlnCooKG8VAm_0wX?J|4dO$HX=s0=HJec1O>Ip?W)cD++6)HTX7jUdU=^|L(5$H8?ECvXjas^Oz}=Qw2rsCBBdqg^bwnbNiXUw~)e=U#FMIOT?R*-9)Fe<qhF>Zb7?<|f3q{LI8X!6l`dn2{7#gkq7Ggmgz?_<OshRQO*e-F8CGEZTRExPvvqLh~3Sx=7x59)L>^6bBRx#yHu=tGL|B!_O|l%YeplK5dJLj&H8njP(uV$q)$l-10=Wnvph#?OtLPZ|8@8i<kA~ay|426ZuNBK%a0nAB1D)C!%JgmgUQvcjeZm5bG}8tlZyp-}Jz*J+@fxpE9VGxUW41xgVNukfFZ)xiKW+g<5#D$ezI^1;sv(BSzUaGiOc`iKOXm(c5a^{yA?nam}M`izBUG57mRMqzm{6*d$~NW)&L@{F9^gJ^t9q*dt9eU3zq9*Af;*F|t_n0ehHe?`jS{EhPaQ-3eJ%X7iKZ9X<f3z?)LBH-~)>BRjb5ou}N!#pq2?`~%(S7Rs^e@Z3#q^HtfJ(^J3>BBiT{8`*ClLr3J})OcQ+Fj)7SZBLN3tko#3E~bKOpc2R{Y;nKd#}zoaDx?lqGJ8t8!&Xe<&-#B)4j0XhGdz^rj($!Y<~K@cTpdI4>_$jyy~bsOlSan-jd<WyiqPlOZB;-WoNw}3!r1PIavDDsWZ@Zib=4>`pUZSs@bAg4S#pLfbC_jMjKNzSc_@ox`MLXbK~$eOC|gAjjL%(xJ^@8Qj*2Dp=_}vkWcGlK{sV|jVcIM<OyN2APX9Tn2tS1VV%DD-BQ}xo+o%~FLgLe}F)B2n@TU&(94ci-g@wsD2W7KCmMT`lEE$kIj%?Xr#)UfK?`pWC&oNYcB=E*QDH`shbay+;qbP_W;FOiM&v@N2Yk-(wUToGZpT4yfI4?D`X_ms8^Wbfo?oPj3*#w_>f)oVL(^`m%-X6Q@bpvkqY0w3)ZUuUhmc?3gGv|s6gUyLFMyic^mRT=#gjFgk2px6^18{N}WMV6C?vIxVkN7#3w!95s$ed>w?%wu<@3$SjpA>P#wU#$jK%8L7Zd}^l27a%G;5@n_Ygi=qKf+lq2&q_L2KvMKh3;wM()!x*qYg<i4;PsU6-xk2@?~k`Zv3rhagZ}?b(1en!hcebB);RWsxDF#g=i(c-R#)(+jmK@j`C*?O7U&egol6<zbcmb<22KVqW&XQVn9Q~%qu4SE7M91_F9@yXWp2Vu_(+d{Hd@CRxxyF4Hr~SJGmz);QdDT`{o+w=a*vd67^smPws7xKR!wmzkGB(2(FV8h@+HEL12DGbVeO~7Sg7ND$Lzz1A5)wdOt>okMqueGJhfYSbf<3(dy-q*l1RdG<kwfl?t`t=37Sd4v7O0JW%z0maD1|Pnyk;UbaH8axXo-Asu+EirTl>6V%}9e?>05Y$_{>pE@!T5Mvq4uYTH`tBd0>yamIDygwWB@ogytc|7j#0ZG?sDBFc_<~|`Hm;FcO(N)e(l-lJs#9{;w9EWLF-vK@?xghIE&RopXdxnZo){uHsq2ds`y7@pXIta897Hw;f*=%qGSQJTVj4`=Nd8t`glri=Q(611wib}`_(@LMO2Q-fX*fY4pnAm*zgpz6aS|0SuecvE-JMt8ppoxcgPhRvv&#H)VLs;i;CVH09aV%|WPopI@QE^gmpO6_^^kwKQ4z`<mYsHL}9UwKk#Jo+i0CB)W6_g&>N@ub#5O<9!NM~4ZRsHN2b3Xa;0H{yp-%tx6Pr~ZZEkG(fvV~wIAeawBuf{-399FY2f(RBp!ioXF7{zP<m(UvX?n?4#3p2o_q0M3tt<knbr^%$>_sn#YE0Q*So~oUv8pq)3{*Cw8i=8?+Lzl`}_h!IZVAfN1N`?~Yk$BRK_r98X1?HxPa>`C2aI=y2G^5)2up1EQ9QVQjeIl|JpqgBIu-)o%jN2nzhDBdub=n}zRr0}70%NpjZumgozEDNueD?(2k_Vrxr!HLeJ2!nc)Cb3bR#XK_PCVnnzG(yLwW2lB-w0<))~trqB*Th!8L~DTR5);K0VJ}abKXm$5p`A7-SQlDv(DALXmj4IhncrsplWqB1u!PQ$7xZ-lMGrO05^Gc#grm9`fz48UD!RT!Q+#P)?jh9tGsXX_EmyyeGnchC~e=SsPx~@7N4aN`ws*pxKIVuG9N!jw)IBu?s{etDvE<^awE#*P0H(h48A|%Gr}@GaU7%pTWD7rd&gR{pc8g)wS>Eg511QxLLTJ~Y%)UvDb;+!$*Z+bs$Ks5Fn09F<<{9h9ieBbfF0{yTPJmjAQXkIly9*XjYnU6v_S(J{wpO?%f2@chKv44`YXhU+$#T;Ra@|m`2b~VgsWJzM-d)hE%|%gUO^FMopEpW-{@P3CXYuqjbkt=MvdNI0D}^so$rD(y+4BpM;xS*Mx;cPDuVV#Dre7nXOtY!Rake6E68ptNLTM$^IdDYOudhabxEqVFe@`CDGSzVY;{$bysQ|ckR{WHqM-&pyn@_R=@v##Rs!YQf3IylsQi@GE7OPA!0C;Dyb(ipBUaw>6!MHLUgp@<U%4fHmf7sWqJjpXZU^Z`U)kn1axzP{_}Ql?d^h5Vz}P?-z<EGm85Pq@u#+T+a3DD?@p#O$PWSy2uu1osM?Uc-kdEkPK-c&SLao@fV$`#Wi&k{?4{{zrr`jq73{mbJbW_aUy#&cyfZ}PHP~m>W2tI+5Hc32VOFMFn9wjf(iwa<9G8tDe&xrV+JVv0ek!!s!Rpqt+l}-FJC(@j4^C+!Im?+Jst{8I^p!`T!@@YbY0GO5i%#e7j9pEr<SMYhwVXqYoGZWwNnDI-~4fB3$2X3&O95l`u0Ahk7Z~@J(d)8HQsbMoWB<T^(+23KZv6wTOKYOL>@HsX9fF*^rG{F6%u{?T`14qA*!qy7s%zLMp!UR~dQfONwu~M_7GeU_cL3A^d!VU<HTG(Y(#cIHlti>hU)^=Pwsl8BLG$IuMAXk<O9#EtxjK4o(yH#a5M3gDQyf1wU_~v5R&qmkKdvTz(rT^<QkTV_OF~HR94e9OZ8{C`|@EYxzC4hq1xZMV-QE5yVZU(y|tcetsLg@qtW?!Oa0zNUuREUmQ<2(YjQed>Y5d!ENH;z%q`i9_)Z0%@He9pZ~i9K;h2_3W+C$?9gs896h#%0c7QRJ&deOT5(FN;(|H0F{~v0sc~bECC_JEy7c2!>7!?gu$D)=ey{nC#96e5GTSVK9(x;CP6pyeHUg0|cbNf`JjezxpL%zshQzs-yU&Q(P({)KYhn!V>2jkBe>AS|@b5^U-p>s+t33awhi=;W>#n!?V}6o_2-o#!^VrjD4Jd?ipTJVil6L)gnQK_vKXYyS}@prfB3yVuH1^!F8ToE2Rea4|o;9`aC7P-NXNx-hA~mU_n7DpE4GPshYk?!drI+eW8F>YPIe43Uf(PnbN*<;q>-UTwlxevP>ZuSTQ!nBwv2-h-Qpb5s+mZ?4f2zMU(DH8QG$z>i-kEs_iJYJ>0Kms5Sb#`L;m&KI{8BIylHKDSEtR*OV|Z@>wGk&OWPbQOTUQd)@1pIz`+~9p-O{eT}5*K094mZ&n>)<latJS0NqzU{P(4K^<GXa5p_`q*GBw%4!>Q)n^eiZ%<#7nDQG?W2XT4!6_D@>FnQxa;0kl0L$wk4r7C=lTF57_-}BZy+-GY1eDfZ5ZtstlO_vX{mOl0`OpNzi`hag+yzbnh*sY=Glv6+Qar_aWg@QyXrrmp-~t*@S6yc$!rJ-~5-GR8yZmPI;!?&Db}^M<#7d(b<HPr%^UnLK=;-b};`?O>gcU^ZDeyFk-hw{TtjlH=yBZ@;sDpW<G+k+OK{IRrMk|!S`K%ZR&?S;$gsq_QnjE!*q^)gzBzdah#L?hscZI|+zZ?{x4%#7ZrWIVT&g(a;rnE6P`p?Mv*tSzJx-h25iMG?Z<)j?%zTvQGW>l)$4aq#XTlVKirpd)MGSBko9S&t>?%gb?_^{nX0)%0efEjxZvEjUn<Bab@xhG0fJh*-xHXtOBjS^<$@JE6Er0*+vpOx}qhbBjZET~GnTqIpEI9)qj$O&li<@?r<KVZu27yVE*Ohg|Fp#)gVw81OS*yk!w8g5Dlndl~)R4y^APpa@ztzMmiq=IcX?VIEx;ayoAn*D%d((%vWYNJOr)}k^$#$wg?P^d@|x5t#al`2!!Y5>V{EGn0{G3mPu-FU+)p6gRqyK65SjunQf@MeMnX2*(_3TN$%zF{qe_AHAVQb@*YX{^MKCFm`6zoG?@nNdlcbT0M9uM&0pB0E{t$1?yS)tE##OXKG_MEq+a%zO${>x!`3bfsK479KUz*O7t}D_aJ&xosCLGz_1L=HRM$*bPMnFA<I&2r`i*dU^Re*su?w=c|h?H=1$Qz3My)d|=nuu%Pb(#(=CJD}6xHhfZa$Q#?AR>s)8V;2qyHWq%1k-3mY0YrMLBfrt*opf&*GjWFNl5m57;IhMiJ@Y2-azM<8!!&VeaaRKa>D>xAKwPtra{=YOcwK1-;2fnfr$?JxRz)TQL>bDeSw@fXWSp;uVPhZwzFF~PBFRJwwsq<Mp$Q#32ce$%Xe?l>z|H6ajxBEl%D+-cu8&yutv+4vHx`@DCk@QPweu1OiRz{dN^68f`)lwwq*e*QnOA$fTp=k~x!sZwV2t2=hUB*t6$}Bpq_zFJP5CPY*DdYOQ$!{Jn{eJs0*080VsiId@<4tkR8SAO%)Q@d+PnePiRx7(DGMU7dE23vXX}y>FxPdm$662SW8XoyrRKMQq7sy>*tXA%H8I^_DtT(<g23StJ|G7-)Ng={;McDuYW)=Dmc=?rv<5g71=NFvC?oTLU0&U)HTIjoj!jZ{ZIC*Ua=06)p++AzFD+GJEDj6nahGq>y$>g24pS2E(6MB-G#pJ<1o4hwBp9f%+*M)K^W6yBt`cTKOM0-@&Q^n3qKJpdo2g~KPeAz0_qkDxonjvD^C<t@5Cxu(FBc7B+k1I{-N!FFL6;uJ=;pIlUoqBbo!nT(CGi|eFAGrZT<>XK}K!f+{G|fFnD~iYLt{<;W<-1MY5tz&A?RQrM)+`jr*668hMj#}Lc0HCum`uxIAK@7GD$m_)1h!4PkEwAm?SKKC14{Db1Ui-uWpwV<U8amLYO!1e0nnf3MdFiLoMZWX>^bfFvrXHGni(pih?*uw_ML9^Z661vxUvy}>`|ibwY`d{BQK(7CKMkCh6)W;nW{L>!+zcB4acL?i<P7j#35d@Q7YT=5_$b|bkBF=b{STUiNJ8%JPC%-6Zj?+Q~$ogjPHQ=2I3SpH&|djJ%wUi6r<^7KcO3zH6>myveS88rz0FT3o&kGX@CooibNXCk(9YFny)N-$?h;%s|z2H6M}|E-CO2Hv82P)AFt>7;j1i6g(J<%PfBUPuJ8+$pzm_5Pnc(1D_f{9VAF44zGhM5!*>cgEt$2by=AYC_s#fu6blpe+Xn28^%)Cq{Ln^ncB02#?W{HsY+u!d@aljmG)9G0OgW^nno0L^F%woMOo86WgVFL_70lN({yl0Md?2SQ@Tk>qt1xQNL376)y_Dran;pjaq2*F25oQ8$|DZr;UVF~(Z#4vOhEeK(C}bn~MCv8(F*s~@q;CO`9lX5o&#i4OG-iIg+~VC~4s@B(A0MPaI0J#1;C>BhzvvqLb=8_>TC^Ym7?#1=6r-xmF`Kgtp(7ZA&2-_nP}JlNGW?z|0lgZYvf_f<wo5&o_*&Jae*%MdvySEFPpz*7!~q7Z%~GUmLd!p@T++PDyfab!%JRtkUhdyr%9N<yL3DucciAM7w!4AV1|(u_?_s@tc-mmFroU|118YgclnLkZfeWqZqiDY?zd15`1zSl!m?C{Qrd6_8gH8{wn$9jgp2sHf^(XhIa5`l{awCKFuj{%Q@RyDDuqZ1_*>{vdgb&v5z{6ZJVN;E<PboU8EV&DlxCyD!o1RnYl{&6_A6HqbS!$3JBG<Wo3Fhxer%zo0e4z$W72Pv;y@!<~TQUZ8$k>hG++gFnG$YJEs3!k8$5b1?YjYvdZ<;jAR0I>*W0-C$g@_pO8kU}|xpYBk`#9zEAeJ)*SvCl~oOe|u_4j1ipF|!%FL;H~i*VLTggW`7DckE>lFzV~7K!`dzgp@fkMh?oq{UH^VBh(tphsj8J)WIs&T{-6><U%qe2wc?mX0S5(rwuNOVR#Z6peo7!$JdfSdzGgbxFcbKW;AAS_~=8#2|^Xv(dp^>@!ax%X)QCy~+Trir-VeQ~Jygy95yG>+px)Ti`6w?a;9gN@biH508RQP))j|Obgk66_*q|JSjJLd>}3YZDep)To~!H9ohXt0=)|x9fgwGyW4Ebncr13vjRWAwdCEfDkl<J7xftw)Jl?#cQ$L5u{ADDW|a3cM9Cu^)+>}paM9bW6nq}FphRf37Z)KtMM^PMKeRox2=O&DCjX^~sxjDMyT<aJKd(L5Gu8_;d)>2*LIg(G+3GCOV{irAH)Ca(VMj}$CjHky0i4g46^uUx0KV9oExvHBdE($4J&~J!n5{Jb!(B8C=C@Fv-@qhH956+zil<51f@{Zx<mJ`ey+`yYzQpZf+9DG?E9AJd)dr(PDbsbM>S0m1y#H!CY)x>Su%FogE^L<8vr#k(<>^PJB}YhcgM4|&mA`NbTu9>|u$NsoR#R(-f3%bH>B>WyD3SwaDN{aOBG8WSX^4sW6}lJAmyTLQ@%xoItA$-bWEd<Ve1^IR5(8>!Lf6egQ(Xz%ZP4-yY=<z_G>F{E1fBC_2xzXrz*RU?!oGZnryK@jBkj?H=${$e;*C`(W$MR?9FHMh{rAOtYv4NJtKq2zffTIGn4;9&y^Jrg&eT8vYvfvwvB~7}m{m?cUi2(HOt>8!bu~n#?f6zbm5lo{L79wse-9N+Dj;?sV0*(k+k6&gQ)p-AZ6m1BjqjNVBCpT|jTTb?k`JmdKy-TkBUU-5wIBi1=ko=eda2cczD%x7V{Z85iBw#vsW^Cq4p%T6q+M#s(BCt|47p{WnP414KVS6+*bf}{2J-t<_&FY}{+w!mf!fe%od!plMERV?Tr^W+gr<k0s{LUW=C`2%Nc~1ptEHfj6mlwJL{vNuqx?7_=<z1-lV;ssSDvkh>zw2V8@b#@IoyM}yo*F=4LPL!K!-uTBt{<(Aed^GU8wXWimUy75g?IAhF6mF;WbC3So}hx2qa1Ry1M30GQjM?>bHA@0&d7-a<;c*m24DP<?3KB($-!GlYO^<FyQv!aF;0#zPJK~awCY(eD}NowyckyJoeZuBf56mI(l?g2-S64tBMSq@G*b-P;RlNSX{LUUX#)6#^2ISR{dz0GarjrvBsi<Dz$&KJAw+zaSX&8RnMVdHXh2j4R;_7q~JKhHhkelIP@QIBZF$2#hHVgMcPTt*<C?mCCkIPREcnTL;<Pd&&i4zN&tk~-Y3*J9eWF|lI(==-YG5haWt<)V+tJO<Yn_vMb}KwC|L#pPstx)+1v9tkUY70+Qsa1u;Goe!S2JY?ZxM~p*c!V$-TAC^w^&HaFkU-wGQ=TI0Z9^owpxAhiBu)viP=xQLEl((bqlL==4$B*0Sz&u1~ANsmq9<*Wod^$lVCYQ<eGljN5h6Pe@FOq+)e*w$DswxSYuO(Znugx_C!_{MxE+JbNvJ$0>$!n>ZWO3p0M49ycSR%%orh1e%Hk|Kf}+1bK1rs{DGJwI|&?mCU}292{+fo@1!)a@5Y-!orW76UY5?qpJ85SYlB&MGM9Kt&@l7A{g2ZG2pk8i`IQW=#M8&I<!s549#Kw82fT?i)`P-rHq7j8VBtT$(MiHW3y}~Ngm#a?xlf!A2b><f9M4lC2r6Zu$o4IWKt%R%#4Nd+!nCA3^rSlgm%a_t5eEWJJaA>L35m~paaL(TR%&`(WS&qRy3%#0+O&)+Z0ybBDX{FUD8ylOb2cJr@U^%y!*KZ3qsEwI4QyAQH5;OktGEt^iHbV6G~HJwL{9B5Qn)?nY)_8zXqV`HIx)15%3rhr*2q9v4D&Z!-#-`R6`JKn}kWX2U7JYrMH+?dQoQcR5W^77MJtBS~w}+AW&BsQGk`pfcA!3A3zv|zmf5SJcGBbh-3>SB=UYGm0g8FBY@B?nSa)|dDDg2fC#!v)?=6nFU=@IDxBXFH4D)*`fORT&R-!}%~1(v%1HJ|uN(}6R92U!D)78)lQ{?okyDQuIwOU#``#li3odM3Y{{ps0IQ3}FFXO<UjWU@hLId$I6tHzjIkybp-A9CFuzW#{_!Y8%nidFMJ)5kbqb4pDKEK@L=u{bC^ofM8oZq7Q<3YCj=1B_?lcRb{9$VyBPDDKtf5>nEZ|<wK*1;4+;L*QnA~j3hC!1FpJ5&ARuGa>x$fF(MBu39nNqZuAkOK?Y4wvTbzg-p$?XK3uDX0`A03RN#Qq4t*ThYYONi~4xg{G3au|nie&X<Rl9c$O{X^nOI=RlEaE-hph>GPo*)2`I$EKiSw+~_CHtJqV!zF#{YuuBbcT0?-?3mZ~v<Egr$!%P8krT(QfYP>?czP4=opAw@E4d>zeYQ6hA4`?XH#usJ7kj+KVQ@H7yDWXu_D#s~h680=u}tQDt9VXS`CGvnLyyF-Tm-J)hA>HIMUsrf8SYM}UP86SG-qC+pRF;G^7g_vVt#X`j(Il!A<P|{g?jwVtT4_3BHTr3joSy?2n!!mYxK`M4b6g1s&x)0V17o1Vq_$WoP9P0>^v}q^Cti%2izG^lXW{8N)t?>D{E}nL{R|iMKU)<Tk`6CY<9H$mLmkjaWh1_N#+dQLf_J9@uVc-7VpnGD<}cl57pcBm@C@>c}382-lh0BD1ScfksA@DHSqUx0XwesgSt-_mqs&PAQ+ZS<B9<RjTLA|{{kws#V%df9TBp!t=_-zJ=s9oxz{t8p}Rod@>0%?NBL&S`OdX^Kl*f>K_v+Jin?xPBLKtcH=1X%T)p-6z7rVyvz}K-AZMiEjC%s0$~=l`f&!FoH5m0MK8@OOwXUTXle3cT&y?Obr`#&z88GCPqP5qHFF^(@LGf#eyfi(XCSioI%p@&7)y%t2$16BN<=gxNTiKqfM_+3Rwv|Qaof`Gi)`0CF&%CBA$LiN47blW0>9BLef)POGo`R`gG1PN>32sD{`7!zq7EvGQA&4Dg$zX#``<$by$l;7-^V=>K)ga_(#(L8#0o*q9X~|oa+XCUPni5h@5GQ%~yWG=TH(1=xg@xvkHgHD1K8eIIz;8|mH!LAU(%2bq1D7zlK1f6#Sdhu7V-Ufc^^J$Lq7s(*)U#YZW1)dQ0u}svKu1cC2mxovkgo@86!l}Kl?Qn<e-G>AD}YPCEM^od2e{y}{v^%xSH6D28io-ZAx>xa?Hbsw5d-Cq$*r#f>>GY|=+7s@X!ED-<4|DfW<4SI>E(}wBSu7K3@gaGUth@e(Rsjn`iLDcH<e0l%<%l{w@f@aJyOy8R`m2W6RrvfthHY0Jl_~|-?k5?vu$bB^VB%Dj$SYh$-*7t%W2Ajz8J{`kNz0+L=j{MZV0|h4Bz-sSU`CgO1=wkRS!)0e-ei3xq+RAHQMs&yO4S&mpYrr>r{U2aXZZs0rzrtr`YEp82zN@d;zJJhy5xd(t!E;9l?*+CQYV%te@<cO-EW|o3FXTQ~vuM!h%@i5N>Y9rS9&JmCAd}y_FXQvjHU}f>W%;)uA*CjR3y)VBxD122WYz>0l(<r@N3<QBE27P5sGzKW<IJNX%R=EwE~Li{VuCw&a(`6hn%U>1yMVRX#`+aiu`w;g{A<PRz$kh04Mu9#?;XqLK>(^Vw{>Pb`=$R&@QW^4riw5p*n6oD#VR*5qOG9y>VWYq|}h*e~svh^3ea@F*({i~O25n$6YI{JnGxEXLzdX)NlMrkL_!#BHLv!#B4gAD!@OeR*O@91rt;Ls`YMXk<78QhN!Bw3Y5yjALFPbHV{YMaZM*qMw^VvMIe!@!Vf_Jy}vBBQZW+M-W#{07)GYBE*D_EJH!E##o5%wQi86q*T5O8YU{;V7<WH8CN(jX5e|-8Tcgy`P$AmZ}?V#_0|0{e0F=uv0UT-XSa2GRwn2Yy5drBF>9w9!QmeLR?kihydrmn+H&DB?ss=}l;ay!=v!m#lWc;{-t-|^;Z}X*jYVyXPc-F1$s4)L@wjkik;lV%7(TJW=tRxmA-X{NE8YZoO$eG8r0>{0T9z8W{n_nUvZ_*b{Bi{6M<ca4s>8^hYD;eHRXV~+;X{}K9-2#xXF|&MAWPxMiF5l}|Nb8<Gv?-$kQaq@%(~yr$2X6c%I4ZZjQj>CH1GVDTK3gMpjG0-rf5y9XUq>HS1YouX_1FMi(@IsUZpxk=Be!E+Z<+LpCC0HeyoY?BVjtVB^i9wGi@rCbFog5b~*U_FY}p}RzP5U;XCmKW<s^Yn9+Qf&8vLyY|{(|^au<mWJ3x9-a&)(V|ygp9+VxYoZM%(#klL}5~Y%lDk%xg@5vnUp@#wn1p{wmToE51&?jX8y@C)j>pF(R>gJg!0>|iw@5@N?(mq@>HEWH^$>||fAhmId6G%md9)z0yB*v}dRo_3vm;c0LdN6SH%2HW*VR=po+ejIjUk>6&CSUsNg5gbC?&@{8YU-fR8WMHsz=bf`v~tSXMYg)R=)oDl_JmuAoNKH(#0xF<GJD?C4W~8wNsO{Y&4-jk$e$TEVo$U*Dq-qQASV4g0t~J(m>9sy+jp$lm&>%6sg=%2@AI;Xf0obOCuUC-A_5`-J3Q<sHlR5YZH?~o{l5=CdU<Tt34G%BEH5wNTV)evY~_DLf^R0<VM;wxu<`!iV2~FK=qdAyRomhem~G|nDTq~X0!9%oAma_K|0;wOj^<eg^`>sLO{&UsN}0c_a)>h)9&L*d#!WV6yG(9z{`7$z-6$(=^2^Bc;5k>pEpnX`-+<@joayE#dW-Kf&>LmQdfoC^xwgx7rMhx~gwAXyJd<U%1w;B2R&s1a!5$vv92#$j$t*2vMvNU}IElEM-(n~Zg?~ymeT?EjY>LK@MQhbBnvPkIqnmkQC1XiFs8I6po9*-b2Op=NewMz!>Wt5wCMXxnrhv-^!pYv%H<qC$1^krq01*jxu>2W=wKs7L3ybKopfCB$B)lQ$_*yO<fbEU&aL9*R@98z<fO|vdm4wNZj~=-ng6;ps4Pvf?Ow!2K28!rIAqKB7K`=8So*(um6GyV))_YMg6>`A{GCzAvDLhmBHG6l0H#&aq$Bho0K`KQ8!jZ#9t=Fgt?Ny-Klo<rN=R6Bqe;rirG;&{?cHHteDB#Bpn&)n#$o+q1{|%yg_7{qWo2S&KTUBOU?Q8IN^RwC-Bo8SpAF5Y`WLm4rQ6D1_SX(&GJFn?e>vWsWHA+w6Sw^ZOuN^?x!~pJ__hu%`D8kMy(j#Z$BcTabU~B)G7q{^Q?PgFCmki)6$B{K#_$b1FF>e!gv;@qpq3NX<Su*0DE(Ay}3^m9WL|4Jhe-qF*TDu{8<>-jTqFa)4oxIpIJIQnwuFbAQ){%$lDyue1?SJFG!Xgu>T|Q8mFqV!dauJJmEVyJAWLcy79({od*DJXHrdAAqPl~1-d5LO9e(g@R88jzw4E@PW@T3{{hY-e57f~uXNc*;i1Cx2%94-lMh^ec~&$*e0_N*cpniUS%t%;|pvlj4Mrpp~i0g+F#0TQceDn=Wex!ERG1H;kCL<XBB6sNkqVe{ds{genMYP!B+3@umAOL~&jxgID~QMD1u;Y(SyM2(X>DQr_;7Rr;Zvo_?7@DGrV69&^)nefd#>;b-372n*9PK4Q+0}_%TzY7~dP^r*}&)_irL$lVy6UbAOdfq3dfmuFwvp7z2P5PJ1lDOBCGHV;{%*PvA;jH~*f0PuRC}IU`+Uv$!LLk4(*802%z52q4PqlG>8;DorBRTFbR&#EcTt&~}4WfG{CBob68uXmKo_g*_*q`_X(BAyVq^Pc43|NAk!JPc-1%3wlSH2x-{IsB@!-k-*Y0-R^q8-UY%<QSPeWZ3c{(tZr4tC@>RUOs*PrfAIw(0P$wX`P?(vi%gNla(#U{M3JownRFRITvUU7!@8;`WfugY+`9*Mkwacy<*w?;ga`N-Ml&s3A=IpuNkWbebt9VjB&X!VvA%Bs^(pE;0>c)9V{6JX3mYNNUr#XjCk*qse7~@CpY7YE7=UXY;?!s#w!fmz$(w5g$upY<I}?8f#4y@|MKx6)(IlEfL<l&nS>!rdb8EoNQQmcli!v^BH*wt|_;x4h-B*!4b%^*i}A-QU9CeEGYPqnL*o-(Tlh8EBii_e0{BN#Oh?QA=SE7epX3%u==MGpdqUg&lQs+T>O=(LK45N0Z^JqJ=J;oAS50<zfQHPedN5zb0kjur>s2NV_u9eb39BHgw_tCk-niZBXjR>OH3^&A{T*4OzvU>cBX`?4mM^VW?z%Ub-Zr{E@jlHIQN*i{!??R_d}jB&~n7MhcLxV8s(-tE8P?E^wiq88G;6o&hP_Cxc^4QqW!98{uXu)kXQGOyhBw448=#Vu5T5QA(X>95m((QO~)!|q;-zLOJYiY$5`;3RDsxLLE^U&!A08Xbze@*H@}qNp<M=@!SkEhFXP2TO_hL&W{6sTUP%E_C96k&gjV7z^=(kA4wT0`wH$^Z7WqP;mA0Hi?thfLD-fn}+cxB>;9S{Pr^WIp9rHkd_1CJufY=ESn-AZ0xQZ-<1a@P>6d^V(T1k}np5o<e)Jlm;8QJ{5poBCUs-J)$AN>g*=bTYAvWOb|f_Ix)n4`D`Z1l+5_c}#5e{iMFUiCjqjM<*m9zM=M`S5GL-bFTuhKy;JH<PPL2`i@CAJh=9l~y?Mz1-pIaoo*f*_$%4m*-n=5oJ<m))3sMn)V$(BlQ+pY%V;Y1+6jctob92U~I#7TTdxHZx69|4>1`Mx?JS8z56;)wyqd}g7WbNOy9vuMZuQl(J=MrXkd)Qa(DWsjuIFL38O^~<eYeWG6+abRZ4GnJqHp7q@6917NU$y<SRI)-WeHqhz<d#e_UlJ3Fiq(m38^G#qSP{`-@upX!r~8XiwZ34iZ$aYXNi-hl|eQ*TrJB?XxED<h(kDDkpJH@o%!z^_k7$69n#~j*~3;Qs>Ti0}b|qs#fPM`LP%vFK4>w{_EB~EkHqivesCR*~lF9=ZF5Ez@QLBPacQkH7iib33OBRA#2#=xt6I)S${B61HjqOM5XBulj&#c9))xf9xH;`EXekjIXc<^^|Gjz!t3V3*afz*1OOd6n{80x>a&(uZFV|#do0*{yL!y5$2XNVNbA9>JwF>Ckb8!UU{cp-d@9=G^Omxo+?Q%VFCm_<AnQz4GrW;Ve`>HrfuO2MWISPHJ^nI5mNy=Z!*}9d(aR?x^j&e#W_Xf=5_iKQEbzNc?;c$U*wjh>6XUpr4{Mz*QAQnrVXnJ%6U2%c;A+!p_pui(C$d<LvX;jD=x=A(bjcJGE+(QGrm<+2*Rn`BSc`UT6qkox{V`gx1*!=uTKhWq%nP|}f@?Q-ol}wU-d@z^K4>4>RLMK<8E8WnkUm{UPCj&;nBW$xi0%%bX*|SGGtGyz6{BDv4(CTCv1^HCN&L~=q+3u~3l9i7`Xt0zA(k5AF!?EGKbtY`ke-l7W!IuCoJh_M_+npuxgESN$eu!rl2pLoGdk;ea=~uP1dgy^1T&n3(%|iA?vIbS=>vGknY{KyPB^O8n+OV@4%cJ&%FtKDFN^%kCSWu-bleq4_zGXgG!SmpxML%}i{X5B@6}oJ=x>vRSWV$wj*E*JB5Ah4P!aC!DQG4{RZpe)5H*Ce&oGWQi0zS5AFmaPki-yyI9iBQSCB%EbwdXm3bJ^Hsl2&*EM*}el-DrzyrRuQWNb!uU<4KHw(vkBmJLed4$*Y`+f0C~$5&qN?puDAC$0go%Kr0v1{;Xdi!1KGd~aEi+Sk=0Szdo+kVQh=k1>0<S|dQ6Eok6P6uYHbUWHL2EcH2g!ZWCXs7VLq&aBkL;PD&2FIUrp$cpGMcK&&VHpcL)UX<V&(`(A(6IfXVzqF6XM1>m5@rlY%Szk#dDi-T$9A@=V&=tGIj1}?>dACt#L7TD~dbM}+nTn*Keq9=(uY*kEqqn`<RRE~oVgigixKc5{wQTrU3Pgy{O%J-t(r~C-LjbM2hl;YQ!Z(k+@SLs=g{#RoI))xs2F(5)gft$s#J+-Q*vBdrj3PL_#!PL{3mwbbFme3qm-Kn$sGvOo1^y4s$)4OWE!rL*jJw87p^P>QDSx!Rsf;~aRst!+J>7X|1p26KjvYQ8E^XUNwd00-(g^wfw!~%2!KB03yC|D6Ii0@Qrg4FqyTuR#8wTB0UErw#R5*Zs9oUG%J4M1A-1Ta6pI7R2^lvRT7tFETzeoh($?#(eXg8uR!oBO9gN^;Ll5eU1sD?@-c_;GOh3V=U`xLifH>2!?#Qg;=_9S<b2i(G<GsxW%V>_S~eS0}SaQ@l~HCxK$RMsGey#RXwr#=Be#F{k%tUN5;JVr3#1!V37?AoG}kQ|x1zbkE=lUrDpV_3;%7|QVz<XW$g>uc_#A7foY_KqqwnCwpS1`0n|;V(5`Z_B{v*ZKlVZ#%G*Q-Jtj(m|ENzr(+VC(&@MXwudvfC|m2A9R}~(7z8p*+0URxqwOXMyhe;_4gO=>hwYLwkk4)SkZ%hR^W-;7_M6^=^n>M5sIN5b_!i0DDtJx78*!R{QK&re2qY*8-uy7A6hQ@amz|g)}e=qhOua&yEi`{a#=8d&?GkPXsjqoI@Bx&M+guI6YldUJ46ghdY1J0BEltvy4f;kbx%ss9Io}=050)i1@kmzJs`2rP}ZdG!Llqq)M>gGg*JRhRH5TB?B+)9VNJNL(#g4*I}#>Bwr7!`iga?SsEH~o&qqZRAVtFVA$@xSIw3m|zC!ZNon}Q>QQG2<RUFWN6V~X#yTsO}OI2s)TP2_ar{DHu>_hbgf^4tdCKj@QwEuH7;7Kb~qTHInk$Z>@(aHKNPKHRnhWG}w^$Uz&WUtgu*6ZzZGyXy1&9?V2z7kUhY|&Tnl;*$S+VXSgb%^=<<ZK&{`<{J3MYP_FsrF>N7+$Wz-wm7fnLi)^MNVOv#XY*F*icKlLI)G|xgV_#PU|@|)5`-E<-t{a1vx_rR=6K-;H^S>`|}3VNB+=C8BjQ>I^P9?ctr}C2J6LPPhb?ZHbGWfKq}Qt{AXr&ta-247LMD4RX}1@nZk4K?@FM+cAUo5*r=!7xfWOykSjfyIXq&yM;EE!f^#UE)guB#&BSz<$q)FKA77R{Zt->O+W}F9<F(pBK|T=fDJijG9u+MIbaEQ2?}3twLkOj3I*MLbw$0d0nu*6Bs$x2ZA0uCT!>J)bV%>6oK|%TlnnyzVv`oK1xku~Cc!LYGMPm;{5!4Rb4;l8N9FRDT-{>;fZ^L!B=Bg1vKJH&gE>RH1N)5m;dE;DuaPeJlH36qq(=hGBwK4|?9rYa)btMb)ZlAtiw6Ue&wX@`@vqwWVlJWGx!QWKfKH6+Rw11kC&mO*l1$lor%u?S<3o-7s9_3pZcW4>Mi0P+jk+#@rtw;5QFF(KJ^UNP!so0nd=4HsZUa78EQP=EoWd{1vv#XRG(dMm3BD>>ibQFWL8X}A_2vS$BuOVJh^e11nhy*LywdwK-Fi57u1r1NIWj`E<n|}aCwMN7AeY2R=k8A$66mevL%S^ZI1QyN>stX8K4<+qj&pISHL6@;n7S3+uok`(_xP>}|A9q&mR$xgS%ZgTY;@Jt3zDY`h_5=NiM?Dj;c&nS~saNW>;lp}}N){P(Sira+8;XtNO~y(qMJ#WZ!|iz^W%mO)5UPc@=oKgA%%nf9-JDq?Jxyx|>xt81Gzqfu#hu8Ew?Qew_u+0ouo@)dUSHkbB|TpcJA=M0%TN+tsT{@3&Fk<tfiZqRC1+9XZXXkT>np_b>oQ+gi42N=0E?16wE9IL#Bj)SR?5Cw(TC;n0)iLB-bV{mtHV4&CJB!1|7}qu5<m#pg#Q&j<|fdcqo^_lXU8@;<a+)WY9vUMgS-BJA}VNO74h?x`Oty#9WoR;55-(bSlw<akwc6g<DB}Gi7Iia;Djls2LX)~pG{SOZf_n_^l`m;n*ecFo5=K*5~gJ;D|+L7b0R3$>{MS<q~5)wKNLGtH-{#734FG~ccliyQRkX~ldppNQ?~G>zvNKE5Iw&YaP<%TAWFux-ZBIbS~)SPe!c09*_e}inLBjF`i2qd_{Td5uih~wx|NC|j#jv&g?r9@uhVBF^NOU+7KUOijrbgDlI)c^ll_RxBf9U3SD*l!{s%+z2iFy(nY;5J{bCJiA4uNsLzI=hLpnN#DGOnMHKh|2c?=hHZSue{9&<A275_^?Rlhz+Bv9>YIaMW_jRV}3C<k5zi97C*F?#-!)XsIB_l4%uo{t?Wg>jAt>mPM899x)Tb~_QU+f7Yd?b4uBv~@<#zlKIclXfCX3+%o`^L!I3WHPG%7VtD%%N31(1iV;#^-&TY9I5KueMz^;5veX{FfPdS;0^yz@>OvylMnTg0aDT|FZ+y3)u$cvV)}=#8=RR6WLSK|ypT<I;(xx^JBKkq%K3yudy(MLad+}K63`$kw`ALx#-QUfm%%Kfu2uV3G=OtU?<kVLvnfI9$xC18zD5b+-_;$zO>dS_taY)Ar{qMKb+2>d+t@~aJ69K-l|Z6E;Sw6j4l;bW5@+J|mHhrRJ{Ye+OaIazLv#Dj$h)mI@jzKOe$2K)D+?fhZ6CH}K=VmY-H0UC=7!lDE^9-Xq3QExb{FRx-Xln|j^w);)k>t+U4;oikYqkvl=z(7-@2V1bf@zZth$@7tR@Z_R6i@S$B*mmfiIrPIl%wjkt4HfHfsjiRfPivV)A`t4v)roT;3eAW)16wj&%uMr>mX0?tkfI&r;+0ofZcQ-Ao}cTVbL*MPMr0X%FQ@2MI8xcqA@M>w~l>udrNpX|_hLRl{3luQd`reQqroK~A{kXy)3JKW4J@nCYe4LCgq1wGH=u0dqhQ8p<=PDL#mqmIi>viUBHP4OCLU(Mn?0*Aj4K;{LySe_}U&;ySBDu$5YX2PK)#R4!8P{UuZXLU5+=No*P>{d3so`x;BK#OYzQFSsWSF0t*?aw0j-{gv(d{~TG+VU0O1$HEZ~9`?0Ie{`5c2ZMwYh=kx30T4A2zCyjdTy*ZV?%O%!^yp{mx=8|$nHDZ?v!DFIOWrrIlo4AJt|E9Nl{5V_TsTweP<Bb{rLS~Wevlgn%|uZABNs`rdEdnG`}hHnn+_$VpWv|peeHq6#b^=v-V^7xwlv^Ve;@Vk6|_s;JNitcCulbz$?0PyOm#d(Lax4kn9z$T10WHdBPj^|1i?c}{OOau*S<wWFj$<kYWF$Q+o#9K5L(Ua$2JvBP9SLp(oK$<kn9M4RNk(G=?Fh_NI<^1oJr_Ry?v2C7y{txCqVmu<<q^lWFx&0uMqS|4|RMm9;Ycgo(@s*@4{<Db1#|PJ%IZAwVN)%>!fA?s!r%M+|QqRfkjRhAJ{?W|7OUM2%a|$3JQgp+Dk&S7-Eah2T24Ke5yQ*Kb855ea*S4T$mi5oJ0xDPHJ)C<0N2*<e}>A12>akisV3n;2>K$<cx-iT_D6`Y{>g%)Wv>W4D8vM44%dqC%XwzsjT_IpqI@xRNLGH8o+Uy8G=`^2#g;Us>tU6u=vdY0JHU~gmC*`KuJZYJ>$1?!30Fis)3wfns=?YtniB4?K!Y{sZM086sQ9_myz5%2FmL};nI2>`qt7sJiUVSR3yAb@7{|_Ihh9iiBB09iu##RjJ}iHPW-4S0ntYOR&H7H-;}mGy-{oF+915vzk>dlv8ZL$${XV|LeEzai~C%b8%e-GS-E6fkbRg%R&hRp72|#6Kr;JB9|1S-gO)}XrCaTxu`qt!dle_+b$=uq2>2zA$4x2@@TikK>7h(Xs5_;kR$FTBJ0xbtL0J5kS|z*az;J$F)GSu})tX?NzjTp*%kMB!K&>E7Na;yK|1MUFSe^38&dKn|mReMHMbBB5FX1H55gY9c3&UInPfONvTdGu5LQHZXLVdrz^AlTiZ&a>~_v<YMeXrAc6&<B~N1!@ovrY5~B!KD4)T`RP@#Em{PnNU-9JE(TE<VvTllDV5$W*Uo6V;Xt8&jaiPdy-3<GaS3;AM1`pJJmDBX|42a;?4ePckQBTNH_#Ype+3knzf^06YMCNkKaZ`R74aSx@A!Z>Es4TEc#UL?>(L6}UIRqjjy=B|by=W(0&R9>5cvg{a;a<H0eb7uQTCi{xsDG101Dnvud!>dMHmwjAB|e}>CxULP(l=SNslMqf=^1BJlI3)smi-P9`37lX(0($JBEhMfLMBKDj1b73&%<4tJRPBqq_@QaycBhY8{_`}yh`ydZ`+9qYE4x#9koEx}I!=);6%F2xI4NsML+S7<QkrnRSj#L+4l?~lWY}>@)WeDe*EU0gylFaw~_gj+RC~6N)d4q#G#{0)dxW_8my;)$R7=g0p%{EUbt{q)tDL%l#EFd86N0LfYx431eAPKT<&`lE|T$4@r5v*IbtJO;4V$1$}D<{!;b{MmJXOL1cuR5`|tY-zD<MqD3U&n0U-d+#=iU*;{oGwmMJ}tA_6b#9#*Z5$D4BNjRdxg`~J@yHF5nVO3^}dY|w@O3y)B7h%q;|6fr-Ss1<sq$JEJiKA;KRU?fIM12m)U3uF~h9M?T6S1{M<zWd7GOQqKhJ@ZNYjI--XRos$u)jq{Pp!iF0^`_<`odQXH8h<Z9OJmC_Yh94UW?zbnXw&d^Hh%O?)tU?2fS$9*Ijl;|QZhnk4JcCbHDIAi0|yw+^-H|Jt4C8IdHnYoFY_3`wQ^R2>V4%_A2!L|o-|8>1szD14{5QotO!4YgfM>E(>VFa1M#CdZXN(#1t^uAG$NjAvS&wQbCWmN<2O!j&h4UjB1tf+v}4q3Ac$+<~90wu@y;uF3j>`o>M9`dxO!N*3^Z&|iLO_^K}D)FXc0CBm+{|BLe=8P-#<Ar-98S8ZkE&L~o<2eSiZr(}7b7V5+)1~abx87%;`-V&dJgf|tCk-`~Z#WnN3}N-c*)=D|wm@OpxfS_nqUf50<&I@(Q|ii|!}gcjjG5~^G0%N=r#=p;Wz6|d`>Uz}o?u|}s~@fGkh0tqOo=x7>maK4j1exOx;!La%}wbd+nN#RyBX_x$&b4rb-adS?>K?}cGx}UF|Ox6pqdSj@*5h)WrJlXiN1Hm<x`9mEn3|C#0Yk}5kwT%&)+|8Uz>aP&oFMs#i%ms%WCC^b!N0`U0n&2)JFLbulTQqW%bm{u)@Q3U!^YE6K)#*85`e^kN=jECPgL1?mE~G=Lf&#V;m+krbo@nKyx3KdKO{XlQCz1NfOnSK%+d8<$W%FQm9Y4z247v_SJsPPP<Y^vgJ?Zx;hwT*jo-|#h*a?6}m;^&iMFOZWQ&2)zQR~yn>XzvSPb>i1)gv?QfaEsGGgm{tIAz%)jaWJ>HMgj(Ea_7fgildZwF^LF2J>tXwH6gVeUP-AI6QV5D(8^`Wt;?*<^VY7y5f#m}#7Df<C#6^h{rEmeHwveEMvqk)kL1n53f`-5v3XKX8g0&JACy%v)I5+C(l51a8mFg>O!<ntr!ZK3d+L;uin4Kse`hfF{T@DHsxVq5VhaS}^_T~Jr_76ohDus)_lP|wEEAGK_FNGE*Z=j2hzPQ&tTkFwk0tmQ4)k-YAsG8VZSWJd`B=qAhco*v;I>5bX3(pmCP81tYWL~PRi6y((1S7MVsbnky108xTAFqpBW>3mP?<udUlNhE8I!`A(Pk4gmD*#GOxi*y&F0D#U*I0LUUtF}z1e*w+4AeTFz6F|+WJb(DQr6uzDGKuM6NnBv7JXJ)4Zi6&;&%`V$h_-9g9!WQ1<YdTsz?g^G-(<)`Pfy)PHo*6yxl-s{$?-wEUnKK4<V3SMuL0}z-aX}d=h%d|2gIv67wexyM%vpD)h7MWr)3s8xR&5k_!#F<-FO=z=lW(4y`cMK$5c0Koku{&m;KW*GUq-#@B^MH_OjF5DeXDbwyH>BJi&rb2q~FKC0kgstvc|i*W7d=iFY`&v2fetmRw{Hp9L#pT0P~b)&#k5u8z+Q2Os@^)gcS5=XPS`@V+kN-**r2&SWIY#z}0Z;T?KwN>O-^NT?+H28YIbHb!s;zgKj*<YkSTU0<Joprc(;^s7<sqc8PGv8g(v*I7o7GwU5Na83wZO`M#S_T{;s2TshLl;!hZjAOXyc50v!oZ4MMc*p}{)AOn~)S}oi5D(Ub9QZfAz%6rSU2(ML+yqde^k;hM`ZK-Aw1rvi_FOlf(0%tT?!kZ7aJW3Zcgo}jz`KWe0x%j5GDXPX(@d-&9(=js(Esy!M0kQa^Rl#hcT|JKlyo6XtnBE(%vZV-xsqyn3uSbUX>2Fb?6q-47|NE)8CYGmYuks$%4PE~3BI2z@)KB<8p@~Zp@QseAnRx%If>A+aNHo&MKH)>D;kfXdE|T<?}#1<TBSA42AwS6;NY3|Cw<kxwBN`=1D^#V*a2X&D+Kan`esT-&u`UWB*9mhLIs%=PG(Dxzc|w&BB~fu^aJ{E(t0mL0?h(Rz<z;FeM6{fNxeyGgj^^=;+7^g!k3C8AeYK&2K;1DM7i4VpqT=X^kG9w>Q6gFZn8*<j@BY9(nk@-y~T(2WlQUv;}|*r5xbqJd^8BdT|(w?bDzcgSX&YDKU02~+W+oA-wx%lrR|qTfXq$xwzxIuoz|PKQx<sT?(85FXC2X9G7K-#R~6rx^`JxJ20qCwxIa?pADN__vgSY2a%Yu$9x~(eah#ip0p1z=K2S7(<XbZ%t!G};>nVza5RmS>yE6+4=c|qG)&ZeqaU)9I@8l0?xvRc&63{g>^J$zB>Vm}3lTP<>uB)*USV8yXZuhl}bfD={eO_3U;PUx74Z8{t(`OsBbUabrQXkPY*3{$ZS4T4aO+|2icFahTFRV&%bPI_0^QNzmb-M^qE&~mmxK8p~S;wa`BQi|)MO!x7@mX&PZ9um(29bcUlBUbj*;AYXwWI!d)43?F)H*LuvuHuAOtS`>^dBwl_?t0g`D%9sl1!Xi->91oyUD|-aO3a<zvW9BZ^QWKGPgSuBAX`BX9=RyMU!taO+3n-Gjv%{uJQZcQ>X5d<C>Yt>E3XK73e%*_d*uM7AGhLzAB@UOTw~yII{MIIJ}<V>g3j2Q7W3HVXx+TqhXMKXfXHzx35eQry3+zPhpBFOrOe9Pdn2ftn4Rh=pfak8oB^*?BrXv4yT<PWcAfW>k**YmFF$kcz0!@Hp1|O3<H99=o7Mgy}GD*+xj(*qa3JgW-Ly;J-Y~-19HY(Gbqk}cTf(J?gI#!nM~>aK&TZ&MR{$OC4bHHDyQ)r<&@kO&?zrY6pV@m+t*oUp#dGcD!9csq1|^+r<?G=$>pybt_p5)su$6nFl%=_2~bj+Ha%|{(3DgB-W<!K`_u$DDMQuWGjVJ{tZZmZZRQ1Q779AR9WV)8O|_)H9^Lh=c7v9Y9C<NvdcDA4iFkenO}hv{5Nvfk6Imqdq~aZg9BuBLIx|6dO+gNSEfk_Px*C1aeA%tJg6L!aBm4Vjk|0wMLQ-$`5a4=~brCJSHPZRD8^AHC`H&M1pwb{%I}PMUd0y;2z?~;jIPY`E7eToE28M0~)<;x~8$ouEvzoY}h_{mgW3>2EE2Z2o)fXug0?%|aHgK_Eto$5UM+@VY&#Wo19q8=Z{02VEvQS=5#>97GdSy(9_WUqHH<G|iPxXP@m(_i=xqIv5q}uQvhT!Ps^OTaYMy*S{rl(Qz2gkB5y8WeeoR$%lN_U7`+!QP`f+f9t)e>hLPofvtg}iNigu}Mv4Z;4m0T+5`y3`wIhQ&fwa}x~nSeLgH*?qR8{@pQ~R0^eQa14C$V(^Yk7LmnQQCmcP9*M$}=EC`jq!bFntWOwC7X$rws()295KJ4+qWOC=ei9|v5!P@X5S`5aqv_~$3OTWAZMhZR#T#*?%GoAvd(BtPw(Rd1iE*&4hlVNJILKHe3lV@%)wl+i_C%CRi<OL&x{|>8TS~FKco+tbWxX`qD_;&0RL3g+)ZNW(=J+A@^(BmlJ+cCf;&&!>;%>yxcY%7nK1qDVvW#MhISyEnPZDg1MmkirakRlAk?Hqui)BbN8mSj2CtDUP5u}MnK<^EnX?p_>OcU`j7bX>wBO~q?`_~bn6MaqtTZNx7)EU~v#sm1U9xm}`oL1<o{Tn@=wt3`5{R+?pOWnw1OP5nk6>m6<DE>}@sM{%Gb6bjSC;W2&h6W}Un>_lko{SBcB2XA~qk>2wHUX<m%f%~Eg{B__+x7c?=>f@SMA3wDFQi;L_anhzQ=AE$qx*g7u2+<M$u7|HpH$U>23nTb{h?&XKHEo)YBvt(NYV@_u%8eZ6fI(kyNWF#yJwb;&4DI@>h!oEf$SFJ`%6l!H=6-ec{2kom}0#83*3=sF8$dNPP#2dbj7#$w+EJm{Cjor%Lf0O%ulPnJF$Cf#1=1dLR82A-u{I629CDP<2I*@2VclqPC);9_^bBi$;$C3?(`zRD2kw?7NgucnkU3;k*>D5<=OShJ)t)##aQ1~CZXc?qe)DBv3A|@%g`@WPl0r^?_O%!<?)95{TuX#8tM?KS2KZ_v{ze(d82Mw{J=T*O=xY4NKkFp$trO?2mjNz7y|x0jC7qXu+s~}UV&Kq^HRU%g)s<DlC>Eoqx%;Ycw=2g=LI~GJ0X6i6ZXpx0eg;#g|5bUn^#wakJ)sGr7j6i6rn&<izC&P8Eo;b2z`b9U_jEZ6{Fr-c7B_|9$^tDMDD@2!-o{GOyX2-M~Y{V<jsOVzh!R|*w3V-Q_JdgOp~#01F_|R=jDoJhGW0XQ&~E8PH1SdLF40nt9K23+%+MMe4;E^BHF3NxA?i#QcI>DSaKO91Wbl}LnWr$N<6_ijsVY_PeO}HAiLk7)HbDlj{T~xp*i7hetOJ73s?j|-xIB-^-UG3Ru;!1_#bJY29KQyu|S_}!H{%G)uzU~pLJX^V62&W==U$^{E=v7rq3!!Xy>@RAFW@p!N1(4Rpo_4vM}WlHhC$_Zh7twH9z7eZ$^&j7Zh&jI~ku~wYtuN&+A(=5VGRE`*4PG2+6sSB2)NK%+52tscjRHAO=0|u%hw&JJXu=K5KW6^d>7;C_@GyaJZetUcBh@);sG3u&Z$(JZh2o(p)MV<V~kE>s3kZCP9IgD`*~uRSpnq@<||A4k7N*l*UZgduq3pHpq_2LR14YxJ}a-Sgbdkf1gmCdh*!I?_PL5YTQ(JeUWeBjIyPj;HA3c9>mg0WSX8S;A&7a>i(})FHGzg!3s1h7WK#sjmnekO8gLSdoqMd*?v)IjE?xsn|n~weOmgz1NBA?BBFAHVe9ND(Z)LkGJQ`w7A&)!N$nkJ2wwUs4Z=y5d^>f=;WS)GryG<RR$^?3l!jT{p13i5?m~A2b<F>-v(HI4Sh()mNvfG4*7}d-)E2jdlbI<SsCJx$=t(2*2&ZU|PJ#Z6nbw<zMzudTnpE`tbe-7)^l8UzFQg(5Q?MAsWIEg*?g*`S>&cF=v)jB;q{z7zGBSr~w!fpzdLEp8h2C9(`1YIBC09gV4ABObKJ)#ji?}Evd~?1Oxv}oAn$7c8Q<Yn%6M7ozIf6k))!|Xw6$8Ky<zDuM3B3|IF5T+@NHey~ea8gwUK=>!gu??+>tiG_I$vY#ft`AFh2ooTCoLlo9=dFCdIw+)rBUdV54{noVuhF9r&Hvt>`i`LiNwy|rR4hZP;x$h6N?GcZZdaZXp}nbD5dp4{t4TH_ScQf8$7S6EBI%?x2d2k%b`7pP;br1_vX>E32r!!F&}z;Z)DRUCi9v;SzaV_pLY(N!*+)t_ehQ<RPA<H42wD5D2T<zHT3CYlrV<T82C3PHft3yx-SZ&vSb;w-dV(r&qa#yZbmJ*)V(9<aOT&^L3C(M+K<l>L`VTe9K+J9o3_zdWcN#Ms1ty)mc<lGM0w(6;)Oag{g{%s#^ix;eUJ%ObC!e#LwQs}^mi?6@~#ARZdE7!)Ax|N4Ogzq*m{7olc0tP@uM<v=RIc9(GJgU7A)j(|HQfJDaceR2TcqGU)F?~4qOn0k~I%3y$8o2N!tnwPd#9uzLDk0At@`d^5YK<Zz8Xq#Fuk;qgaoq`T0NNBv8X4(a!q0NmH2PQQG(iyuBdGNDSv}rd|%H-dzavI|~ltmobU70a0NiOU&%xNo*T~qc5cbBClw%ri32Uevkq!c6?uDmf?puPLePvuxuM#(duLg?js7g!L(TVDYPIj#2i`Rp3J4o&#t_{5CBe;6T*oA66!FebIlHf*1_!61Vk3z8sgtF8Ly1M$b{KMh5EeKuRW58_jiPN9|zFlsA?E@|3pU4;1-XG=vcA!((H*YD#RHqNzsqfsEKw3`TozWtRg?&g1|9V-+}-1Q#!tOj~$j2bmTx6q4B0c@t}X9F=f@6nEeRenv0lwaiVwn9Xk5AfqY79iF~6PM?ohz2;uNYt7ZQ*J35N??3S@*?o;#ZTGU8pyZfK2^BhT?d8TmSVAG^z9djB{!&~Y;6;sR`F^*5yd{IZ=vck?-h3VE3igVCIv7|NSI{Dngy=4+0btC0{&M4^8kZ9EC_k`tj+09C5<TFTS2CbG;qGA!Rp~nY=w?>e#+^V$c{xeOG7v-N&9bO@_BigSDbCHo-yn%)rY1YYCn{(UBK(vwQ#Wrejai&7N|B~9@JzP_~J+4a|s~>V|SaQ80tknsNbt(-SAD?ug*JY(H>yxyLmzCRrbDo$@l;uXw*kMesOE`;=KTIlm&~qYpV+mYW2Bv+?FUs*Bd;Aa-fMD;5FV)@e;*NV#dK_abf7oQJITA%Dh#DXmU|pcHT`1Ig$qW!2xi6UtU-f*A`B4ry1<aQ@MKCWhn1_zqu{H5OK{tErVx-{J`xvLL#pG}(aDS2NQYxSVvUca+OSx3cI2+H0{iIKsstTZ@sb2n%yBR!V<*=9ezm{Y>tU8LozQ2d#df8SKsy;ybHtoe&_Cpzv;GWQ#l~yeVq(@qZILv2Z81!RD%DG!H8n2`pNLmUNb6d_c6Iq>aw%Z(W=Umhwv-wptS*SK&R3w!Q%!G9wf}Acs<*A`pOD3P=Vcxt*?JJ>4QzN)vzoA@Z1xUUxJ<UzgW<@87L5Pm)w)XHT+x5_yT1eq}gHU(z3p{zyIZz5WAWJU|i4f9aG+yk}6Z5MpRUsvI#BKH^d>RLkYlEi5;x_Rt=><HRQ#4Q+{p=PJ0d6{f>D>Q=rjU0yi&lHUj#)WIFBF>`{R=LGWYyi$(fTU}cI<o9JchebLa<L5hFzQ@xE?3pYjM9JuqQ!z$Ejy!6IY#c3y1^*b0$O*7KWs)THLAh9ow?ER6xwLNKmaNA%TRhOW5?&zo*oC={zY`Xf80B_mMdeks2vJo2)o=!hvxuk(dRqI0FAirfFJI8}Vfgn0Xx(`})h<n15o19G0r>j1(+5CzE^tsgZSRqw@djRBN#jk^>k?pLz4j&3b&yw=!kP1OP>~WN2tY8FG%?>_n1nAA5IpfQ&CVw-?!lowYk?5b^xhN03lyIJ*sWQ7J~8+~#Pi?7n=JSB3gQATjQeDcBeVWsTrRn?>NFO+!R!gqhR7vr&7>feO^)3z&I>1j9(EzoYSF^=NLBQ;u7}@mE#qp!|#7Pz|gbVgBC9GD<H+Twt3e2#@bJU^^=1O!Z;DXiNFck47_CF-y^p6^vpF_~)g-1>;YFGG8DEW!Jq%Ia64(fvAkdAip>Hc(o`PFX%{O@WEDh@vx#vbWeDZ%G>N)Q=nA`r)zJwAd#jql*$bR@0Hy!&&ANm;_Ll83lX{;L7!rAjmy!2feRpERB~N5s4i0-kyvm7ge>46F}5$iVen|dOlw5YKZ7{oM+E{(^lw`(<gpQAxZ1+Yk)#d6-g^RVJ%Dn`0ZHc94QYjIg+#oDq)2gCyK!<&XG18vNntrb*TE}}ALL?W79{h(^S%zx3Hy)eg^aiGsIFhJHkHCrFWWmBJ%vqGt#tULC>$k8NO~?d9_W^LP#-`Eqi^(FbaGFsC!O23aGjfzEDu=xA9!Fq$Z9n3Rs*XcI$K^%C>58Cs>yJ%)PrqaKcEb>-H8BRp&x8ihe>imn%Z9liHx+6^0od4kg3W083-7!U4VVSYFMs3hIQ`>2%CGbrxRjF*PJOCvq@pr=gM)eR8iHV>tY{X3Co{|td+?3WTMYnqgU_*@Kp6hz@n{P*tA;!)xrF1%c4*V9n$}K{|&C*Yy>752at~_NxrQpx;iq<TNeZq)^m3*%tLge5zl941xk{XN3-qu=*`MWbuuW`;9me0r-=z2?S-MjqV3zwfylL!COqCwn<~ekbRYggVVDH`Bbs0Ghy-)j(dX3=KLlDMfhkw}E{L&PaC`y&lA&`;d^`K%STGek2g$P5WF;Modeg=Tc3{cGLK`bWOkUg~cIFsl{9R(6+OuBQyyz>d{81Hu@MBxQ&!l!^miv`<Z&YHEFVpYYbno?9%|B!j0v~{kUwVMeF<?H*w>(xAQ~Z(SL-~$r*f38>67qKWoK=sV;mt3YqWnmfMQW`EARv%*$WF1b(Q_S{qa9AAg7&<+@zl496nqAOOkF-RWitv!&Lyl+x^zq_d<ccd2Rcyas?3gQBtus8&i0YTUj$o{y|PDaH%}eKTM&<8Pv*4AObX=m1S<i`8G1O?z0Gt6(KAH}S~BZYcrJ_kL25jbuuR8|k!g3VS|&nbl8hH}!SLT60o0>+@>naf^+L7x?+O@(sdClRg~fv*>1GgE)3-1?r;xE+x^Msh00000ptC1B)$~DR00G6x0rJQTWtDjRvBYQl0ssI200dcD'

def load_template_document():
    """Memuat FORMAT ST.docx dari disk, atau otomatis merekonstruksi dari memori jika file belum diunggah ke GitHub."""
    for candidate in ("FORMAT ST.docx", "format st.docx", "FORMAT_ST.docx", "template_st.docx"):
        if os.path.exists(candidate):
            return docx.Document(candidate)
    raw_docx = lzma.decompress(base64.b85decode(EMBEDDED_TEMPLATE_B85))
    try:
        with open("FORMAT ST.docx", "wb") as f_out:
            f_out.write(raw_docx)
    except Exception:
        pass
    return docx.Document(io.BytesIO(raw_docx))

def generate_docx_buffer(data_gen):
    doc = load_template_document()

    for p in doc.paragraphs:
        if "Segala biaya yang timbul" in p.text:
            p.paragraph_format.first_line_indent = Inches(0.5)
        if "Surat Tugas ini disusun" in p.text:
            p.paragraph_format.first_line_indent = Inches(0.5)

    # 1. Paragraf Konsideran
    teks_pt2 = ""
    if data_gen["pt_2"]:
        teks_pt2 = f" serta surat permohonan {data_gen['pt_2']} nomor {data_gen['no_surat_2']} tanggal {data_gen['tgl_surat_2']} hal {data_gen['permohonan_2']}"

    for p in doc.paragraphs:
        if "Sehubungan dengan surat permohonan" in p.text:
            p.text = f"\tSehubungan dengan surat permohonan {data_gen['pt_1']} nomor {data_gen['no_surat_1']} tanggal {data_gen['tgl_surat_1']} hal {data_gen['permohonan_1']}{teks_pt2},  kami menugasi :"
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.name = "Arial"
            break

    # 2. Table 0 (Petugas)
    t_petugas = doc.tables[0]
    while len(t_petugas.rows) > 0:
        t_petugas._tbl.remove(t_petugas.rows[0]._tr)

    petugas_list = data_gen["daftar_petugas"]
    num_petugas = len(petugas_list)
    for idx_p, pet in enumerate(petugas_list, start=1):
        lbl_nama = f"{idx_p}. Nama" if num_petugas > 1 else "Nama"
        items = [
            (lbl_nama, f": {pet['nama']}"),
            ("Pangkat / Gol", f": {pet['pangkat']}"),
            ("Jabatan", f": {pet['jabatan']}")
        ]
        for lbl, val in items:
            row = t_petugas.add_row()
            c0 = row.cells[0]
            c1 = row.cells[1]
            c0.width = Inches(1.13)
            c1.width = Inches(4.38)

            p0 = c0.paragraphs[0]
            p0.paragraph_format.space_before = Pt(0)
            p0.paragraph_format.space_after = Pt(0)
            p0.paragraph_format.line_spacing = 1.05
            r0 = p0.add_run(lbl)
            r0.font.name = "Arial"

            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(0)
            p1.paragraph_format.line_spacing = 1.05
            r1 = p1.add_run(val)
            r1.font.name = "Arial"

    # 3. Paragraf Kegiatan
    teks_keg2 = f" dan {data_gen['kegiatan_2']}" if data_gen["kegiatan_2"] else ""
    for p in doc.paragraphs:
        if "untuk melaksanakan" in p.text:
            p.text = f"untuk melaksanakan {data_gen['kegiatan_1']}{teks_keg2} dan melakukan tugas lain terkait kepabeanan dan cukai pada:"
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.name = "Arial"
            break

    # 4. Table 1 (Hari / Tanggal & Tempat - Pabrik 1 & Pabrik 2 Sama-sama BOLD)
    t_waktu = doc.tables[1]
    c_waktu = t_waktu.rows[0].cells[2]
    c_waktu.text = f"{data_gen['hari']} / {data_gen['tgl_pelaksanaan']}"
    for p in c_waktu.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"

    c_tempat = t_waktu.rows[1].cells[2]
    c_tempat.text = ""
    p_t = c_tempat.paragraphs[0]

    r_pt1 = p_t.add_run(data_gen["pt_1"])
    r_pt1.font.name = "Arial"
    r_pt1.font.bold = True
    r_koma = p_t.add_run(", ")
    r_koma.font.name = "Arial"
    r_alm = p_t.add_run(data_gen["alamat_pabrik_1"])
    r_alm.font.name = "Arial"

    if data_gen["pt_2"] and data_gen["alamat_pabrik_2"]:
        r_serta = p_t.add_run(" serta ")
        r_serta.font.name = "Arial"
        r_pt2 = p_t.add_run(data_gen["pt_2"])
        r_pt2.font.name = "Arial"
        r_pt2.font.bold = True
        r_koma2 = p_t.add_run(", ")
        r_koma2.font.name = "Arial"
        r_alm2 = p_t.add_run(data_gen["alamat_pabrik_2"])
        r_alm2.font.name = "Arial"

    # 5. Table 2 (Tanda Tangan - Pemenggalan Baku Anti-Offside & Rapat ke Tanggal Malang)
    t_ttd = doc.tables[2]
    while len(t_ttd.rows) > 4:
        t_ttd._tbl.remove(t_ttd.rows[2]._tr)

    c_tgl = t_ttd.rows[0].cells[1]
    p_tgl = c_tgl.paragraphs[0]
    p_tgl.paragraph_format.space_before = Pt(0)
    p_tgl.paragraph_format.space_after = Pt(2)
    p_tgl.paragraph_format.line_spacing = 1.05
    for r in p_tgl.runs:
        r.font.name = "Arial"

    c_prefix = t_ttd.rows[1].cells[0]
    c_jab = t_ttd.rows[1].cells[1]

    jab_kk_baku = "Kepala Kantor Pengawasan dan\nPelayanan Bea dan Cukai Tipe\nMadya Cukai Malang"

    if data_gen["tipe_ttd"] == "Kepala Kantor Langsung":
        if data_gen["is_plh"]:
            c_prefix.text = "Plh."
            c_jab.text = jab_kk_baku
        else:
            c_prefix.text = ""
            c_jab.text = jab_kk_baku

        for p in c_prefix.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for r in p.runs:
                r.font.name = "Arial"
        for p in c_jab.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Arial"
    else:
        if data_gen["is_plh"]:
            new_row = t_ttd.add_row()
            t_ttd.rows[1]._tr.addnext(new_row._tr)

            c_prefix.text = "a.n."
            c_jab.text = jab_kk_baku

            c_plh = t_ttd.rows[2].cells[0]
            c_subjab = t_ttd.rows[2].cells[1]
            c_plh.text = "Plh."
            c_subjab.text = "Kepala Subbagian Umum"

            for p in c_prefix.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.font.name = "Arial"
            for p in c_jab.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = "Arial"
            for p in c_plh.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.font.name = "Arial"
            for p in c_subjab.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = "Arial"
        else:
            c_prefix.text = "a.n."
            c_jab.text = f"{jab_kk_baku}\nKepala Subbagian Umum"
            for p in c_prefix.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.font.name = "Arial"
            for p in c_jab.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    r.font.name = "Arial"

    idx_elek = 3 if (data_gen["tipe_ttd"] != "Kepala Kantor Langsung" and data_gen["is_plh"]) else 2
    idx_nama = idx_elek + 1

    c_elek = t_ttd.rows[idx_elek].cells[1]
    c_elek.text = "\n\n\n\n\nDitandatangani secara elektronik"
    for p in c_elek.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.color.rgb = RGBColor(191, 191, 191)

    c_nama = t_ttd.rows[idx_nama].cells[1]
    c_nama.text = data_gen["nama_pejabat_ttd"]
    for p in c_nama.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.bold = False

    # 6. Table 3 (Tembusan)
    t_tembusan = doc.tables[3]
    c_temb = t_tembusan.rows[1].cells[0]
    if data_gen["tipe_ttd"] == "Kepala Kantor Langsung":
        c_temb.text = "1. Kepala Subbagian Umum\n2. Kepala Seksi Pelayanan Kepabeanan dan Cukai II\n3. Kepala Seksi Kepatuhan Internal"
    else:
        c_temb.text = "1. Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang\n2. Kepala Subbagian Umum\n3. Kepala Seksi Pelayanan Kepabeanan dan Cukai II\n4. Kepala Seksi Kepatuhan Internal"
    for p in c_temb.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"

    docx_buffer = io.BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer

# ==========================================
# FUNGSI RENDER PRATINJAU DOKUMEN (HTML 100% PRESISI)
# ==========================================
def render_preview_html(data):
    petugas_rows = ""
    num_p = len(data["daftar_petugas"])
    for idx, p in enumerate(data["daftar_petugas"], start=1):
        lbl = f"{idx}. Nama" if num_p > 1 else "Nama"
        petugas_rows += f"""
        <tr>
            <td style="width: 130px;">{lbl}</td>
            <td>: {p['nama']}</td>
        </tr>
        <tr>
            <td>Pangkat / Gol</td>
            <td>: {p['pangkat']}</td>
        </tr>
        <tr>
            <td>Jabatan</td>
            <td>: {p['jabatan']}</td>
        </tr>
        """
        if idx < num_p:
            petugas_rows += "<tr><td colspan='2' style='height: 8px;'></td></tr>"

    teks_p2 = ""
    if data["pt_2"]:
        teks_p2 = f" serta surat permohonan {data['pt_2']} nomor {data['no_surat_2']} tanggal {data['tgl_surat_2']} hal {data['permohonan_2']}"

    teks_keg2 = f" dan {data['kegiatan_2']}" if data["kegiatan_2"] else ""

    tempat_full = f"<b>{data['pt_1']}</b>, {data['alamat_pabrik_1']}"
    if data["pt_2"] and data["alamat_pabrik_2"]:
        tempat_full += f" serta <b>{data['pt_2']}</b>, {data['alamat_pabrik_2']}"

    jab_kk_html = "Kepala Kantor Pengawasan dan<br>Pelayanan Bea dan Cukai Tipe<br>Madya Cukai Malang"

    if data["tipe_ttd"] == "Kepala Kantor Langsung":
        prefix_1 = "Plh." if data["is_plh"] else ""
        jab_1 = jab_kk_html
        baris_2_html = ""
        temb_txt = "1. Kepala Subbagian Umum<br>2. Kepala Seksi Pelayanan Kepabeanan dan Cukai II<br>3. Kepala Seksi Kepatuhan Internal"
    else:
        prefix_1 = "a.n."
        temb_txt = "1. Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang<br>2. Kepala Subbagian Umum<br>3. Kepala Seksi Pelayanan Kepabeanan dan Cukai II<br>4. Kepala Seksi Kepatuhan Internal"
        if data["is_plh"]:
            jab_1 = jab_kk_html
            baris_2_html = """
            <tr>
              <td style="width: 38px; text-align: right; vertical-align: top; padding-right: 8px; white-space: nowrap;">
                Plh.
              </td>
              <td style="text-align: left; vertical-align: top;">
                Kepala Subbagian Umum
              </td>
            </tr>
            """
        else:
            jab_1 = f"{jab_kk_html}<br>Kepala Subbagian Umum"
            baris_2_html = ""

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{
    background-color: #525659;
    margin: 0;
    padding: 20px 10px;
    font-family: Arial, sans-serif;
  }}
  .paper {{
    background: #ffffff;
    color: #111111;
    max-width: 680px;
    margin: 0 auto;
    padding: 40px 45px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    border-radius: 2px;
    font-size: 13.5px;
    line-height: 1.35;
    box-sizing: border-box;
  }}
  .title {{
    text-align: center;
    font-weight: bold;
    font-size: 15px;
    margin-bottom: 2px;
    letter-spacing: 0.5px;
  }}
  .nomor {{
    text-align: center;
    font-size: 13.5px;
    margin-bottom: 22px;
  }}
  p {{
    text-align: justify;
    margin: 10px 0;
  }}
  .indent {{
    text-indent: 32px;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 6px 0;
    font-size: 13.5px;
  }}
  td {{
    padding: 1.5px 0;
    vertical-align: top;
  }}
  .tembusan {{
    font-size: 11.5px;
    line-height: 1.4;
    margin-top: 25px;
  }}
</style>
</head>
<body>
  <div class="paper">
    <div class="title">SURAT TUGAS</div>
    <div class="nomor">NOMOR [@NomorND]</div>

    <p class="indent">
      Sehubungan dengan surat permohonan {data['pt_1']} nomor {data['no_surat_1']} tanggal {data['tgl_surat_1']} hal {data['permohonan_1']}{teks_p2},  kami menugasi :
    </p>

    <table>
      {petugas_rows}
    </table>

    <p>
      untuk melaksanakan {data['kegiatan_1']}{teks_keg2} dan melakukan tugas lain terkait kepabeanan dan cukai pada:
    </p>

    <table>
      <tr>
        <td style="width: 130px;">hari / tanggal</td>
        <td style="width: 15px;">:</td>
        <td>{data['hari']} / {data['tgl_pelaksanaan']}</td>
      </tr>
      <tr>
        <td>tempat</td>
        <td>:</td>
        <td>{tempat_full}</td>
      </tr>
    </table>

    <p class="indent">
      Segala biaya yang timbul sebagai akibat dilaksanakannya surat ini dibebankan pada DIPA Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang
    </p>

    <p class="indent">
      Surat Tugas ini disusun untuk dilaksanakan dan setelah selesai dilaksanakan, pelaksana segera menyampaikan laporan. Kepada instansi terkait, kami mohon bantuan demi kelancaran pelaksanaan tugas tersebut.
    </p>

    <table style="width: 100%; border-collapse: collapse; margin-top: 25px; margin-bottom: 20px;">
      <tr>
        <td style="width: 44%;"></td>
        <td style="width: 56%;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="width: 38px;"></td>
              <td style="text-align: left; padding-bottom: 3px; line-height: 1.15;">Malang, [@TanggalND]</td>
            </tr>
            <tr>
              <td style="width: 38px; text-align: right; vertical-align: top; padding-right: 8px; white-space: nowrap;">
                {prefix_1}
              </td>
              <td style="text-align: left; vertical-align: top; line-height: 1.25;">
                {jab_1}
              </td>
            </tr>
            {baris_2_html}
            <tr>
              <td></td>
              <td style="text-align: left; vertical-align: top;">
                <div style="margin-top: 75px; color: #bfbfbf; font-size: 11.5px;">Ditandatangani secara elektronik</div>
                <div style="margin-top: 3px; font-weight: normal; color: #000;">{data['nama_pejabat_ttd']}</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>

    <div class="tembusan">
      <b>Tembusan:</b><br>
      {temb_txt}
    </div>
  </div>
</body>
</html>"""
    return html

# ==========================================
# HELPER ANALITIK & VISUALISASI DASHBOARD EKSEKUTIF (100% IN-MEMORY & BEBAS BOCOR HTML/SVG)
# ==========================================
def parse_rupiah_to_int(val_str: str) -> int:
    if not val_str:
        return 0
    digits = re.sub(r'[^\d]', '', str(val_str))
    return int(digits) if digits else 0

def format_rupiah_short(nom: int) -> str:
    if nom >= 1_000_000:
        jt = nom / 1_000_000
        return f"Rp {jt:,.2f} Jt".replace(",", ".")
    return f"Rp {nom:,}".replace(",", ".")

def format_rupiah_full(nom: int) -> str:
    return f"Rp {nom:,}".replace(",", ".")

def compute_perjadin_nominal(p: dict) -> int:
    """Menghitung nominal Rp satu baris Perjadin (mengutamakan kolom G atau fallback sesuai aturan bulan)."""
    val_rp = parse_rupiah_to_int(p.get("nominal", ""))
    if val_rp == 0:
        d_o = p.get("date_obj")
        if p.get("luar_kota"):
            val_rp = 100000 if (d_o and d_o.month == 10) else 150000
        elif p.get("dalam_kota"):
            val_rp = 50000
    return val_rp

def render_copy_tsv_button(tsv_text: str, button_label: str = "📋 Copy Tabel ke Clipboard (1-Klik)"):
    """Komponen tombol 1-Click Copy to Clipboard berbasis HTML/JS ringan untuk laporan Bagian Umum."""
    escaped_json = json.dumps(tsv_text)
    html_copy = f"""
    <html>
    <head><meta charset="utf-8"></head>
    <body style="margin:0;padding:0;background:transparent;font-family:'Plus Jakarta Sans',sans-serif;">
        <button id="copy-btn" onclick="copyTsv()" style="
            width:100%; min-height:44px; border-radius:11px; border:1.5px solid #2563EB;
            border-left:5px solid #F59E0B; background:linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%);
            color:#FFFFFF; font-weight:800; font-size:13.5px; cursor:pointer;
            display:flex; align-items:center; justify-content:center; gap:8px;
            box-shadow:0 4px 12px rgba(37,99,235,0.25); transition:all 0.16s ease;">
            {button_label}
        </button>
        <script>
        function copyTsv() {{
            var text = {escaped_json};
            var btn = document.getElementById("copy-btn");
            if (navigator.clipboard && window.isSecureContext) {{
                navigator.clipboard.writeText(text).then(function() {{
                    btn.innerHTML = "✅ Berhasil Disalin ke Clipboard! Siap Paste di WA / Excel";
                    btn.style.background = "linear-gradient(135deg,#047857 0%,#10B981 100%)";
                    setTimeout(function() {{
                        btn.innerHTML = "{button_label}";
                        btn.style.background = "linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%)";
                    }}, 2500);
                }});
            }} else {{
                var ta = document.createElement("textarea");
                ta.value = text;
                ta.style.position = "fixed";
                ta.style.left = "-9999px";
                document.body.appendChild(ta);
                ta.focus();
                ta.select();
                try {{
                    document.execCommand("copy");
                    btn.innerHTML = "✅ Berhasil Disalin ke Clipboard! Siap Paste di WA / Excel";
                    btn.style.background = "linear-gradient(135deg,#047857 0%,#10B981 100%)";
                    setTimeout(function() {{
                        btn.innerHTML = "{button_label}";
                        btn.style.background = "linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%)";
                    }}, 2500);
                }} catch (e) {{}}
                document.body.removeChild(ta);
            }}
        }}
        </script>
    </body>
    </html>
    """
    components.html(html_copy, height=50, scrolling=False)

def build_svg_smooth_path(pts: list) -> str:
    if not pts:
        return ""
    if len(pts) == 1:
        return f"M {pts[0][0]:.1f},{pts[0][1]:.1f}"
    d = [f"M {pts[0][0]:.1f},{pts[0][1]:.1f}"]
    for i in range(len(pts) - 1):
        x0, y0 = pts[i]
        x1, y1 = pts[i + 1]
        cx1 = x0 + (x1 - x0) * 0.45
        cy1 = y0
        cx2 = x0 + (x1 - x0) * 0.55
        cy2 = y1
        d.append(f"C {cx1:.1f},{cy1:.1f} {cx2:.1f},{cy2:.1f} {x1:.1f},{y1:.1f}")
    return " ".join(d)

def get_dynamic_2026_period_options(riwayat_list: list, perjadin_list: list) -> list:
    """Menyusun opsi bulan Tahun 2026 (mendukung otomatis semua bulan di 2026 yang memiliki data)."""
    months_present = {7, 8, 9, 10, 11, 12}
    for it in riwayat_list:
        dl = parse_indo_date_range(it.get("tgl_pelaksanaan", ""))
        d0 = dl[0] if dl else parse_indo_date(it.get("tgl_surat_1", ""))
        if d0 and d0.year == 2026:
            months_present.add(d0.month)
    for p in perjadin_list:
        do = p.get("date_obj")
        if do and do.year == 2026 and (p["luar_kota"] or p["dalam_kota"] or p["no_st"]):
            months_present.add(do.month)
    ordered_m = sorted(list(months_present))
    opts = ["Semua Bulan (2026)"] + [f"{BULAN_INDONESIA[m]} 2026" for m in ordered_m]
    return opts

def render_dashboard_analytics(
    riwayat_list: list,
    perjadin_list: list,
    rekap_list: list,
    perusahaan_dict: dict,
    period_options_2026: list,
    theme_mode: str
):
    is_dark = ("Dark" in str(theme_mode))
    selected_period = st.session_state.get("active_period_2026", period_options_2026[0])
    if selected_period not in period_options_2026:
        selected_period = period_options_2026[0]

    target_month = None
    target_year = 2026
    for m_num, m_name in BULAN_INDONESIA.items():
        if m_name.lower() in selected_period.lower():
            target_month = m_num
            break

    filtered_st = []
    for it in riwayat_list:
        d_list = parse_indo_date_range(it.get("tgl_pelaksanaan", ""))
        d_first = d_list[0] if d_list else parse_indo_date(it.get("tgl_surat_1", ""))
        if target_month is None or (d_first and d_first.month == target_month):
            filtered_st.append(it)

    filtered_perj = []
    for p in perjadin_list:
        d_obj = p.get("date_obj")
        if target_month is None or (d_obj and d_obj.month == target_month):
            filtered_perj.append(p)

    rekap_map_by_no = {str(r["no"]).strip(): r for r in rekap_list}

    kpi_total_st = len(filtered_st)
    kpi_hari_aktif = sum(
        1 for p in filtered_perj
        if p["luar_kota"] or p["dalam_kota"] or p["no_st"] or p["keterangan"] == "BELUM ADA PERJADIN"
    )
    kpi_luar_kota = sum(1 for p in filtered_perj if p["luar_kota"])
    kpi_dalam_kota = sum(1 for p in filtered_perj if p["dalam_kota"])
    kpi_belum_nost = sum(
        1 for it in filtered_st
        if not (it.get("nomor_st_resmi") or rekap_map_by_no.get(str(it["no_urut"]).strip(), {}).get("nomor_st", ""))
    )
    kpi_belum_id = sum(1 for p in filtered_perj if p["keterangan"] == "BELUM ADA PERJADIN")

    # ---------------------------------------------------------
    # ZONA A: 4 KARTU KPI SUMMARY ROW (BENTO RINGKAS)
    # ---------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4, gap="small")
    with k1:
        st.markdown(clean_html(f"""
        <div class="dash-kpi-card accent-navy">
            <div>
                <div class="dash-kpi-title">📌 Total Surat Tugas</div>
                <div class="dash-kpi-value">{kpi_total_st} ST</div>
                <div class="dash-kpi-sub">ST Berikutnya: <b>#{next_no_preview}</b></div>
            </div>
            <div class="dash-kpi-icon" style="background: rgba(251, 191, 36, 0.2); color: #FBBF24;">📋</div>
        </div>
        """), unsafe_allow_html=True)
    with k2:
        st.markdown(clean_html(f"""
        <div class="dash-kpi-card">
            <div>
                <div class="dash-kpi-title">🔥 Hari Penugasan Aktif</div>
                <div class="dash-kpi-value" style="color: #2563EB;">{kpi_hari_aktif} Hari</div>
                <div class="dash-kpi-sub">🚗 {kpi_luar_kota} Luar Kota • 🏙️ {kpi_dalam_kota} Dalam Kota</div>
            </div>
            <div class="dash-kpi-icon" style="background: rgba(37, 99, 235, 0.14); color: #2563EB;">🗓️</div>
        </div>
        """), unsafe_allow_html=True)
    with k3:
        st.markdown(clean_html(f"""
        <div class="dash-kpi-card">
            <div>
                <div class="dash-kpi-title">⏳ Antrean Nomor ST</div>
                <div class="dash-kpi-value" style="color: #F59E0B;">{kpi_belum_nost} ST</div>
                <div class="dash-kpi-sub">Menunggu Nomor Satu Kemenkeu</div>
            </div>
            <div class="dash-kpi-icon" style="background: rgba(245, 158, 11, 0.15); color: #F59E0B;">🔢</div>
        </div>
        """), unsafe_allow_html=True)
    with k4:
        st.markdown(clean_html(f"""
        <div class="dash-kpi-card">
            <div>
                <div class="dash-kpi-title">⚠️ Antrean ID Perjadin</div>
                <div class="dash-kpi-value" style="color: #8B5CF6;">{kpi_belum_id} Hari</div>
                <div class="dash-kpi-sub">Berstatus BELUM ADA PERJADIN</div>
            </div>
            <div class="dash-kpi-icon" style="background: rgba(139, 92, 246, 0.15); color: #8B5CF6;">⚡</div>
        </div>
        """), unsafe_allow_html=True)

    st.markdown(clean_html("<div style='height: 8px;'></div>"), unsafe_allow_html=True)

    # ---------------------------------------------------------
    # LAYOUT ASIMETRIS 70% (KOLOM KIRI UTAMA) : 30% (PANEL SAMPING KANAN)
    # ---------------------------------------------------------
    col_main, col_side = st.columns([2.35, 1.05], gap="medium")

    with col_main:
        # =====================================================
        # B.1 KARTU GRAFIK TREN PENUGASAN & PERJADIN (BEBAS BUG KODE SVG BOCOR)
        # =====================================================
        with st.container(border=True):
            months_order = [7, 8, 9, 10, 11, 12]
            st_per_m = {m: 0 for m in months_order}
            lk_per_m = {m: 0 for m in months_order}
            dk_per_m = {m: 0 for m in months_order}

            for it in riwayat_list:
                dl = parse_indo_date_range(it.get("tgl_pelaksanaan", ""))
                d0 = dl[0] if dl else parse_indo_date(it.get("tgl_surat_1", ""))
                if d0 and d0.month in st_per_m:
                    st_per_m[d0.month] += 1

            for p in perjadin_list:
                do = p.get("date_obj")
                if do and do.month in lk_per_m:
                    if p["luar_kota"]:
                        lk_per_m[do.month] += 1
                    elif p["dalam_kota"]:
                        dk_per_m[do.month] += 1

            active_months = [7, 8, 9, 10, 11, 12]
            labels_m = [f"{BULAN_INDONESIA[m][:3]} 2026" for m in active_months]
            vals_st = [st_per_m[m] for m in active_months]
            vals_lk = [lk_per_m[m] for m in active_months]
            vals_dk = [dk_per_m[m] for m in active_months]

            max_y = max(vals_st + vals_lk + vals_dk + [10])
            svg_w, svg_h = 680, 210
            pad_l, pad_r, pad_t, pad_b = 36, 26, 34, 34
            plot_w = svg_w - pad_l - pad_r
            plot_h = svg_h - pad_t - pad_b

            def to_xy(vals):
                pts = []
                n = len(vals)
                for idx, v in enumerate(vals):
                    x = pad_l + (idx / max(1, n - 1)) * plot_w
                    y = pad_t + plot_h - (v / max_y) * plot_h
                    pts.append((x, y))
                return pts

            pts_st = to_xy(vals_st)
            pts_lk = to_xy(vals_lk)
            pts_dk = to_xy(vals_dk)

            path_st = build_svg_smooth_path(pts_st)
            path_lk = build_svg_smooth_path(pts_lk)
            path_dk = build_svg_smooth_path(pts_dk)
            area_st = f"{path_st} L {pts_st[-1][0]:.1f},{pad_t + plot_h:.1f} L {pts_st[0][0]:.1f},{pad_t + plot_h:.1f} Z"

            peak_idx = 0
            peak_val = -1
            for idx_v, v_curr in enumerate(vals_st):
                if v_curr >= peak_val and v_curr > 0:
                    peak_val = v_curr
                    peak_idx = idx_v
            px, py = pts_st[peak_idx]

            grid_stroke = "#334155" if is_dark else "#E2E8F0"
            lbl_fill = "#94A3B8" if is_dark else "#64748B"
            callout_bg = "#0F172A" if not is_dark else "#3B82F6"

            grid_lines_svg = ""
            for g_i in range(4):
                gy = pad_t + (g_i / 3) * plot_h
                g_val = int(round(max_y * (1 - g_i / 3)))
                grid_lines_svg += f'<line x1="{pad_l}" y1="{gy:.1f}" x2="{svg_w - pad_r}" y2="{gy:.1f}" stroke="{grid_stroke}" stroke-dasharray="4 4" stroke-width="1"/>'
                grid_lines_svg += f'<text x="{pad_l - 8}" y="{gy + 4:.1f}" text-anchor="end" fill="{lbl_fill}" font-size="10.5" font-weight="600">{g_val}</text>'

            x_labels_svg = ""
            dots_svg = ""
            for idx_m, m_lbl in enumerate(labels_m):
                xx, yy_st = pts_st[idx_m]
                _, yy_lk = pts_lk[idx_m]
                x_labels_svg += f'<text x="{xx:.1f}" y="{svg_h - 8}" text-anchor="middle" fill="{lbl_fill}" font-size="11" font-weight="700">{m_lbl}</text>'
                dots_svg += f'<circle cx="{xx:.1f}" cy="{yy_st:.1f}" r="4" fill="#3B82F6" stroke="#FFFFFF" stroke-width="1.8"><title>{m_lbl}: {vals_st[idx_m]} ST, {vals_lk[idx_m]} Luar Kota, {vals_dk[idx_m]} Dalam Kota</title></circle>'
                dots_svg += f'<circle cx="{xx:.1f}" cy="{yy_lk:.1f}" r="3.2" fill="#F59E0B" stroke="#FFFFFF" stroke-width="1.4"><title>{m_lbl}: {vals_lk[idx_m]} Hari Luar Kota</title></circle>'

            callout_svg = (
                f'<circle cx="{px:.1f}" cy="{py:.1f}" r="7.5" fill="rgba(59, 130, 246, 0.25)"/>'
                f'<circle cx="{px:.1f}" cy="{py:.1f}" r="4.8" fill="#3B82F6" stroke="#FFFFFF" stroke-width="2"/>'
                f'<rect x="{px - 28:.1f}" y="{max(2, py - 34):.1f}" width="56" height="22" rx="6" fill="{callout_bg}"/>'
                f'<text x="{px:.1f}" y="{max(16, py - 19):.1f}" text-anchor="middle" fill="#FFFFFF" font-size="11" font-weight="800">{vals_st[peak_idx]} ST</text>'
            )

            chart_card_html = f"""
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:6px;">
                <div>
                    <span style="font-size:1.03rem; font-weight:800; color:var(--text-main);">📈 Tren Penugasan & Perjadin Bulanan</span>
                    <div style="font-size:0.77rem; color:var(--text-sub);">Perbandingan Total ST Terbit vs Hari Perjadin Luar Kota & Dalam Kota (Jul – Des 2026)</div>
                </div>
                <div style="display:flex; gap:12px; font-size:0.75rem; font-weight:700; color:var(--text-sub);">
                    <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#3B82F6;margin-right:4px;"></span>Total ST</span>
                    <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#F59E0B;margin-right:4px;"></span>Hari Luar Kota</span>
                    <span><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#10B981;margin-right:4px;"></span>Hari Dalam Kota</span>
                </div>
            </div>
            <svg viewBox="0 0 {svg_w} {svg_h}" style="width:100%; height:auto; overflow:visible; display:block;">
                <defs>
                    <linearGradient id="stAreaGrad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stop-color="#3B82F6" stop-opacity="0.32"/>
                        <stop offset="100%" stop-color="#3B82F6" stop-opacity="0.0"/>
                    </linearGradient>
                </defs>
                {grid_lines_svg}
                <path d="{area_st}" fill="url(#stAreaGrad)"/>
                <path d="{path_dk}" fill="none" stroke="#10B981" stroke-width="2.2" stroke-dasharray="5 3" stroke-linecap="round"/>
                <path d="{path_lk}" fill="none" stroke="#F59E0B" stroke-width="2.6" stroke-linecap="round"/>
                <path d="{path_st}" fill="none" stroke="#3B82F6" stroke-width="3.2" stroke-linecap="round"/>
                {dots_svg}
                {callout_svg}
                {x_labels_svg}
            </svg>
            """
            st.markdown(clean_html(chart_card_html), unsafe_allow_html=True)

        # =====================================================
        # B.2 BARIS BAWAH KOLOM KIRI (2 KARTU SEJAJAR 1 : 1)
        # =====================================================
        sub_c1, sub_c2 = st.columns(2, gap="medium")

        # KARTU KIRI: TOP PETUGAS PEMERIKSA
        with sub_c1:
            with st.container(border=True):
                st.markdown(clean_html("""
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <span style="font-size:0.96rem; font-weight:800; color:var(--text-main);">👮 Top Petugas Pemeriksa</span>
                    <span style="font-size:0.73rem; font-weight:700; color:var(--text-sub);">Frekuensi ST</span>
                </div>
                """), unsafe_allow_html=True)

                petugas_counter = Counter()
                petugas_meta = {}
                for it in filtered_st:
                    for p in it.get("daftar_petugas", []):
                        nm = p["nama"].strip()
                        if nm:
                            petugas_counter[nm] += 1
                            if nm not in petugas_meta:
                                petugas_meta[nm] = p.get("pangkat") or p.get("jabatan") or "Pemeriksa Bea & Cukai"

                top_petugas = petugas_counter.most_common(6)
                max_p_cnt = top_petugas[0][1] if top_petugas else 1
                bar_colors = ["#10B981", "#3B82F6", "#8B5CF6", "#F59E0B", "#EC4899", "#06B6D4"]

                if not top_petugas:
                    st.caption("Belum ada data penugasan petugas pada periode ini.")
                else:
                    rows_html = []
                    for idx_u, (nm_u, cnt_u) in enumerate(top_petugas):
                        pct_u = max(12, int(round((cnt_u / max_p_cnt) * 100)))
                        col_u = bar_colors[idx_u % len(bar_colors)]
                        parts_nm = [w[0].upper() for w in nm_u.split() if w][:2]
                        initials = "".join(parts_nm) or "BC"
                        sub_info = petugas_meta.get(nm_u, "Pemeriksa")
                        rows_html.append(f"""
                        <div class="user-row-item">
                            <div style="display:flex; align-items:center; justify-content:space-between; gap:10px;">
                                <div style="display:flex; align-items:center; gap:10px; min-width:0;">
                                    <div class="user-avatar-circle" style="background:{col_u};">{initials}</div>
                                    <div style="min-width:0;">
                                        <div style="font-weight:800; font-size:0.84rem; color:var(--text-main); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{nm_u}</div>
                                        <div style="font-size:0.72rem; color:var(--text-sub); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{sub_info}</div>
                                    </div>
                                </div>
                                <span style="font-size:0.77rem; font-weight:800; color:var(--text-main); flex-shrink:0;">{cnt_u} ST</span>
                            </div>
                            <div class="prog-track">
                                <div class="prog-fill" style="width:{pct_u}%; background:{col_u};"></div>
                            </div>
                        </div>
                        """)
                    st.markdown(clean_html("".join(rows_html)), unsafe_allow_html=True)

        # KARTU TENGAH: DISTRIBUSI WILAYAH & TOP KECAMATAN/PABRIK (BEBAS BUG KODE HTML BOCOR)
        with sub_c2:
            with st.container(border=True):
                st.markdown(clean_html("""
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span style="font-size:0.96rem; font-weight:800; color:var(--text-main);">🗺️ Distribusi Wilayah & Pabrik</span>
                    <span style="font-size:0.73rem; font-weight:700; color:var(--text-sub);">Proporsi ST</span>
                </div>
                <div style="display:flex; justify-content:space-between; font-size:0.68rem; font-weight:700; color:var(--text-sub); margin-bottom:6px; padding:0 2px;">
                    <span>0%</span><span>25%</span><span>50%</span><span>75%</span><span>100%</span>
                </div>
                """), unsafe_allow_html=True)

                kec_counter = Counter()
                pt_counter = Counter()
                wil_counter = Counter()

                for it in filtered_st:
                    pt1_n = it.get("pt_1", "").strip()
                    if pt1_n:
                        pt_counter[pt1_n] += 1
                    lok_w, kec_w = extract_lokasi_kecamatan(it.get("alamat_pabrik_1", ""), pt1_n, perusahaan_dict)
                    if "kota malang" in lok_w.lower():
                        wil_counter["Dalam Kota (Kota Malang)"] += 1
                    else:
                        wil_counter["Luar Kota (Kab. Malang / Batu)"] += 1
                    if kec_w and kec_w != "-":
                        kec_counter[kec_w] += 1

                tot_wil = max(1, sum(wil_counter.values()))
                lk_st_cnt = wil_counter["Luar Kota (Kab. Malang / Batu)"]
                dk_st_cnt = wil_counter["Dalam Kota (Kota Malang)"]
                lk_pct = round((lk_st_cnt / tot_wil) * 100, 1)
                dk_pct = round((dk_st_cnt / tot_wil) * 100, 1)

                top_kecs = kec_counter.most_common(4)
                top_pts = pt_counter.most_common(3)
                city_colors = ["#3B82F6", "#10B981", "#8B5CF6", "#F59E0B", "#EC4899"]

                bars_html = f"""
                <div style="margin-bottom:8px;">
                    <div style="display:flex; justify-content:space-between; font-size:0.76rem; font-weight:700; margin-bottom:3px;">
                        <span style="color:var(--text-main);">🚗 Luar Kota (Kab. Malang)</span>
                        <span style="color:#3B82F6;">{lk_pct}% ({lk_st_cnt} ST)</span>
                    </div>
                    <div class="prog-track"><div class="prog-fill" style="width:{max(6, int(lk_pct))}%; background:#3B82F6;"></div></div>
                </div>
                <div style="margin-bottom:12px;">
                    <div style="display:flex; justify-content:space-between; font-size:0.76rem; font-weight:700; margin-bottom:3px;">
                        <span style="color:var(--text-main);">🏙️ Dalam Kota (Kota Malang)</span>
                        <span style="color:#10B981;">{dk_pct}% ({dk_st_cnt} ST)</span>
                    </div>
                    <div class="prog-track"><div class="prog-fill" style="width:{max(6, int(dk_pct))}%; background:#10B981;"></div></div>
                </div>
                <div style="font-size:0.74rem; font-weight:800; color:var(--text-sub); text-transform:uppercase; margin:10px 0 6px 0;">📍 Top Kecamatan & Pabrik Tujuan</div>
                """

                list_items_html = []
                for idx_k, (k_nm, k_cnt) in enumerate(top_kecs):
                    k_pct = round((k_cnt / tot_wil) * 100, 1)
                    c_dot = city_colors[idx_k % len(city_colors)]
                    list_items_html.append(f"""
                    <div style="display:flex; justify-content:space-between; align-items:center; padding:5px 0; border-bottom:1px dashed var(--border-card); font-size:0.79rem;">
                        <span style="display:flex; align-items:center; gap:8px; color:var(--text-main); font-weight:600;">
                            <span style="width:9px; height:9px; border-radius:50%; border:2.5px solid {c_dot}; display:inline-block;"></span>
                            Kec. {k_nm}
                        </span>
                        <span style="font-weight:800; color:var(--text-main);">{k_cnt} ST ({k_pct}%)</span>
                    </div>
                    """)

                for idx_pt, (pt_nm, pt_cnt) in enumerate(top_pts):
                    short_pt = pt_nm if len(pt_nm) <= 28 else pt_nm[:25] + "..."
                    list_items_html.append(f"""
                    <div style="display:flex; justify-content:space-between; align-items:center; padding:5px 0; font-size:0.77rem;">
                        <span style="color:var(--text-sub); font-weight:600;">🏭 {short_pt}</span>
                        <span style="font-weight:800; color:#F59E0B;">{pt_cnt} ST</span>
                    </div>
                    """)

                st.markdown(clean_html(bars_html + "".join(list_items_html)), unsafe_allow_html=True)

    # =========================================================
    # KOLOM KANAN PANEL SAMPING (30% - FINANSIAL PER BULAN, KALENDER, ACTION CENTER DEEP-LINK)
    # =========================================================
    with col_side:
        # -----------------------------------------------------
        # C.1 PILIHAN BULAN (TAHUN 2026) & KARTU FINANSIAL BIRU GRADIEN + TABEL REKAP BULANAN
        # -----------------------------------------------------
        st.selectbox(
            "📅 Pilih Bulan Pantauan Finansial & Kalender (2026):",
            options=period_options_2026,
            key="fin_period_select",
            on_change=on_fin_period_change
        )

        nom_total = 0
        nom_cair = 0
        nom_belum = 0
        hari_cair = 0
        hari_belum_cair = 0
        spark_vals = []

        for p in filtered_perj:
            if p["luar_kota"] or p["dalam_kota"] or p["no_st"]:
                val_rp = compute_perjadin_nominal(p)
                nom_total += val_rp
                spark_vals.append((val_rp, p["sudah_masuk"]))
                if p["sudah_masuk"]:
                    nom_cair += val_rp
                    hari_cair += 1
                else:
                    nom_belum += val_rp
                    hari_belum_cair += 1

        total_hari_fin = hari_cair + hari_belum_cair
        spark_tail = spark_vals[-24:] if spark_vals else [(100000, False)] * 12
        max_sp = max([x[0] for x in spark_tail] + [150000])
        bars_div = "".join([
            f'<div class="fin-bar-col {"active" if is_cair else ""}" style="height:{max(22, int((v / max_sp) * 100))}%;"></div>'
            for v, is_cair in spark_tail
        ])

        st.markdown(clean_html(f"""
        <div class="fin-gradient-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.84rem; font-weight:800; letter-spacing:0.3px; opacity:0.95;">💳 Pantauan Finansial Perjadin</span>
                <span style="font-size:0.72rem; background:rgba(255,255,255,0.2); padding:3px 10px; border-radius:999px; font-weight:800;">{selected_period}</span>
            </div>
            <div style="margin-top:9px; display:flex; justify-content:space-between; align-items:baseline;">
                <div>
                    <div style="font-size:0.73rem; opacity:0.88; font-weight:600;">Total Estimasi Perjadin Aktif</div>
                    <div style="font-size:1.55rem; font-weight:800; letter-spacing:-0.3px; margin-top:2px;">{format_rupiah_full(nom_total)}</div>
                </div>
                <span style="font-size:0.78rem; font-weight:800; background:rgba(15,23,42,0.32); padding:4px 9px; border-radius:8px;">🔥 {total_hari_fin} Hari Aktif</span>
            </div>
            <div class="fin-bars-wrap">
                {bars_div}
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:9px; font-size:0.77rem;">
                <div style="background:rgba(15, 23, 42, 0.3); padding:8px 10px; border-radius:10px;">
                    <div style="opacity:0.88; font-size:0.71rem;">✅ Sudah Cair ({hari_cair} Hari)</div>
                    <div style="font-weight:800; font-size:0.88rem; margin-top:2px;">{format_rupiah_full(nom_cair)}</div>
                </div>
                <div style="background:rgba(15, 23, 42, 0.3); padding:8px 10px; border-radius:10px;">
                    <div style="opacity:0.88; font-size:0.71rem;">⏳ Menunggu ({hari_belum_cair} Hari)</div>
                    <div style="font-weight:800; font-size:0.88rem; margin-top:2px;">{format_rupiah_full(nom_belum)}</div>
                </div>
            </div>
        </div>
        """), unsafe_allow_html=True)

        with st.expander("📊 Lihat Tabel Rekapitulasi Finansial Per Bulan (2026)", expanded=False):
            monthly_fin_rows = []
            grand_hari = 0
            grand_cair = 0
            grand_belum = 0
            grand_tot = 0
            for m_num in range(1, 13):
                m_perj = [
                    p for p in perjadin_list
                    if p.get("date_obj") and p["date_obj"].year == 2026 and p["date_obj"].month == m_num
                    and (p["luar_kota"] or p["dalam_kota"] or p["no_st"])
                ]
                if not m_perj and m_num < 7:
                    continue
                h_akt = len(m_perj)
                rp_c = sum(compute_perjadin_nominal(p) for p in m_perj if p["sudah_masuk"])
                rp_b = sum(compute_perjadin_nominal(p) for p in m_perj if not p["sudah_masuk"])
                rp_t = rp_c + rp_b
                if h_akt > 0 or m_num in (7, 8, 9, 10):
                    grand_hari += h_akt
                    grand_cair += rp_c
                    grand_belum += rp_b
                    grand_tot += rp_t
                    monthly_fin_rows.append({
                        "Bulan": f"{BULAN_INDONESIA[m_num]} 2026",
                        "Hari Aktif": f"{h_akt} Hari",
                        "Sudah Cair (Rp)": format_rupiah_full(rp_c),
                        "Menunggu (Rp)": format_rupiah_full(rp_b),
                        "Total (Rp)": format_rupiah_full(rp_t)
                    })
            monthly_fin_rows.append({
                "Bulan": "TOTAL 2026",
                "Hari Aktif": f"{grand_hari} Hari",
                "Sudah Cair (Rp)": format_rupiah_full(grand_cair),
                "Menunggu (Rp)": format_rupiah_full(grand_belum),
                "Total (Rp)": format_rupiah_full(grand_tot)
            })
            st.dataframe(monthly_fin_rows, use_container_width=True, hide_index=True)

        # -----------------------------------------------------
        # C.2 KARTU KALENDER MINI PENUGASAN BULAN INI
        # -----------------------------------------------------
        with st.container(border=True):
            cal_month = target_month if target_month is not None else 10
            cal_month_name = f"{BULAN_INDONESIA[cal_month]} {target_year}"

            day_status_map = {}
            for p in perjadin_list:
                do = p.get("date_obj")
                if do and do.year == target_year and do.month == cal_month:
                    if p["keterangan"] == "BELUM ADA PERJADIN":
                        day_status_map[do.day] = ("belum-id", f"{p['tanggal_str']}: BELUM ADA ID PERJADIN ({p['no_st'] or 'Tanpa No.ST'})")
                    elif p["luar_kota"]:
                        day_status_map[do.day] = ("luar-kota", f"{p['tanggal_str']}: Penugasan Luar Kota ({p['no_st'] or 'Ada ST'})")
                    elif p["dalam_kota"]:
                        day_status_map[do.day] = ("dalam-kota", f"{p['tanggal_str']}: Penugasan Dalam Kota ({p['no_st'] or 'Ada ST'})")

            cal_matrix = calendar.monthcalendar(target_year, cal_month)
            cal_cells_html = []
            for week in cal_matrix:
                for d_num in week:
                    if d_num == 0:
                        cal_cells_html.append('<div class="cal-day-box empty"></div>')
                    else:
                        cls_d, tip_d = day_status_map.get(d_num, ("", f"{d_num} {cal_month_name}"))
                        cal_cells_html.append(f'<div class="cal-day-box {cls_d}" title="{tip_d}">{d_num:02d}</div>')

            st.markdown(clean_html(f"""
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-size:0.94rem; font-weight:800; color:var(--text-main);">🗓️ Kalender Penugasan</span>
                <span style="font-size:0.75rem; font-weight:800; color:#3B82F6;">{cal_month_name}</span>
            </div>
            <div class="cal-grid-header">
                <span>S</span><span>S</span><span>R</span><span>K</span><span>J</span><span>S</span><span>M</span>
            </div>
            <div class="cal-grid-days">
                {"".join(cal_cells_html)}
            </div>
            <div style="display:flex; flex-wrap:wrap; gap:10px; margin-top:10px; font-size:0.71rem; font-weight:700; color:var(--text-sub);">
                <span><span style="display:inline-block;width:9px;height:9px;border-radius:3px;background:#F59E0B;margin-right:4px;"></span>Luar Kota</span>
                <span><span style="display:inline-block;width:9px;height:9px;border-radius:3px;background:#10B981;margin-right:4px;"></span>Dalam Kota</span>
                <span><span style="display:inline-block;width:9px;height:9px;border-radius:3px;background:#8B5CF6;margin-right:4px;"></span>Belum ID Perjadin</span>
            </div>
            """), unsafe_allow_html=True)

        # -----------------------------------------------------
        # C.3 KARTU ACTION CENTER INTERAKTIF (1-KLIK AUTO-JUMP & AUTO-SELECT)
        # -----------------------------------------------------
        with st.container(border=True):
            st.markdown(clean_html("""
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-size:0.94rem; font-weight:800; color:var(--text-main);">⚡ Action Center Interaktif</span>
                <span style="font-size:0.72rem; font-weight:700; color:#F59E0B;">Klik untuk Eksekusi</span>
            </div>
            """), unsafe_allow_html=True)

            pending_nost_items = [
                it for it in filtered_st
                if not (it.get("nomor_st_resmi") or rekap_map_by_no.get(str(it["no_urut"]).strip(), {}).get("nomor_st", ""))
            ][:4]
            pending_id_items = [
                p for p in filtered_perj
                if p["keterangan"] == "BELUM ADA PERJADIN"
            ][:5]

            if not pending_nost_items and not pending_id_items:
                st.markdown(clean_html( clean_html('<div class="info-pill-ok">🎉 <b>Semua Beres!</b> Tidak ada tunggakan Nomor ST maupun ID Perjadin pada periode ini.</div>')), unsafe_allow_html=True)
            else:
                if pending_nost_items:
                    st.caption(f"🔢 **Antrean Nomor ST Satu Kemenkeu ({len(pending_nost_items)} ditampilkan):**")
                    for it_p in pending_nost_items:
                        pt_short = it_p["pt_1"] if len(it_p["pt_1"]) <= 20 else it_p["pt_1"][:18] + ".."
                        st.button(
                            f"🔢 Isi No.ST #{it_p['no_urut']} — {pt_short} ({it_p['tgl_pelaksanaan']})",
                            key=f"ac_nost_{it_p['no_urut']}",
                            on_click=jump_to_nost_action,
                            args=(str(it_p["no_urut"]),),
                            use_container_width=True
                        )

                if pending_id_items:
                    st.caption(f"🚗 **Antrean Hari Belum Ada ID Perjadin ({len(pending_id_items)} ditampilkan):**")
                    for p_id in pending_id_items:
                        no_st_raw = p_id["no_st"] if p_id["no_st"] else "Tanpa No.ST"
                        no_st_short = no_st_raw.split("/")[0] if "/" in no_st_raw else (no_st_raw[:14] if len(no_st_raw) > 14 else no_st_raw)
                        st.button(
                            f"⚡ Isi ID: {p_id['tanggal_str']} ({no_st_short})",
                            key=f"ac_perj_{p_id['row_num']}",
                            on_click=jump_to_perj_id_action,
                            args=(int(p_id["row_num"]),),
                            use_container_width=True
                        )

# ==========================================
# SIAPKAN DAFTAR PERIODE DINAMIS 2026
# ==========================================
dynamic_periods_2026 = get_dynamic_2026_period_options(riwayat_list_global, perjadin_list_global)
if st.session_state["active_period_2026"] not in dynamic_periods_2026:
    st.session_state["active_period_2026"] = dynamic_periods_2026[0]
    st.session_state["sb_period_select"] = dynamic_periods_2026[0]
    st.session_state["fin_period_select"] = dynamic_periods_2026[0]

# ==========================================
# SIDEBAR EXECUTIVE (4 MENU ALUR KERJA, TEMA, PERIODE & STATUS ANTREAN)
# ==========================================
with st.sidebar:
    st.markdown(clean_html("""
    <div style="padding: 6px 4px 10px 4px; border-bottom: 1px solid #1E293B; margin-bottom: 10px;">
        <div style="font-size: 0.72rem; font-weight: 800; color: #FBBF24; letter-spacing: 0.8px; text-transform: uppercase;">🏛️ KPPBC TMC MALANG</div>
        <div style="font-size: 1.02rem; font-weight: 800; color: #FFFFFF; margin-top: 2px;">Otomasi ST & Perjadin</div>
        <div style="font-size: 0.74rem; color: #94A3B8; margin-top: 2px;">Arsitektur 3-Sheet Terpadu</div>
    </div>
    """), unsafe_allow_html=True)

    st.radio(
        "📌 ALUR KERJA UTAMA (4 MENU):",
        options=MENU_OPTIONS,
        key="nav_sidebar_radio",
        on_change=on_sidebar_nav_change
    )

    st.radio(
        "🎨 TEMA TAMPILAN:",
        options=THEME_OPTIONS,
        horizontal=True,
        key="sidebar_theme_radio",
        on_change=on_theme_change
    )

    st.selectbox(
        "📅 FILTER BULAN (TAHUN 2026):",
        options=dynamic_periods_2026,
        key="sb_period_select",
        on_change=on_sb_period_change
    )

    st.markdown(clean_html(f"""
    <div class="sb-queue-box">
        <div style="font-weight:800; color:#FBBF24; font-size:0.75rem; text-transform:uppercase; margin-bottom:5px;">🔔 Status Antrean Aktif</div>
        <div style="display:flex; justify-content:space-between; margin-bottom:3px;">
            <span>• Belum No. ST:</span>
            <b style="color:#FDE68A;">{jml_st_belum_nomor} ST</b>
        </div>
        <div style="display:flex; justify-content:space-between; margin-bottom:3px;">
            <span>• Belum ID Perjadin:</span>
            <b style="color:#C4B5FD;">{jml_belum_perjadin} Hari</b>
        </div>
        <div style="display:flex; justify-content:space-between;">
            <span>• Menunggu Cair:</span>
            <b style="color:#93C5FD;">{jml_belum_cair} Hari</b>
        </div>
    </div>
    """), unsafe_allow_html=True)

    sb_c1, sb_c2 = st.columns(2, gap="small")
    with sb_c1:
        if st.button("🔄 Refresh", use_container_width=True, key="sb_btn_refresh"):
            st.cache_data.clear()
            st.toast("Data berhasil disegarkan dari Google Sheets!", icon="✅")
            st.rerun()
    with sb_c2:
        if st.button("🧹 Reset", use_container_width=True, key="sb_btn_reset"):
            reset_form_input(list_pegawai)
            st.toast("Form input telah direset!", icon="🧹")
            st.rerun()

# ==========================================
# ULTRA-CLEAN TOP HEADER BAR (1 BARIS RAMPING & BEBAS DUPLIKASI METRIK)
# ==========================================
st.markdown(clean_html(f"""
<div class="bc-ultraclean-header">
    <div>
        <div class="bc-ultraclean-title">🏛️ KPPBC TMC Malang — Sistem Otomasi ST & Monitoring Perjadin</div>
        <div class="bc-ultraclean-sub">Arsitektur Terpadu 3-Sheet (Database ST • Rekap Bagian Umum • Data Perjadin)</div>
    </div>
    <div class="bc-online-badge">
        <span class="bc-online-dot"></span>
        <span>Online • Next ST <b>#{next_no_preview}</b></span>
    </div>
</div>
"""), unsafe_allow_html=True)

# Quick Navigation & Utility Bar (Tersinkronisasi 2 Arah dengan Sidebar, Ramah HP & PC)
qb_left, qb_right = st.columns([2.1, 1.05], gap="small")
with qb_left:
    st.selectbox(
        "📱 Pindah Cepat Menu Utama:",
        options=MENU_OPTIONS,
        key="nav_top_select",
        on_change=on_top_nav_change,
        label_visibility="collapsed"
    )
with qb_right:
    qa_b1, qa_b2 = st.columns(2, gap="small")
    with qa_b1:
        if st.button("🔄 Refresh Sheets", use_container_width=True, key="btn_top_refresh"):
            st.cache_data.clear()
            st.toast("Sinkronisasi ulang Google Sheets selesai!", icon="✅")
            st.rerun()
    with qa_b2:
        if st.button("🧹 Reset Form", use_container_width=True, key="btn_top_reset"):
            reset_form_input(list_pegawai)
            st.toast("Form siap untuk pengisian baru!", icon="✨")
            st.rerun()

# ==========================================
# ROUTING 4 MENU ALUR KERJA UTAMA (TANPA NESTED TABS YANG MEMBINGUNGKAN)
# ==========================================
active_menu = st.session_state["active_main_menu"]

# ------------------------------------------------------------------
# MENU 1: 📊 1. DASHBOARD & ANALITIK
# ------------------------------------------------------------------
if active_menu == MENU_OPTIONS[0]:
    render_dashboard_analytics(
        riwayat_list=riwayat_list_global,
        perjadin_list=perjadin_list_global,
        rekap_list=rekap_list_global,
        perusahaan_dict=perusahaan_data,
        period_options_2026=dynamic_periods_2026,
        theme_mode=st.session_state["app_theme_mode"]
    )

# ------------------------------------------------------------------
# MENU 2: 📝 2. BUAT SURAT TUGAS BARU
# ------------------------------------------------------------------
elif active_menu == MENU_OPTIONS[1]:
    # ==========================================
    # TAB 1: FORM INPUT SURAT TUGAS BARU (OTOMATIS MASUK KE 3 SHEET)
    # ==========================================
    # -------------------------------------------------------------
    # KARTU 1: DATA SURAT PEMOHON & LOKASI PABRIK (SMART FILTER + QUICK PRESET)
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown(clean_html('<div class="section-badge">1️⃣ Data Surat Pemohon & Lokasi Pabrik</div>'), unsafe_allow_html=True)

        st.markdown("**🏢 Pabrik 1 (Pemohon Utama)**")
        pt_1 = st.selectbox(
            "Nama Perusahaan / Pabrik 1*",
            options=[""] + list_perusahaan,
            index=0,
            key="sel_pt1",
            help="Ketik sebagian nama PT atau pabrik untuk mencari dengan cepat."
        )

        c_p1_1, c_p1_2 = st.columns([1.3, 1])
        with c_p1_1:
            no_surat_1 = st.text_input("Nomor Surat Pemohon 1*", placeholder="Contoh: 63/IX/BBSA/2026", key="in_nosurat1")
        with c_p1_2:
            tgl_surat_1 = st.date_input("Tanggal Surat 1*", value=datetime.date.today(), key="in_tglsurat1")

        # Quick Preset Permohonan 1 (Hemat Ketik di HP)
        st.selectbox(
            "⚡ Quick Preset Permohonan 1 (Pilih untuk Isi Otomatis atau Ketik Manual di Bawah):",
            options=PRESET_PERMOHONAN_OPTIONS,
            index=0,
            key="preset_permohonan_1",
            on_change=on_preset_permohonan_1_change
        )

        permohonan_1 = st.text_area(
            "Perihal / Permohonan 1*",
            key="in_permohonan_1",
            placeholder="Contoh: Permohonan Pelayanan Pelekatan Tanda Pengaman...",
            on_change=on_permohonan_1_change,
            height=76
        )

        kegiatan_1 = st.text_input(
            "Kegiatan Baku 1 (Otomatis dari Permohonan, Bisa Diedit & Jadi PERIHAL Rekap Umum)*",
            key="in_kegiatan_1",
            placeholder="Pelayanan Pelekatan Tanda Pengaman..."
        )

        # Smart Filter Dropdown Titik Lokasi Pabrik 1 (Hanya tampilkan lokasi yang tidak kosong)
        if pt_1 and pt_1 in perusahaan_data:
            valid_locs_1 = [idx for idx in range(1, 13) if perusahaan_data[pt_1].get(idx, "").strip()]
            if not valid_locs_1:
                valid_locs_1 = [1]
        else:
            valid_locs_1 = [1]

        def fmt_lokasi_1(loc_num):
            if not pt_1 or pt_1 not in perusahaan_data:
                return f"📍 Lokasi {loc_num} — (Pilih Pabrik terlebih dahulu)"
            alm = perusahaan_data[pt_1].get(loc_num, "").strip()
            if alm:
                short_alm = alm if len(alm) <= 72 else alm[:69] + "..."
                return f"📍 Lokasi {loc_num} — {short_alm}"
            return f"⚪ Lokasi {loc_num} — (Kosong)"

        pilihan_lokasi_1 = st.selectbox(
            f"Pilihan Titik Lokasi Pabrik 1 ({len(valid_locs_1)} titik tersedia)*",
            options=valid_locs_1,
            index=0,
            format_func=fmt_lokasi_1,
            key="sel_lok1"
        )

        alamat_pabrik_1 = ""
        lokasi_kab_1, kec_1 = "Kabupaten Malang", "-"
        if pt_1 and pt_1 in perusahaan_data:
            alamat_pabrik_1 = perusahaan_data[pt_1].get(pilihan_lokasi_1, "").strip()
            lokasi_kab_1, kec_1 = extract_lokasi_kecamatan(alamat_pabrik_1, pt_1, perusahaan_data)
            if alamat_pabrik_1:
                kat_perjadin = "🚗 LUAR KOTA (Rp 150.000)" if "kota malang" not in lokasi_kab_1.lower() else "🏙️ DALAM KOTA (Rp 50.000)"
                st.markdown(
                    clean_html(f'<div class="info-pill">📍 <b>Alamat Lengkap Pabrik 1 (Titik {pilihan_lokasi_1}):</b><br>{alamat_pabrik_1}<br>'
                    f'🗺️ <b>Deteksi Otomatis Rekap & Perjadin:</b> Lokasi = <b>{lokasi_kab_1}</b> | Kecamatan = <b>{kec_1}</b> | Kategori = <b>{kat_perjadin}</b></div>'), unsafe_allow_html=True)
            else:
                st.markdown(clean_html(f'<div class="info-pill-warn">⚠️ Semua titik lokasi pada <b>{pt_1}</b> masih kosong di database DATA NPPBKC.</div>'), unsafe_allow_html=True)

        tambah_pt2 = st.checkbox("➕ Tambah Pabrik Kedua (Jika 1 Surat Tugas mencakup 2 Pabrik)", value=False, key="chk_pt2")

        pt_2, no_surat_2, tgl_surat_2, permohonan_2, kegiatan_2, alamat_pabrik_2 = "", "", None, "", "", ""
        pilihan_lokasi_2 = 1
        if tambah_pt2:
            st.divider()
            st.markdown("**🏢 Pabrik 2 (Pemohon Kedua)**")
            pt_2 = st.selectbox("Nama Perusahaan / Pabrik 2*", options=[""] + list_perusahaan, index=0, key="sel_pt2")
            c_p2_1, c_p2_2 = st.columns([1.3, 1])
            with c_p2_1:
                no_surat_2 = st.text_input("Nomor Surat Pemohon 2*", placeholder="Nomor surat pabrik kedua", key="in_nosurat2")
            with c_p2_2:
                tgl_surat_2 = st.date_input("Tanggal Surat 2*", value=datetime.date.today(), key="in_tglsurat2")

            st.selectbox(
                "⚡ Quick Preset Permohonan 2 (Opsional):",
                options=PRESET_PERMOHONAN_OPTIONS,
                index=0,
                key="preset_permohonan_2",
                on_change=on_preset_permohonan_2_change
            )

            permohonan_2 = st.text_area(
                "Perihal / Permohonan 2*",
                key="in_permohonan_2",
                placeholder="Permohonan pabrik kedua...",
                on_change=on_permohonan_2_change,
                height=76
            )
            kegiatan_2 = st.text_input(
                "Kegiatan Baku 2 (Otomatis dari Permohonan 2, Bisa Diedit)*",
                key="in_kegiatan_2"
            )

            # Smart Filter Dropdown Titik Lokasi Pabrik 2
            if pt_2 and pt_2 in perusahaan_data:
                valid_locs_2 = [idx for idx in range(1, 13) if perusahaan_data[pt_2].get(idx, "").strip()]
                if not valid_locs_2:
                    valid_locs_2 = [1]
            else:
                valid_locs_2 = [1]

            def fmt_lokasi_2(loc_num):
                if not pt_2 or pt_2 not in perusahaan_data:
                    return f"📍 Lokasi {loc_num} — (Pilih Pabrik 2 terlebih dahulu)"
                alm = perusahaan_data[pt_2].get(loc_num, "").strip()
                if alm:
                    short_alm = alm if len(alm) <= 72 else alm[:69] + "..."
                    return f"📍 Lokasi {loc_num} — {short_alm}"
                return f"⚪ Lokasi {loc_num} — (Kosong)"

            pilihan_lokasi_2 = st.selectbox(
                f"Pilihan Titik Lokasi Pabrik 2 ({len(valid_locs_2)} titik tersedia)*",
                options=valid_locs_2,
                index=0,
                format_func=fmt_lokasi_2,
                key="sel_lok2"
            )
            if pt_2 and pt_2 in perusahaan_data:
                alamat_pabrik_2 = perusahaan_data[pt_2].get(pilihan_lokasi_2, "").strip()
                if alamat_pabrik_2:
                    st.markdown(clean_html(f'<div class="info-pill">📍 <b>Alamat Lengkap Pabrik 2 (Titik {pilihan_lokasi_2}):</b><br>{alamat_pabrik_2}</div>'), unsafe_allow_html=True)
                else:
                    st.markdown(clean_html(f'<div class="info-pill-warn">⚠️ Titik Lokasi {pilihan_lokasi_2} pada <b>{pt_2}</b> kosong di database.</div>'), unsafe_allow_html=True)

    # -------------------------------------------------------------
    # KARTU 2: WAKTU PELAKSANAAN (MENDUKUNG 1 HARI MAUPUN RENTANG MULTI-HARI) & NOMOR ST
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown(clean_html('<div class="section-badge">2️⃣ Waktu Pelaksanaan (1 Hari / Multi-Hari) & Nomor ST</div>'), unsafe_allow_html=True)
        st.caption("💡 **Tips Cepat HP:** Jika kegiatan **1 hari**, cukup pilih **Tanggal Mulai** saja (Tanggal Selesai otomatis mengikuti). Jika **lebih dari 1 hari**, ubah **Tanggal Selesai** sesuai akhir kegiatan.")

        if "in_tgl_pelaks_mulai" not in st.session_state:
            st.session_state["in_tgl_pelaks_mulai"] = datetime.date.today()
        if "in_tgl_pelaks_akhir" not in st.session_state:
            st.session_state["in_tgl_pelaks_akhir"] = st.session_state["in_tgl_pelaks_mulai"]

        # Jaga agar tanggal akhir tidak pernah lebih kecil dari tanggal mulai
        if st.session_state["in_tgl_pelaks_akhir"] < st.session_state["in_tgl_pelaks_mulai"]:
            st.session_state["in_tgl_pelaks_akhir"] = st.session_state["in_tgl_pelaks_mulai"]

        c_wkt1, c_wkt2, c_wkt3 = st.columns([1, 1, 1.35])
        with c_wkt1:
            tgl_pelaksanaan_mulai = st.date_input(
                "Tanggal Mulai Kegiatan*",
                key="in_tgl_pelaks_mulai",
                on_change=on_tgl_mulai_change
            )
        with c_wkt2:
            tgl_pelaksanaan_akhir = st.date_input(
                "Tanggal Akhir Kegiatan*",
                min_value=tgl_pelaksanaan_mulai,
                key="in_tgl_pelaks_akhir"
            )
        with c_wkt3:
            no_st_opsional = st.text_input(
                "No. ST Satu Kemenkeu (Opsional):",
                placeholder="Kosongkan / ketik misal: 2013",
                key="in_nost_opsional",
                help="Karena Nomor ST baru didapat dari Satu Kemenkeu setelah TTE, kamu boleh mengosongkan kolom ini sekarang dan mengisinya nanti lewat Menu 3 (Tindak Lanjut) dengan 1 klik."
            )

        tgl_pelaksanaan = tgl_pelaksanaan_mulai
        nama_hari_id, tgl_indo_str, durasi_hari = format_hari_tanggal_pelaksanaan(
            tgl_pelaksanaan_mulai,
            tgl_pelaksanaan_akhir
        )
        tgl_st_indo_str = tgl_indo_str
        cur_tipe_ttd_k2 = st.session_state.get("in_tipe_ttd", "a.n. Kepala Kantor (Kepala Subbagian Umum)")
        pola_kode_st_k2 = "ST-xx/KBC.1201/2026" if "kepala kantor langsung" in cur_tipe_ttd_k2.lower() else "ST-xx/KBC.1201/KBC.120101/2026"
        preview_no_st_fmt = (
            format_nomor_st_resmi(no_st_opsional, tgl_pelaksanaan_mulai.year, cur_tipe_ttd_k2, auto_adjust_unit=True)
            if no_st_opsional
            else f"( Pola Aktif: {pola_kode_st_k2} — Bisa diisi nanti di Menu 3 )"
        )

        badge_durasi = (
            f"<span style='background:#DBEAFE; color:#1E3A8A; padding:2px 8px; border-radius:6px; font-weight:800;'>1 Hari</span>"
            if durasi_hari == 1
            else f"<span style='background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:6px; font-weight:800;'>{durasi_hari} Hari Berturut-turut (Otomatis Centang {durasi_hari} Baris di DATA PERJADIN)</span>"
        )
        st.markdown(clean_html(f'<div class="info-pill">🗓️ <b>Tampilan di Surat Tugas & Database:</b><br>'
            f'• <b>hari / tanggal :</b> <b>{nama_hari_id} / {tgl_indo_str}</b> &nbsp;{badge_durasi}<br>'
            f'• <b>Format Nomor ST :</b> {preview_no_st_fmt}</div>'), unsafe_allow_html=True)

    # -------------------------------------------------------------
    # KARTU 3: TIM PETUGAS PEMERIKSA (MULTISELECT INSTAN + BENTO MINI-CARDS)
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown(clean_html('<div class="section-badge">3️⃣ Tim Petugas Pemeriksa (Instan Tanpa Loading Berulang)</div>'), unsafe_allow_html=True)

        default_petugas_init = get_default_petugas(list_pegawai)
        if "ms_petugas_st" not in st.session_state:
            st.session_state["ms_petugas_st"] = default_petugas_init

        selected_names_ms = st.multiselect(
            "👮 Pilih Tim Petugas Pemeriksa (Urutan Pilihan = Urutan di ST, Maks. 6 Orang):",
            options=list_pegawai,
            max_selections=6,
            key="ms_petugas_st",
            placeholder="Ketik nama pegawai untuk memilih (Maksimal 6 orang)..."
        )

        data_petugas_terpilih = []
        bento_cards_html = []

        for idx_p, sel_nama in enumerate(selected_names_ms, start=1):
            detail = pegawai_data.get(sel_nama, {})
            p_nip = detail.get("nip", "-")
            p_pangkat = detail.get("pangkat", "-")
            p_jabatan = detail.get("jabatan", "-")
            data_petugas_terpilih.append({
                "no": idx_p,
                "nama": sel_nama,
                "nip": p_nip,
                "pangkat": p_pangkat,
                "jabatan": p_jabatan
            })
            bento_cards_html.append(
                f'<div class="petugas-bento-card">'
                f'<div class="p-title">#{idx_p}. {sel_nama}</div>'
                f'<div class="p-sub">🎖️ <b>{p_pangkat}</b> | 💼 {p_jabatan}<br>🆔 NIP: {p_nip}</div>'
                f'</div>'
            )

        if bento_cards_html:
            st.markdown(clean_html(f'<div class="petugas-bento-grid">{"".join(bento_cards_html)}</div>'), unsafe_allow_html=True)
        else:
            st.markdown(clean_html('<div class="info-pill-warn">⚠️ Belum ada petugas yang dipilih. Silakan pilih minimal 1 pegawai pada kotak di atas.</div>'), unsafe_allow_html=True)

    # -------------------------------------------------------------
    # KARTU 4: KONFIGURASI PENANDATANGAN
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown(clean_html('<div class="section-badge">4️⃣ Pejabat Penandatangan Surat Tugas</div>'), unsafe_allow_html=True)

        c_ttd_left, c_ttd_right = st.columns([1.2, 1])
        with c_ttd_left:
            tipe_ttd = st.radio(
                "Format Jabatan Penandatangan (Sekaligus Menentukan Format Kode Nomor ST):",
                options=["a.n. Kepala Kantor (Kepala Subbagian Umum)", "Kepala Kantor Langsung"],
                index=0,
                key="in_tipe_ttd"
            )
            is_plh = st.checkbox("Bertindak sebagai Pelaksana Harian (Plh.)", value=False)

        with c_ttd_right:
            if tipe_ttd == "Kepala Kantor Langsung":
                if is_plh:
                    nama_pejabat_ttd = st.selectbox(
                        "Pilih Pejabat Plh. Kepala Kantor*:",
                        options=[""] + list_pegawai,
                        index=0
                    )
                    lbl_jab_ringkas = "Plh. Kepala Kantor KPPBC TMC Malang"
                else:
                    nama_pejabat_ttd = st.text_input("Nama Kepala Kantor Definitif:", value="Johan Pandores")
                    lbl_jab_ringkas = "Kepala Kantor KPPBC TMC Malang"
            else:
                if is_plh:
                    nama_pejabat_ttd = st.selectbox(
                        "Pilih Pejabat Plh. Kasubag Umum*:",
                        options=[""] + list_pegawai,
                        index=0
                    )
                    lbl_jab_ringkas = "a.n. Kepala Kantor — Plh. Kepala Subbagian Umum"
                else:
                    nama_pejabat_ttd = st.text_input("Nama Kasubag Umum Definitif:", value="Kukuh Hanna Prapanca")
                    lbl_jab_ringkas = "a.n. Kepala Kantor — Kepala Subbagian Umum"

            fmt_st_k4 = f"ST-xx/KBC.1201/{tgl_pelaksanaan.year}" if tipe_ttd == "Kepala Kantor Langsung" else f"ST-xx/KBC.1201/KBC.120101/{tgl_pelaksanaan.year}"
            st.markdown(clean_html(f'<div class="info-pill">✍️ <b>Format TTE:</b> {lbl_jab_ringkas}<br>'
                f'👤 <b>Pejabat:</b> {nama_pejabat_ttd or "(Belum dipilih)"}<br>'
                f'🔢 <b>Format Nomor ST:</b> <code>{fmt_st_k4}</code></div>'), unsafe_allow_html=True)

    # -------------------------------------------------------------
    # TOMBOL AKSI UTAMA INPUT BARU (OTOMATIS MASUK 3 SHEET!)
    # -------------------------------------------------------------
    btn_simpan_eksekusi = st.button(
        f"💾 Simpan Sekaligus ke 3 Sheet (Database ST #{next_no_preview} + Rekap Umum + Monitoring Perjadin) & Buat Word",
        type="primary",
        use_container_width=True
    )

    if btn_simpan_eksekusi:
        if not pt_1:
            st.error("⚠️ Nama Perusahaan / Pabrik 1 wajib dipilih!")
        elif not no_surat_1.strip():
            st.error("⚠️ Nomor Surat Pemohon 1 wajib diisi!")
        elif len(data_petugas_terpilih) == 0:
            st.error("⚠️ Minimal harus memilih 1 Petugas Pemeriksa!")
        elif not nama_pejabat_ttd.strip():
            st.error("⚠️ Nama Pejabat Penandatangan tidak boleh kosong!")
        else:
            try:
                with st.spinner("Menyimpan ke Sheet1, REKAP BAGIAN UMUM, dan DATA PERJADIN serta merakit Word..."):
                    ss = get_spreadsheet()
                    ws_target = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])

                    # Pastikan baris yang sempat bergeser ke kolom AG (jika ada) dirapikan lebih dulu
                    repair_shifted_rows_in_sheet1(ws_target)
                    col_no = ws_target.col_values(1)
                    existing_numbers = [int(x) for x in col_no if x.strip().isdigit()]
                    next_no = max(existing_numbers) + 1 if existing_numbers else 1
                    target_row_s1 = len(col_no) + 1

                    petugas_row_data = []
                    for p_idx in range(6):
                        if p_idx < len(data_petugas_terpilih):
                            p = data_petugas_terpilih[p_idx]
                            petugas_row_data.extend([p["nama"], p["pangkat"], p["jabatan"]])
                        else:
                            petugas_row_data.extend(["", "", ""])

                    nomor_st_clean = format_nomor_st_resmi(
                        no_st_opsional,
                        tgl_pelaksanaan.year,
                        tipe_ttd=tipe_ttd,
                        auto_adjust_unit=True
                    )

                    row_payload = [
                        next_no,
                        pt_1,
                        no_surat_1,
                        format_tanggal_indo(tgl_surat_1),
                        pt_2 if pt_2 else "",
                        no_surat_2 if no_surat_2 else "",
                        format_tanggal_indo(tgl_surat_2) if (pt_2 and tgl_surat_2) else "",
                        permohonan_1,
                        permohonan_2 if pt_2 else "",
                        *petugas_row_data,
                        kegiatan_1,
                        kegiatan_2 if pt_2 else "",
                        nama_hari_id,
                        tgl_indo_str,
                        pilihan_lokasi_1,
                        alamat_pabrik_1,
                        pilihan_lokasi_2 if pt_2 else "",
                        alamat_pabrik_2 if pt_2 else "",
                        nomor_st_clean,
                        tgl_st_indo_str
                    ]

                    # 1. Simpan ke Sheet1 secara eksplisit di Kolom A:AK (100% Anti Bergeser ke Kolom AG!)
                    if target_row_s1 > ws_target.row_count:
                        ws_target.add_rows(50)
                    ws_target.update(
                        f"A{target_row_s1}:AK{target_row_s1}",
                        [row_payload],
                        value_input_option="USER_ENTERED"
                    )

                    # 2. Siapkan data gabungan untuk REKAP BAGIAN UMUM & DATA PERJADIN
                    perihal_rekap = kegiatan_1.strip() if kegiatan_1.strip() else strip_permohonan(permohonan_1)
                    if pt_2 and kegiatan_2.strip() and kegiatan_2.strip().lower() != perihal_rekap.lower():
                        perihal_rekap = f"{perihal_rekap} dan {kegiatan_2.strip()}"

                    nama_pabrik_gabung = f"{pt_1}, {pt_2}" if pt_2 else pt_1
                    lokasi_kab_final, kec_final = extract_lokasi_kecamatan(alamat_pabrik_1, pt_1, perusahaan_data)
                    if pt_2:
                        _, kec_2_val = extract_lokasi_kecamatan(alamat_pabrik_2, pt_2, perusahaan_data)
                        if kec_2_val != "-" and kec_2_val.lower() not in kec_final.lower():
                            kec_final = f"{kec_final}, {kec_2_val}"

                    petugas_gabung_str = ", ".join([p["nama"] for p in data_petugas_terpilih])

                    sync_res = sync_st_to_rekap_and_perjadin(
                        ss=ss,
                        no_urut=next_no,
                        tgl_st_str=tgl_st_indo_str,
                        perihal_str=perihal_rekap,
                        nama_pabrik_gabung=nama_pabrik_gabung,
                        lokasi_kab_kota=lokasi_kab_final,
                        kecamatan_str=kec_final,
                        petugas_gabung_str=petugas_gabung_str,
                        tgl_pelaksanaan_str=tgl_indo_str,
                        nomor_st_resmi=nomor_st_clean
                    )

                    st.toast(f"ST #{next_no} tersimpan di 3 Sheet sekaligus!", icon="🎉")
                    jml_hari_tercentang = len(sync_res.get("perjadin_rows", []))
                    msg_perjadin = (
                        f"**{jml_hari_tercentang} Hari Pelaksanaan** (`{nama_hari_id} / {tgl_indo_str}`, Baris {sync_res.get('perjadin_row')}) di `DATA PERJADIN` otomatis tercentang!"
                        if sync_res.get("perjadin_ok")
                        else "Tanggal pelaksanaan berada di luar rentang kalender `DATA PERJADIN` saat ini."
                    )
                    st.success(
                        f"✅ **Berhasil Simpan Sekali Klik ke 3 Sheet!**\n"
                        f"- **Sheet1 (Database ST):** Tersimpan sebagai No. Urut **#{next_no}** ({nama_pabrik_gabung} — `{nama_hari_id} / {tgl_indo_str}`)\n"
                        f"- **Sheet `{REKAP_SHEET_NAME}`:** Otomatis terisi (`{lokasi_kab_final}` — Kec. `{kec_final}`)\n"
                        f"- **Sheet `DATA PERJADIN`:** {msg_perjadin}"
                    )
                    st.cache_data.clear()

                    st.session_state["last_submitted_data"] = {
                        "no_urut": next_no,
                        "pt_1": pt_1,
                        "no_surat_1": no_surat_1,
                        "tgl_surat_1": format_tanggal_indo(tgl_surat_1),
                        "pt_2": pt_2 if pt_2 else "",
                        "no_surat_2": no_surat_2 if pt_2 else "",
                        "tgl_surat_2": format_tanggal_indo(tgl_surat_2) if (pt_2 and tgl_surat_2) else "",
                        "permohonan_1": permohonan_1,
                        "permohonan_2": permohonan_2 if pt_2 else "",
                        "kegiatan_1": kegiatan_1,
                        "kegiatan_2": kegiatan_2 if pt_2 else "",
                        "hari": nama_hari_id,
                        "tgl_pelaksanaan": tgl_indo_str,
                        "alamat_pabrik_1": alamat_pabrik_1,
                        "alamat_pabrik_2": alamat_pabrik_2 if pt_2 else "",
                        "daftar_petugas": data_petugas_terpilih,
                        "tipe_ttd": tipe_ttd,
                        "is_plh": is_plh,
                        "nama_pejabat_ttd": nama_pejabat_ttd
                    }
                    st.rerun()
            except Exception as err:
                st.error(f"Gagal menyimpan ke Google Sheets: {err}")

    # Kartu Unduh Langsung & Pratinjau Opsional (Menggunakan Base64 Murni Anti-Gagal Download)
    if "last_submitted_data" in st.session_state:
        data_gen = st.session_state["last_submitted_data"]
        with st.container(border=True):
            st.markdown(clean_html(f'<div class="section-badge">📄 Dokumen Siap Diunduh — ST No. #{data_gen["no_urut"]} ({data_gen["pt_1"]})</div>'), unsafe_allow_html=True)

            try:
                buf = generate_docx_buffer(data_gen)
                docx_bytes_val = buf.getvalue()
                safe_tgl_fn = data_gen['tgl_pelaksanaan'].replace('s.d.', 'sd').replace('s.d', 'sd').replace(' ', '_')
                file_name_out = f"ST_{data_gen['no_urut']}_{data_gen['pt_1'].replace(' ', '_')}_{safe_tgl_fn}.docx"
                col_dl1, col_dl2 = st.columns([2, 1])
                with col_dl1:
                    render_direct_docx_download(
                        docx_bytes=docx_bytes_val,
                        filename=file_name_out,
                        button_label=f"📥 Unduh Surat Tugas (.docx) — {file_name_out}"
                    )
                with col_dl2:
                    if st.button("✨ Buat ST Baru Lagi (Reset)", use_container_width=True, key="btn_reset_after_dl"):
                        reset_form_input(list_pegawai)
                        st.rerun()
            except Exception as e_dl:
                st.error(f"Gagal membuat file Word: {e_dl}")

            with st.expander("👁️ Intip Pratinjau Kertas A4 (Opsional)", expanded=False):
                components.html(render_preview_html(data_gen), height=850, scrolling=True)

# ------------------------------------------------------------------
# MENU 3: ⚡ 3. TINDAK LANJUT ST & PERJADIN (DESAIN HYBRID SATUAN vs TABEL MASSAL)
# ------------------------------------------------------------------
elif active_menu == MENU_OPTIONS[2]:
    focus_opt_a = f"🔢 A. Input Nomor ST & Rekap Umum ({jml_st_belum_nomor} Antrean)"
    focus_opt_b = f"🚗 B. Update ID & Pencairan Perjadin ({jml_belum_perjadin} Belum ID)"
    focus_options = [focus_opt_a, focus_opt_b]
    default_focus_idx = 1 if st.session_state.get("tl_focus_mode") == "B" else 0

    selected_focus = st.radio(
        "🎯 Pilih Fokus Kerja Tindak Lanjut (Segmented Pill Switcher):",
        options=focus_options,
        index=default_focus_idx,
        horizontal=True,
        label_visibility="collapsed"
    )
    st.session_state["tl_focus_mode"] = "B" if selected_focus == focus_opt_b else "A"

    # ==============================================================
    # FOKUS A: 🔢 INPUT NOMOR ST & REKAP BAGIAN UMUM (HYBRID SATUAN & MASSAL)
    # ==============================================================
    if st.session_state["tl_focus_mode"] == "A":
        with st.container(border=True):
            hd_a1, hd_a2 = st.columns([1.35, 1], gap="small")
            with hd_a1:
                st.markdown(clean_html('<div class="section-badge">🔢 A.1 — Input Nomor ST Satu Kemenkeu (Otomatis Sinkron 3 Sheet)</div>'), unsafe_allow_html=True)
                st.caption("Pilih **Mode Form Satuan** untuk pengisian presisi di HP, atau **Mode Tabel Edit Massal** untuk mengetik banyak Nomor ST sekaligus seperti di Excel.")
            with hd_a2:
                nost_mode_opts = ["📱 Mode Form Satuan", "⚡ Mode Tabel Edit Massal"]
                cur_nost_mode = st.session_state.get("nost_work_mode", nost_mode_opts[0])
                nost_mode_sel = st.radio(
                    "Mode Input Nomor ST:",
                    options=nost_mode_opts,
                    index=1 if cur_nost_mode == nost_mode_opts[1] else 0,
                    horizontal=True,
                    label_visibility="collapsed"
                )
                st.session_state["nost_work_mode"] = nost_mode_sel

            rekap_map_by_no = {str(r["no"]).strip(): r for r in rekap_list_global}

            # ------------------------------------------------------
            # MODE 1: 📱 MODE FORM SATUAN (PRESISI & AUTO-SELECT DARI ACTION CENTER)
            # ------------------------------------------------------
            if st.session_state["nost_work_mode"] == "📱 Mode Form Satuan":
                if not riwayat_list_global:
                    st.info("Belum ada data Surat Tugas yang tersimpan di Sheet1.")
                else:
                    jump_no = st.session_state.get("jump_target_nost_urut")
                    default_filter_idx = 0
                    if jump_no:
                        for it_chk in riwayat_list_global:
                            if str(it_chk["no_urut"]).strip() == str(jump_no):
                                c_n = it_chk.get("nomor_st_resmi") or rekap_map_by_no.get(str(jump_no), {}).get("nomor_st", "")
                                if c_n:
                                    default_filter_idx = 1
                                break

                    filter_st_mode = st.radio(
                        "Filter Daftar Surat Tugas:",
                        options=[
                            f"⏳ Belum Ada Nomor ST ({jml_st_belum_nomor} ST)",
                            f"🗂️ Semua Riwayat ST ({len(riwayat_list_global)} ST)"
                        ],
                        index=default_filter_idx,
                        horizontal=True,
                        key="pill_filter_st_mode"
                    )

                    kandidat_st = []
                    for item in riwayat_list_global:
                        no_u = str(item["no_urut"]).strip()
                        cur_nost = item.get("nomor_st_resmi") or rekap_map_by_no.get(no_u, {}).get("nomor_st", "")
                        item_copy = dict(item)
                        item_copy["cur_nost"] = cur_nost
                        if filter_st_mode.startswith("⏳"):
                            if not cur_nost:
                                kandidat_st.append(item_copy)
                        else:
                            kandidat_st.append(item_copy)

                    if not kandidat_st:
                        st.success("🎉 Seluruh Surat Tugas sudah memiliki Nomor ST resmi! Menampilkan 25 ST terbaru:")
                        kandidat_st = [
                            dict(item, cur_nost=(item.get("nomor_st_resmi") or rekap_map_by_no.get(str(item["no_urut"]).strip(), {}).get("nomor_st", "")))
                            for item in riwayat_list_global[:25]
                        ]

                    opsi_st_update = [
                        f"No. #{it['no_urut']} | {it['tgl_pelaksanaan']} | {it['pt_1']} ({'✅ ' + it['cur_nost'] if it['cur_nost'] else '⏳ Belum Ada Nomor ST'})"
                        for it in kandidat_st
                    ]

                    default_st_idx = 0
                    if jump_no:
                        for idx_k, it_k in enumerate(kandidat_st):
                            if str(it_k["no_urut"]).strip() == str(jump_no):
                                default_st_idx = idx_k
                                break

                    pil_st_lbl = st.selectbox(
                        f"Pilih Surat Tugas ({len(kandidat_st)} ST ditampilkan):",
                        options=opsi_st_update,
                        index=default_st_idx
                    )
                    sel_st_obj = kandidat_st[opsi_st_update.index(pil_st_lbl)]

                    lok_preview_a, kec_preview_a = extract_lokasi_kecamatan(
                        sel_st_obj.get("alamat_pabrik_1", ""),
                        sel_st_obj.get("pt_1", ""),
                        perusahaan_data
                    )
                    st.markdown(clean_html(f"""
                    <div class="info-pill">
                        📌 <b>ST Terpilih:</b> No. Urut <b>#{sel_st_obj["no_urut"]}</b> &nbsp;|&nbsp;
                        🏢 <b>Pabrik:</b> {sel_st_obj["pt_1"]}{(" & " + sel_st_obj["pt_2"]) if sel_st_obj.get("pt_2") else ""}<br>
                        🗓️ <b>Waktu Pelaksanaan:</b> {sel_st_obj["hari"]} / <b>{sel_st_obj["tgl_pelaksanaan"]}</b> &nbsp;|&nbsp;
                        📍 <b>Wilayah:</b> {lok_preview_a} (Kec. {kec_preview_a})
                    </div>
                    """), unsafe_allow_html=True)

                    def_ttd_fast = detect_tipe_ttd_from_nost(sel_st_obj.get("cur_nost", ""))
                    pil_ttd_fast = st.radio(
                        "Pejabat Penandatangan ST (Penentu Format Kode Nomor ST):",
                        options=[
                            "a.n. Kepala Kantor (ST-xx/KBC.1201/KBC.120101/2026)",
                            "Kepala Kantor Langsung (ST-xx/KBC.1201/2026)"
                        ],
                        index=1 if def_ttd_fast == "Kepala Kantor Langsung" else 0,
                        horizontal=True,
                        key=f"fast_ttd_mode_{sel_st_obj['no_urut']}"
                    )
                    tipe_ttd_fast_clean = (
                        "Kepala Kantor Langsung"
                        if "Kepala Kantor Langsung" in pil_ttd_fast
                        else "a.n. Kepala Kantor (Kepala Subbagian Umum)"
                    )

                    default_petugas_rekap = rekap_map_by_no.get(str(sel_st_obj["no_urut"]).strip(), {}).get("petugas", "")
                    if not default_petugas_rekap:
                        default_petugas_rekap = ", ".join([p["nama"] for p in sel_st_obj["daftar_petugas"]])

                    c_fn1, c_fn2 = st.columns([1.1, 1.1])
                    with c_fn1:
                        input_angka_st = st.text_input(
                            "🔢 Ketik Nomor ST dari Satu Kemenkeu*:",
                            value=sel_st_obj.get("cur_nost", ""),
                            placeholder="Ketik angka saja misal: 2041 atau ST-2041/...",
                            key=f"fast_nost_in_{sel_st_obj['no_urut']}"
                        )
                    with c_fn2:
                        input_petugas_rekap = st.text_input(
                            "👮 Nama Petugas di Rekap Bagian Umum (Bisa Disingkat):",
                            value=default_petugas_rekap,
                            key=f"fast_pet_in_{sel_st_obj['no_urut']}"
                        )
                    input_tgl_st = sel_st_obj["tgl_pelaksanaan"]

                    d_pel_prev = parse_indo_date(sel_st_obj["tgl_pelaksanaan"])
                    yr_pel_prev = d_pel_prev.year if d_pel_prev else 2026
                    live_prev_nost = format_nomor_st_resmi(
                        input_angka_st,
                        yr_pel_prev,
                        tipe_ttd=tipe_ttd_fast_clean,
                        auto_adjust_unit=True
                    )
                    if live_prev_nost:
                        st.markdown(clean_html( clean_html(f'<div class="info-pill-ok">✅ <b>Format Nomor ST Resmi yang Akan Disimpan ke 3 Sheet:</b> <code>{live_prev_nost}</code></div>')), unsafe_allow_html=True)

                    if st.button("💾 Simpan Nomor ST ke Sheet1 + REKAP BAGIAN UMUM + DATA PERJADIN", type="primary", use_container_width=True):
                        if not input_angka_st.strip():
                            st.error("⚠️ Masukkan Nomor ST terlebih dahulu!")
                        else:
                            try:
                                with st.spinner("Menyimpan Nomor ST ke Sheet1, REKAP BAGIAN UMUM, dan DATA PERJADIN..."):
                                    ss = get_spreadsheet()
                                    d_pel = parse_indo_date(sel_st_obj["tgl_pelaksanaan"])
                                    yr_pel = d_pel.year if d_pel else 2026
                                    no_st_final = format_nomor_st_resmi(
                                        input_angka_st,
                                        yr_pel,
                                        tipe_ttd=tipe_ttd_fast_clean,
                                        auto_adjust_unit=True
                                    )

                                    ws_s1 = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                                    ws_s1.update(
                                        f"AJ{sel_st_obj['row_num']}:AK{sel_st_obj['row_num']}",
                                        [[no_st_final, input_tgl_st]],
                                        value_input_option="USER_ENTERED"
                                    )

                                    perihal_r = sel_st_obj["kegiatan_1"].strip() or strip_permohonan(sel_st_obj["permohonan_1"])
                                    if sel_st_obj["pt_2"] and sel_st_obj["kegiatan_2"].strip():
                                        perihal_r = f"{perihal_r} dan {sel_st_obj['kegiatan_2'].strip()}"
                                    pt_gab = f"{sel_st_obj['pt_1']}, {sel_st_obj['pt_2']}" if sel_st_obj["pt_2"] else sel_st_obj["pt_1"]
                                    lok_r, kec_r = extract_lokasi_kecamatan(sel_st_obj["alamat_pabrik_1"], sel_st_obj["pt_1"], perusahaan_data)

                                    sync_st_to_rekap_and_perjadin(
                                        ss=ss,
                                        no_urut=int(sel_st_obj["no_urut"]),
                                        tgl_st_str=input_tgl_st,
                                        perihal_str=perihal_r,
                                        nama_pabrik_gabung=pt_gab,
                                        lokasi_kab_kota=lok_r,
                                        kecamatan_str=kec_r,
                                        petugas_gabung_str=input_petugas_rekap,
                                        tgl_pelaksanaan_str=sel_st_obj["tgl_pelaksanaan"],
                                        nomor_st_resmi=no_st_final
                                    )
                                    st.session_state["jump_target_nost_urut"] = None
                                    st.toast(f"Nomor ST {no_st_final} berhasil disinkronkan ke 3 Sheet!", icon="✅")
                                    st.success(f"✅ **Nomor ST `{no_st_final}`** berhasil disimpan ke `Sheet1`, `{REKAP_SHEET_NAME}`, dan `DATA PERJADIN`!")
                                    st.cache_data.clear()
                                    st.rerun()
                            except Exception as e_upst:
                                st.error(f"Gagal menyimpan Nomor ST: {e_upst}")

            # ------------------------------------------------------
            # MODE 2: ⚡ MODE TABEL EDIT MASSAL (ST.DATA_EDITOR SPREADSHEET MODE)
            # ------------------------------------------------------
            else:
                scope_mass_nost = st.radio(
                    "Cakupan Baris Tabel Edit Massal:",
                    options=[
                        f"⏳ Hanya ST yang Belum Bernomor ({jml_st_belum_nomor} ST)",
                        "🗂️ 25 Surat Tugas Terbaru"
                    ],
                    index=0 if jml_st_belum_nomor > 0 else 1,
                    horizontal=True,
                    key="scope_mass_nost_radio"
                )

                mass_candidates = []
                for item in riwayat_list_global:
                    no_u = str(item["no_urut"]).strip()
                    cur_nost = item.get("nomor_st_resmi") or rekap_map_by_no.get(no_u, {}).get("nomor_st", "")
                    item_c = dict(item, cur_nost=cur_nost)
                    if scope_mass_nost.startswith("⏳"):
                        if not cur_nost:
                            mass_candidates.append(item_c)
                    else:
                        mass_candidates.append(item_c)

                if scope_mass_nost.startswith("🗂️"):
                    mass_candidates = mass_candidates[:25]
                elif not mass_candidates:
                    st.info("Seluruh ST sudah memiliki Nomor ST resmi. Menampilkan 20 ST terbaru untuk edit cepat:")
                    mass_candidates = [
                        dict(item, cur_nost=(item.get("nomor_st_resmi") or rekap_map_by_no.get(str(item["no_urut"]).strip(), {}).get("nomor_st", "")))
                        for item in riwayat_list_global[:20]
                    ]

                ttd_opts_short = [
                    "a.n. Kepala Kantor (KBC.1201/KBC.120101)",
                    "Kepala Kantor Langsung (KBC.1201)"
                ]

                editor_rows_nost = []
                orig_nost_map = {}
                for it in mass_candidates:
                    no_u = str(it["no_urut"]).strip()
                    cur_nost = it["cur_nost"]
                    det_ttd = detect_tipe_ttd_from_nost(cur_nost)
                    ttd_lbl = ttd_opts_short[1] if det_ttd == "Kepala Kantor Langsung" else ttd_opts_short[0]
                    tgl_lapor_def = it["tgl_pelaksanaan"]
                    pet_def = rekap_map_by_no.get(no_u, {}).get("petugas", "") or ", ".join([p["nama"] for p in it["daftar_petugas"]])
                    pt_gab = f"{it['pt_1']}, {it['pt_2']}" if it["pt_2"] else it["pt_1"]

                    orig_nost_map[no_u] = {
                        "item": it,
                        "nost": cur_nost,
                        "ttd_lbl": ttd_lbl,
                        "tgl_lapor": tgl_lapor_def,
                        "petugas": pet_def
                    }
                    editor_rows_nost.append({
                        "NO": no_u,
                        "TGL PELAKSANAAN": it["tgl_pelaksanaan"],
                        "NAMA PABRIK": pt_gab,
                        "TIPE TTD": ttd_lbl,
                        "NOMOR ST (ANGKA / LENGKAP)": cur_nost,
                        "PETUGAS REKAP": pet_def
                    })

                st.caption("💡 **Tips:** Ketik langsung angka Nomor ST (misal `2041`) pada kolom **NOMOR ST**, pilih **TIPE TTD**, lalu klik tombol simpan di bawah (Kolom `TANGGAL` di Rekap Bagian Umum otomatis disamakan dengan `TGL PELAKSANAAN`):")
                edited_nost_df = st.data_editor(
                    editor_rows_nost,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "NO": st.column_config.TextColumn("NO", disabled=True, width="small"),
                        "TGL PELAKSANAAN": st.column_config.TextColumn("TGL PELAKSANAAN", disabled=True),
                        "NAMA PABRIK": st.column_config.TextColumn("NAMA PABRIK", disabled=True),
                        "TIPE TTD": st.column_config.SelectboxColumn("TIPE TTD", options=ttd_opts_short, required=True),
                        "NOMOR ST (ANGKA / LENGKAP)": st.column_config.TextColumn("NOMOR ST (KETIK ANGKA)"),
                        "PETUGAS REKAP": st.column_config.TextColumn("PETUGAS REKAP"),
                    },
                    key="data_editor_mass_nost"
                )

                if st.button("💾 Simpan Semua Nomor ST Sekaligus (Batch 3-Sheet)", type="primary", use_container_width=True):
                    changed_items = []
                    for row_ed in edited_nost_df:
                        no_u = str(row_ed["NO"]).strip()
                        orig = orig_nost_map.get(no_u)
                        if not orig:
                            continue
                        new_raw_nost = str(row_ed.get("NOMOR ST (ANGKA / LENGKAP)", "") or "").strip()
                        new_ttd_lbl = str(row_ed.get("TIPE TTD", orig["ttd_lbl"]) or orig["ttd_lbl"])
                        new_tgl_lapor = orig["item"]["tgl_pelaksanaan"]
                        new_pet = str(row_ed.get("PETUGAS REKAP", orig["petugas"]) or "").strip()

                        it_obj = orig["item"]
                        d_pel = parse_indo_date(it_obj["tgl_pelaksanaan"])
                        yr_pel = d_pel.year if d_pel else 2026
                        tipe_ttd_clean = "Kepala Kantor Langsung" if "Kepala Kantor Langsung" in new_ttd_lbl else "a.n. Kepala Kantor"
                        formatted_nost = format_nomor_st_resmi(new_raw_nost, yr_pel, tipe_ttd=tipe_ttd_clean, auto_adjust_unit=True) if new_raw_nost else ""

                        if (formatted_nost != orig["nost"]) or (new_tgl_lapor != orig["tgl_lapor"]) or (new_pet != orig["petugas"]):
                            changed_items.append((it_obj, formatted_nost, new_tgl_lapor, new_pet))

                    if not changed_items:
                        st.warning("⚠️ Belum ada perubahan Nomor ST pada tabel di atas.")
                    else:
                        try:
                            with st.spinner(f"Menyimpan {len(changed_items)} baris Nomor ST secara Batch ke 3 Sheet..."):
                                ss = get_spreadsheet()
                                ws_s1 = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                                s1_batch = []
                                for it_obj, fmt_nost, tgl_lap, pet_str in changed_items:
                                    s1_batch.append({
                                        "range": f"AJ{it_obj['row_num']}:AK{it_obj['row_num']}",
                                        "values": [[fmt_nost, tgl_lap]]
                                    })
                                    perihal_r = it_obj["kegiatan_1"].strip() or strip_permohonan(it_obj["permohonan_1"])
                                    if it_obj["pt_2"] and it_obj["kegiatan_2"].strip():
                                        perihal_r = f"{perihal_r} dan {it_obj['kegiatan_2'].strip()}"
                                    pt_gab = f"{it_obj['pt_1']}, {it_obj['pt_2']}" if it_obj["pt_2"] else it_obj["pt_1"]
                                    lok_r, kec_r = extract_lokasi_kecamatan(it_obj["alamat_pabrik_1"], it_obj["pt_1"], perusahaan_data)
                                    sync_st_to_rekap_and_perjadin(
                                        ss=ss,
                                        no_urut=int(it_obj["no_urut"]),
                                        tgl_st_str=tgl_lap,
                                        perihal_str=perihal_r,
                                        nama_pabrik_gabung=pt_gab,
                                        lokasi_kab_kota=lok_r,
                                        kecamatan_str=kec_r,
                                        petugas_gabung_str=pet_str,
                                        tgl_pelaksanaan_str=it_obj["tgl_pelaksanaan"],
                                        nomor_st_resmi=fmt_nost
                                    )
                                if s1_batch:
                                    ws_s1.batch_update(s1_batch, value_input_option="USER_ENTERED")
                                st.toast(f"{len(changed_items)} Nomor ST berhasil disimpan ke 3 Sheet!", icon="🎉")
                                st.success(f"✅ Berhasil menyimpan **{len(changed_items)} Nomor ST** sekaligus ke `Sheet1`, `{REKAP_SHEET_NAME}`, dan `DATA PERJADIN`!")
                                st.cache_data.clear()
                                st.rerun()
                        except Exception as e_mass_nost:
                            st.error(f"Gagal menyimpan Nomor ST secara massal: {e_mass_nost}")

        # ------------------------------------------------------
        # KARTU A.2: TABEL REKAP BAGIAN UMUM & FILTER COPY PINTAR (1-CLICK COPY JS)
        # ------------------------------------------------------
        with st.container(border=True):
            st.markdown(clean_html('<div class="section-badge">📊 A.2 — Tabel Sheet REKAP BAGIAN UMUM & Filter Copy Pintar</div>'), unsafe_allow_html=True)
            if not rekap_list_global:
                st.warning("Sheet `REKAP BAGIAN UMUM` masih kosong. Gunakan utilitas Sinkronisasi Massal di bagian bawah halaman ini untuk menarik data dari `Sheet1`.")
            else:
                fc1, fc2 = st.columns(2, gap="small")
                with fc1:
                    copy_month_filter = st.selectbox(
                        "📅 Filter Bulan Rekap Bagian Umum:",
                        options=dynamic_periods_2026,
                        index=0,
                        key="rekap_copy_month_filter"
                    )
                with fc2:
                    copy_limit_filter = st.selectbox(
                        "🔢 Jumlah Baris Terakhir untuk Di-Copy:",
                        options=["20 Baris Terakhir", "10 Baris Terakhir", "50 Baris Terakhir", "Semua Baris"],
                        index=0,
                        key="rekap_copy_limit_filter"
                    )

                target_m_rek = None
                for m_num, m_name in BULAN_INDONESIA.items():
                    if m_name.lower() in copy_month_filter.lower():
                        target_m_rek = m_num
                        break

                filtered_rekap_rows = []
                for r in rekap_list_global:
                    if target_m_rek is not None:
                        d_r = parse_indo_date(r.get("tgl_pelaksanaan", "")) or parse_indo_date(r.get("tanggal", ""))
                        if not d_r or d_r.month != target_m_rek:
                            continue
                    filtered_rekap_rows.append(r)

                if copy_limit_filter.startswith("10"):
                    sliced_rekap = filtered_rekap_rows[-10:]
                elif copy_limit_filter.startswith("20"):
                    sliced_rekap = filtered_rekap_rows[-20:]
                elif copy_limit_filter.startswith("50"):
                    sliced_rekap = filtered_rekap_rows[-50:]
                else:
                    sliced_rekap = filtered_rekap_rows

                st.dataframe(
                    [
                        {
                            "NO": r["no"],
                            "TANGGAL": r["tanggal"],
                            "PERIHAL": r["perihal"],
                            "Nama Pabrik": r["nama_pabrik"],
                            "LOKASI": r["lokasi"],
                            "Kecamatan": r["kecamatan"],
                            "PETUGAS": r["petugas"],
                            "TANGGAL PELAKSANAAN": r["tgl_pelaksanaan"],
                            "NOMOR ST": r["nomor_st"]
                        }
                        for r in reversed(sliced_rekap)
                    ],
                    use_container_width=True,
                    hide_index=True
                )

                tsv_lines = ["\t".join(REKAP_HEADERS)]
                for r in sliced_rekap:
                    tsv_lines.append("\t".join([
                        str(r["no"]), r["tanggal"], r["perihal"], r["nama_pabrik"],
                        r["lokasi"], r["kecamatan"], r["petugas"], r["tgl_pelaksanaan"], r["nomor_st"]
                    ]))
                tsv_payload = "\n".join(tsv_lines)

                render_copy_tsv_button(
                    tsv_text=tsv_payload,
                    button_label=f"📋 Copy {len(sliced_rekap)} Baris Laporan ke Clipboard (Siap Paste ke Excel / WhatsApp Bagian Umum)"
                )
                with st.expander("👁️ Lihat Teks Mentah Format TSV (Opsional)", expanded=False):
                    st.code(tsv_payload, language="text")

        # ------------------------------------------------------
        # EXPANDER A.3: UTILITAS SINKRONISASI MASSAL & BACKFILL 3-SHEET (TERSEMBUNYI RAPI)
        # ------------------------------------------------------
        with st.expander("🔄 Utilitas Sinkronisasi Massal & Backfill 3-Sheet (Batch API)", expanded=False):
            st.markdown("""
            Gunakan tombol di bawah jika kamu ingin merapikan ulang baris `Sheet1` yang bergeser, menyusun ulang seluruh tabel **`REKAP BAGIAN UMUM`**, dan mencentang otomatis seluruh tanggal penugasan di **`DATA PERJADIN`** dalam 1x klik Batch API:
            """)
            if st.button("🚀 Jalankan Sinkronisasi Otomatis Seluruh Data Sheet1 -> REKAP BAGIAN UMUM & DATA PERJADIN", type="primary", use_container_width=True):
                try:
                    with st.spinner("Sedang memproses & menyinkronkan seluruh riwayat ST secara Batch ke REKAP BAGIAN UMUM dan DATA PERJADIN..."):
                        ss = get_spreadsheet()
                        ws_s1_sync = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                        repair_shifted_rows_in_sheet1(ws_s1_sync)

                        ws_rekap = get_or_create_rekap_worksheet(ss)
                        ws_perj = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])

                        raw_perj_all = ws_perj.get_all_values()
                        date_to_nost_perj = {}
                        date_to_perj_row = {}
                        for idx_p, rp in enumerate(raw_perj_all[1:], start=2):
                            rp_pad = rp + [""] * (11 - len(rp))
                            d_p = parse_indo_date(rp_pad[0])
                            if d_p:
                                date_to_perj_row[d_p] = (idx_p, rp_pad)
                                if rp_pad[8].strip():
                                    date_to_nost_perj[d_p] = rp_pad[8].strip()

                        sorted_st = sorted(
                            riwayat_list_global,
                            key=lambda x: int(x["no_urut"]) if str(x["no_urut"]).isdigit() else 999999
                        )

                        rekap_batch = [REKAP_HEADERS]
                        batch_perjadin_map = {}

                        for item in sorted_st:
                            no_u = item["no_urut"]
                            tgl_st_val = item["tgl_pelaksanaan"]
                            perihal_val = item["kegiatan_1"].strip() if item["kegiatan_1"].strip() else strip_permohonan(item["permohonan_1"])
                            if item["pt_2"] and item["kegiatan_2"].strip() and item["kegiatan_2"].strip().lower() != perihal_val.lower():
                                perihal_val = f"{perihal_val} dan {item['kegiatan_2'].strip()}"

                            pt_gab_val = f"{item['pt_1']}, {item['pt_2']}" if item["pt_2"] else item["pt_1"]
                            lok_val, kec_val = extract_lokasi_kecamatan(item["alamat_pabrik_1"], item["pt_1"], perusahaan_data)
                            if item["pt_2"]:
                                _, kec2 = extract_lokasi_kecamatan(item["alamat_pabrik_2"], item["pt_2"], perusahaan_data)
                                if kec2 != "-" and kec2.lower() not in kec_val.lower():
                                    kec_val = f"{kec_val}, {kec2}"

                            pet_gab_val = ", ".join([p["nama"] for p in item["daftar_petugas"]])
                            tgl_pel_val = item["tgl_pelaksanaan"]
                            d_pel_list = parse_indo_date_range(tgl_pel_val)

                            no_st_val = item.get("nomor_st_resmi", "").strip()
                            if not no_st_val and d_pel_list:
                                for d_cand in d_pel_list:
                                    if d_cand in date_to_nost_perj:
                                        no_st_val = date_to_nost_perj[d_cand]
                                        break

                            rekap_batch.append([
                                int(no_u) if str(no_u).isdigit() else no_u,
                                tgl_st_val,
                                perihal_val,
                                pt_gab_val,
                                lok_val,
                                kec_val,
                                pet_gab_val,
                                tgl_pel_val,
                                no_st_val
                            ])

                            for d_pel_obj in d_pel_list:
                                if d_pel_obj and d_pel_obj in date_to_perj_row:
                                    p_rnum, p_rvals = date_to_perj_row[d_pel_obj]
                                    cur_b = p_rvals[1].strip().upper() == "TRUE"
                                    cur_c = p_rvals[2].strip().upper() == "TRUE"
                                    cur_d = p_rvals[3].strip().upper() == "TRUE"
                                    cur_i = p_rvals[8].strip()
                                    is_dk = ("kota malang" in lok_val.lower())
                                    need_b = True if (not is_dk or cur_b) else False
                                    need_c = False if need_b else True
                                    if (need_b != cur_b) or (need_c != cur_c) or cur_d:
                                        p_rvals[1] = "TRUE" if need_b else "FALSE"
                                        p_rvals[2] = "TRUE" if need_c else "FALSE"
                                        p_rvals[3] = "FALSE"
                                        batch_perjadin_map[f"BCD_{p_rnum}"] = {
                                            "range": f"B{p_rnum}:D{p_rnum}",
                                            "values": [[need_b, need_c, False]]
                                        }
                                    if no_st_val:
                                        merged_i = merge_nomor_st_cell(cur_i, no_st_val)
                                        if merged_i != cur_i:
                                            p_rvals[8] = merged_i
                                            batch_perjadin_map[f"I_{p_rnum}"] = {
                                                "range": f"I{p_rnum}",
                                                "values": [[merged_i]]
                                            }

                        batch_perjadin_payload = list(batch_perjadin_map.values())
                        if batch_perjadin_payload:
                            ws_perj.batch_update(batch_perjadin_payload, value_input_option="USER_ENTERED")

                        ws_rekap.clear()
                        ws_rekap.update(f"A1:I{len(rekap_batch)}", rekap_batch, value_input_option="USER_ENTERED")
                        try:
                            ws_rekap.format("A1:I1", {
                                "backgroundColor": {"red": 0.04, "green": 0.10, "blue": 0.18},
                                "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
                                "horizontalAlignment": "CENTER"
                            })
                        except Exception:
                            pass

                        st.toast("Sinkronisasi massal Batch API selesai!", icon="🎉")
                        st.success(
                            f"✅ **Sinkronisasi Massal Batch Selesai!**\n"
                            f"- **{len(rekap_batch) - 1} Surat Tugas** berhasil disusun rapi ke dalam sheet **`{REKAP_SHEET_NAME}`**.\n"
                            f"- **{len(batch_perjadin_payload)} Update Sel** di **`DATA PERJADIN`** otomatis diperbarui sekaligus via `batch_update`."
                        )
                        st.cache_data.clear()
                        st.rerun()
                except Exception as e_sync:
                    st.error(f"Gagal menjalankan sinkronisasi massal: {e_sync}")

    # ==============================================================
    # FOKUS B: 🚗 UPDATE ID & PENCAIRAN PERJADIN (HYBRID SATUAN & SPREADSHEET MODE)
    # ==============================================================
    else:
        with st.container(border=True):
            hd_b1, hd_b2 = st.columns([1.35, 1], gap="small")
            with hd_b1:
                st.markdown(clean_html('<div class="section-badge">🚗 B.1 — Update ID Kegiatan & Pencairan Perjadin</div>'), unsafe_allow_html=True)
                st.caption("Pilih filter antrean di bawah. Gunakan **Mode Form Satuan** untuk update detail / multi-hari, atau **Mode Tabel Edit Massal** untuk paste banyak ID Kegiatan sekaligus.")
            with hd_b2:
                perj_mode_opts = ["📱 Mode Form Satuan", "⚡ Mode Tabel Edit Massal"]
                cur_perj_mode = st.session_state.get("perj_work_mode", perj_mode_opts[0])
                perj_mode_sel = st.radio(
                    "Mode Kerja Perjadin:",
                    options=perj_mode_opts,
                    index=1 if cur_perj_mode == perj_mode_opts[1] else 0,
                    horizontal=True,
                    label_visibility="collapsed"
                )
                st.session_state["perj_work_mode"] = perj_mode_sel

            mode_filter_perj = st.radio(
                "Filter Antrean Hari Penugasan:",
                options=[
                    f"⚠️ Belum Ada ID Perjadin ({jml_belum_perjadin} Hari)",
                    f"⏳ Belum Cair ({jml_belum_cair} Hari)",
                    f"🔥 Semua Hari Penugasan Aktif ({jml_hari_penugasan_aktif} Hari)"
                ],
                index=0 if jml_belum_perjadin > 0 else (1 if jml_belum_cair > 0 else 2),
                horizontal=True,
                key="pill_filter_perj_mode"
            )

            if mode_filter_perj.startswith("⚠️"):
                filtered_perj = [p for p in perjadin_list_global if p["keterangan"] == "BELUM ADA PERJADIN"]
            elif mode_filter_perj.startswith("⏳"):
                filtered_perj = [
                    p for p in perjadin_list_global
                    if (p["luar_kota"] or p["dalam_kota"] or p["no_st"]) and not p["sudah_masuk"]
                ]
            else:
                filtered_perj = [
                    p for p in perjadin_list_global
                    if p["luar_kota"] or p["dalam_kota"] or p["no_st"] or p["keterangan"] == "BELUM ADA PERJADIN"
                ]

            if not filtered_perj:
                st.success("🎉 Tidak ada antrean pada filter ini! Menampilkan 20 hari penugasan aktif terakhir:")
                filtered_perj = [
                    p for p in perjadin_list_global
                    if p["luar_kota"] or p["dalam_kota"] or p["no_st"]
                ][-20:]

            # ------------------------------------------------------
            # MODE 1: 📱 MODE FORM SATUAN (DETAIL & MULTI-HARI + AUTO-SELECT DARI ACTION CENTER)
            # ------------------------------------------------------
            if st.session_state["perj_work_mode"] == "📱 Mode Form Satuan":
                jump_row_p = st.session_state.get("jump_target_perj_row")
                if jump_row_p is not None and not any(p["row_num"] == jump_row_p for p in filtered_perj):
                    for p_all in perjadin_list_global:
                        if p_all["row_num"] == jump_row_p:
                            filtered_perj = [p_all] + filtered_perj
                            break

                opsi_perj = []
                default_perj_idx = 0
                for idx_fp, p in enumerate(filtered_perj):
                    if jump_row_p is not None and p["row_num"] == jump_row_p:
                        default_perj_idx = idx_fp
                    kat = "🚗 Luar Kota" if p["luar_kota"] else ("🏙️ Dalam Kota" if p["dalam_kota"] else "⚪ Belum Centang")
                    ket_badge = (
                        "⚠️ BELUM ADA PERJADIN"
                        if p["keterangan"] == "BELUM ADA PERJADIN"
                        else (f"✅ ID: {p['id_kegiatan']}" if p["id_kegiatan"] else (p["status"] or "Aman"))
                    )
                    st_info = p["no_st"] if p["no_st"] else "Belum No.ST"
                    opsi_perj.append(f"{p['tanggal_str']} | {kat} | {st_info} | {ket_badge}")

                pil_perj_lbl = st.selectbox(
                    f"Pilih Tanggal Pelaksanaan ({len(filtered_perj)} hari ditampilkan):",
                    options=opsi_perj,
                    index=default_perj_idx
                )
                sel_perj = filtered_perj[opsi_perj.index(pil_perj_lbl)]

                ket_pill_class = "info-pill-warn" if sel_perj["keterangan"] == "BELUM ADA PERJADIN" else "info-pill-ok"
                kat_teks = "🚗 LUAR KOTA" if sel_perj["luar_kota"] else ("🏙️ DALAM KOTA" if sel_perj["dalam_kota"] else "⚪ BELUM TERCENTANG")
                st.markdown(clean_html(f"""
                <div class="{ket_pill_class}">
                    🗓️ <b>Tanggal Terpilih:</b> <b>{sel_perj["tanggal_str"]}</b> (Baris GSheet: <b>#{sel_perj["row_num"]}</b>) &nbsp;|&nbsp;
                    📍 <b>Kategori:</b> <b>{kat_teks}</b> &nbsp;|&nbsp;
                    💰 <b>Nominal:</b> <b>Rp {sel_perj["nominal"] or "-"}</b><br>
                    🔢 <b>Nomor ST:</b> <code>{sel_perj["no_st"] or "(Belum diisi)"}</code> &nbsp;|&nbsp;
                    📌 <b>Status Peringatan:</b> <b>{sel_perj["keterangan"] or "✅ AMAN (SUDAH ADA ID PERJADIN)"}</b>
                </div>
                """), unsafe_allow_html=True)

                with st.form(f"form_perjadin_{sel_perj['row_num']}", clear_on_submit=False):
                    cp1, cp2 = st.columns(2, gap="medium")

                    with cp1:
                        with st.container(border=True):
                            st.markdown("**🔑 1. Identitas Nomor ST & ID Kegiatan**")
                            def_ttd_perj = detect_tipe_ttd_from_nost(sel_perj["no_st"])
                            in_ttd_perj = st.radio(
                                "Format Kode Nomor ST:",
                                options=[
                                    "a.n. Kepala Kantor (KBC.1201/KBC.120101)",
                                    "Kepala Kantor Langsung (KBC.1201)"
                                ],
                                index=1 if def_ttd_perj == "Kepala Kantor Langsung" else 0,
                                horizontal=True,
                                key=f"perj_ttd_fmt_{sel_perj['row_num']}"
                            )
                            in_no_st_perj = st.text_input(
                                "Nomor ST (Kolom I — Ketik angka / nomor lengkap):",
                                value=sel_perj["no_st"],
                                placeholder="Contoh: 2041 atau ST-2041/KBC.1201/2026"
                            )
                            in_id_keg = st.text_input(
                                "ID Kegiatan Perjadin (Kolom J)*:",
                                value=sel_perj["id_kegiatan"],
                                placeholder="Paste ID 24 karakter dari Bagian Umum..."
                            )
                            opsi_status_perj = [
                                "KIRIM PERTANGGUNGJAWABAN",
                                "PROSES PENCAIRAN",
                                "SELESAI",
                                "PERJADIN 1",
                                "(Biarkan Kosong / Otomatis Rumus)"
                            ]
                            idx_st_def = opsi_status_perj.index(sel_perj["status"]) if sel_perj["status"] in opsi_status_perj else 0
                            in_status_perj = st.selectbox(
                                "Status Perjadin (Kolom E):",
                                options=opsi_status_perj,
                                index=idx_st_def
                            )

                    with cp2:
                        with st.container(border=True):
                            st.markdown("**📍 2. Kategori Wilayah & Status Pencairan**")
                            st.caption("Sentuh kotak di bawah untuk mencentang kategori wilayah:")
                            cb_col1, cb_col2 = st.columns(2)
                            with cb_col1:
                                in_luar = st.checkbox(
                                    "🚗 Luar Kota (Kab. Malang / Batu)",
                                    value=sel_perj["luar_kota"],
                                    key=f"chk_lk_{sel_perj['row_num']}"
                                )
                            with cb_col2:
                                in_dalam = st.checkbox(
                                    "🏙️ Dalam Kota (Kota Malang)",
                                    value=sel_perj["dalam_kota"],
                                    key=f"chk_dk_{sel_perj['row_num']}"
                                )

                            in_sudah_masuk = st.checkbox(
                                "💰 Dana Sudah Masuk Rekening? (Kolom F)",
                                value=sel_perj["sudah_masuk"],
                                key=f"chk_cair_{sel_perj['row_num']}"
                            )
                            in_tgl_cair = st.text_input(
                                "🗓️ Tanggal Cair (Kolom H — Opsional):",
                                value=sel_perj["tgl_cair"],
                                placeholder="Contoh: 20 Oktober 2026"
                            )
                            apply_same_nost = st.checkbox(
                                "⚡ Terapkan juga ke tanggal lain dengan Nomor ST yang sama (ST Multi-Hari)",
                                value=bool(sel_perj["no_st"].strip()),
                                key=f"chk_multi_{sel_perj['row_num']}"
                            )

                    btn_save_perj = st.form_submit_button(
                        "💾 Simpan ID Perjadin & Perbarui Status",
                        type="primary",
                        use_container_width=True
                    )

                if btn_save_perj:
                    try:
                        with st.spinner("Memperbarui baris di sheet DATA PERJADIN..."):
                            ss = get_spreadsheet()
                            ws_p = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
                            r_p = sel_perj["row_num"]

                            yr_p_row = sel_perj["date_obj"].year if sel_perj.get("date_obj") else 2026
                            no_st_clean_p = format_nomor_st_resmi(
                                in_no_st_perj,
                                yr_p_row,
                                tipe_ttd=("Kepala Kantor Langsung" if "Kepala Kantor Langsung" in in_ttd_perj else "a.n. Kepala Kantor"),
                                auto_adjust_unit=True
                            )
                            tidak_ada_st_val = False if (in_luar or in_dalam or no_st_clean_p or in_id_keg.strip()) else True

                            target_rows_perj = [sel_perj]
                            if apply_same_nost and no_st_clean_p:
                                for cand_p in perjadin_list_global:
                                    if cand_p["row_num"] != r_p and no_st_clean_p in cand_p["no_st"]:
                                        target_rows_perj.append(cand_p)

                            batch_update_perj_b = []
                            for tp in target_rows_perj:
                                tr = tp["row_num"]
                                if in_status_perj == "(Biarkan Kosong / Otomatis Rumus)":
                                    status_cell_val = f'=IF(A{tr}="", "", IF(D{tr}=TRUE, IF(WEEKDAY(A{tr}, 2)>=6, "PREI BOS! MANDEK!", "NGANTOR BOS!"), ""))'
                                else:
                                    status_cell_val = in_status_perj

                                is_primary_row = (tr == r_p)
                                b_val = in_luar if is_primary_row else tp["luar_kota"]
                                c_val = in_dalam if is_primary_row else tp["dalam_kota"]
                                d_val = tidak_ada_st_val if is_primary_row else False
                                i_val = no_st_clean_p if is_primary_row else tp["no_st"]

                                batch_update_perj_b.append({
                                    "range": f"B{tr}:F{tr}",
                                    "values": [[b_val, c_val, d_val, status_cell_val, in_sudah_masuk]]
                                })
                                batch_update_perj_b.append({
                                    "range": f"H{tr}:J{tr}",
                                    "values": [[in_tgl_cair.strip(), i_val, in_id_keg.strip()]]
                                })

                            ws_p.batch_update(batch_update_perj_b, value_input_option="USER_ENTERED")
                            st.session_state["jump_target_perj_row"] = None
                            tgl_updated_list = ", ".join([tp["tanggal_str"] for tp in target_rows_perj])
                            st.toast(f"Data Perjadin ({len(target_rows_perj)} hari) berhasil diperbarui!", icon="🚗")
                            st.success(f"✅ Berhasil mengupdate Perjadin untuk **{len(target_rows_perj)} hari ({tgl_updated_list})**! Status **BELUM ADA PERJADIN** otomatis terupdate.")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as e_sp:
                        st.error(f"Gagal memperbarui DATA PERJADIN: {e_sp}")

            # ------------------------------------------------------
            # MODE 2: ⚡ MODE TABEL EDIT MASSAL (SPREADSHEET MODE - ST.DATA_EDITOR)
            # ------------------------------------------------------
            else:
                st.caption("💡 **Spreadsheet Mode:** Kamu bisa langsung mengetik/paste banyak **ID KEGIATAN** (24 karakter), mengubah **STATUS**, mencentang **SUDAH CAIR?**, dan mengisi **TGL CAIR** sekaligus, lalu klik tombol simpan di bawah:")
                status_opts_editor = [
                    "KIRIM PERTANGGUNGJAWABAN",
                    "PROSES PENCAIRAN",
                    "SELESAI",
                    "PERJADIN 1",
                    "(Otomatis Rumus)"
                ]

                editor_perj_rows = []
                orig_perj_map = {}
                for p in filtered_perj:
                    r_num = int(p["row_num"])
                    wil_str = "🚗 Luar Kota" if p["luar_kota"] else ("🏙️ Dalam Kota" if p["dalam_kota"] else "-")
                    cur_st = p["status"] if p["status"] in status_opts_editor else "(Otomatis Rumus)"
                    orig_perj_map[r_num] = {
                        "p": p,
                        "no_st": p["no_st"],
                        "id_keg": p["id_kegiatan"],
                        "status_ui": cur_st,
                        "cair": bool(p["sudah_masuk"]),
                        "tgl_cair": p["tgl_cair"]
                    }
                    editor_perj_rows.append({
                        "BARIS": r_num,
                        "TANGGAL": p["tanggal_str"],
                        "WILAYAH": wil_str,
                        "NOMINAL": p["nominal"] or str(compute_perjadin_nominal(p)),
                        "NO ST": p["no_st"],
                        "ID KEGIATAN": p["id_kegiatan"],
                        "STATUS": cur_st,
                        "SUDAH CAIR?": bool(p["sudah_masuk"]),
                        "TGL CAIR": p["tgl_cair"],
                        "KETERANGAN": p["keterangan"] or "✅ Aman"
                    })

                edited_perj_df = st.data_editor(
                    editor_perj_rows,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "BARIS": st.column_config.NumberColumn("BARIS", disabled=True, width="small"),
                        "TANGGAL": st.column_config.TextColumn("TANGGAL", disabled=True),
                        "WILAYAH": st.column_config.TextColumn("WILAYAH", disabled=True),
                        "NOMINAL": st.column_config.TextColumn("NOMINAL", disabled=True),
                        "NO ST": st.column_config.TextColumn("NO ST"),
                        "ID KEGIATAN": st.column_config.TextColumn("ID KEGIATAN (PASTE 24 DIGIT)"),
                        "STATUS": st.column_config.SelectboxColumn("STATUS", options=status_opts_editor, required=True),
                        "SUDAH CAIR?": st.column_config.CheckboxColumn("SUDAH CAIR?"),
                        "TGL CAIR": st.column_config.TextColumn("TGL CAIR"),
                        "KETERANGAN": st.column_config.TextColumn("KETERANGAN", disabled=True),
                    },
                    key="data_editor_mass_perjadin"
                )

                if st.button("💾 Simpan Seluruh Perubahan Tabel ke DATA PERJADIN (1x Klik Batch API)", type="primary", use_container_width=True):
                    batch_mass_perj = []
                    updated_dates_mass = []

                    for row_ed in edited_perj_df:
                        tr = int(row_ed["BARIS"])
                        orig = orig_perj_map.get(tr)
                        if not orig:
                            continue
                        p_orig = orig["p"]
                        new_nost_raw = str(row_ed.get("NO ST", "") or "").strip()
                        new_idkeg = str(row_ed.get("ID KEGIATAN", "") or "").strip()
                        new_stat_ui = str(row_ed.get("STATUS", "(Otomatis Rumus)") or "(Otomatis Rumus)")
                        new_cair = bool(row_ed.get("SUDAH CAIR?", False))
                        new_tgl_cair = str(row_ed.get("TGL CAIR", "") or "").strip()

                        if new_idkeg and not orig["id_keg"] and new_stat_ui == "(Otomatis Rumus)":
                            new_stat_ui = "KIRIM PERTANGGUNGJAWABAN"

                        has_changed = (
                            new_nost_raw != orig["no_st"]
                            or new_idkeg != orig["id_keg"]
                            or new_stat_ui != orig["status_ui"]
                            or new_cair != orig["cair"]
                            or new_tgl_cair != orig["tgl_cair"]
                        )
                        if not has_changed:
                            continue

                        yr_r = p_orig["date_obj"].year if p_orig.get("date_obj") else 2026
                        det_ttd = detect_tipe_ttd_from_nost(new_nost_raw or p_orig["no_st"])
                        fmt_nost_p = format_nomor_st_resmi(new_nost_raw, yr_r, tipe_ttd=det_ttd, auto_adjust_unit=False) if new_nost_raw else ""

                        if new_stat_ui == "(Otomatis Rumus)":
                            status_cell_val = f'=IF(A{tr}="", "", IF(D{tr}=TRUE, IF(WEEKDAY(A{tr}, 2)>=6, "PREI BOS! MANDEK!", "NGANTOR BOS!"), ""))'
                        else:
                            status_cell_val = new_stat_ui

                        d_val = False if (p_orig["luar_kota"] or p_orig["dalam_kota"] or fmt_nost_p or new_idkeg) else True

                        batch_mass_perj.append({
                            "range": f"B{tr}:F{tr}",
                            "values": [[p_orig["luar_kota"], p_orig["dalam_kota"], d_val, status_cell_val, new_cair]]
                        })
                        batch_mass_perj.append({
                            "range": f"H{tr}:J{tr}",
                            "values": [[new_tgl_cair, fmt_nost_p, new_idkeg]]
                        })
                        updated_dates_mass.append(p_orig["tanggal_str"])

                    if not updated_dates_mass:
                        st.warning("⚠️ Belum ada perubahan yang terdeteksi pada tabel di atas.")
                    else:
                        try:
                            with st.spinner(f"Menyimpan perubahan {len(updated_dates_mass)} hari ke DATA PERJADIN via Batch API..."):
                                ss = get_spreadsheet()
                                ws_p = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
                                ws_p.batch_update(batch_mass_perj, value_input_option="USER_ENTERED")
                                st.toast(f"{len(updated_dates_mass)} baris DATA PERJADIN berhasil diupdate!", icon="🎉")
                                st.success(f"✅ Berhasil menyimpan perubahan untuk **{len(updated_dates_mass)} Hari Penugasan** sekaligus via Batch API!")
                                st.cache_data.clear()
                                st.rerun()
                        except Exception as e_mass_p:
                            st.error(f"Gagal menyimpan perubahan tabel ke DATA PERJADIN: {e_mass_p}")

        with st.expander(f"📊 Lihat Tabel Monitoring Lengkap DATA PERJADIN ({jml_hari_penugasan_aktif} Hari Aktif)", expanded=False):
            active_perj_table = [
                {
                    "TANGGAL": p["tanggal_str"],
                    "LUAR KOTA": "✅" if p["luar_kota"] else "",
                    "DALAM KOTA": "✅" if p["dalam_kota"] else "",
                    "STATUS": p["status"],
                    "NOMINAL": p["nominal"],
                    "NO ST": p["no_st"],
                    "ID KEGIATAN": p["id_kegiatan"],
                    "KETERANGAN": p["keterangan"],
                    "CAIR": "✅ " + p["tgl_cair"] if p["sudah_masuk"] else "⏳ Belum"
                }
                for p in perjadin_list_global
                if p["luar_kota"] or p["dalam_kota"] or p["no_st"] or p["keterangan"] == "BELUM ADA PERJADIN"
            ]
            st.dataframe(list(reversed(active_perj_table)), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# MENU 4: 🗂️ 4. RIWAYAT, CETAK ULANG & EDIT
# ------------------------------------------------------------------
else:
    # ==========================================
    # TAB 2: MANAJEMEN DATA (RIWAYAT, CETAK ULANG, EDIT & HAPUS)
    # ==========================================
    with st.container(border=True):
        st.markdown(clean_html('<div class="section-badge">🔍 Cari & Pilih Riwayat Surat Tugas</div>'), unsafe_allow_html=True)

        riwayat_list = riwayat_list_global
        if not riwayat_list:
            st.info("Belum ada riwayat surat tugas di database Sheet1 atau data sedang dimuat.")
        else:
            kata_kunci = st.text_input(
                "Filter Cepat (Ketik Nama Pabrik / Nama Petugas / Tanggal / Nomor ST):",
                placeholder="Contoh: Gudang Baru atau Derry atau Oktober...",
                key="filter_riwayat_kw"
            ).strip().lower()

            filtered_riwayat = [
                item for item in riwayat_list
                if not kata_kunci
                or kata_kunci in item["label"].lower()
                or kata_kunci in item["kegiatan_1"].lower()
                or kata_kunci in item["no_surat_1"].lower()
                or kata_kunci in item.get("nomor_st_resmi", "").lower()
            ]

            if not filtered_riwayat:
                st.warning(f"Tidak ditemukan Surat Tugas yang cocok dengan kata kunci '{kata_kunci}'.")
                selected_item = None
            else:
                label_options = [item["label"] for item in filtered_riwayat]
                selected_label = st.selectbox(
                    f"Pilih Surat Tugas ({len(filtered_riwayat)} data ditampilkan):",
                    options=label_options,
                    index=0
                )
                selected_item = next(item for item in filtered_riwayat if item["label"] == selected_label)

    if riwayat_list and selected_item:
        petugas_str = ", ".join([f"{p['nama']} ({p['pangkat']})" for p in selected_item["daftar_petugas"]]) or "-"
        info_pt2 = f"<br>🏢 <b>Pabrik 2:</b> {selected_item['pt_2']} (No: {selected_item['no_surat_2']})" if selected_item["pt_2"] else ""
        no_st_badge = selected_item.get("nomor_st_resmi") or "(Belum diisi dari Satu Kemenkeu)"
        st.markdown(clean_html(f"""
        <div class="info-pill" style="background:#EFF6FF; border-left: 4px solid #1E3A8A; padding: 12px 15px;">
            📌 <b>Ringkasan ST No. #{selected_item['no_urut']}</b> (Baris Spreadsheet: {selected_item['row_num']}) &nbsp;|&nbsp; 🔢 <b>Nomor ST Resmi:</b> {no_st_badge}<br>
            🏢 <b>Pabrik 1:</b> {selected_item['pt_1']} &nbsp;|&nbsp; 📄 <b>No. Surat:</b> {selected_item['no_surat_1']} ({selected_item['tgl_surat_1']}){info_pt2}<br>
            🛠️ <b>Kegiatan:</b> {selected_item['kegiatan_1']} &nbsp;|&nbsp; 🗓️ <b>Waktu:</b> {selected_item['hari']}, {selected_item['tgl_pelaksanaan']}<br>
            👮 <b>Tim Petugas ({len(selected_item['daftar_petugas'])} orang):</b> {petugas_str}
        </div>
        """), unsafe_allow_html=True)

        subtab_cetak, subtab_edit, subtab_hapus = st.tabs([
            "📥 Cetak Ulang Dokumen", 
            "✏️ Edit Data Penugasan", 
            "🗑️ Hapus dari Database"
        ])

        with subtab_cetak:
            with st.container(border=True):
                st.markdown("**Pengaturan Penandatangan untuk Cetak Ulang:**")
                col_r_ttd1, col_r_ttd2 = st.columns(2)

                with col_r_ttd1:
                    tipe_ttd_r = st.radio(
                        "Format Penandatangan:",
                        options=["a.n. Kepala Kantor (Kepala Subbagian Umum)", "Kepala Kantor Langsung"],
                        index=0,
                        key="r_tipe_ttd"
                    )
                    is_plh_r = st.checkbox("Bertindak sebagai Pelaksana Harian (Plh.)", value=False, key="r_is_plh")

                with col_r_ttd2:
                    if tipe_ttd_r == "Kepala Kantor Langsung":
                        if is_plh_r:
                            nama_ttd_r = st.selectbox(
                                "Pilih Pejabat Plh. Kepala Kantor*:",
                                options=[""] + list_pegawai,
                                index=0,
                                key="r_plh_kk"
                            )
                        else:
                            nama_ttd_r = st.text_input("Nama Kepala Kantor Definitif:", value="Johan Pandores", key="r_def_kk")
                    else:
                        if is_plh_r:
                            nama_ttd_r = st.selectbox(
                                "Pilih Pejabat Plh. Kasubag Umum*:",
                                options=[""] + list_pegawai,
                                index=0,
                                key="r_plh_subag"
                            )
                        else:
                            nama_ttd_r = st.text_input("Nama Kasubag Umum Definitif:", value="Kukuh Hanna Prapanca", key="r_def_subag")

                data_cetak_ulang = dict(selected_item)
                data_cetak_ulang["tipe_ttd"] = tipe_ttd_r
                data_cetak_ulang["is_plh"] = is_plh_r
                data_cetak_ulang["nama_pejabat_ttd"] = nama_ttd_r if nama_ttd_r.strip() else "Kukuh Hanna Prapanca"

                try:
                    buf_r = generate_docx_buffer(data_cetak_ulang)
                    file_name_r = f"Cetak_Ulang_ST_{data_cetak_ulang['no_urut']}_{data_cetak_ulang['pt_1'].replace(' ', '_')}.docx"
                    render_direct_docx_download(
                        docx_bytes=buf_r.getvalue(),
                        filename=file_name_r,
                        button_label=f"📥 Unduh Ulang Surat Tugas (.docx) — No. #{data_cetak_ulang['no_urut']}"
                    )
                except Exception as e_r:
                    st.error(f"Gagal menyiapkan dokumen cetak ulang: {e_r}")

                with st.expander("👁️ Pratinjau Dokumen Cetak Ulang (Klik untuk Buka/Tutup)", expanded=True):
                    components.html(render_preview_html(data_cetak_ulang), height=850, scrolling=True)

        with subtab_edit:
            with st.container(border=True):
                st.markdown(f"**✏️ Edit Data Surat Tugas No. #{selected_item['no_urut']} (Baris Spreadsheet: {selected_item['row_num']})**")
                st.caption("Perubahan akan langsung memperbarui `Sheet1`, `REKAP BAGIAN UMUM`, dan `DATA PERJADIN` sekaligus.")

                with st.form(f"form_edit_{selected_item['no_urut']}", clear_on_submit=False):
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        idx_pt1 = list_perusahaan.index(selected_item["pt_1"]) if selected_item["pt_1"] in list_perusahaan else 0
                        edit_pt1 = st.selectbox("Nama Perusahaan / Pabrik 1*", options=list_perusahaan, index=idx_pt1)
                        edit_no_surat1 = st.text_input("Nomor Surat Pemohon 1*", value=selected_item["no_surat_1"])
                        edit_tgl_surat1 = st.text_input("Tanggal Surat 1*", value=selected_item["tgl_surat_1"])
                        edit_permohonan1 = st.text_area("Perihal / Permohonan 1*", value=selected_item["permohonan_1"], height=76)
                        edit_kegiatan1 = st.text_input("Kegiatan Baku 1*", value=selected_item["kegiatan_1"])
                        edit_alamat1 = st.text_area("Alamat Pabrik 1*", value=selected_item["alamat_pabrik_1"], height=76)

                    with col_e2:
                        parsed_existing_range = parse_indo_date_range(selected_item["tgl_pelaksanaan"])
                        def_edit_start = parsed_existing_range[0] if parsed_existing_range else datetime.date.today()
                        def_edit_end = parsed_existing_range[-1] if parsed_existing_range else def_edit_start
                        st.markdown("**🗓️ Ubah Cepat Tanggal Pelaksanaan (1 Hari / Multi-Hari):**")
                        ce_d1, ce_d2 = st.columns(2)
                        with ce_d1:
                            edit_d_mulai = st.date_input("Tgl Mulai:", value=def_edit_start, key=f"ed_dm_{selected_item['no_urut']}")
                        with ce_d2:
                            edit_d_akhir = st.date_input("Tgl Akhir:", value=max(def_edit_start, def_edit_end), min_value=edit_d_mulai, key=f"ed_da_{selected_item['no_urut']}")
                        use_auto_date_edit = st.checkbox(
                            "Gunakan hasil kalkulasi Hari & Tanggal dari pemilih di atas saat Simpan",
                            value=False,
                            key=f"ed_use_auto_{selected_item['no_urut']}",
                            help="Centang ini jika kamu ingin mengubah tanggal pelaksanaan ST menjadi rentang tanggal baru secara otomatis."
                        )
                        auto_h_edit, auto_t_edit, _ = format_hari_tanggal_pelaksanaan(edit_d_mulai, edit_d_akhir)
                        st.caption(f"Pratinjau kalkulasi otomatis: **{auto_h_edit} / {auto_t_edit}**")
                        edit_hari = st.text_input("Hari Pelaksanaan (Teks di ST)*", value=selected_item["hari"])
                        edit_tgl_pelaks = st.text_input("Tanggal Pelaksanaan Tugas (Teks di ST)*", value=selected_item["tgl_pelaksanaan"])
                        def_ttd_edit = detect_tipe_ttd_from_nost(selected_item.get("nomor_st_resmi", ""))
                        opsi_ttd_edit = [
                            "a.n. Kepala Kantor (Kepala Subbagian Umum)",
                            "Kepala Kantor Langsung"
                        ]
                        edit_tipe_ttd_nost = st.radio(
                            "Penandatangan ST (Penentu Format Nomor ST):",
                            options=opsi_ttd_edit,
                            index=1 if def_ttd_edit == "Kepala Kantor Langsung" else 0,
                            horizontal=True,
                            key=f"ed_ttd_nost_{selected_item['no_urut']}"
                        )
                        edit_no_st_resmi = st.text_input(
                            "Nomor ST Resmi (Satu Kemenkeu — Cukup ketik angka misal 2041 atau nomor lengkap):",
                            value=selected_item.get("nomor_st_resmi", "")
                        )

                        cur_p_names = [p["nama"] for p in selected_item["daftar_petugas"] if p["nama"] in list_pegawai]
                        edit_selected_names = st.multiselect(
                            "👮 Tim Petugas yang Ditugaskan (Maks. 6 Orang):",
                            options=list_pegawai,
                            default=cur_p_names,
                            max_selections=6,
                            key=f"edit_ms_p_{selected_item['no_urut']}"
                        )

                    btn_simpan_edit = st.form_submit_button("💾 Simpan Perubahan ke 3 Sheet Sekaligus", type="primary", use_container_width=True)

                if btn_simpan_edit:
                    try:
                        with st.spinner("Memperbarui baris di Sheet1, REKAP BAGIAN UMUM, dan DATA PERJADIN..."):
                            ss = get_spreadsheet()
                            ws = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                            r_num = selected_item["row_num"]

                            edit_petugas_list = []
                            for p_nama_edit in edit_selected_names:
                                dtl = pegawai_data.get(p_nama_edit, {})
                                edit_petugas_list.append({
                                    "nama": p_nama_edit,
                                    "pangkat": dtl.get("pangkat", ""),
                                    "jabatan": dtl.get("jabatan", "")
                                })

                            pad_p = []
                            for pi in range(6):
                                if pi < len(edit_petugas_list):
                                    pad_p.extend([edit_petugas_list[pi]["nama"], edit_petugas_list[pi]["pangkat"], edit_petugas_list[pi]["jabatan"]])
                                else:
                                    pad_p.extend(["", "", ""])

                            if use_auto_date_edit:
                                edit_hari, edit_tgl_pelaks, _ = format_hari_tanggal_pelaksanaan(edit_d_mulai, edit_d_akhir)
                            d_ed_yr = parse_indo_date(edit_tgl_pelaks)
                            yr_ed = d_ed_yr.year if d_ed_yr else 2026
                            no_st_fmt_edit = format_nomor_st_resmi(
                                edit_no_st_resmi,
                                yr_ed,
                                tipe_ttd=edit_tipe_ttd_nost,
                                auto_adjust_unit=True
                            )
                            tgl_st_rec = edit_tgl_pelaks

                            row_update = [
                                selected_item["no_urut"],
                                edit_pt1,
                                edit_no_surat1,
                                edit_tgl_surat1,
                                selected_item["pt_2"],
                                selected_item["no_surat_2"],
                                selected_item["tgl_surat_2"],
                                edit_permohonan1,
                                selected_item["permohonan_2"],
                                *pad_p,
                                edit_kegiatan1,
                                selected_item["kegiatan_2"],
                                edit_hari,
                                edit_tgl_pelaks,
                                selected_item.get("lokasi_idx_1", "1"),
                                edit_alamat1,
                                selected_item.get("lokasi_idx_2", "") if selected_item["pt_2"] else "",
                                selected_item["alamat_pabrik_2"],
                                no_st_fmt_edit,
                                tgl_st_rec
                            ]

                            ws.update(f"A{r_num}:AK{r_num}", [row_update], value_input_option="USER_ENTERED")

                            lok_e, kec_e = extract_lokasi_kecamatan(edit_alamat1, edit_pt1, perusahaan_data)
                            pt_gab_e = f"{edit_pt1}, {selected_item['pt_2']}" if selected_item["pt_2"] else edit_pt1
                            pet_gab_e = ", ".join([p["nama"] for p in edit_petugas_list])
                            sync_st_to_rekap_and_perjadin(
                                ss=ss,
                                no_urut=selected_item["no_urut"],
                                tgl_st_str=tgl_st_rec,
                                perihal_str=edit_kegiatan1,
                                nama_pabrik_gabung=pt_gab_e,
                                lokasi_kab_kota=lok_e,
                                kecamatan_str=kec_e,
                                petugas_gabung_str=pet_gab_e,
                                tgl_pelaksanaan_str=edit_tgl_pelaks,
                                nomor_st_resmi=no_st_fmt_edit
                            )

                            st.toast(f"ST No. #{selected_item['no_urut']} berhasil diperbarui di 3 Sheet!", icon="✅")
                            st.success(f"Berhasil memperbarui data Surat Tugas No. {selected_item['no_urut']}!")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as err_edit:
                        st.error(f"Gagal memperbarui Google Sheets: {err_edit}")

        with subtab_hapus:
            with st.container(border=True):
                st.warning(f"⚠️ **Tindakan Penghapusan Permanen:** Data Surat Tugas **No. #{selected_item['no_urut']} ({selected_item['pt_1']})** pada baris ke-{selected_item['row_num']} spreadsheet akan dihapus secara permanen dari Google Sheets.")

                cek_konfirm = st.checkbox(
                    f"Saya yakin ingin menghapus Surat Tugas No. #{selected_item['no_urut']} dari database.",
                    value=False,
                    key=f"del_chk_{selected_item['no_urut']}"
                )

                if st.button("🗑️ Hapus Baris Ini Sekarang", disabled=not cek_konfirm, type="primary", use_container_width=True):
                    try:
                        with st.spinner("Menghapus baris dari Google Sheets..."):
                            ss = get_spreadsheet()
                            ws = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                            ws.delete_rows(selected_item["row_num"])
                            st.toast(f"ST No. #{selected_item['no_urut']} telah dihapus!", icon="🗑️")
                            st.success(f"Berhasil menghapus Surat Tugas No. {selected_item['no_urut']} dari Google Sheets!")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as err_del:
                        st.error(f"Gagal menghapus baris dari Google Sheets: {err_del}")