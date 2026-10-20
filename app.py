import os
import io
import re
import lzma
import base64
import datetime
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
    initial_sidebar_state="collapsed"
)

# ==========================================
# ARSITEKTUR UI/UX: BENTO GRID + ANTI-DARK-MODE BLEED (100% KONTRAS TINGGI DI HP)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    /* 1. Kunci Palet Light Mode & Variabel Internal Streamlit (100% Kebal Dark Mode Browser HP) */
    :root, html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        color-scheme: light !important;
        --primary-color: #1E3A8A !important;
        --background-color: #F4F7FB !important;
        --secondary-background-color: #FFFFFF !important;
        --text-color: #0F172A !important;
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        background-color: #F4F7FB !important;
        color: #0F172A !important;
    }

    /* Header Top Bar Streamlit agar bersih & tidak memotong judul di HP */
    header[data-testid="stHeader"] {
        background: rgba(244, 247, 251, 0.88) !important;
        backdrop-filter: blur(8px) !important;
        -webkit-backdrop-filter: blur(8px) !important;
        border-bottom: 1px solid #E2E8F0 !important;
    }

    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 1140px !important;
    }

    /* 2. Paksa Seluruh Label Form, Teks, Radio, Checkbox & Heading Kontras Tinggi (Gelap di Atas Putih) */
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
    div[data-testid="stExpander"] summary span {
        color: #0F172A !important;
    }

    div[data-testid="stWidgetLabel"] p {
        font-weight: 700 !important;
        font-size: 0.9rem !important;
        color: #0A192F !important;
        margin-bottom: 2px !important;
    }

    div[data-testid="stCaptionContainer"] p,
    div[data-testid="stCaptionContainer"] span {
        color: #475569 !important;
        font-size: 0.82rem !important;
    }

    /* 3. Arsitektur Bento Grid Sejati via Native Container (Kompatibel Semua Versi Streamlit) */
    div[data-testid="stVerticalBlockBorderWrapper"],
    div[data-testid="stLayoutWrapper"] > div[style*="border"] {
        background-color: #FFFFFF !important;
        border-radius: 18px !important;
        border: 1.5px solid #E2E8F0 !important;
        box-shadow: 0 8px 24px -4px rgba(15, 23, 42, 0.06) !important;
        padding: 18px !important;
        margin-bottom: 0.9rem !important;
    }

    /* 4. Banner Header Instansi & Metric Bento Strip (Soft Glassmorphism) */
    .bc-header-card {
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%) !important;
        color: #FFFFFF !important;
        padding: 1.3rem 1.45rem !important;
        border-radius: 20px !important;
        border-left: 6px solid #F59E0B !important;
        border-bottom: 2px solid rgba(245, 158, 11, 0.5) !important;
        box-shadow: 0 12px 28px -6px rgba(10, 25, 47, 0.25) !important;
        margin-bottom: 0.95rem !important;
    }
    .bc-header-card .bc-header-title {
        font-size: 1.3rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        letter-spacing: -0.2px !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        line-height: 1.35 !important;
    }
    .bc-header-card .bc-header-sub {
        font-size: 0.85rem !important;
        color: #DBEAFE !important;
        -webkit-text-fill-color: #DBEAFE !important;
        margin-top: 0.35rem !important;
        margin-bottom: 0 !important;
        font-weight: 500 !important;
        line-height: 1.4 !important;
    }
    .bc-header-card .metric-strip {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 8px !important;
        margin-top: 0.9rem !important;
    }
    .bc-header-card .metric-pill {
        background: rgba(255, 255, 255, 0.14) !important;
        backdrop-filter: blur(10px) !important;
        -webkit-backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(255, 255, 255, 0.28) !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        padding: 5px 12px !important;
        border-radius: 999px !important;
        font-size: 0.78rem !important;
        font-weight: 600 !important;
    }
    .bc-header-card .metric-pill b {
        color: #FBBF24 !important;
        -webkit-text-fill-color: #FBBF24 !important;
        font-weight: 800 !important;
    }

    /* 5. Header Seksi Bento di Dalam Kartu */
    .section-badge {
        display: inline-flex !important;
        align-items: center !important;
        gap: 8px !important;
        background: linear-gradient(90deg, #EFF6FF 0%, #F8FAFC 100%) !important;
        color: #0A192F !important;
        -webkit-text-fill-color: #0A192F !important;
        padding: 7px 14px !important;
        border-radius: 10px !important;
        font-weight: 800 !important;
        font-size: 0.93rem !important;
        border: 1px solid #DBEAFE !important;
        border-left: 4px solid #1E3A8A !important;
        margin-bottom: 0.8rem !important;
    }

    /* 6. Bento Mini-Cards & Info Pills (Pegawai, Lokasi, Status) */
    .info-pill {
        background-color: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        border-left: 4px solid #1E3A8A !important;
        padding: 10px 14px !important;
        border-radius: 10px !important;
        font-size: 0.84rem !important;
        color: #1E293B !important;
        margin-top: 4px !important;
        margin-bottom: 10px !important;
        line-height: 1.5 !important;
    }
    .info-pill-warn {
        background-color: #FFFBEB !important;
        border: 1px solid #FCD34D !important;
        border-left: 4px solid #F59E0B !important;
        padding: 10px 14px !important;
        border-radius: 10px !important;
        font-size: 0.84rem !important;
        color: #92400E !important;
        margin-top: 4px !important;
        margin-bottom: 10px !important;
        line-height: 1.5 !important;
    }
    .info-pill-ok {
        background-color: #ECFDF5 !important;
        border: 1px solid #6EE7B7 !important;
        border-left: 4px solid #10B981 !important;
        padding: 10px 14px !important;
        border-radius: 10px !important;
        font-size: 0.84rem !important;
        color: #065F46 !important;
        margin-top: 4px !important;
        margin-bottom: 10px !important;
        line-height: 1.5 !important;
    }
    .petugas-bento-grid {
        display: grid !important;
        grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)) !important;
        gap: 10px !important;
        margin-top: 8px !important;
        margin-bottom: 6px !important;
    }
    .petugas-bento-card {
        background: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        border-left: 4px solid #F59E0B !important;
        border-radius: 12px !important;
        padding: 10px 13px !important;
        font-size: 0.83rem !important;
        line-height: 1.45 !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04) !important;
    }
    .petugas-bento-card .p-title {
        font-weight: 800 !important;
        color: #0A192F !important;
        font-size: 0.88rem !important;
        margin-bottom: 3px !important;
    }
    .petugas-bento-card .p-sub {
        color: #334155 !important;
        font-size: 0.79rem !important;
    }

    /* 7. Kolom Input, Selectbox, DateInput, Textarea & Menu Dropdown (Terang, Jelas & Anti Auto-Zoom 16px) */
    input,
    textarea,
    div[data-baseweb="input"],
    div[data-baseweb="base-input"],
    div[data-baseweb="textarea"],
    div[data-baseweb="select"] {
        font-size: 16px !important;
        border-radius: 12px !important;
    }

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"],
    div[data-baseweb="textarea"] {
        min-height: 48px !important;
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1.5px solid #94A3B8 !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04) !important;
    }

    div[data-baseweb="select"] > div:focus-within,
    div[data-baseweb="input"] > div:focus-within,
    div[data-baseweb="textarea"]:focus-within {
        background-color: #FFFFFF !important;
        border-color: #1E3A8A !important;
        box-shadow: 0 0 0 3px rgba(30, 58, 138, 0.15) !important;
    }

    input,
    textarea,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        background-color: transparent !important;
    }

    div[data-baseweb="select"] svg {
        fill: #1E293B !important;
        color: #1E293B !important;
    }

    input::placeholder,
    textarea::placeholder {
        color: #64748B !important;
        -webkit-text-fill-color: #64748B !important;
        opacity: 1 !important;
    }

    /* Tag Multiselect Petugas (Biru Muda Kontras Tinggi) */
    span[data-baseweb="tag"] {
        background-color: #DBEAFE !important;
        border: 1px solid #60A5FA !important;
        border-radius: 8px !important;
        padding: 3px 8px !important;
    }
    span[data-baseweb="tag"] span {
        color: #0A192F !important;
        -webkit-text-fill-color: #0A192F !important;
        font-weight: 700 !important;
        font-size: 0.82rem !important;
    }

    /* Daftar Pilihan Dropdown saat Dibuka di HP */
    div[data-baseweb="popover"],
    ul[data-baseweb="menu"],
    div[role="listbox"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px rgba(15, 23, 42, 0.14) !important;
    }
    li[role="option"],
    div[role="option"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 0.92rem !important;
        padding: 10px 14px !important;
    }
    li[role="option"]:hover,
    li[role="option"][aria-selected="true"],
    div[role="option"]:hover,
    div[role="option"][aria-selected="true"] {
        background-color: #EFF6FF !important;
        color: #1E3A8A !important;
        -webkit-text-fill-color: #1E3A8A !important;
        font-weight: 700 !important;
    }

    /* 8. Segmented Pill Control untuk Tab Navigasi (100% Terbaca Jelas di HP & Bisa Di-Swipe) */
    div[data-baseweb="tab-list"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        white-space: nowrap !important;
        scrollbar-width: none !important;
        -ms-overflow-style: none !important;
        gap: 8px !important;
        background-color: #E2E8F0 !important;
        padding: 6px !important;
        border-radius: 14px !important;
        border: 1px solid #CBD5E1 !important;
        margin-bottom: 1rem !important;
    }
    div[data-baseweb="tab-list"]::-webkit-scrollbar {
        display: none !important;
    }
    button[data-baseweb="tab"] {
        font-weight: 700 !important;
        font-size: 0.88rem !important;
        padding: 10px 16px !important;
        border-radius: 10px !important;
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important;
        transition: all 0.18s ease !important;
        min-height: 44px !important;
        flex-shrink: 0 !important;
    }
    button[data-baseweb="tab"] p,
    button[data-baseweb="tab"] span,
    button[data-baseweb="tab"] div {
        color: #1E293B !important;
        -webkit-text-fill-color: #1E293B !important;
        font-weight: 700 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #0A192F !important;
        border-bottom: 3px solid #F59E0B !important;
        box-shadow: 0 4px 12px rgba(10, 25, 47, 0.22) !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p,
    button[data-baseweb="tab"][aria-selected="true"] span,
    button[data-baseweb="tab"][aria-selected="true"] div {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 800 !important;
    }
    div[data-baseweb="tab-highlight"],
    div[data-baseweb="tab-border"] {
        display: none !important;
    }

    /* 9. Tombol Sekunder (Refresh & Reset) & Tombol Eksekusi Utama (Primary) */
    div.stButton > button,
    div.stDownloadButton > button {
        min-height: 48px !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 0.92rem !important;
        background: linear-gradient(180deg, #FFFFFF 0%, #F1F5F9 100%) !important;
        color: #0A192F !important;
        border: 1.5px solid #94A3B8 !important;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06) !important;
        transition: all 0.18s ease-in-out !important;
    }
    div.stButton > button p,
    div.stButton > button span,
    div.stButton > button div,
    div.stDownloadButton > button p,
    div.stDownloadButton > button span {
        color: #0A192F !important;
        -webkit-text-fill-color: #0A192F !important;
        font-weight: 700 !important;
    }
    div.stButton > button:hover {
        background: #EFF6FF !important;
        border-color: #1E3A8A !important;
    }

    /* Tombol Primary (Simpan & Eksekusi) */
    div.stButton > button[kind="primary"],
    div[data-testid="stFormSubmitButton"] > button,
    div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
        min-height: 54px !important;
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #1E3A8A !important;
        border-bottom: 3.5px solid #F59E0B !important;
        box-shadow: 0 8px 20px -4px rgba(30, 58, 138, 0.35) !important;
        font-size: 0.95rem !important;
    }
    div.stButton > button[kind="primary"] p,
    div.stButton > button[kind="primary"] span,
    div.stButton > button[kind="primary"] div,
    div[data-testid="stFormSubmitButton"] > button p,
    div[data-testid="stFormSubmitButton"] > button span,
    div[data-testid="stFormSubmitButton"] > button div {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 800 !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div[data-testid="stFormSubmitButton"] > button:hover {
        background: linear-gradient(135deg, #112240 0%, #2563EB 100%) !important;
        box-shadow: 0 10px 24px -4px rgba(37, 99, 235, 0.45) !important;
        transform: translateY(-1px);
    }

    /* Sidebar tetap bersih & terang */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }

    @media (max-width: 640px) {
        .block-container {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 3.3rem !important;
        }
        .bc-header-card {
            padding: 1.15rem 1.15rem !important;
            border-radius: 16px !important;
        }
        .bc-header-card .bc-header-title {
            font-size: 1.12rem !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            padding: 15px !important;
            border-radius: 16px !important;
        }
        button[data-baseweb="tab"] {
            padding: 9px 13px !important;
            font-size: 0.83rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)

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
    "jul": 7, "agu": 8, "ags": 8, "sep": 9, "okt": 10, "nov": 11, "des": 12
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
    """Mengubah berbagai format string tanggal Indonesia / ISO menjadi datetime.date (mengambil tanggal pertama jika berupa rentang)."""
    if not val:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if not s:
        return None
    # Coba format "15 Oktober 2026"
    m = re.search(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s)
    if m:
        day = int(m.group(1))
        mon_str = m.group(2).lower()
        yr = int(m.group(3))
        if mon_str in BULAN_MAP_REV:
            try:
                return datetime.date(yr, BULAN_MAP_REV[mon_str], day)
            except Exception:
                pass
    # Coba format rentang ringkas "21 s.d. 23 Oktober 2026"
    m_short = re.search(r'(\d{1,2})\s*(?:s\.?d\.?|s/d|-|–)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s, flags=re.IGNORECASE)
    if m_short:
        day1 = int(m_short.group(1))
        mon_str = m_short.group(3).lower()
        yr = int(m_short.group(4))
        if mon_str in BULAN_MAP_REV:
            try:
                return datetime.date(yr, BULAN_MAP_REV[mon_str], day1)
            except Exception:
                pass
    # Coba format ISO "2026-10-15"
    m2 = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m2:
        try:
            return datetime.date(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except Exception:
            pass
    # Coba format DD/MM/YYYY
    m3 = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})', s)
    if m3:
        try:
            return datetime.date(int(m3.group(3)), int(m3.group(2)), int(m3.group(1)))
        except Exception:
            pass
    return None

def parse_indo_date_range(val) -> list:
    """
    Mengekstrak seluruh daftar datetime.date harian dari string tanggal tunggal maupun rentang tanggal.
    Contoh:
    - '21 Oktober 2026' -> [date(2026, 10, 21)]
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

    # 1. Cek jika ada >= 2 tanggal lengkap (misal: "21 Oktober 2026 s.d. 23 Oktober 2026" atau "30 Oktober 2026 s.d. 2 November 2026")
    all_full = re.findall(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s)
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
    m_short = re.search(r'(\d{1,2})\s*(?:s\.?d\.?|s/d|-|–)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', s, flags=re.IGNORECASE)
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

    # 3. Fallback tanggal tunggal
    d_one = parse_indo_date(s)
    return [d_one] if d_one else []

def strip_permohonan(text: str) -> str:
    if not text:
        return ""
    clean = re.sub(r'^(permohonan\s+|permohonan:\s*)', '', text.strip(), flags=re.IGNORECASE).strip()
    return clean

def format_nomor_st_resmi(raw_st: str, tahun: int = 2026) -> str:
    """
    Memudahkan user: cukup ketik angka (misal '2013' atau 'ST-2013')
    maka otomatis menjadi 'ST-2013/KBC.1201/KBC.120101/2026'.
    Jika sudah lengkap, biarkan apa adanya.
    """
    if not raw_st or not str(raw_st).strip():
        return ""
    s = str(raw_st).strip()
    if "/" in s:
        return s
    m = re.match(r'^(?:ST[-\s]*)?(\d+)$', s, flags=re.IGNORECASE)
    if m:
        angka = m.group(1)
        return f"ST-{angka}/KBC.1201/KBC.120101/{tahun}"
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
        "in_tgl_pelaks_mulai", "in_tgl_pelaks_akhir", "in_tgl_st_rekap"
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

        rekap_payload = [
            no_urut,
            tgl_st_str,
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
            ws_rekap.append_row(rekap_payload, value_input_option="USER_ENTERED")
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

                    new_luar = True if not is_dalam_kota else cur_luar
                    new_dalam = True if is_dalam_kota else cur_dalam

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
                        if cur_no_st and nomor_st_resmi not in cur_no_st:
                            combined_st = f"{cur_no_st}, {nomor_st_resmi}"
                        else:
                            combined_st = nomor_st_resmi
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

@st.cache_data(ttl=45)
def load_riwayat_data():
    try:
        ss = get_spreadsheet()
        ws = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
        all_rows = ws.get_all_values()
        if len(all_rows) <= 1:
            return []
        
        data_rows = all_rows[1:]
        riwayat = []
        for orig_idx, r_raw in enumerate(data_rows):
            if not r_raw or not r_raw[0].strip():
                continue
            sheet_row_num = orig_idx + 2
            r = r_raw + [""] * (37 - len(r_raw))
            
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

@st.cache_data(ttl=45)
def load_perjadin_and_rekap_data():
    """Memuat data dari sheet DATA PERJADIN dan REKAP BAGIAN UMUM untuk Tab 3."""
    ss = get_spreadsheet()
    perjadin_rows = []
    rekap_rows = []

    # 1. Baca DATA PERJADIN
    try:
        ws_p = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
        raw_p = ws_p.get_all_values()
        for idx, r in enumerate(raw_p[1:], start=2):
            r_pad = r + [""] * (11 - len(r))
            tgl_str = r_pad[0].strip()
            if not tgl_str or tgl_str.upper() == "TOTAL":
                continue
            d_obj = parse_indo_date(tgl_str)
            perjadin_rows.append({
                "row_num": idx,
                "tanggal_str": tgl_str,
                "date_obj": d_obj,
                "luar_kota": r_pad[1].strip().upper() == "TRUE",
                "dalam_kota": r_pad[2].strip().upper() == "TRUE",
                "tidak_ada_st": r_pad[3].strip().upper() == "TRUE",
                "status": r_pad[4].strip(),
                "sudah_masuk": r_pad[5].strip().upper() == "TRUE",
                "nominal": r_pad[6].strip(),
                "tgl_cair": r_pad[7].strip(),
                "no_st": r_pad[8].strip(),
                "id_kegiatan": r_pad[9].strip(),
                "keterangan": r_pad[10].strip()
            })
    except Exception:
        pass

    # 2. Baca REKAP BAGIAN UMUM (jika sudah ada)
    try:
        all_titles = [w.title.strip().upper() for w in ss.worksheets()]
        if any(t in ("REKAP BAGIAN UMUM", "LAPORAN BAGIAN UMUM", "REKAP ST") for t in all_titles):
            ws_r = get_or_create_rekap_worksheet(ss)
            raw_r = ws_r.get_all_values()
            for idx, r in enumerate(raw_r[1:], start=2):
                r_pad = r + [""] * (9 - len(r))
                if not r_pad[0].strip() and not r_pad[3].strip():
                    continue
                rekap_rows.append({
                    "row_num": idx,
                    "no": r_pad[0].strip(),
                    "tanggal": r_pad[1].strip(),
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
jml_belum_perjadin = sum(1 for p in perjadin_list_global if p["keterangan"] == "BELUM ADA PERJADIN")
jml_st_belum_nomor = sum(1 for r in rekap_list_global if not r["nomor_st"])

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
# HEADER BANNER EKSEKUTIF & QUICK ACTION BAR
# ==========================================
st.markdown(f"""
<div class="bc-header-card">
    <div class="bc-header-title">📋 Sistem Otomasi ST & Monitoring Perjadin — KPPBC TMC Malang</div>
    <div class="bc-header-sub">Arsitektur 1x Input Terpadu: Database ST (Sheet1) • Sheet REKAP BAGIAN UMUM • Sheet DATA PERJADIN</div>
    <div class="metric-strip">
        <span class="metric-pill">📌 No. ST Berikutnya: <b>#{next_no_preview}</b></span>
        <span class="metric-pill">🗂️ Total Riwayat: <b>{len(riwayat_list_global)} ST</b></span>
        <span class="metric-pill">⚠️ Belum Ada ID Perjadin: <b>{jml_belum_perjadin} Hari</b></span>
        <span class="metric-pill">📝 Menunggu No. ST Satu Kemenkeu: <b>{jml_st_belum_nomor} ST</b></span>
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Kontrol & Utilitas")
    if st.button("🔄 Segarkan Data Spreadsheet", use_container_width=True):
        st.cache_data.clear()
        st.toast("Data berhasil disegarkan dari Google Sheets!", icon="✅")
        st.rerun()
    if st.button("🧹 Bersihkan / Reset Form Input", use_container_width=True):
        reset_form_input(list_pegawai)
        st.toast("Form input telah direset!", icon="🧹")
        st.rerun()
    st.divider()
    st.markdown("**💡 Alur Otomasi 3-Sheet (Zero Double Work):**")
    st.caption("1️⃣ **Saat Buat ST Baru:** Data otomatis masuk ke `Sheet1` + `REKAP BAGIAN UMUM` (Lokasi Kab/Kota & Kecamatan terisi otomatis) + mencentang tanggal di `DATA PERJADIN` (otomatis memicu *BELUM ADA PERJADIN*).")
    st.caption("2️⃣ **Saat Nomor ST Keluar dari Satu Kemenkeu:** Buka Tab 3, ketik angkanya saja (misal `2013`), klik Simpan -> otomatis terisi di `REKAP BAGIAN UMUM` & `DATA PERJADIN`, siap di-copy untuk Bagian Umum.")
    st.caption("3️⃣ **Saat ID Perjadin Keluar:** Input ID Kegiatan di Tab 3 -> tulisan *BELUM ADA PERJADIN* otomatis hilang & status berubah!")

# Quick Action Bar Ringkas & Elegan
with st.container(border=True):
    qa_col1, qa_col2, qa_col3 = st.columns([2.2, 1, 1])
    with qa_col1:
        st.markdown(
            f"<div style='padding-top:6px; font-size:0.86rem; color:#334155;'>"
            f"⚡ <b>Quick Action Bar:</b> Siap membuat Surat Tugas <b>#{next_no_preview}</b> &nbsp;|&nbsp; "
            f"<span style='color:#059669; font-weight:700;'>Terhubung ke Google Sheets</span></div>",
            unsafe_allow_html=True
        )
    with qa_col2:
        if st.button("🔄 Refresh Data Sheets", use_container_width=True, key="btn_top_refresh"):
            st.cache_data.clear()
            st.toast("Sinkronisasi ulang Google Sheets selesai!", icon="✅")
            st.rerun()
    with qa_col3:
        if st.button("🧹 Reset Form Baru", use_container_width=True, key="btn_top_reset"):
            reset_form_input(list_pegawai)
            st.toast("Form siap untuk pengisian baru!", icon="✨")
            st.rerun()

tab_input, tab_riwayat, tab_perjadin = st.tabs([
    "📝 1. Buat Surat Tugas Baru",
    f"🗂️ 2. Riwayat & Cetak Ulang ({len(riwayat_list_global)})",
    f"🚗 3. Monitoring Perjadin & Rekap Umum ({jml_belum_perjadin} Belum ID)"
])

# ==========================================
# TAB 1: FORM INPUT SURAT TUGAS BARU (OTOMATIS MASUK KE 3 SHEET)
# ==========================================
with tab_input:
    # -------------------------------------------------------------
    # KARTU 1: DATA SURAT PEMOHON & LOKASI PABRIK (SMART FILTER + QUICK PRESET)
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown('<div class="section-badge">1️⃣ Data Surat Pemohon & Lokasi Pabrik</div>', unsafe_allow_html=True)
        
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
                    f'<div class="info-pill">📍 <b>Alamat Lengkap Pabrik 1 (Titik {pilihan_lokasi_1}):</b><br>{alamat_pabrik_1}<br>'
                    f'🗺️ <b>Deteksi Otomatis Rekap & Perjadin:</b> Lokasi = <b>{lokasi_kab_1}</b> | Kecamatan = <b>{kec_1}</b> | Kategori = <b>{kat_perjadin}</b></div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(f'<div class="info-pill-warn">⚠️ Semua titik lokasi pada <b>{pt_1}</b> masih kosong di database DATA NPPBKC.</div>', unsafe_allow_html=True)

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
                    st.markdown(f'<div class="info-pill">📍 <b>Alamat Lengkap Pabrik 2 (Titik {pilihan_lokasi_2}):</b><br>{alamat_pabrik_2}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="info-pill-warn">⚠️ Titik Lokasi {pilihan_lokasi_2} pada <b>{pt_2}</b> kosong di database.</div>', unsafe_allow_html=True)

    # -------------------------------------------------------------
    # KARTU 2: WAKTU PELAKSANAAN (MENDUKUNG 1 HARI MAUPUN RENTANG MULTI-HARI) & NOMOR ST
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown('<div class="section-badge">2️⃣ Waktu Pelaksanaan (1 Hari / Multi-Hari) & Nomor ST</div>', unsafe_allow_html=True)
        st.caption("💡 **Tips Cepat HP:** Jika kegiatan **1 hari**, cukup pilih **Tanggal Mulai** saja (Tanggal Selesai otomatis mengikuti). Jika **lebih dari 1 hari**, ubah **Tanggal Selesai** sesuai akhir kegiatan.")

        if "in_tgl_pelaks_mulai" not in st.session_state:
            st.session_state["in_tgl_pelaks_mulai"] = datetime.date.today()
        if "in_tgl_pelaks_akhir" not in st.session_state:
            st.session_state["in_tgl_pelaks_akhir"] = st.session_state["in_tgl_pelaks_mulai"]

        # Jaga agar tanggal akhir tidak pernah lebih kecil dari tanggal mulai
        if st.session_state["in_tgl_pelaks_akhir"] < st.session_state["in_tgl_pelaks_mulai"]:
            st.session_state["in_tgl_pelaks_akhir"] = st.session_state["in_tgl_pelaks_mulai"]

        c_wkt1, c_wkt2, c_wkt3, c_wkt4 = st.columns([1, 1, 1, 1.25])
        with c_wkt1:
            tgl_st_rekap = st.date_input(
                "Tanggal ST / Lapor*",
                value=datetime.date.today(),
                key="in_tgl_st_rekap"
            )
        with c_wkt2:
            tgl_pelaksanaan_mulai = st.date_input(
                "Tanggal Mulai Kegiatan*",
                key="in_tgl_pelaks_mulai",
                on_change=on_tgl_mulai_change
            )
        with c_wkt3:
            tgl_pelaksanaan_akhir = st.date_input(
                "Tanggal Akhir Kegiatan*",
                min_value=tgl_pelaksanaan_mulai,
                key="in_tgl_pelaks_akhir"
            )
        with c_wkt4:
            no_st_opsional = st.text_input(
                "No. ST Satu Kemenkeu (Opsional):",
                placeholder="Kosongkan / ketik misal: 2013",
                key="in_nost_opsional",
                help="Karena Nomor ST baru didapat dari Satu Kemenkeu setelah TTE, kamu boleh mengosongkan kolom ini sekarang dan mengisinya nanti lewat Tab 3 dengan 1 klik."
            )

        tgl_pelaksanaan = tgl_pelaksanaan_mulai
        nama_hari_id, tgl_indo_str, durasi_hari = format_hari_tanggal_pelaksanaan(
            tgl_pelaksanaan_mulai,
            tgl_pelaksanaan_akhir
        )
        tgl_st_indo_str = format_tanggal_indo(tgl_st_rekap)
        preview_no_st_fmt = format_nomor_st_resmi(no_st_opsional, tgl_pelaksanaan_mulai.year) if no_st_opsional else "(Menunggu dari Satu Kemenkeu — Bisa diupdate di Tab 3)"

        badge_durasi = (
            f"<span style='background:#DBEAFE; color:#1E3A8A; padding:2px 8px; border-radius:6px; font-weight:800;'>1 Hari</span>"
            if durasi_hari == 1
            else f"<span style='background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:6px; font-weight:800;'>{durasi_hari} Hari Berturut-turut (Otomatis Centang {durasi_hari} Baris di DATA PERJADIN)</span>"
        )
        st.markdown(
            f'<div class="info-pill">🗓️ <b>Tampilan di Surat Tugas & Database:</b><br>'
            f'• <b>hari / tanggal :</b> <b>{nama_hari_id} / {tgl_indo_str}</b> &nbsp;{badge_durasi}<br>'
            f'• <b>Format Nomor ST :</b> {preview_no_st_fmt}</div>',
            unsafe_allow_html=True
        )

    # -------------------------------------------------------------
    # KARTU 3: TIM PETUGAS PEMERIKSA (MULTISELECT INSTAN + BENTO MINI-CARDS)
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown('<div class="section-badge">3️⃣ Tim Petugas Pemeriksa (Instan Tanpa Loading Berulang)</div>', unsafe_allow_html=True)
        
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
            st.markdown(
                f'<div class="petugas-bento-grid">{"".join(bento_cards_html)}</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="info-pill-warn">⚠️ Belum ada petugas yang dipilih. Silakan pilih minimal 1 pegawai pada kotak di atas.</div>',
                unsafe_allow_html=True
            )

    # -------------------------------------------------------------
    # KARTU 4: KONFIGURASI PENANDATANGAN
    # -------------------------------------------------------------
    with st.container(border=True):
        st.markdown('<div class="section-badge">4️⃣ Pejabat Penandatangan Surat Tugas</div>', unsafe_allow_html=True)
        
        c_ttd_left, c_ttd_right = st.columns([1.2, 1])
        with c_ttd_left:
            tipe_ttd = st.radio(
                "Format Jabatan Penandatangan:",
                options=["a.n. Kepala Kantor (Kepala Subbagian Umum)", "Kepala Kantor Langsung"],
                index=0
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
            
            st.markdown(
                f'<div class="info-pill">✍️ <b>Format TTE:</b> {lbl_jab_ringkas}<br>👤 <b>Pejabat:</b> {nama_pejabat_ttd or "(Belum dipilih)"}</div>',
                unsafe_allow_html=True
            )

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
                    
                    col_no = ws_target.col_values(1)
                    existing_numbers = [int(x) for x in col_no if x.strip().isdigit()]
                    next_no = max(existing_numbers) + 1 if existing_numbers else 1

                    petugas_row_data = []
                    for p_idx in range(6):
                        if p_idx < len(data_petugas_terpilih):
                            p = data_petugas_terpilih[p_idx]
                            petugas_row_data.extend([p["nama"], p["pangkat"], p["jabatan"]])
                        else:
                            petugas_row_data.extend(["", "", ""])

                    nomor_st_clean = format_nomor_st_resmi(no_st_opsional, tgl_pelaksanaan.year)

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

                    # 1. Simpan ke Sheet1 (Database Surat Tugas)
                    ws_target.append_row(row_payload, value_input_option="USER_ENTERED")

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
            except Exception as err:
                st.error(f"Gagal menyimpan ke Google Sheets: {err}")

    # Kartu Unduh Langsung & Pratinjau Opsional (Menggunakan Base64 Murni Anti-Gagal Download)
    if "last_submitted_data" in st.session_state:
        data_gen = st.session_state["last_submitted_data"]
        with st.container(border=True):
            st.markdown(f'<div class="section-badge">📄 Dokumen Siap Diunduh — ST No. #{data_gen["no_urut"]} ({data_gen["pt_1"]})</div>', unsafe_allow_html=True)
            
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

# ==========================================
# TAB 2: MANAJEMEN DATA (RIWAYAT, CETAK ULANG, EDIT & HAPUS)
# ==========================================
with tab_riwayat:
    with st.container(border=True):
        st.markdown('<div class="section-badge">🔍 Cari & Pilih Riwayat Surat Tugas</div>', unsafe_allow_html=True)
        
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
        st.markdown(f"""
        <div class="info-pill" style="background:#EFF6FF; border-left: 4px solid #1E3A8A; padding: 12px 15px;">
            📌 <b>Ringkasan ST No. #{selected_item['no_urut']}</b> (Baris Spreadsheet: {selected_item['row_num']}) &nbsp;|&nbsp; 🔢 <b>Nomor ST Resmi:</b> {no_st_badge}<br>
            🏢 <b>Pabrik 1:</b> {selected_item['pt_1']} &nbsp;|&nbsp; 📄 <b>No. Surat:</b> {selected_item['no_surat_1']} ({selected_item['tgl_surat_1']}){info_pt2}<br>
            🛠️ <b>Kegiatan:</b> {selected_item['kegiatan_1']} &nbsp;|&nbsp; 🗓️ <b>Waktu:</b> {selected_item['hari']}, {selected_item['tgl_pelaksanaan']}<br>
            👮 <b>Tim Petugas ({len(selected_item['daftar_petugas'])} orang):</b> {petugas_str}
        </div>
        """, unsafe_allow_html=True)

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
                        edit_no_st_resmi = st.text_input("Nomor ST Resmi (Satu Kemenkeu):", value=selected_item.get("nomor_st_resmi", ""))
                        
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
                            no_st_fmt_edit = format_nomor_st_resmi(edit_no_st_resmi)
                            tgl_st_rec = selected_item.get("tgl_st_resmi") or edit_tgl_surat1 or edit_tgl_pelaks

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

# ==========================================
# TAB 3: MONITORING PERJADIN & REKAP BAGIAN UMUM (ZERO DOUBLE WORK)
# ==========================================
with tab_perjadin:
    st.markdown("""
    <div class="info-pill" style="background:#EFF6FF; border-left: 4px solid #1E3A8A; padding: 11px 15px;">
        💡 <b>Arsitektur Zero-Double-Work (Cukup 1x Kerja):</b><br>
        • <b>Tahap 1 (Saat Buat ST di Tab 1):</b> Data otomatis masuk ke <code>REKAP BAGIAN UMUM</code> & mencentang tanggal di <code>DATA PERJADIN</code> (memicu pengingat <b>BELUM ADA PERJADIN</b>).<br>
        • <b>Tahap 2 (Saat Nomor ST Keluar dari Satu Kemenkeu):</b> Ketik angkanya saja di bawah (misal <code>2013</code>) -> otomatis masuk ke <code>REKAP BAGIAN UMUM</code> & <code>DATA PERJADIN</code>, lalu tinggal Copy tabelnya untuk dikirim ke Bagian Umum.<br>
        • <b>Tahap 3 (Saat ID Kegiatan Perjadin Keluar dari Bagian Umum):</b> Input <code>ID KEGIATAN</code> di bawah -> peringatan <b>BELUM ADA PERJADIN</b> otomatis hilang & status berubah!
    </div>
    """, unsafe_allow_html=True)

    sub_rekap_st, sub_id_perjadin, sub_sync_massal = st.tabs([
        f"📋 A. Input Nomor ST & Copy Laporan Bagian Umum ({len(rekap_list_global)} Baris)",
        f"🚗 B. Update ID Kegiatan & Status Perjadin ({jml_belum_perjadin} Belum ID)",
        "🔄 C. Sinkronisasi / Tarik Otomatis Data Lama"
    ])

    # ------------------------------------------------------------------
    # SUB-TAB A: INPUT CEPAT NOMOR ST (SATU KEMENKEU) & COPY KE BAGIAN UMUM
    # ------------------------------------------------------------------
    with sub_rekap_st:
        with st.container(border=True):
            st.markdown('<div class="section-badge">1️⃣ Isi Cepat Nomor ST dari Satu Kemenkeu (Otomatis Update 3 Sheet)</div>', unsafe_allow_html=True)
            st.caption("Pilih ST yang baru mendapat nomor dari aplikasi Satu Kemenkeu. Cukup ketik angkanya saja (contoh: `2013`), sistem otomatis melengkapi jadi `ST-2013/KBC.1201/KBC.120101/2026` dan mengupdate `Sheet1`, `REKAP BAGIAN UMUM`, serta `DATA PERJADIN`!")

            if not riwayat_list_global:
                st.info("Belum ada data Surat Tugas.")
            else:
                filter_st_mode = st.radio(
                    "Tampilkan Daftar ST:",
                    options=["Hanya yang Belum Ada Nomor ST / Baru", "Semua Riwayat ST"],
                    index=0,
                    horizontal=True
                )

                rekap_map_by_no = {str(r["no"]).strip(): r for r in rekap_list_global}
                kandidat_st = []
                for item in riwayat_list_global:
                    no_u = str(item["no_urut"]).strip()
                    cur_nost = item.get("nomor_st_resmi") or rekap_map_by_no.get(no_u, {}).get("nomor_st", "")
                    item_copy = dict(item)
                    item_copy["cur_nost"] = cur_nost
                    if filter_st_mode == "Hanya yang Belum Ada Nomor ST / Baru":
                        if not cur_nost:
                            kandidat_st.append(item_copy)
                    else:
                        kandidat_st.append(item_copy)

                if not kandidat_st:
                    st.success("🎉 Mantap bro! Semua Surat Tugas sudah memiliki Nomor ST resmi.")
                    kandidat_st = [dict(item, cur_nost=rekap_map_by_no.get(str(item["no_urut"]).strip(), {}).get("nomor_st", "")) for item in riwayat_list_global[:20]]

                opsi_st_update = [
                    f"No. #{it['no_urut']} | {it['tgl_pelaksanaan']} | {it['pt_1']} ({'✅ ' + it['cur_nost'] if it['cur_nost'] else '⏳ Belum Ada Nomor ST'})"
                    for it in kandidat_st
                ]
                pil_st_lbl = st.selectbox("Pilih Surat Tugas:", options=opsi_st_update, index=0, key="sel_fast_nost")
                sel_st_obj = kandidat_st[opsi_st_update.index(pil_st_lbl)]

                c_fn1, c_fn2, c_fn3 = st.columns([1.2, 1, 1])
                with c_fn1:
                    input_angka_st = st.text_input(
                        "Ketik Nomor ST dari Satu Kemenkeu*:",
                        value=sel_st_obj.get("cur_nost", ""),
                        placeholder="Ketik angka saja misal: 2013 atau ST-2013/...",
                        key=f"fast_nost_in_{sel_st_obj['no_urut']}"
                    )
                with c_fn2:
                    default_tgl_lapor = sel_st_obj.get("tgl_st_resmi") or sel_st_obj.get("tgl_surat_1") or sel_st_obj.get("tgl_pelaksanaan")
                    input_tgl_st = st.text_input(
                        "Tanggal ST (Kolom TANGGAL Rekap):",
                        value=default_tgl_lapor,
                        key=f"fast_tglst_in_{sel_st_obj['no_urut']}"
                    )
                with c_fn3:
                    default_petugas_rekap = rekap_map_by_no.get(str(sel_st_obj["no_urut"]).strip(), {}).get("petugas", "")
                    if not default_petugas_rekap:
                        default_petugas_rekap = ", ".join([p["nama"] for p in sel_st_obj["daftar_petugas"]])
                    input_petugas_rekap = st.text_input(
                        "Nama Petugas di Rekap Umum (Bisa Disingkat):",
                        value=default_petugas_rekap,
                        help="Contoh: Gasyudha, Davin, Aji, Yudha",
                        key=f"fast_pet_in_{sel_st_obj['no_urut']}"
                    )

                if st.button("💾 Simpan Nomor ST ke Sheet1 + REKAP BAGIAN UMUM + DATA PERJADIN", type="primary", use_container_width=True):
                    if not input_angka_st.strip():
                        st.error("⚠️ Masukkan Nomor ST terlebih dahulu!")
                    else:
                        try:
                            with st.spinner("Mengupdate Nomor ST di ketiga sheet sekaligus..."):
                                ss = get_spreadsheet()
                                d_pel = parse_indo_date(sel_st_obj["tgl_pelaksanaan"])
                                yr_pel = d_pel.year if d_pel else 2026
                                no_st_final = format_nomor_st_resmi(input_angka_st, yr_pel)

                                # 1. Update Sheet1 kolom AJ & AK (Kolom 36 & 37)
                                ws_s1 = find_worksheet_flexible(ss, ["Sheet1", "SURAT TUGAS", "DATABASE ST"])
                                ws_s1.update(
                                    f"AJ{sel_st_obj['row_num']}:AK{sel_st_obj['row_num']}",
                                    [[no_st_final, input_tgl_st]],
                                    value_input_option="USER_ENTERED"
                                )

                                # 2. Update REKAP BAGIAN UMUM & DATA PERJADIN
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
                                st.toast(f"Nomor ST {no_st_final} berhasil disinkronkan!", icon="✅")
                                st.success(f"✅ **Nomor ST `{no_st_final}`** berhasil disimpan ke `Sheet1`, `{REKAP_SHEET_NAME}`, dan `DATA PERJADIN`!")
                                st.cache_data.clear()
                                st.rerun()
                        except Exception as e_upst:
                            st.error(f"Gagal menyimpan Nomor ST: {e_upst}")

        # TABEL REKAP BAGIAN UMUM & FITUR COPY CEPAT
        with st.container(border=True):
            st.markdown('<div class="section-badge">2️⃣ Tabel Sheet REKAP BAGIAN UMUM (Siap Lapor / Copy ke Bagian Umum)</div>', unsafe_allow_html=True)
            if not rekap_list_global:
                st.warning("Sheet `REKAP BAGIAN UMUM` masih kosong. Klik tombol di Sub-Tab C (`🔄 Sinkronisasi / Tarik Otomatis Data Lama`) untuk menarik seluruh data ST dari `Sheet1` secara otomatis!")
            else:
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
                        for r in reversed(rekap_list_global)
                    ],
                    use_container_width=True,
                    hide_index=True
                )

                with st.expander("📋 Salin / Copy Teks Laporan ke Bagian Umum (Format Tabel Excel / WhatsApp)", expanded=False):
                    tsv_lines = ["\t".join(REKAP_HEADERS)]
                    for r in rekap_list_global[-20:]:
                        tsv_lines.append("\t".join([
                            str(r["no"]), r["tanggal"], r["perihal"], r["nama_pabrik"],
                            r["lokasi"], r["kecamatan"], r["petugas"], r["tgl_pelaksanaan"], r["nomor_st"]
                        ]))
                    st.caption("Copy blok di bawah ini (20 ST terbaru) untuk ditempel langsung ke Excel / WhatsApp Bagian Umum:")
                    st.code("\n".join(tsv_lines), language="text")

    # ------------------------------------------------------------------
    # SUB-TAB B: UPDATE ID KEGIATAN & STATUS PERJADIN (OTOMATIS HILANGKAN 'BELUM ADA PERJADIN')
    # ------------------------------------------------------------------
    with sub_id_perjadin:
        with st.container(border=True):
            st.markdown('<div class="section-badge">🚗 Update ID Kegiatan & Status Perjadin (Sheet DATA PERJADIN)</div>', unsafe_allow_html=True)
            st.caption("Begitu Bagian Umum menerbitkan **ID Kegiatan Perjadin**, masukkan di sini. Saat `ID KEGIATAN` terisi, tulisan **BELUM ADA PERJADIN** otomatis hilang dan Status berubah!")

            mode_filter_perj = st.radio(
                "Filter Hari Penugasan di DATA PERJADIN:",
                options=[
                    f"⚠️ Hanya yang 'BELUM ADA PERJADIN' ({jml_belum_perjadin} Hari)",
                    "🚗 Semua Hari yang Ada ST (Luar Kota / Dalam Kota)",
                    "📅 Cari Berdasarkan Semua Tanggal Kalender"
                ],
                index=0 if jml_belum_perjadin > 0 else 1,
                horizontal=True
            )

            if mode_filter_perj.startswith("⚠️"):
                filtered_perj = [p for p in perjadin_list_global if p["keterangan"] == "BELUM ADA PERJADIN"]
            elif mode_filter_perj.startswith("🚗"):
                filtered_perj = [p for p in perjadin_list_global if p["luar_kota"] or p["dalam_kota"] or p["no_st"]]
            else:
                filtered_perj = perjadin_list_global

            if not filtered_perj:
                st.success("🎉 Tidak ada antrean 'BELUM ADA PERJADIN'! Semua penugasan sudah memiliki ID Perjadin.")
                filtered_perj = [p for p in perjadin_list_global if p["luar_kota"] or p["dalam_kota"]][-15:]

            if filtered_perj:
                opsi_perj = []
                for p in filtered_perj:
                    kat = "Luar Kota" if p["luar_kota"] else ("Dalam Kota" if p["dalam_kota"] else "-")
                    ket_badge = "⚠️ BELUM ADA PERJADIN" if p["keterangan"] == "BELUM ADA PERJADIN" else (f"✅ ID: {p['id_kegiatan']}" if p["id_kegiatan"] else p["status"])
                    st_info = p["no_st"] if p["no_st"] else "Belum No.ST"
                    opsi_perj.append(f"{p['tanggal_str']} | {kat} | {st_info} | {ket_badge}")

                pil_perj_lbl = st.selectbox("Pilih Tanggal Pelaksanaan di DATA PERJADIN:", options=opsi_perj, index=0)
                sel_perj = filtered_perj[opsi_perj.index(pil_perj_lbl)]

                st.markdown(
                    f'<div class="info-pill">🗓️ <b>Tanggal:</b> {sel_perj["tanggal_str"]} (Baris {sel_perj["row_num"]}) &nbsp;|&nbsp; '
                    f'💰 <b>Nominal Otomatis:</b> {sel_perj["nominal"] or "-"} &nbsp;|&nbsp; '
                    f'📌 <b>Keterangan Saat Ini:</b> <b>{sel_perj["keterangan"] or "Aman (Sudah Ada Perjadin)"}</b></div>',
                    unsafe_allow_html=True
                )

                with st.form(f"form_perjadin_{sel_perj['row_num']}", clear_on_submit=False):
                    cp1, cp2 = st.columns(2)
                    with cp1:
                        in_no_st_perj = st.text_input("Nomor ST (Kolom I):", value=sel_perj["no_st"], placeholder="Contoh: ST-2013/KBC.1201/KBC.120101/2026")
                        in_id_keg = st.text_input("ID Kegiatan Perjadin (Kolom J)*:", value=sel_perj["id_kegiatan"], placeholder="Masukkan ID 24 karakter dari Bagian Umum...")
                        opsi_status_perj = [
                            "KIRIM PERTANGGUNGJAWABAN",
                            "PROSES PENCAIRAN",
                            "SELESAI",
                            "PERJADIN 1",
                            "(Biarkan Kosong / Otomatis Rumus)"
                        ]
                        idx_st_def = opsi_status_perj.index(sel_perj["status"]) if sel_perj["status"] in opsi_status_perj else 0
                        in_status_perj = st.selectbox("Status Perjadin (Kolom E):", options=opsi_status_perj, index=idx_st_def)
                    with cp2:
                        in_luar = st.checkbox("Luar Kota (Kabupaten Malang / Batu)", value=sel_perj["luar_kota"])
                        in_dalam = st.checkbox("Dalam Kota (Kota Malang)", value=sel_perj["dalam_kota"])
                        in_sudah_masuk = st.checkbox("Dana Sudah Masuk Rekening? (Kolom F)", value=sel_perj["sudah_masuk"])
                        in_tgl_cair = st.text_input("Tanggal Cair (Kolom H - Opsional):", value=sel_perj["tgl_cair"], placeholder="Contoh: 15 Oktober 2026")
                        apply_same_nost = st.checkbox(
                            "⚡ Terapkan juga ID Kegiatan & Status ini ke seluruh tanggal lain yang memiliki Nomor ST yang sama (Praktis untuk ST Multi-Hari)",
                            value=bool(sel_perj["no_st"].strip())
                        )

                    btn_save_perj = st.form_submit_button("💾 Simpan ID Perjadin & Perbarui Status", type="primary", use_container_width=True)

                if btn_save_perj:
                    try:
                        with st.spinner("Memperbarui baris di sheet DATA PERJADIN..."):
                            ss = get_spreadsheet()
                            ws_p = find_worksheet_flexible(ss, ["DATA PERJADIN", "PERJADIN"])
                            r_p = sel_perj["row_num"]

                            no_st_clean_p = format_nomor_st_resmi(in_no_st_perj)
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

                                # Update Kolom B, C, D, E, F (Jangan timpa Kolom G karena berisi rumus NOMINAL!)
                                batch_update_perj_b.append({
                                    "range": f"B{tr}:F{tr}",
                                    "values": [[b_val, c_val, d_val, status_cell_val, in_sudah_masuk]]
                                })
                                # Update Kolom H, I, J (Jangan timpa Kolom K karena berisi rumus KETERANGAN!)
                                batch_update_perj_b.append({
                                    "range": f"H{tr}:J{tr}",
                                    "values": [[in_tgl_cair.strip(), i_val, in_id_keg.strip()]]
                                })

                            ws_p.batch_update(batch_update_perj_b, value_input_option="USER_ENTERED")
                            tgl_updated_list = ", ".join([tp["tanggal_str"] for tp in target_rows_perj])
                            st.toast(f"Data Perjadin ({len(target_rows_perj)} hari) berhasil diperbarui!", icon="🚗")
                            st.success(f"✅ Berhasil mengupdate Perjadin untuk **{len(target_rows_perj)} hari ({tgl_updated_list})**! Status **BELUM ADA PERJADIN** otomatis terupdate.")
                            st.cache_data.clear()
                            st.rerun()
                    except Exception as e_sp:
                        st.error(f"Gagal memperbarui DATA PERJADIN: {e_sp}")

            st.divider()
            st.markdown("**📊 Tabel Monitoring DATA PERJADIN (Hari Penugasan Aktif):**")
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
    # SUB-TAB C: SINKRONISASI MASSAL DATA LAMA (BATCH API OPTIMIZED)
    # ------------------------------------------------------------------
    with sub_sync_massal:
        with st.container(border=True):
            st.markdown('<div class="section-badge">🔄 Tarik & Sinkronkan Otomatis Seluruh Riwayat ST (Batch API 10x Lebih Cepat)</div>', unsafe_allow_html=True)
            st.markdown("""
            Fitur ini dirancang khusus agar kamu **tidak perlu mengetik ulang** data Surat Tugas yang sudah pernah diinput di `Sheet1`:
            1. Otomatis membuat sheet baru bernama **`REKAP BAGIAN UMUM`** (jika belum ada) lengkap dengan 9 kolom (`NO`, `TANGGAL`, `PERIHAL`, `Nama Pabrik`, `LOKASI`, `Kecamatan`, `PETUGAS`, `TANGGAL PELAKSANAAN`, `NOMOR ST`).
            2. Mengekstrak otomatis **Kabupaten/Kota** dan **Kecamatan** dari alamat pabrik untuk seluruh riwayat ST di `Sheet1`.
            3. Mencocokkan `NOMOR ST` yang sudah tercatat di `DATA PERJADIN` agar langsung masuk ke `REKAP BAGIAN UMUM`.
            4. Menggunakan **Batch API Update (`ws_perj.batch_update`)** untuk mencentang sekaligus seluruh tanggal `LUAR KOTA` / `DALAM KOTA` di `DATA PERJADIN` dalam 1x panggilan jaringan (10x lebih cepat & bebas limit kuota API Google Sheets).
            """)

            if st.button("🚀 Jalankan Sinkronisasi Otomatis Seluruh Data Sheet1 -> REKAP BAGIAN UMUM & DATA PERJADIN", type="primary", use_container_width=True):
                try:
                    with st.spinner("Sedang memproses & menyinkronkan seluruh riwayat ST secara Batch ke REKAP BAGIAN UMUM dan DATA PERJADIN..."):
                        ss = get_spreadsheet()
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
                            tgl_st_val = item.get("tgl_st_resmi") or item.get("tgl_surat_1") or item.get("tgl_pelaksanaan")
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
                                    is_dk = ("kota malang" in lok_val.lower())
                                    need_b = True if not is_dk else cur_b
                                    need_c = True if is_dk else cur_c
                                    if (need_b != cur_b) or (need_c != cur_c) or cur_d:
                                        p_rvals[1] = "TRUE" if need_b else "FALSE"
                                        p_rvals[2] = "TRUE" if need_c else "FALSE"
                                        p_rvals[3] = "FALSE"
                                        batch_perjadin_map[p_rnum] = {
                                            "range": f"B{p_rnum}:D{p_rnum}",
                                            "values": [[need_b, need_c, False]]
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
                            f"- **{len(batch_perjadin_payload)} Tanggal Pelaksanaan** di **`DATA PERJADIN`** otomatis diperbarui sekaligus via `batch_update`."
                        )
                        st.cache_data.clear()
                        st.rerun()
                except Exception as e_sync:
                    st.error(f"Gagal menjalankan sinkronisasi massal: {e_sync}")
