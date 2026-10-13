import io
import re
import datetime
import streamlit as st
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# KONFIGURASI HALAMAN (MOBILE-FIRST)
# ==========================================
st.set_page_config(
    page_title="Generator Surat Tugas KPPBC TMC Malang",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

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

def format_tanggal_indo(d: datetime.date) -> str:
    if not d:
        return ""
    return f"{d.day} {BULAN_INDONESIA[d.month]} {d.year}"

def strip_permohonan(text: str) -> str:
    if not text:
        return ""
    # Menghapus kata "permohonan" atau "permohonan:" di awal teks
    clean = re.sub(r'^(permohonan\s+|permohonan:\s*)', '', text.strip(), flags=re.IGNORECASE).strip()
    return clean

# Inisialisasi Session State untuk Efisiensi Mobile & Otomatisasi
if "num_petugas" not in st.session_state:
    st.session_state["num_petugas"] = 1

if "in_kegiatan_1" not in st.session_state:
    st.session_state["in_kegiatan_1"] = ""
if "prev_permohonan_1" not in st.session_state:
    st.session_state["prev_permohonan_1"] = ""

if "in_kegiatan_2" not in st.session_state:
    st.session_state["in_kegiatan_2"] = ""
if "prev_permohonan_2" not in st.session_state:
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

# ==========================================
# KONEKSI & CACHING GOOGLE SHEETS
# ==========================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
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
    sheet_id = st.secrets.get("SPREADSHEET_ID", "1lMZmVmMXZO1SRq2JuAWLLN3EhoNHLuYWvd265WySVww")
    return client.open_by_key(sheet_id)

@st.cache_data(ttl=600)
def load_master_data():
    ss = get_spreadsheet()
    
    # 1. Master Pegawai (DATA PEGAWAI: B4:F111)
    ws_pegawai = ss.worksheet("DATA PEGAWAI")
    raw_pegawai = ws_pegawai.get("B4:F111")
    
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

    # 2. Master Perusahaan (DATA NPPBKC: B2:AK179)
    ws_nppbkc = ss.worksheet("DATA NPPBKC")
    raw_nppbkc = ws_nppbkc.get("B2:AK179")
    
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

@st.cache_data(ttl=60)
def load_riwayat_data():
    try:
        ss = get_spreadsheet()
        ws = ss.worksheet("Sheet1")
        all_rows = ws.get_all_values()
        if len(all_rows) <= 1:
            return []
        
        data_rows = all_rows[1:]
        riwayat = []
        for r_raw in reversed(data_rows):
            if not r_raw or not r_raw[0].strip():
                continue
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
                "alamat_pabrik_1": r[32].strip(),
                "alamat_pabrik_2": r[34].strip(),
                "daftar_petugas": daftar_petugas
            })
        return riwayat
    except Exception as e:
        return []

try:
    pegawai_data, perusahaan_data = load_master_data()
except Exception as e:
    st.error(f"Gagal memuat master data dari Google Sheets: {e}")
    pegawai_data, perusahaan_data = {}, {}

list_perusahaan = sorted(list(perusahaan_data.keys())) if perusahaan_data else []
list_pegawai = sorted(list(pegawai_data.keys())) if pegawai_data else []

# ==========================================
# FUNGSI HELPER: GENERATE WORD (.DOCX)
# ==========================================
def generate_docx_buffer(data_gen):
    template_file = "FORMAT ST.docx"
    doc = docx.Document(template_file)
    
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

    # 4. Table 1 (Hari / Tanggal & Tempat)
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
        r_dan = p_t.add_run(f" serta {data_gen['pt_2']}, {data_gen['alamat_pabrik_2']}")
        r_dan.font.name = "Arial"

    # 5. Table 2 (Tanda Tangan)
    t_ttd = doc.tables[2]
    c_prefix = t_ttd.rows[1].cells[0]
    c_jab = t_ttd.rows[1].cells[1]
    
    if data_gen["tipe_ttd"] == "Kepala Kantor Langsung":
        if data_gen["is_plh"]:
            c_prefix.text = "Plh."
            c_jab.text = "Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang"
        else:
            c_prefix.text = ""
            c_jab.text = "Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang"
    else:
        c_prefix.text = "a.n"
        if data_gen["is_plh"]:
            c_jab.text = "Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang\nPlh. Kepala Subbagian Umum"
        else:
            c_jab.text = "Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang\nKepala Subbagian Umum"
            
    # Set RATA KANAN pada c_prefix agar a.n/Plh menempel tepat di samping c_jab
    for p in c_prefix.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for r in p.runs:
            r.font.name = "Arial"

    for p in c_jab.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs:
            r.font.name = "Arial"
                
    c_nama = t_ttd.rows[3].cells[1]
    c_nama.text = data_gen["nama_pejabat_ttd"]
    for p in c_nama.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"

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
        tempat_full += f" serta {data['pt_2']}, {data['alamat_pabrik_2']}"

    if data["tipe_ttd"] == "Kepala Kantor Langsung":
        prefix_txt = "Plh." if data["is_plh"] else ""
        jab_txt = "Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang"
        temb_txt = "1. Kepala Subbagian Umum<br>2. Kepala Seksi Pelayanan Kepabeanan dan Cukai II<br>3. Kepala Seksi Kepatuhan Internal"
    else:
        prefix_txt = "a.n"
        sub_jab = "Plh. Kepala Subbagian Umum" if data["is_plh"] else "Kepala Subbagian Umum"
        jab_txt = f"Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang<br>{sub_jab}"
        temb_txt = "1. Kepala Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang<br>2. Kepala Subbagian Umum<br>3. Kepala Seksi Pelayanan Kepabeanan dan Cukai II<br>4. Kepala Seksi Kepatuhan Internal"

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
    
    <p>
      Segala biaya yang timbul sebagai akibat dilaksanakannya surat ini dibebankan pada DIPA Kantor Pengawasan dan Pelayanan Bea dan Cukai Tipe Madya Cukai Malang
    </p>
    
    <p class="indent">
      Surat Tugas ini disusun untuk dilaksanakan dan setelah selesai dilaksanakan, pelaksana segera menyampaikan laporan. Kepada instansi terkait, kami mohon bantuan demi kelancaran pelaksanaan tugas tersebut.
    </p>
    
    <!-- BLOK TANDA TANGAN 100% PRESISI VERTIKAL -->
    <table style="width: 100%; border-collapse: collapse; margin-top: 25px; margin-bottom: 20px;">
      <tr>
        <td style="width: 42%;"></td>
        <td style="width: 58%;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="width: 35px;"></td>
              <td style="text-align: left; padding-bottom: 12px;">Malang, [@TanggalND]</td>
            </tr>
            <tr>
              <td style="width: 35px; text-align: right; vertical-align: top; padding-right: 8px; white-space: nowrap;">
                {prefix_txt}
              </td>
              <td style="text-align: left; vertical-align: top;">
                {jab_txt}
                <div style="margin-top: 25px; color: #555; font-size: 11.5px;">Ditandatangani secara elektronik</div>
                <div style="margin-top: 3px; font-weight: bold; color: #000;">{data['nama_pejabat_ttd']}</div>
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
# INTERFACE UTAMA
# ==========================================
st.title("📋 Sistem Otomasi Surat Tugas KPPBC TMC Malang")
st.caption("Input Real-Time Database Spreadsheet & Generator Dokumen Word (.docx) Resmi")

with st.sidebar:
    st.header("⚙️ Pengaturan & Kontrol")
    if st.button("🔄 Segarkan Data (Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.markdown("**Petunjuk Pengisian Cepat:**")
    st.caption("1. Isi Permohonan, maka Kegiatan Baku otomatis terisi.")
    st.caption("2. Gunakan tombol [+] jika ada lebih dari 1 petugas atau 2 pabrik.")
    st.caption("3. Klik Simpan & Unduh Dokumen.")

tab_input, tab_riwayat = st.tabs(["📝 Input Surat Tugas Baru", "🕒 Riwayat & Cetak Ulang"])

# ==========================================
# TAB 1: FORM INPUT SURAT TUGAS BARU (EFISIEN MOBILE)
# ==========================================
with tab_input:
    # -------------------------------------------------------------
    # 1. DATA SURAT PEMOHON
    # -------------------------------------------------------------
    st.subheader("1. Data Surat Pemohon")
    
    st.markdown("##### 🏢 Pabrik 1 (Pemohon Utama)")
    pt_1 = st.selectbox("Nama Perusahaan / Pabrik 1*", options=[""] + list_perusahaan, index=0)
    
    c_p1_1, c_p1_2 = st.columns([1.2, 1])
    with c_p1_1:
        no_surat_1 = st.text_input("Nomor Surat Pemohon 1*", placeholder="Contoh: 63/IX/BBSA/2026")
    with c_p1_2:
        tgl_surat_1 = st.date_input("Tanggal Surat 1*", value=datetime.date.today())
        
    permohonan_1 = st.text_area(
        "Perihal / Permohonan 1*",
        key="in_permohonan_1",
        placeholder="Permohonan Pembukaan Segel Hasil Tembakau...",
        on_change=on_permohonan_1_change,
        height=75
    )
    
    kegiatan_1 = st.text_input(
        "Kegiatan Baku 1 (Otomatis dari Permohonan, Bisa Diedit)*",
        key="in_kegiatan_1",
        placeholder="Pembukaan Segel Hasil Tembakau..."
    )
    
    pilihan_lokasi_1 = st.selectbox("Pilihan Titik Lokasi Pabrik 1 (1 - 12)", options=list(range(1, 13)), index=0)
    alamat_pabrik_1 = ""
    if pt_1 and pt_1 in perusahaan_data:
        alamat_pabrik_1 = perusahaan_data[pt_1].get(pilihan_lokasi_1, "")
    st.text_area("Alamat Terpilih Pabrik 1", value=alamat_pabrik_1, disabled=True, height=65)

    # OPSI PABRIK 2 (DINAMIS UNTUK MOBILE)
    tambah_pt2 = st.checkbox("➕ Tambah Pabrik Kedua (Jika 1 Surat Tugas untuk 2 Pabrik)", value=False)
    
    pt_2, no_surat_2, tgl_surat_2, permohonan_2, kegiatan_2, alamat_pabrik_2 = "", "", None, "", "", ""
    if tambah_pt2:
        st.markdown("---")
        st.markdown("##### 🏢 Pabrik 2 (Pemohon Kedua)")
        pt_2 = st.selectbox("Nama Perusahaan / Pabrik 2*", options=[""] + list_perusahaan, index=0)
        c_p2_1, c_p2_2 = st.columns([1.2, 1])
        with c_p2_1:
            no_surat_2 = st.text_input("Nomor Surat Pemohon 2*", placeholder="Nomor surat kedua")
        with c_p2_2:
            tgl_surat_2 = st.date_input("Tanggal Surat 2*", value=None)
            
        permohonan_2 = st.text_area(
            "Perihal / Permohonan 2*",
            key="in_permohonan_2",
            placeholder="Permohonan pabrik kedua...",
            on_change=on_permohonan_2_change,
            height=75
        )
        kegiatan_2 = st.text_input(
            "Kegiatan Baku 2 (Otomatis dari Permohonan 2, Bisa Diedit)*",
            key="in_kegiatan_2"
        )
        pilihan_lokasi_2 = st.selectbox("Pilihan Titik Lokasi Pabrik 2 (1 - 12)", options=list(range(1, 13)), index=0)
        if pt_2 and pt_2 in perusahaan_data:
            alamat_pabrik_2 = perusahaan_data[pt_2].get(pilihan_lokasi_2, "")
        st.text_area("Alamat Terpilih Pabrik 2", value=alamat_pabrik_2, disabled=True, height=65)

    st.divider()

    # -------------------------------------------------------------
    # 2. WAKTU PELAKSANAAN TUGAS
    # -------------------------------------------------------------
    st.subheader("2. Waktu Pelaksanaan Tugas")
    c_wkt1, c_wkt2 = st.columns(2)
    with c_wkt1:
        tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Tugas*", value=datetime.date.today())
    with c_wkt2:
        nama_hari_en = tgl_pelaksanaan.strftime("%A")
        nama_hari_id = HARI_INDONESIA.get(nama_hari_en, nama_hari_en)
        st.text_input("Hari Pelaksanaan (Otomatis)", value=nama_hari_id, disabled=True)

    st.divider()

    # -------------------------------------------------------------
    # 3. PERSONEL TIM PETUGAS (DINAMIS & EFISIEN MOBILE)
    # -------------------------------------------------------------
    st.subheader("3. Tim Petugas Pemeriksa")
    st.caption("Secara bawaan hanya menampilkan Petugas 1. Gunakan tombol [+] jika ada petugas tambahan.")

    data_petugas_terpilih = []
    
    # Render hanya sejumlah petugas yang aktif
    for i in range(1, st.session_state["num_petugas"] + 1):
        st.markdown(f"**👤 Petugas {i}**")
        sel_nama = st.selectbox(
            f"Nama Petugas {i}*",
            options=["-- Pilih Pegawai --"] + list_pegawai,
            index=0,
            key=f"dyn_petugas_{i}"
        )
        
        p_nip, p_pangkat, p_jabatan = "", "", ""
        if sel_nama and sel_nama != "-- Pilih Pegawai --":
            detail = pegawai_data.get(sel_nama, {})
            p_nip = detail.get("nip", "")
            p_pangkat = detail.get("pangkat", "")
            p_jabatan = detail.get("jabatan", "")
            data_petugas_terpilih.append({
                "no": len(data_petugas_terpilih) + 1,
                "nama": sel_nama,
                "nip": p_nip,
                "pangkat": p_pangkat,
                "jabatan": p_jabatan
            })
            
        cp1, cp2 = st.columns(2)
        cp1.text_input(f"Pangkat/Gol {i}", value=p_pangkat, disabled=True, key=f"dyn_pkt_{i}")
        cp2.text_input(f"Jabatan {i}", value=p_jabatan, disabled=True, key=f"dyn_jbt_{i}")
        st.write("")

    # Tombol Tambah/Kurangi Petugas (Mobile Friendly)
    col_tambah_p, col_kurang_p = st.columns(2)
    with col_tambah_p:
        if st.button("➕ Tambah Petugas", disabled=(st.session_state["num_petugas"] >= 6), use_container_width=True):
            st.session_state["num_petugas"] += 1
            st.rerun()
    with col_kurang_p:
        if st.button("➖ Kurangi Petugas", disabled=(st.session_state["num_petugas"] <= 1), use_container_width=True):
            st.session_state["num_petugas"] -= 1
            st.rerun()

    st.divider()

    # -------------------------------------------------------------
    # 4. KONFIGURASI PENANDATANGAN
    # -------------------------------------------------------------
    st.subheader("4. Konfigurasi Penandatangan Surat Tugas")
    tipe_ttd = st.radio(
        "Format Penandatangan:",
        options=["a.n. Kepala Kantor (Kepala Subbagian Umum)", "Kepala Kantor Langsung"],
        index=0
    )
    is_plh = st.checkbox("Bertindak sebagai Pelaksana Harian (Plh.)", value=False)

    if tipe_ttd == "Kepala Kantor Langsung":
        if is_plh:
            st.info("Penandatangan: Plh. Kepala Kantor")
            nama_pejabat_ttd = st.selectbox(
                "Pilih Pejabat yang Ditunjuk sebagai Plh. Kepala Kantor*:",
                options=[""] + list_pegawai,
                index=0
            )
        else:
            nama_pejabat_ttd = st.text_input("Nama Kepala Kantor Definitif:", value="Johan Pandores")
    else:
        if is_plh:
            st.info("Penandatangan: a.n. Kepala Kantor ... Plh. Kepala Subbagian Umum")
            nama_pejabat_ttd = st.selectbox(
                "Pilih Pejabat yang Ditunjuk sebagai Plh. Kasubag Umum*:",
                options=[""] + list_pegawai,
                index=0
            )
        else:
            nama_pejabat_ttd = st.text_input("Nama Kasubag Umum Definitif:", value="Kukuh Hanna Prapanca")

    st.divider()
    
    # -------------------------------------------------------------
    # TOMBOL AKSI UTAMA
    # -------------------------------------------------------------
    btn_simpan_eksekusi = st.button("💾 Simpan ke Database Spreadsheet", type="primary", use_container_width=True)

    if btn_simpan_eksekusi:
        if not pt_1:
            st.error("Nama Perusahaan / Pabrik 1 wajib dipilih!")
        elif not no_surat_1:
            st.error("Nomor Surat Pemohon 1 wajib diisi!")
        elif len(data_petugas_terpilih) == 0:
            st.error("Minimal harus memilih Petugas 1!")
        elif not nama_pejabat_ttd.strip():
            st.error("Nama Pejabat Penandatangan tidak boleh kosong!")
        else:
            try:
                with st.spinner("Menyimpan ke Google Sheets..."):
                    ss = get_spreadsheet()
                    ws_target = ss.worksheet("Sheet1")
                    
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

                    row_payload = [
                        next_no,
                        pt_1,
                        no_surat_1,
                        format_tanggal_indo(tgl_surat_1),
                        pt_2 if pt_2 else "",
                        no_surat_2 if no_surat_2 else "",
                        format_tanggal_indo(tgl_surat_2) if tgl_surat_2 else "",
                        permohonan_1,
                        permohonan_2 if permohonan_2 else "",
                        *petugas_row_data,
                        kegiatan_1,
                        kegiatan_2 if kegiatan_2 else "",
                        nama_hari_id,
                        format_tanggal_indo(tgl_pelaksanaan),
                        pilihan_lokasi_1,
                        alamat_pabrik_1,
                        pilihan_lokasi_2 if pt_2 else "",
                        alamat_pabrik_2 if pt_2 else "",
                        "",
                        ""
                    ]

                    ws_target.append_row(row_payload, value_input_option="USER_ENTERED")
                    st.success(f"Berhasil menyimpan Surat Tugas No. Urut {next_no} ke Database Spreadsheet!")
                    st.cache_data.clear()
                    
                    st.session_state["last_submitted_data"] = {
                        "no_urut": next_no,
                        "pt_1": pt_1,
                        "no_surat_1": no_surat_1,
                        "tgl_surat_1": format_tanggal_indo(tgl_surat_1),
                        "pt_2": pt_2,
                        "no_surat_2": no_surat_2,
                        "tgl_surat_2": format_tanggal_indo(tgl_surat_2) if tgl_surat_2 else "",
                        "permohonan_1": permohonan_1,
                        "permohonan_2": permohonan_2,
                        "kegiatan_1": kegiatan_1,
                        "kegiatan_2": kegiatan_2,
                        "hari": nama_hari_id,
                        "tgl_pelaksanaan": format_tanggal_indo(tgl_pelaksanaan),
                        "alamat_pabrik_1": alamat_pabrik_1,
                        "alamat_pabrik_2": alamat_pabrik_2,
                        "daftar_petugas": data_petugas_terpilih,
                        "tipe_ttd": tipe_ttd,
                        "is_plh": is_plh,
                        "nama_pejabat_ttd": nama_pejabat_ttd
                    }
            except Exception as err:
                st.error(f"Gagal menyimpan ke Google Sheets: {err}")

    # Pratinjau & Tombol Unduh untuk Input Baru
    if "last_submitted_data" in st.session_state:
        data_gen = st.session_state["last_submitted_data"]
        st.divider()
        st.subheader("📄 Pratinjau & Unduh Surat Tugas Baru")
        
        with st.expander("👁️ Pratinjau Dokumen Surat Tugas (Klik untuk Buka/Tutup)", expanded=True):
            components.html(render_preview_html(data_gen), height=850, scrolling=True)
            
        try:
            buf = generate_docx_buffer(data_gen)
            file_name_out = f"ST_{data_gen['pt_1'].replace(' ', '_')}_{data_gen['tgl_pelaksanaan'].replace(' ', '_')}.docx"
            st.download_button(
                label="📥 Unduh Surat Tugas (.docx)",
                data=buf,
                file_name=file_name_out,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                key="btn_dl_new"
            )
        except Exception as e_dl:
            st.error(f"Gagal membuat file Word: {e_dl}")

# ==========================================
# TAB 2: RIWAYAT & CETAK ULANG SURAT TUGAS
# ==========================================
with tab_riwayat:
    st.subheader("🕒 Riwayat Penugasan & Cetak Ulang")
    st.caption("Pilih penugasan dari database yang pernah direkam untuk melihat pratinjau dan mengunduh ulang dokumen resminya.")
    
    riwayat_list = load_riwayat_data()
    
    if not riwayat_list:
        st.info("Belum ada riwayat surat tugas di database Sheet1 atau data sedang dimuat.")
    else:
        label_options = [item["label"] for item in riwayat_list]
        selected_label = st.selectbox("Cari & Pilih Surat Tugas:", options=label_options, index=0)
        
        selected_item = next(item for item in riwayat_list if item["label"] == selected_label)
        
        st.divider()
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

        with st.expander("👁️ Pratinjau Dokumen Cetak Ulang (Klik untuk Buka/Tutup)", expanded=True):
            components.html(render_preview_html(data_cetak_ulang), height=850, scrolling=True)

        try:
            buf_r = generate_docx_buffer(data_cetak_ulang)
            file_name_r = f"Cetak_Ulang_ST_{data_cetak_ulang['no_urut']}_{data_cetak_ulang['pt_1'].replace(' ', '_')}.docx"
            st.download_button(
                label="📥 Unduh Ulang Surat Tugas (.docx)",
                data=buf_r,
                file_name=file_name_r,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                key="btn_dl_reprint"
            )
        except Exception as e_r:
            st.error(f"Gagal menyiapkan dokumen cetak ulang: {e_r}")
