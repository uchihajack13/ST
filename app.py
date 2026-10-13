import io
import datetime
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docxtpl import DocxTemplate

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Generator Surat Tugas KPPBC TMC Malang",
    page_icon="📋",
    layout="wide"
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

# Inisialisasi Data
try:
    pegawai_data, perusahaan_data = load_master_data()
except Exception as e:
    st.error(f"Gagal memuat master data dari Google Sheets: {e}")
    st.info("Pastikan credentials.json atau secrets.toml sudah terkonfigurasi dan Service Account sudah di-share ke spreadsheet.")
    pegawai_data, perusahaan_data = {}, {}

# ==========================================
# INTERFACE UTAMA
# ==========================================
st.title("📋 Sistem Otomasi Surat Tugas KPPBC TMC Malang")
st.caption("Input Real-Time Database Spreadsheet & Generator Dokumen Word (.docx) Resmi")

with st.sidebar:
    st.header("⚙️ Pengaturan & Kontrol")
    if st.button("🔄 Segarkan Master Data (Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.markdown("**Petunjuk Cepat:**")
    st.caption("1. Isi form di samping kanan secara lengkap.")
    st.caption("2. Klik 'Simpan ke Database' untuk append ke Sheet1.")
    st.caption("3. Klik 'Unduh Surat Tugas (.docx)' untuk download draf Word.")

list_perusahaan = sorted(list(perusahaan_data.keys())) if perusahaan_data else []
list_pegawai = sorted(list(pegawai_data.keys())) if pegawai_data else []

with st.form("form_surat_tugas", clear_on_submit=False):
    # -------------------------------------------------------------
    # 1. DATA SURAT PEMOHON
    # -------------------------------------------------------------
    st.subheader("1. Data Surat Pemohon")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Pabrik 1 (Pemohon Utama)**")
        pt_1 = st.selectbox("Nama Perusahaan / Pabrik 1*", options=[""] + list_perusahaan, index=0)
        no_surat_1 = st.text_input("Nomor Surat Pemohon 1*", placeholder="Contoh: 078/BIT-IX/2026")
        tgl_surat_1 = st.date_input("Tanggal Surat Pemohon 1*", value=datetime.date.today())
        permohonan_1 = st.text_area("Perihal / Permohonan 1*", placeholder="Permohonan Pemeriksaan TIS Retur")
        kegiatan_1 = st.text_input("Kegiatan Baku 1 (Pelaksanaan Tugas)*", placeholder="Pemeriksaan TIS Retur")
        
        pilihan_lokasi_1 = st.selectbox("Pilihan Lokasi Pabrik 1 (1 - 12)", options=list(range(1, 13)), index=0)
        alamat_pabrik_1 = ""
        if pt_1 and pt_1 in perusahaan_data:
            alamat_pabrik_1 = perusahaan_data[pt_1].get(pilihan_lokasi_1, "")
        st.text_area("Alamat Terpilih Pabrik 1", value=alamat_pabrik_1, disabled=True, height=70)

    with col2:
        st.markdown("**Pabrik 2 (Opsional / Jika Ada)**")
        pt_2 = st.selectbox("Nama Perusahaan / Pabrik 2", options=[""] + list_perusahaan, index=0)
        no_surat_2 = st.text_input("Nomor Surat Pemohon 2", placeholder="Kosongkan jika tidak ada")
        tgl_surat_2 = st.date_input("Tanggal Surat Pemohon 2", value=None)
        permohonan_2 = st.text_area("Perihal / Permohonan 2", placeholder="Kosongkan jika tidak ada")
        kegiatan_2 = st.text_input("Kegiatan Baku 2", placeholder="Kosongkan jika tidak ada")
        
        pilihan_lokasi_2 = st.selectbox("Pilihan Lokasi Pabrik 2 (1 - 12)", options=list(range(1, 13)), index=0)
        alamat_pabrik_2 = ""
        if pt_2 and pt_2 in perusahaan_data:
            alamat_pabrik_2 = perusahaan_data[pt_2].get(pilihan_lokasi_2, "")
        st.text_area("Alamat Terpilih Pabrik 2", value=alamat_pabrik_2, disabled=True, height=70)

    st.divider()

    # -------------------------------------------------------------
    # 2. WAKTU PELAKSANAAN TUGAS
    # -------------------------------------------------------------
    st.subheader("2. Waktu Pelaksanaan Tugas")
    col_tgl1, col_tgl2 = st.columns(2)
    with col_tgl1:
        tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan Tugas*", value=datetime.date.today())
    with col_tgl2:
        nama_hari_en = tgl_pelaksanaan.strftime("%A")
        nama_hari_id = HARI_INDONESIA.get(nama_hari_en, nama_hari_en)
        st.text_input("Hari Pelaksanaan (Otomatis)", value=nama_hari_id, disabled=True)

    st.divider()

    # -------------------------------------------------------------
    # 3. PERSONEL TIM PETUGAS (1 s.d. 6)
    # -------------------------------------------------------------
    st.subheader("3. Tim Petugas Pemeriksa")
    st.caption("Pilih minimal 1 orang Petugas. Petugas berikutnya bersifat opsional.")

    data_petugas_terpilih = []
    for i in range(1, 7):
        col_p1, col_p2 = st.columns([1.5, 2])
        with col_p1:
            lbl = f"Nama Petugas {i}*" if i == 1 else f"Nama Petugas {i}"
            sel_nama = st.selectbox(lbl, options=["-- Tidak Ditugaskan --"] + list_pegawai, index=0, key=f"petugas_{i}")
        
        p_nip, p_pangkat, p_jabatan = "", "", ""
        if sel_nama and sel_nama != "-- Tidak Ditugaskan --":
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
            
        with col_p2:
            cp1, cp2 = st.columns(2)
            cp1.text_input(f"Pangkat {i}", value=p_pangkat, disabled=True, key=f"pkt_{i}")
            cp2.text_input(f"Jabatan {i}", value=p_jabatan, disabled=True, key=f"jbt_{i}")

    st.divider()

    # -------------------------------------------------------------
    # 4. KONFIGURASI PENANDATANGAN SURAT TUGAS
    # -------------------------------------------------------------
    st.subheader("4. Konfigurasi Penandatangan Surat Tugas")
    col_ttd1, col_ttd2 = st.columns(2)
    
    with col_ttd1:
        tipe_ttd = st.radio(
            "Format Penandatangan:",
            options=["a.n. Kepala Kantor (Kepala Subbagian Umum)", "Kepala Kantor Langsung"],
            index=0
        )
        is_plh = st.checkbox("Bertindak sebagai Pelaksana Harian (Plh.)", value=False)

    with col_ttd2:
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
    submit_simpan = st.form_submit_button("💾 Simpan ke Database Spreadsheet", use_container_width=True, type="primary")

# ==========================================
# EKSEKUSI 1: APPEND KE GOOGLE SHEETS
# ==========================================
if submit_simpan:
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

# ==========================================
# EKSEKUSI 2: GENERATE & DOWNLOAD WORD
# ==========================================
st.divider()
st.subheader("📄 Generate Draf Dokumen Word (.docx)")

if "last_submitted_data" in st.session_state:
    data_gen = st.session_state["last_submitted_data"]
    st.write(f"Data Siap Cetak: **{data_gen['pt_1']}** (Pelaksanaan: {data_gen['hari']}, {data_gen['tgl_pelaksanaan']})")
    
    try:
        # Pilih template file sesuai tipe penandatangan
        if data_gen["tipe_ttd"] == "Kepala Kantor Langsung":
            template_path = "template_st_kk.docx"
            prefix_ttd = "Plh." if data_gen["is_plh"] else ""
            jabatan_struktural = ""
        else:
            template_path = "template_st_an.docx"
            prefix_ttd = "a.n."
            jabatan_struktural = "Plh. Kepala Subbagian Umum" if data_gen["is_plh"] else "Kepala Subbagian Umum"

        doc = DocxTemplate(template_path)
        
        context = {
            "NOMOR_ND": "[@NomorND]",  # Siap dinomori otomatis di Nadine
            "TANGGAL_ND": "[@TanggalND]",
            "NAMA_PERUSAHAAN": data_gen["pt_1"],
            "NOMOR_SURAT": data_gen["no_surat_1"],
            "TANGGAL_SURAT": data_gen["tgl_surat_1"],
            "PERMOHONAN": data_gen["permohonan_1"],
            "KEGIATAN": data_gen["kegiatan_1"],
            "HARI": data_gen["hari"],
            "TANGGAL_PELAKSANAAN": data_gen["tgl_pelaksanaan"],
            "ALAMAT_PABRIK": data_gen["alamat_pabrik_1"],
            "ADA_PABRIK_2": bool(data_gen["pt_2"]),
            "NAMA_PERUSAHAAN_2": data_gen["pt_2"],
            "NOMOR_SURAT_2": data_gen["no_surat_2"],
            "TANGGAL_SURAT_2": data_gen["tgl_surat_2"],
            "PERMOHONAN_2": data_gen["permohonan_2"],
            "KEGIATAN_2": data_gen["kegiatan_2"],
            "ALAMAT_PABRIK_2": data_gen["alamat_pabrik_2"],
            "DAFTAR_PETUGAS": data_gen["daftar_petugas"],
            "PREFIX_TTD": prefix_ttd,
            "JABATAN_STRUKTURAL": jabatan_struktural,
            "NAMA_PENANDATANGAN": data_gen["nama_pejabat_ttd"]
        }
        
        doc.render(context)
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        nama_file_unduh = f"ST_{data_gen['pt_1'].replace(' ', '_')}_{data_gen['tgl_pelaksanaan'].replace(' ', '_')}.docx"
        
        st.download_button(
            label="📥 Unduh Surat Tugas (.docx)",
            data=docx_buffer,
            file_name=nama_file_unduh,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
    except FileNotFoundError:
        st.warning(f"⚠️ File `{template_path}` tidak ditemukan. Pastikan file template berada di satu folder dengan `app.py`.")
    except Exception as e_docx:
        st.error(f"Gagal membuat dokumen Word: {e_docx}")
else:
    st.info("Kirim/simpan form di atas terlebih dahulu untuk mengaktifkan tombol unduh dokumen Word.")
