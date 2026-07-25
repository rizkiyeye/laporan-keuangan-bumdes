# -*- coding: utf-8 -*-
"""
streamlit_app.py
------------------------------------------------------------------
Versi WEB dari Sistem Cetak Laporan Keuangan BUMDes, dibangun dengan
Streamlit. Bisa dibuka dari browser HP (Android/iPhone) maupun laptop
manapun lewat 1 alamat/link, tanpa perlu install aplikasi apapun.

Menggunakan ulang seluruh logika bisnis yang sama dengan versi desktop
(excel_reader.py, laporan.py, word_generator.py, pdf_generator.py) —
hanya tampilannya saja yang diganti dari Tkinter ke web.

Cara menjalankan LOKAL (di laptop, untuk uji coba):
    pip install -r requirements.txt
    streamlit run streamlit_app.py

Cara PUBLISH GRATIS supaya bisa diakses dari HP di mana saja:
    Lihat panduan lengkap di PANDUAN_PUBLIKASI.md bagian "Opsi 3 - Versi Web"
------------------------------------------------------------------
"""

import os
import tempfile
from datetime import datetime

import streamlit as st

from excel_reader import ExcelReader, ExcelReadError
from laporan import LaporanKeuangan
from word_generator import WordGenerator
from pdf_generator import PDFGenerator

# ------------------------------------------------------------------
# KONFIGURASI HALAMAN
# ------------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Cetak Laporan Keuangan BUMDes",
    page_icon="📊",
    layout="centered",
)

HIJAU_TUA = "#1B5E20"

# Styling minimal & aman — hanya untuk tombol (warna hijau tema BUMDes).
# Warna latar & warna teks utama diatur lewat .streamlit/config.toml
# (cara resmi Streamlit), BUKAN lewat override CSS mentah, supaya tidak
# bentrok dengan warna teks bawaan tema dan menyebabkan tulisan pudar/
# sulit dibaca.
st.markdown(f"""
<style>
    div.stButton > button {{
        background-color: {HIJAU_TUA}; color: white !important; font-weight: bold;
        border-radius: 6px; border: none; padding: 0.5em 1em;
    }}
    div.stButton > button:hover {{ background-color: #2E7D32; color: white !important; }}
    div.stButton > button p {{ color: white !important; }}
</style>
""", unsafe_allow_html=True)

st.markdown(
    f"""
    <div style="background-color:{HIJAU_TUA};padding:20px;border-radius:8px;margin-bottom:20px;">
        <h2 style="color:white;margin:0;">📊 Sistem Cetak Laporan Keuangan</h2>
        <p style="color:#C8E6C9;margin:0;">Badan Usaha Milik Desa (BUMDes)</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------
# STATE AWAL
# ------------------------------------------------------------------
if "laporan_siap" not in st.session_state:
    st.session_state.laporan_siap = False


# ------------------------------------------------------------------
# 1. SUMBER DATA
# ------------------------------------------------------------------
st.subheader("1. Sumber Data")
file_excel = st.file_uploader("File Excel Data Transaksi *", type=["xlsx", "xls"])
file_logo = st.file_uploader("Logo BUMDes (opsional)", type=["png", "jpg", "jpeg"])

# ------------------------------------------------------------------
# 2. IDENTITAS BUMDES
# ------------------------------------------------------------------
st.subheader("2. Identitas BUMDes")
col1, col2 = st.columns(2)
with col1:
    nama_bumdes = st.text_input("Nama BUMDes *")
    nama_desa = st.text_input("Nama Desa *")
    kecamatan = st.text_input("Kecamatan *")
    kabupaten = st.text_input("Kabupaten *")
    telepon = st.text_input("Nomor Telepon")
with col2:
    email = st.text_input("Email")
    website = st.text_input("Website")
    nama_direktur = st.text_input("Nama Direktur *")
    nama_bendahara = st.text_input("Nama Bendahara *")

# ------------------------------------------------------------------
# 3. DATA DOKUMEN
# ------------------------------------------------------------------
st.subheader("3. Data Dokumen")
col3, col4 = st.columns(2)
bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
              "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
with col3:
    bulan_terpilih = st.selectbox("Bulan Periode *", bulan_list, index=datetime.now().month - 1)
    tahun_terpilih = st.number_input(
        "Tahun Periode *", min_value=2020, max_value=2035, value=datetime.now().year, step=1
    )
with col4:
    nomor_surat = st.text_input("Nomor Surat *", placeholder="045/BUMDES-MS/VI/2026")

periode = f"{bulan_terpilih} {tahun_terpilih}"
st.caption(f"Periode laporan: **{periode}**")


# ------------------------------------------------------------------
# HELPER
# ------------------------------------------------------------------
def validasi_form():
    wajib = {
        "File Excel Data Transaksi": file_excel,
        "Nama BUMDes": nama_bumdes,
        "Nama Desa": nama_desa,
        "Kecamatan": kecamatan,
        "Kabupaten": kabupaten,
        "Nomor Surat": nomor_surat,
        "Nama Direktur": nama_direktur,
        "Nama Bendahara": nama_bendahara,
    }
    kosong = [label for label, nilai in wajib.items() if not nilai]
    return kosong


def simpan_upload_ke_temp(uploaded_file, suffix):
    """Menyimpan file yang diupload pengguna ke file sementara di server, mengembalikan path-nya."""
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


def kumpulkan_info_bumdes():
    return {
        "nama_bumdes": nama_bumdes.strip(),
        "nama_desa": nama_desa.strip(),
        "kecamatan": kecamatan.strip(),
        "kabupaten": kabupaten.strip(),
        "telepon": telepon.strip(),
        "email": email.strip(),
        "website": website.strip(),
        "periode": periode,
        "nomor_surat": nomor_surat.strip(),
        "nama_direktur": nama_direktur.strip(),
        "nama_bendahara": nama_bendahara.strip(),
    }


def baca_dan_bangun_laporan():
    path_excel = simpan_upload_ke_temp(file_excel, ".xlsx")
    reader = ExcelReader(path_excel)
    transaksi = reader.baca()
    os.remove(path_excel)
    info = kumpulkan_info_bumdes()
    return LaporanKeuangan(transaksi, info)


# ------------------------------------------------------------------
# 4. AKSI
# ------------------------------------------------------------------
st.subheader("4. Aksi")
kosong = validasi_form()

col_a, col_b, col_c = st.columns(3)

with col_a:
    if st.button("👁️ Preview Laporan", use_container_width=True):
        if kosong:
            st.warning("Lengkapi dulu: " + ", ".join(kosong))
        else:
            try:
                with st.spinner("Membaca data Excel..."):
                    laporan = baca_dan_bangun_laporan()
                r = laporan.get_ringkasan_format()
                st.success("Preview siap!")
                st.markdown(f"""
                **BUMDes:** {nama_bumdes}
                **Periode:** {periode}
                **Nomor Surat:** {nomor_surat}
                **Tanggal Cetak:** {laporan.tanggal_cetak()}

                | Ringkasan | Nilai |
                |---|---|
                | Jumlah Transaksi | {r['jumlah_transaksi']} |
                | Total Pemasukan | {r['total_pemasukan']} |
                | Total Pengeluaran | {r['total_pengeluaran']} |
                | Saldo Akhir | {r['saldo_akhir']} |
                | Transaksi Terbesar | {r['transaksi_terbesar']} |
                | Transaksi Terkecil | {r['transaksi_terkecil']} |
                """)
            except ExcelReadError as e:
                st.error(f"Kesalahan Data Excel:\n\n{e}")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

with col_b:
    if st.button("📄 Buat Word (.docx)", use_container_width=True):
        if kosong:
            st.warning("Lengkapi dulu: " + ", ".join(kosong))
        else:
            try:
                with st.spinner("Membuat dokumen Word..."):
                    laporan = baca_dan_bangun_laporan()
                    logo_path = None
                    if file_logo is not None:
                        logo_path = simpan_upload_ke_temp(file_logo, ".png")
                    tmp_out = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
                    tmp_out.close()
                    WordGenerator(laporan, logo_path=logo_path).generate(tmp_out.name)
                    with open(tmp_out.name, "rb") as f:
                        data_docx = f.read()
                    os.remove(tmp_out.name)
                    if logo_path:
                        os.remove(logo_path)
                st.success("Dokumen Word berhasil dibuat!")
                st.download_button(
                    "⬇️ Download Laporan.docx", data=data_docx,
                    file_name=f"Laporan_Keuangan_BUMDes_{periode.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except ExcelReadError as e:
                st.error(f"Kesalahan Data Excel:\n\n{e}")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

with col_c:
    if st.button("🖨️ Buat PDF", use_container_width=True):
        if kosong:
            st.warning("Lengkapi dulu: " + ", ".join(kosong))
        else:
            try:
                with st.spinner("Membuat dokumen PDF..."):
                    laporan = baca_dan_bangun_laporan()
                    logo_path = None
                    if file_logo is not None:
                        logo_path = simpan_upload_ke_temp(file_logo, ".png")
                    tmp_out = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                    tmp_out.close()
                    PDFGenerator(laporan, logo_path=logo_path).generate(tmp_out.name)
                    with open(tmp_out.name, "rb") as f:
                        data_pdf = f.read()
                    os.remove(tmp_out.name)
                    if logo_path:
                        os.remove(logo_path)
                st.success("Dokumen PDF berhasil dibuat!")
                st.download_button(
                    "⬇️ Download Laporan.pdf", data=data_pdf,
                    file_name=f"Laporan_Keuangan_BUMDes_{periode.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except ExcelReadError as e:
                st.error(f"Kesalahan Data Excel:\n\n{e}")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

st.divider()
st.caption(
    "Sistem Cetak Laporan Keuangan BUMDes — versi web. "
    "Kolom Excel wajib: Tanggal, Nomor Transaksi, Uraian, Kategori, "
    "Pemasukan, Pengeluaran, Saldo, Keterangan."
)
