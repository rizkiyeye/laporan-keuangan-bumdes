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
    layout="wide",
)

HIJAU_TUA = "#1B5E20"
HIJAU_MUDA = "#2E7D32"
MERAH = "#B71C1C"
ABU_GELAP = "#37474F"

# ------------------------------------------------------------------
# STYLING — hanya elemen yang aman di-custom (tombol, kartu statistik).
# Warna latar & warna teks utama tetap diatur lewat .streamlit/config.toml
# supaya tidak bentrok dengan tema bawaan (mencegah teks pudar/sulit dibaca).
# ------------------------------------------------------------------
st.markdown(f"""
<style>
    div.stButton > button {{
        background-color: {HIJAU_TUA}; color: white !important; font-weight: 600;
        border-radius: 8px; border: none; padding: 0.6em 1em;
        transition: background-color 0.15s ease;
    }}
    div.stButton > button:hover {{ background-color: {HIJAU_MUDA}; color: white !important; }}
    div.stButton > button p {{ color: white !important; font-size: 0.95rem; }}

    div[data-testid="stDownloadButton"] > button {{
        background-color: {ABU_GELAP}; color: white !important; font-weight: 600;
        border-radius: 8px; border: none;
    }}
    div[data-testid="stDownloadButton"] > button:hover {{ background-color: #263238; }}
    div[data-testid="stDownloadButton"] > button p {{ color: white !important; }}

    div[data-testid="stFileUploaderDropzone"] {{
        border-radius: 10px;
    }}

    .kartu-statistik {{
        border-radius: 10px; padding: 16px 10px; text-align: center; color: white;
    }}
    .kartu-statistik .label {{
        font-size: 0.72rem; font-weight: 700; letter-spacing: 0.04em;
        text-transform: uppercase; opacity: 0.9; margin-bottom: 6px;
    }}
    .kartu-statistik .nilai {{ font-size: 1.35rem; font-weight: 700; }}

    h3 {{ margin-top: 0 !important; }}
</style>
""", unsafe_allow_html=True)


def kartu_statistik(label, nilai, warna):
    return f"""
    <div class="kartu-statistik" style="background-color:{warna};">
        <div class="label">{label}</div>
        <div class="nilai">{nilai}</div>
    </div>
    """


# ------------------------------------------------------------------
# HEADER
# ------------------------------------------------------------------
st.markdown(
    f"""
    <div style="background-color:{HIJAU_TUA};padding:22px 26px;border-radius:10px;margin-bottom:22px;">
        <h2 style="color:white;margin:0;">📊 Sistem Cetak Laporan Keuangan</h2>
        <p style="color:#C8E6C9;margin:2px 0 0 0;">Badan Usaha Milik Desa (BUMDes) — Versi Web</p>
    </div>
    """,
    unsafe_allow_html=True,
)

kolom_utama, kolom_bantuan = st.columns([2.4, 1], gap="large")

with kolom_utama:
    # ------------------------------------------------------------------
    # 1. SUMBER DATA
    # ------------------------------------------------------------------
    with st.container(border=True):
        st.markdown("### 📁 1. Sumber Data")
        c1, c2 = st.columns(2)
        with c1:
            file_excel = st.file_uploader("File Excel Data Transaksi *", type=["xlsx", "xls"])
        with c2:
            file_logo = st.file_uploader("Logo BUMDes (opsional)", type=["png", "jpg", "jpeg"])
            if file_logo is not None:
                st.image(file_logo, width=70)

    # ------------------------------------------------------------------
    # 2. IDENTITAS BUMDES
    # ------------------------------------------------------------------
    with st.container(border=True):
        st.markdown("### 🏢 2. Identitas BUMDes")
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
    with st.container(border=True):
        st.markdown("### 📄 3. Data Dokumen")
        col3, col4, col5 = st.columns([1.3, 1, 1.7])
        bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        with col3:
            bulan_terpilih = st.selectbox("Bulan Periode *", bulan_list, index=datetime.now().month - 1)
        with col4:
            tahun_terpilih = st.number_input(
                "Tahun Periode *", min_value=2020, max_value=2035, value=datetime.now().year, step=1
            )
        with col5:
            nomor_surat = st.text_input("Nomor Surat *", placeholder="045/BUMDES-MS/VI/2026")

        periode = f"{bulan_terpilih} {tahun_terpilih}"
        st.caption(f"📌 Periode laporan: **{periode}**")

with kolom_bantuan:
    with st.container(border=True):
        st.markdown("### ℹ️ Panduan Singkat")
        st.markdown("""
        1. Upload file Excel data transaksi
        2. Lengkapi identitas BUMDes
        3. Pilih periode & isi nomor surat
        4. Klik **Preview** untuk cek ringkasan
        5. Klik **Buat Word/PDF** untuk download

        **Kolom Excel wajib:**
        `Tanggal, Nomor Transaksi, Uraian, Kategori, Pemasukan, Pengeluaran, Saldo, Keterangan`
        """)


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
    return [label for label, nilai in wajib.items() if not nilai]


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
st.markdown("### ⚡ 4. Aksi")
kosong = validasi_form()

col_a, col_b, col_c = st.columns(3)
tombol_preview = col_a.button("👁️ Preview Laporan", use_container_width=True)
tombol_word = col_b.button("📄 Buat Word (.docx)", use_container_width=True)
tombol_pdf = col_c.button("🖨️ Buat PDF", use_container_width=True)

if (tombol_preview or tombol_word or tombol_pdf) and kosong:
    st.warning("⚠️ Lengkapi dulu: " + ", ".join(kosong))

# --- Preview ---
if tombol_preview and not kosong:
    try:
        with st.spinner("Membaca data Excel..."):
            laporan = baca_dan_bangun_laporan()
        r = laporan.get_ringkasan_format()

        st.success("✅ Preview siap!")
        st.markdown(f"""
        **BUMDes:** {nama_bumdes} &nbsp;|&nbsp;
        **Periode:** {periode} &nbsp;|&nbsp;
        **Nomor Surat:** {nomor_surat} &nbsp;|&nbsp;
        **Tanggal Cetak:** {laporan.tanggal_cetak()}
        """)

        k1, k2, k3 = st.columns(3)
        k1.markdown(kartu_statistik("Total Pemasukan", r["total_pemasukan"], HIJAU_TUA), unsafe_allow_html=True)
        k2.markdown(kartu_statistik("Total Pengeluaran", r["total_pengeluaran"], MERAH), unsafe_allow_html=True)
        k3.markdown(kartu_statistik("Saldo Akhir", r["saldo_akhir"], ABU_GELAP), unsafe_allow_html=True)

        st.write("")
        detail1, detail2 = st.columns(2)
        with detail1:
            st.metric("Jumlah Transaksi", r["jumlah_transaksi"])
            st.metric("Rata-rata Pemasukan", r["rata_pemasukan"])
        with detail2:
            st.metric("Rata-rata Pengeluaran", r["rata_pengeluaran"])
        st.caption(f"📈 Transaksi Terbesar: {r['transaksi_terbesar']}")
        st.caption(f"📉 Transaksi Terkecil: {r['transaksi_terkecil']}")

    except ExcelReadError as e:
        st.error(f"❌ Kesalahan Data Excel:\n\n{e}")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {e}")

# --- Buat Word ---
if tombol_word and not kosong:
    try:
        with st.spinner("Membuat dokumen Word..."):
            laporan = baca_dan_bangun_laporan()
            logo_path = simpan_upload_ke_temp(file_logo, ".png") if file_logo is not None else None
            tmp_out = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            tmp_out.close()
            WordGenerator(laporan, logo_path=logo_path).generate(tmp_out.name)
            with open(tmp_out.name, "rb") as f:
                data_docx = f.read()
            os.remove(tmp_out.name)
            if logo_path:
                os.remove(logo_path)
        st.success("✅ Dokumen Word berhasil dibuat!")
        st.download_button(
            "⬇️ Download Laporan.docx", data=data_docx,
            file_name=f"Laporan_Keuangan_BUMDes_{periode.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    except ExcelReadError as e:
        st.error(f"❌ Kesalahan Data Excel:\n\n{e}")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {e}")

# --- Buat PDF ---
if tombol_pdf and not kosong:
    try:
        with st.spinner("Membuat dokumen PDF..."):
            laporan = baca_dan_bangun_laporan()
            logo_path = simpan_upload_ke_temp(file_logo, ".png") if file_logo is not None else None
            tmp_out = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp_out.close()
            PDFGenerator(laporan, logo_path=logo_path).generate(tmp_out.name)
            with open(tmp_out.name, "rb") as f:
                data_pdf = f.read()
            os.remove(tmp_out.name)
            if logo_path:
                os.remove(logo_path)
        st.success("✅ Dokumen PDF berhasil dibuat!")
        st.download_button(
            "⬇️ Download Laporan.pdf", data=data_pdf,
            file_name=f"Laporan_Keuangan_BUMDes_{periode.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except ExcelReadError as e:
        st.error(f"❌ Kesalahan Data Excel:\n\n{e}")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {e}")

st.divider()
st.caption("Sistem Cetak Laporan Keuangan BUMDes — versi web. Dibuat untuk mendukung transparansi pengelolaan BUMDes.")
