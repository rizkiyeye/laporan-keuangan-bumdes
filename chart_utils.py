# -*- coding: utf-8 -*-
"""
chart_utils.py
------------------------------------------------------------------
Modul bantuan (utility) untuk menghasilkan aset visual laporan:
1. Grafik batang perbandingan pemasukan vs pengeluaran.
2. QR Code verifikasi dokumen.

Dipakai bersama oleh word_generator.py maupun pdf_generator.py agar
tidak ada duplikasi kode.
------------------------------------------------------------------
"""

import matplotlib
matplotlib.use("Agg")  # Mode non-GUI, aman dipakai di background thread
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import qrcode


def buat_grafik_pemasukan_pengeluaran(total_pemasukan, total_pengeluaran, output_path):
    """
    Membuat grafik batang perbandingan total pemasukan dan pengeluaran,
    lalu menyimpannya sebagai file gambar (PNG) di output_path.
    """
    labels = ["Pemasukan", "Pengeluaran"]
    nilai = [total_pemasukan, total_pengeluaran]
    warna = ["#1B5E20", "#B71C1C"]  # hijau tua & merah tua

    fig, ax = plt.subplots(figsize=(5.2, 3))
    bars = ax.bar(labels, nilai, color=warna, width=0.45)
    ax.set_title("Perbandingan Pemasukan dan Pengeluaran", fontsize=11, fontweight="bold")
    ax.set_ylabel("Jumlah (Rp)")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # Nonaktifkan notasi ilmiah (mis. "1e7") pada sumbu Y, ganti dengan
    # format angka biasa bertitik ribuan (format Rupiah Indonesia).
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: "{:,.0f}".format(x).replace(",", "."))
    )

    # Beri ruang ekstra di atas batang tertinggi supaya label angka tidak kepotong.
    batas_atas = max(nilai) * 1.18 if max(nilai) > 0 else 1
    ax.set_ylim(0, batas_atas)

    for bar in bars:
        tinggi = bar.get_height()
        label = "Rp {:,.0f}".format(tinggi).replace(",", ".")
        ax.annotate(
            label,
            xy=(bar.get_x() + bar.get_width() / 2, tinggi),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center",
            fontsize=8,
            fontweight="bold",
        )

    plt.tight_layout()
    fig.savefig(output_path, dpi=160)
    plt.close(fig)
    return output_path


def buat_qrcode(data_text, output_path):
    """Membuat QR Code verifikasi dokumen dan menyimpannya sebagai PNG."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=6,
        border=2,
    )
    qr.add_data(data_text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1B5E20", back_color="white")
    img.save(output_path)
    return output_path
