# -*- coding: utf-8 -*-
"""
excel_reader.py
------------------------------------------------------------------
Modul untuk membaca dan memvalidasi data transaksi keuangan dari
file Excel (.xlsx). Bertugas memastikan struktur kolom sesuai,
membersihkan data, lalu mengembalikan list transaksi siap pakai.
------------------------------------------------------------------
"""

import os
import pandas as pd


class ExcelReadError(Exception):
    """Exception khusus untuk kesalahan pembacaan/validasi file Excel."""
    pass


class ExcelReader:
    """Membaca file Excel transaksi BUMDes dan memvalidasi strukturnya."""

    # Kolom wajib sesuai struktur yang ditentukan.
    KOLOM_WAJIB = [
        "Tanggal", "Nomor Transaksi", "Uraian", "Kategori",
        "Pemasukan", "Pengeluaran", "Saldo", "Keterangan",
    ]

    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None

    def baca(self):
        """
        Membaca file Excel, memvalidasi, membersihkan, lalu mengembalikan
        data dalam bentuk list of dict (satu dict = satu baris transaksi).
        """
        if not self.filepath or not os.path.exists(self.filepath):
            raise ExcelReadError(f"File Excel tidak ditemukan:\n{self.filepath}")

        try:
            df = pd.read_excel(self.filepath, engine="openpyxl")
        except Exception as e:
            raise ExcelReadError(f"Gagal membaca file Excel.\nDetail: {e}")

        # Bersihkan nama kolom dari spasi berlebih
        df.columns = [str(c).strip() for c in df.columns]

        if df.dropna(how="all").empty:
            raise ExcelReadError(
                "File Excel kosong (tidak ada data transaksi). "
                "Pastikan file berisi minimal satu baris data."
            )

        self._validasi_kolom(df)
        df = self._bersihkan_data(df)
        self.df = df
        return self._ke_list_transaksi(df)

    def _validasi_kolom(self, df):
        """Memastikan seluruh kolom wajib ada pada file Excel."""
        kolom_ada = list(df.columns)
        kolom_hilang = [k for k in self.KOLOM_WAJIB if k not in kolom_ada]
        if kolom_hilang:
            raise ExcelReadError(
                "Struktur kolom Excel tidak sesuai.\n\n"
                "Kolom yang hilang: " + ", ".join(kolom_hilang) + "\n\n"
                "Kolom wajib yang harus ada:\n" + ", ".join(self.KOLOM_WAJIB)
            )

    def _bersihkan_data(self, df):
        """Membersihkan tipe data agar konsisten (angka, tanggal, teks)."""
        df = df.dropna(how="all").copy()

        for kolom in ["Pemasukan", "Pengeluaran", "Saldo"]:
            df[kolom] = pd.to_numeric(df[kolom], errors="coerce").fillna(0)

        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")

        for kolom in ["Nomor Transaksi", "Uraian", "Kategori", "Keterangan"]:
            df[kolom] = df[kolom].fillna("-").astype(str)

        # Buang baris yang benar-benar tidak punya uraian/nomor transaksi
        df = df[~((df["Uraian"] == "-") & (df["Nomor Transaksi"] == "-"))]

        return df.reset_index(drop=True)

    def _ke_list_transaksi(self, df):
        """Mengubah DataFrame menjadi list of dict yang siap dipakai modul lain."""
        hasil = []
        for _, row in df.iterrows():
            hasil.append({
                "tanggal": row["Tanggal"],
                "nomor_transaksi": row["Nomor Transaksi"],
                "uraian": row["Uraian"],
                "kategori": row["Kategori"],
                "pemasukan": float(row["Pemasukan"]),
                "pengeluaran": float(row["Pengeluaran"]),
                "saldo": float(row["Saldo"]),
                "keterangan": row["Keterangan"],
            })
        if not hasil:
            raise ExcelReadError("Tidak ada baris transaksi yang valid ditemukan di file Excel.")
        return hasil
