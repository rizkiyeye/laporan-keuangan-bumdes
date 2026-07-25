# -*- coding: utf-8 -*-
"""
laporan.py
------------------------------------------------------------------
Modul ini berisi kelas LaporanKeuangan yang bertugas mengolah data
transaksi (hasil bacaan dari Excel) menjadi ringkasan keuangan yang
siap ditampilkan pada dokumen Word/PDF.

Semua perhitungan otomatis (total pemasukan, total pengeluaran,
saldo akhir, rata-rata, transaksi terbesar/terkecil) dilakukan di
kelas ini agar terpisah dari logika GUI maupun logika pembuatan
dokumen (Word/PDF).
------------------------------------------------------------------
"""

from datetime import datetime


class LaporanKeuangan:
    """
    Merepresentasikan satu laporan keuangan BUMDes berdasarkan
    data transaksi yang sudah dibaca dari file Excel.
    """

    def __init__(self, transaksi_list, info_bumdes=None):
        """
        Parameters
        ----------
        transaksi_list : list[dict]
            Daftar transaksi hasil dari ExcelReader.
        info_bumdes : dict
            Informasi umum BUMDes (nama, alamat, direktur, dll)
            yang diisi lewat form GUI.
        """
        self.transaksi_list = transaksi_list
        self.info_bumdes = info_bumdes or {}
        self.ringkasan = {}
        self._hitung_ringkasan()

    # ------------------------------------------------------------------
    # PERHITUNGAN OTOMATIS
    # ------------------------------------------------------------------
    def _hitung_ringkasan(self):
        """Menghitung seluruh nilai ringkasan keuangan secara otomatis."""
        jumlah_transaksi = len(self.transaksi_list)
        total_pemasukan = sum(t["pemasukan"] for t in self.transaksi_list)
        total_pengeluaran = sum(t["pengeluaran"] for t in self.transaksi_list)

        # Saldo akhir diambil dari kolom saldo transaksi terakhir bila ada,
        # jika tidak tersedia maka dihitung dari selisih pemasukan-pengeluaran.
        if self.transaksi_list:
            saldo_akhir = self.transaksi_list[-1]["saldo"]
        else:
            saldo_akhir = 0

        rata_pemasukan = (total_pemasukan / jumlah_transaksi) if jumlah_transaksi else 0
        rata_pengeluaran = (total_pengeluaran / jumlah_transaksi) if jumlah_transaksi else 0

        transaksi_terbesar = None
        transaksi_terkecil = None
        if self.transaksi_list:
            def nilai_transaksi(t):
                # Nilai transaksi dianggap sebagai nilai absolut terbesar
                # antara pemasukan dan pengeluaran pada baris tersebut.
                return max(t["pemasukan"], t["pengeluaran"])

            transaksi_terbesar = max(self.transaksi_list, key=nilai_transaksi)
            transaksi_terkecil = min(self.transaksi_list, key=nilai_transaksi)

        self.ringkasan = {
            "jumlah_transaksi": jumlah_transaksi,
            "total_pemasukan": total_pemasukan,
            "total_pengeluaran": total_pengeluaran,
            "saldo_akhir": saldo_akhir,
            "rata_pemasukan": rata_pemasukan,
            "rata_pengeluaran": rata_pengeluaran,
            "transaksi_terbesar": transaksi_terbesar,
            "transaksi_terkecil": transaksi_terkecil,
        }

    # ------------------------------------------------------------------
    # UTILITAS FORMAT
    # ------------------------------------------------------------------
    @staticmethod
    def format_rupiah(nilai):
        """
        Mengubah angka menjadi format Rupiah Indonesia, contoh: Rp 1.250.000.
        Memakai non-breaking space (bukan spasi biasa) antara "Rp" dan angka,
        supaya saat teks di-wrap di dalam sel tabel (PDF/Word), "Rp" tidak
        pernah terpisah sendirian dari angkanya di baris berikutnya.
        """
        try:
            nilai = float(nilai)
        except (ValueError, TypeError):
            nilai = 0
        angka = "{:,.0f}".format(nilai).replace(",", ".")
        return "Rp\u00A0" + angka

    @staticmethod
    def format_tanggal(tanggal):
        """Mengubah objek tanggal menjadi format dd-mm-yyyy."""
        try:
            return tanggal.strftime("%d-%m-%Y")
        except Exception:
            return str(tanggal) if tanggal else "-"

    def get_ringkasan_format(self):
        """Mengembalikan ringkasan keuangan dalam format string Rupiah, siap tampil."""
        r = self.ringkasan
        hasil = {
            "jumlah_transaksi": str(r["jumlah_transaksi"]),
            "total_pemasukan": self.format_rupiah(r["total_pemasukan"]),
            "total_pengeluaran": self.format_rupiah(r["total_pengeluaran"]),
            "saldo_akhir": self.format_rupiah(r["saldo_akhir"]),
            "rata_pemasukan": self.format_rupiah(r["rata_pemasukan"]),
            "rata_pengeluaran": self.format_rupiah(r["rata_pengeluaran"]),
        }
        tb = r["transaksi_terbesar"]
        tk = r["transaksi_terkecil"]
        hasil["transaksi_terbesar"] = (
            f"{tb['uraian']} - {self.format_rupiah(max(tb['pemasukan'], tb['pengeluaran']))}"
            if tb else "-"
        )
        hasil["transaksi_terkecil"] = (
            f"{tk['uraian']} - {self.format_rupiah(max(tk['pemasukan'], tk['pengeluaran']))}"
            if tk else "-"
        )
        return hasil

    def paragraf_pembuka(self):
        return (
            "Laporan keuangan ini disusun sebagai bentuk pertanggungjawaban "
            "pengelolaan keuangan Badan Usaha Milik Desa (BUMDes) selama periode "
            "yang telah ditentukan. Seluruh transaksi telah direkap berdasarkan "
            "data yang tersimpan dalam sistem."
        )

    def paragraf_penutup(self):
        return (
            "Demikian laporan keuangan ini dibuat dengan sebenar-benarnya sebagai "
            "bentuk pertanggungjawaban pengelolaan keuangan BUMDes. Semoga laporan "
            "ini dapat menjadi bahan evaluasi dan pengambilan keputusan dalam "
            "pengelolaan usaha desa."
        )

    def tanggal_cetak(self):
        """Tanggal saat laporan dibuat/dicetak (otomatis, hari ini)."""
        bulan_indonesia = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember",
        ]
        now = datetime.now()
        return f"{now.day} {bulan_indonesia[now.month - 1]} {now.year}"
