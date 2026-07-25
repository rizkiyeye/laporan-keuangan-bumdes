# Sistem Cetak Laporan Keuangan BUMDes

Aplikasi desktop Python untuk menghasilkan laporan keuangan resmi BUMDes
secara otomatis dari data Excel, lengkap dengan kop surat, logo, tabel
transaksi, ringkasan keuangan, grafik, watermark, QR code verifikasi,
dan tanda tangan — bisa disimpan sebagai Word (.docx) maupun PDF.

## 1. Instalasi

Pastikan Python 3.9+ sudah terpasang, lalu jalankan:

```bash
pip install -r requirements.txt
```

> Catatan: `tkinter` biasanya sudah bawaan Python di Windows/macOS.
> Di Linux (Ubuntu/Debian) jika belum ada, install dengan:
> `sudo apt install python3-tk`

## 2. (Opsional) Buat Data Contoh

Untuk langsung mencoba aplikasi tanpa menyiapkan data sendiri:

```bash
python buat_data_contoh.py
```

Perintah ini akan membuat `data/laporan.xlsx` berisi 15 baris contoh transaksi.

## 3. Menjalankan Aplikasi

```bash
python main.py
```

## 4. Cara Pakai

1. **Pilih File Excel** — pilih file data transaksi (.xlsx). Kolom wajib:
   `Tanggal, Nomor Transaksi, Uraian, Kategori, Pemasukan, Pengeluaran, Saldo, Keterangan`.
2. **Pilih Logo BUMDes** (opsional) — jika tidak dipilih, logo default dipakai.
3. Isi **Identitas BUMDes** (nama, desa, kecamatan, kabupaten, kontak) dan
   **Data Dokumen** (periode, nomor surat, nama direktur & bendahara).
4. Klik **Preview Laporan** untuk melihat ringkasan sebelum dicetak.
5. Klik **Cetak ke Word** atau **Cetak ke PDF** untuk menghasilkan dokumen —
   hasil disimpan otomatis ke folder `output/`.
6. Klik **Print Langsung** untuk mengirim laporan langsung ke printer default.

## 5. Struktur Proyek

```
project/
├── main.py              # Tampilan GUI (Tkinter) — entry point aplikasi
├── laporan.py            # Kelas LaporanKeuangan — perhitungan & ringkasan
├── excel_reader.py       # Baca & validasi file Excel
├── word_generator.py     # Generator dokumen Word (.docx)
├── pdf_generator.py      # Generator dokumen PDF
├── chart_utils.py        # Bantuan pembuatan grafik & QR code
├── buat_data_contoh.py   # Script pembuat data Excel contoh
├── assets/
│   ├── logo_default.png
│   └── icon.ico
├── data/
│   └── laporan.xlsx
├── output/
│   ├── laporan.docx
│   └── laporan.pdf
└── requirements.txt
```

## 6. Perhitungan Otomatis

Aplikasi menghitung otomatis dari data Excel:
Jumlah Transaksi, Total Pemasukan, Total Pengeluaran, Saldo Akhir,
Rata-rata Pemasukan, Rata-rata Pengeluaran, Transaksi Terbesar, dan
Transaksi Terkecil — semua ditampilkan dalam format Rupiah (Rp).

## 7. Validasi Bawaan

- Menolak file Excel kosong.
- Menolak file dengan kolom yang tidak sesuai (menampilkan kolom yang hilang).
- Semua field identitas wajib diisi sebelum laporan bisa dicetak.

## 8. Troubleshooting

- **"No module named tkinter"** → install `python3-tk` (Linux) atau gunakan
  installer Python resmi dari python.org (Windows/macOS sudah termasuk tkinter).
- **Print Langsung tidak berfungsi** → pastikan printer default sudah
  terpasang di sistem operasi Anda (Windows/macOS/Linux dengan CUPS).
- **Logo tidak tampil** → pastikan format file logo adalah .png/.jpg.
