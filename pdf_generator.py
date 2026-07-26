# -*- coding: utf-8 -*-
"""
pdf_generator.py
------------------------------------------------------------------
Modul untuk menghasilkan laporan keuangan resmi dalam format PDF
menggunakan reportlab. Dibuat terpisah dari word_generator.py agar
PDF tidak bergantung pada Microsoft Word / LibreOffice terpasang.

Mencakup: kop surat + logo, judul, tabel transaksi, ringkasan,
grafik, paragraf pembuka/penutup, tanda tangan, watermark,
header/footer + nomor halaman otomatis, dan QR Code verifikasi.
------------------------------------------------------------------
"""

import os
import sys
import tempfile

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
    TableStyle, Image, NextPageTemplate,
)
from reportlab.pdfgen import canvas

from chart_utils import buat_grafik_pemasukan_pengeluaran, buat_qrcode

HIJAU_TUA = colors.HexColor("#1B5E20")
_BASE_DIR_PDF = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))
ABU_MUDA = colors.HexColor("#F2F2F2")
UKURAN_HALAMAN = landscape(A4)  # Landscape (mendatar) — mirip tampilan lembar Excel,
                                 # supaya baris transaksi lebih sering muat 1 baris saja.


class _KanvasBernomor(canvas.Canvas):
    """
    Canvas kustom agar footer bisa menampilkan 'Halaman X dari Y'.
    ReportLab secara default hanya tahu nomor halaman saat ini saat
    dokumen sedang dibangun (belum tahu total halaman) — kelas ini
    menyimpan semua halaman dulu, baru menggambar footer setelah
    total halaman diketahui.
    """

    def __init__(self, *args, **kwargs):
        self._gambar_dasar = kwargs.pop("gambar_dasar")
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._state_halaman = []

    def showPage(self):
        self._state_halaman.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_halaman = len(self._state_halaman)
        for state in self._state_halaman:
            self.__dict__.update(state)
            self._gambar_dasar(self, total_halaman)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)


class PDFGenerator:
    """Bertanggung jawab membangun dokumen PDF laporan keuangan BUMDes."""

    def __init__(self, laporan, logo_path=None):
        self.laporan = laporan
        self.info = laporan.info_bumdes
        self.logo_path = logo_path if logo_path and os.path.exists(logo_path) else \
            os.path.join(_BASE_DIR_PDF, "assets", "logo_default.png")
        self._temp_files = []
        self.styles = getSampleStyleSheet()
        self._buat_style_kustom()

    def _buat_style_kustom(self):
        self.styles.add(ParagraphStyle(
            "Judul", fontSize=17, textColor=HIJAU_TUA, alignment=TA_CENTER,
            fontName="Helvetica-Bold", spaceAfter=10,
        ))
        self.styles.add(ParagraphStyle(
            "NamaBumdes", fontSize=15, textColor=HIJAU_TUA, leading=18,
            fontName="Helvetica-Bold", spaceBefore=2, spaceAfter=5,
        ))
        self.styles.add(ParagraphStyle(
            "LabelBumdes", fontSize=9.5, textColor=HIJAU_TUA, leading=12,
            fontName="Helvetica-Bold", spaceAfter=2,
        ))
        self.styles.add(ParagraphStyle(
            "InfoKecil", fontSize=8.5, leading=12, textColor=colors.HexColor("#333333"),
        ))
        self.styles.add(ParagraphStyle(
            "Isi", fontSize=10, alignment=TA_JUSTIFY, leading=14, spaceAfter=7,
        ))
        self.styles.add(ParagraphStyle(
            "SubJudul", fontSize=13, textColor=HIJAU_TUA,
            fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=8,
        ))
        self.styles.add(ParagraphStyle(
            "TTD", fontSize=10, alignment=TA_CENTER,
        ))
        self.styles.add(ParagraphStyle(
            "TTDNama", fontSize=10, alignment=TA_CENTER,
            fontName="Helvetica-Bold", underlineWidth=1,
        ))
        self.styles.add(ParagraphStyle(
            "TempatTanggal", fontSize=10, alignment=TA_RIGHT, spaceAfter=6,
        ))

    # ------------------------------------------------------------------
    # ENTRY POINT
    # ------------------------------------------------------------------
    def generate(self, output_path):
        # Margin dipersempit karena sudah landscape (lebih lebar secara alami)
        self._margin_kiri = 1.6 * cm
        self._margin_kanan = 1.6 * cm
        self._margin_atas = 1.6 * cm
        self._margin_bawah = 1.6 * cm

        doc = BaseDocTemplate(
            output_path, pagesize=UKURAN_HALAMAN,
            topMargin=self._margin_atas, bottomMargin=self._margin_bawah,
            leftMargin=self._margin_kiri, rightMargin=self._margin_kanan,
        )
        frame = Frame(
            doc.leftMargin, doc.bottomMargin,
            doc.width, doc.height, id="normal",
        )
        template = PageTemplate(id="laporan", frames=frame)
        doc.addPageTemplates([template])

        story = []
        story += self._blok_kop_surat()
        story.append(self._garis_pembatas())
        story.append(Spacer(1, 6))
        story += self._blok_judul_meta()
        story.append(Paragraph(self.laporan.paragraf_pembuka(), self.styles["Isi"]))
        story.append(self._tabel_transaksi())
        story.append(Spacer(1, 10))
        story += self._blok_ringkasan()
        story.append(Spacer(1, 6))
        story.append(self._gambar_grafik())
        story.append(Spacer(1, 6))
        story.append(Paragraph(self.laporan.paragraf_penutup(), self.styles["Isi"]))
        story += self._blok_tanda_tangan()
        story += self._blok_qrcode()

        def pembuat_kanvas(*args, **kwargs):
            return _KanvasBernomor(*args, gambar_dasar=self._gambar_dasar_halaman, **kwargs)

        doc.build(story, canvasmaker=pembuat_kanvas)
        self._bersihkan_file_sementara()
        return output_path

    # ------------------------------------------------------------------
    # HEADER, FOOTER, WATERMARK, AKSEN, NOMOR HALAMAN (digambar tiap halaman)
    # ------------------------------------------------------------------
    def _gambar_dasar_halaman(self, c: canvas.Canvas, total_halaman):
        width, height = UKURAN_HALAMAN
        c.saveState()

        # --- Aksen garis warna di sisi kiri halaman (kesan lebih premium) ---
        c.setFillColor(HIJAU_TUA)
        c.rect(0, 0, 0.3 * cm, height, fill=1, stroke=0)

        # --- Watermark diagonal ---
        c.saveState()
        c.setFont("Helvetica-Bold", 46)
        c.setFillColor(HIJAU_TUA, alpha=0.07)
        c.translate(width / 2, height / 2)
        c.rotate(35)
        c.drawCentredString(0, 0, "DOKUMEN RESMI BUMDES")
        c.restoreState()

        # --- Footer: nama BUMDes kiri, "Halaman X dari Y" kanan ---
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#606060"))
        c.drawString(
            self._margin_kiri, 1.1 * cm,
            f"{self.info.get('nama_bumdes', 'BUMDes')} — Dokumen Resmi",
        )
        c.drawRightString(
            width - self._margin_kanan, 1.1 * cm,
            f"Halaman {c.getPageNumber()} dari {total_halaman}",
        )
        c.setStrokeColor(colors.HexColor("#CCCCCC"))
        c.line(self._margin_kiri, 1.4 * cm, width - self._margin_kanan, 1.4 * cm)

        c.restoreState()

    # ------------------------------------------------------------------
    # KOP SURAT
    # ------------------------------------------------------------------
    def _blok_kop_surat(self):
        elemen = []
        try:
            logo = Image(self.logo_path, width=2.4 * cm, height=2.4 * cm)
        except Exception:
            logo = Paragraph("[LOGO]", self.styles["InfoKecil"])

        info_lines = [
            Paragraph("BADAN USAHA MILIK DESA (BUMDes)", self.styles["LabelBumdes"]),
            Paragraph(self.info.get("nama_bumdes", "-"), self.styles["NamaBumdes"]),
            Paragraph(
                f"Desa {self.info.get('nama_desa', '-')}, "
                f"Kec. {self.info.get('kecamatan', '-')}, "
                f"Kab. {self.info.get('kabupaten', '-')}",
                self.styles["InfoKecil"],
            ),
        ]
        kontak = []
        if self.info.get("telepon"):
            kontak.append(f"Telp: {self.info.get('telepon')}")
        if self.info.get("email"):
            kontak.append(f"Email: {self.info.get('email')}")
        if self.info.get("website"):
            kontak.append(f"Website: {self.info.get('website')}")
        if kontak:
            info_lines.append(Paragraph(" | ".join(kontak), self.styles["InfoKecil"]))

        tabel = Table([[logo, info_lines]], colWidths=[3 * cm, 23.5 * cm])
        tabel.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ]))
        elemen.append(tabel)
        elemen.append(Spacer(1, 6))
        return elemen

    def _garis_pembatas(self):
        t = Table([[""]], colWidths=[26.5 * cm], rowHeights=[0.5])
        t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1.5, HIJAU_TUA)]))
        return t

    # ------------------------------------------------------------------
    # JUDUL + META
    # ------------------------------------------------------------------
    def _blok_judul_meta(self):
        elemen = [Paragraph("LAPORAN KEUANGAN", self.styles["Judul"])]
        meta = [
            ("Nomor Surat", self.info.get("nomor_surat", "-")),
            ("Periode", self.info.get("periode", "-")),
            ("Tanggal Cetak", self.laporan.tanggal_cetak()),
        ]
        data = [[Paragraph(f"<b>{l}</b>", self.styles["InfoKecil"]),
                 Paragraph(f": {v}", self.styles["InfoKecil"])] for l, v in meta]
        tabel = Table(data, colWidths=[3.5 * cm, 10 * cm])
        tabel.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        elemen.append(tabel)
        elemen.append(Spacer(1, 10))
        return elemen

    # ------------------------------------------------------------------
    # TABEL TRANSAKSI
    # ------------------------------------------------------------------
    def _tabel_transaksi(self):
        # PENTING: sel tabel HARUS berupa objek Paragraph, bukan string biasa.
        # ReportLab tidak mem-wrap string biasa di dalam Table — kalau
        # dibiarkan string, teks panjang (mis. kolom "Uraian") akan
        # meluber ke kolom sebelah alih-alih turun ke baris baru.
        gaya_header = ParagraphStyle(
            "HeaderSel", fontSize=8.5, leading=10, fontName="Helvetica-Bold",
            textColor=colors.white, alignment=TA_CENTER,
        )
        gaya_isi_tengah = ParagraphStyle(
            "IsiTengah", fontSize=8.5, leading=10.5, fontName="Helvetica",
            alignment=TA_CENTER,
        )
        gaya_isi_kiri = ParagraphStyle(
            "IsiKiri", fontSize=8.5, leading=10.5, fontName="Helvetica",
            alignment=TA_LEFT,
        )
        gaya_isi_kanan = ParagraphStyle(
            "IsiKanan", fontSize=8.5, leading=10.5, fontName="Helvetica",
            alignment=TA_RIGHT,
        )

        headers = ["No", "Tanggal", "No. Transaksi", "Uraian", "Kategori",
                   "Pemasukan", "Pengeluaran", "Saldo", "Keterangan"]
        data = [[Paragraph(h, gaya_header) for h in headers]]

        # Kolom rata-tengah: No, Tanggal, No. Transaksi, Kategori
        # Kolom rata-kiri: Uraian, Keterangan
        # Kolom rata-kanan: Pemasukan, Pengeluaran, Saldo (nominal uang)
        for idx, t in enumerate(self.laporan.transaksi_list, start=1):
            data.append([
                Paragraph(str(idx), gaya_isi_tengah),
                Paragraph(self.laporan.format_tanggal(t["tanggal"]), gaya_isi_tengah),
                Paragraph(t["nomor_transaksi"], gaya_isi_tengah),
                Paragraph(t["uraian"], gaya_isi_kiri),
                Paragraph(t["kategori"], gaya_isi_tengah),
                Paragraph(self.laporan.format_rupiah(t["pemasukan"]) if t["pemasukan"] else "-", gaya_isi_kanan),
                Paragraph(self.laporan.format_rupiah(t["pengeluaran"]) if t["pengeluaran"] else "-", gaya_isi_kanan),
                Paragraph(self.laporan.format_rupiah(t["saldo"]), gaya_isi_kanan),
                Paragraph(t["keterangan"], gaya_isi_kiri),
            ])

        # Total 26.0 cm — memanfaatkan lebar penuh kertas landscape (26.5cm
        # usable), supaya baris transaksi mayoritas muat dalam SATU baris
        # saja (mirip tampilan spreadsheet Excel), bukan wrap ke banyak baris.
        col_widths = [1.0, 2.1, 2.3, 7.0, 2.1, 2.6, 2.6, 2.6, 3.2]
        col_widths = [c * cm for c in col_widths]

        tabel = Table(data, colWidths=col_widths, repeatRows=1)
        estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), HIJAU_TUA),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]
        for i in range(1, len(data)):
            if i % 2 == 0:
                estilo.append(("BACKGROUND", (0, i), (-1, i), ABU_MUDA))
        tabel.setStyle(TableStyle(estilo))
        return tabel

    # ------------------------------------------------------------------
    # RINGKASAN
    # ------------------------------------------------------------------
    def _blok_ringkasan(self):
        elemen = [Paragraph("Ringkasan Keuangan", self.styles["SubJudul"])]
        r = self.laporan.get_ringkasan_format()

        gaya_label_kartu = ParagraphStyle(
            "LabelKartu", fontSize=9, leading=11, fontName="Helvetica-Bold",
            textColor=colors.white, alignment=TA_CENTER,
        )
        gaya_nilai_kartu = ParagraphStyle(
            "NilaiKartu", fontSize=15, leading=18, fontName="Helvetica-Bold",
            textColor=colors.white, alignment=TA_CENTER, spaceBefore=2,
        )

        def buat_kartu(label, nilai, warna_latar):
            isi = [[Paragraph(label.upper(), gaya_label_kartu)],
                   [Paragraph(nilai, gaya_nilai_kartu)]]
            t = Table(isi, colWidths=[8.6 * cm])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), warna_latar),
                ("TOPPADDING", (0, 0), (-1, 0), 9),
                ("BOTTOMPADDING", (0, -1), (-1, -1), 11),
                ("TOPPADDING", (0, -1), (-1, -1), 2),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]))
            return t

        baris_kartu = Table(
            [[
                buat_kartu("Total Pemasukan", r["total_pemasukan"], HIJAU_TUA),
                buat_kartu("Total Pengeluaran", r["total_pengeluaran"], colors.HexColor("#B71C1C")),
                buat_kartu("Saldo Akhir", r["saldo_akhir"], colors.HexColor("#37474F")),
            ]],
            colWidths=[8.83 * cm, 8.83 * cm, 8.83 * cm],
        )
        baris_kartu.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("RIGHTPADDING", (0, 0), (0, 0), 4),
            ("LEFTPADDING", (1, 0), (1, 0), 4),
            ("RIGHTPADDING", (1, 0), (1, 0), 4),
            ("LEFTPADDING", (2, 0), (2, 0), 4),
            ("RIGHTPADDING", (2, 0), (2, 0), 0),
        ]))
        # Beri jarak antar kartu dengan padding kiri/kanan tipis per kolom
        elemen.append(baris_kartu)
        elemen.append(Spacer(1, 10))

        gaya_label_detail = ParagraphStyle(
            "LabelDetail", fontSize=9.5, fontName="Helvetica-Bold", leading=12,
        )
        gaya_nilai_detail = ParagraphStyle(
            "NilaiDetail", fontSize=9.5, leading=12,
        )
        data = [
            [Paragraph("Jumlah Transaksi", gaya_label_detail), Paragraph(r["jumlah_transaksi"], gaya_nilai_detail)],
            [Paragraph("Rata-rata Pemasukan", gaya_label_detail), Paragraph(r["rata_pemasukan"], gaya_nilai_detail)],
            [Paragraph("Rata-rata Pengeluaran", gaya_label_detail), Paragraph(r["rata_pengeluaran"], gaya_nilai_detail)],
            [Paragraph("Transaksi Terbesar", gaya_label_detail), Paragraph(r["transaksi_terbesar"], gaya_nilai_detail)],
            [Paragraph("Transaksi Terkecil", gaya_label_detail), Paragraph(r["transaksi_terkecil"], gaya_nilai_detail)],
        ]
        tabel = Table(data, colWidths=[6 * cm, 20.5 * cm])
        tabel.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E8F5E9")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ]))
        elemen.append(tabel)
        return elemen

    # ------------------------------------------------------------------
    # GRAFIK
    # ------------------------------------------------------------------
    def _gambar_grafik(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.close()
        self._temp_files.append(tmp.name)
        buat_grafik_pemasukan_pengeluaran(
            self.laporan.ringkasan["total_pemasukan"],
            self.laporan.ringkasan["total_pengeluaran"],
            tmp.name,
        )
        img = Image(tmp.name, width=13.5 * cm, height=7.5 * cm)
        img.hAlign = "CENTER"
        return img

    # ------------------------------------------------------------------
    # TANDA TANGAN
    # ------------------------------------------------------------------
    def _blok_tanda_tangan(self):
        elemen = [Spacer(1, 6)]
        elemen.append(Paragraph(
            f"{self.info.get('nama_desa', '-')}, {self.laporan.tanggal_cetak()}",
            self.styles["TempatTanggal"],
        ))

        gaya_nama_ttd = ParagraphStyle(
            "NamaTTD", fontSize=10, alignment=TA_CENTER, fontName="Helvetica-Bold",
        )

        def blok_kolom_ttd(jabatan, nama):
            garis = Table([[""]], colWidths=[4.5 * cm], rowHeights=[0.3])
            garis.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), 0.8, colors.black)]))
            garis.hAlign = "CENTER"
            return [
                Paragraph(jabatan, self.styles["TTD"]),
                Spacer(1, 45),
                garis,
                Spacer(1, 3),
                Paragraph(nama, gaya_nama_ttd),
            ]

        data = [[
            blok_kolom_ttd("Direktur BUMDes", self.info.get("nama_direktur", "-")),
            blok_kolom_ttd("Bendahara", self.info.get("nama_bendahara", "-")),
        ]]
        tabel = Table(data, colWidths=[13 * cm, 13 * cm])
        tabel.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        elemen.append(tabel)
        return elemen

    # ------------------------------------------------------------------
    # QR CODE
    # ------------------------------------------------------------------
    def _blok_qrcode(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.close()
        self._temp_files.append(tmp.name)

        data_verifikasi = (
            f"BUMDes:{self.info.get('nama_bumdes', '-')}|"
            f"Nomor:{self.info.get('nomor_surat', '-')}|"
            f"Periode:{self.info.get('periode', '-')}|"
            f"Cetak:{self.laporan.tanggal_cetak()}"
        )
        buat_qrcode(data_verifikasi, tmp.name)

        img = Image(tmp.name, width=2.2 * cm, height=2.2 * cm)
        img.hAlign = "RIGHT"
        keterangan = Paragraph("Kode verifikasi dokumen", self.styles["InfoKecil"])
        keterangan.alignment = TA_RIGHT
        return [Spacer(1, 10), keterangan, img]

    def _bersihkan_file_sementara(self):
        for f in self._temp_files:
            try:
                os.remove(f)
            except OSError:
                pass
