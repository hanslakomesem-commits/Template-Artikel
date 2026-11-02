# Audit Kesalahan Word dan Uji v6 — ZAIN.NET Multi Template

## Laporan sumber masalah
Screenshot menunjukkan Word berhenti bekerja ketika membuka hasil RABBANI v5. Ini **tidak membuktikan** penyebab tunggal, karena Word/Office lokal juga bisa mengalami crash.

Namun inspeksi OOXML menemukan **kesalahan nyata dari mesin v5**:
- Template RABBANI asli: 0 properti tidak terurut menurut skema OOXML pada file Word asli.
- Hasil AUTO RABBANI v5: 26 grup properti tidak terurut: `w:pPr` (16), `w:rPr` (9), `w:tblPr` (1).
- Contoh: `<w:jc/><w:spacing/>` (seharusnya `spacing` lebih dahulu), `<w:sz/><w:b/>` (seharusnya `b` sebelum `sz`), `tblLayout` sebelum `tblBorders` (terbalik).

Word 2010 dapat lebih sensitif terhadap struktur OOXML yang salah daripada LibreOffice. V6 memperbaiki **mesin pembuat Word** untuk semua jurnal, bukan hanya RABBANI.

## Cakupan pengujian otomatis v6

| Pengujian | Hasil |
|---|---|
| 15 jurnal × 2 artikel (biasa dan dengan judul panjang, gambar, tabel) | 30/30 DOCX terbentuk |
| ZIP CRC, seluruh XML dapat diparse, relasi internal menuju bagian yang tersedia | 30/30 lolos |
| Dokumen dapat dibuka melalui python-docx | 30/30 lolos |
| Footnote dari artikel sumber (8 buah) | 30/30 tetap terdeteksi |
| Gambar dan tabel sumber pada kasus kompleks | 15/15 dipertahankan secara struktural |
| Urutan properti OOXML `pPr/rPr/tblPr/tcPr` pada semua file uji | 30/30 tanpa pelanggaran yang terdeteksi |
| Seluruh 15 file hasil biasa dikonversi menjadi PDF di LibreOffice | 15/15 berhasil |
| Dibuka langsung dengan Microsoft Word 2010 di PC pengguna | **Belum dapat diuji di lingkungan ini** |

**Penting:** Validitas file secara struktural dan keberhasilan LibreOffice tidak sama dengan jaminan Word 2010 terbebas dari crash. Komponen Word/Office 2010, add-ins, font, mode kompatibilitas, atau instalasi yang bermasalah masih dapat menyebabkan crash.

## Cara menguji di komputer pengguna

1. Coba buka terlebih dahulu **Template Jurnal Rabbani New 2023(1).docx asli**. Jika ikut crash, investigasi Word/Office lokal.
2. Kemudian buka `RABBANI_HASIL_V6_UJI_WORD2010.docx` (contoh file hasil engine yang sudah diperbaiki).
3. Jika masih crash, uji file alternatif `RABBANI_ALTERNATIF_DISIMPAN_ULANG_LO.docx` (disimpan ulang menggunakan LibreOffice, bisa mengubah beberapa pemisahan halaman).
4. Jika Word tetap crash, tutup semua Word. Tekan `Win+R` lalu ketik `winword /safe`, Enter. Coba buka file tadi.
5. Bila Word masih tidak stabil, gunakan **File → Open → panah di samping Open → Open and Repair** pada Word 2010.
6. Periksa apakah Word juga crash saat membuka DOCX baru kosong. Jika iya, masalah mungkin berada di aplikasi Word, bukan hanya file.

Versi v6 adalah source GitHub Pages Only. Tidak perlu backend Python/URL API.
