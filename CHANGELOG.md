# ZAIN.NET v6.0 — Perbaikan kompatibilitas DOCX

- Memperbaiki urutan properti OOXML (`w:pPr`, `w:rPr`, `w:tblPr`, `w:tcPr`) untuk file hasil **semua 15 jurnal**. V5 menulis properti tidak sesuai urutan standar.
- Pada contoh RABBANI v5 terdapat 26 pelanggaran urutan (16 paragraph, 9 run, 1 tabel); setelah koreksi menjadi 0.
- Menangani properti bawaan di `word/footer*.xml` dan `word/numbering.xml` pada template hasil konversi .doc lama.
- MIME DOCX diatur benar dan asset JavaScript diberi cache-buster agar GitHub Pages memakai kode baru.
- 30 pengujian pembuatan DOCX (15 template × 2 skenario): berhasil.
- 30 pemeriksaan ZIP CRC, XML, relasi part, dan pembukaan python-docx: berhasil.
- 15 uji konversi hasil ke PDF LibreOffice: berhasil.
- **Belum diuji langsung di Microsoft Word 2010**; hindari klaim kompatibilitas 100% sampai diverifikasi pada PC pengguna.
