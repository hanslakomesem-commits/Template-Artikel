# v10.0 — Final DOCX diproses dengan LibreOffice server

- Backend FastAPI + Docker LibreOffice Writer; file DOCX akhir dihasilkan melalui Save As DOCX pada server, bukan diunduh langsung dari browser.
- Semua 15 master template dan aturan lama dipertahankan dari v9.
- RABBANI tetap hanya mengambil abstrak Indonesia (permintaan khusus ZAIN.NET).
- Tidak ada fallback mengunduh DOCX mentah jika konversi server gagal.
- File sementara dihapus setelah request, tidak ada penyimpanan database atau pencatatan isi artikel.
- API memvalidasi DOCX ZIP/XML, memeriksa catatan kaki dan gambar, serta memberi pesan error bila konversi gagal.
- `config.js` untuk memasukkan URL server HTTPS setelah deploy.
- Uji konektivitas `/api/health`, validasi file rusak, round-trip LibreOffice dan alur unduhan Chromium pada 15 jurnal.
- Penanganan khusus MABNY: wrapper Google Docs kosong yang menyebabkan LibreOffice menghasilkan XML tidak valid dibersihkan sebelum penyimpanan ulang.

**Batas kepastian**: sukses dalam LibreOffice/struktur DOCX belum membuktikan 100% kesesuaian visual maupun kompatibilitas Word 2010 semua dokumen. Pengguna pernah membuktikan LibreOffice-resaved RABBANI terbuka, tetapi versi v10 perlu dicoba langsung.
