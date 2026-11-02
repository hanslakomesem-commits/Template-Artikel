# Hasil pengujian (2 November 2026)

- Paket berisi 14 template jurnal, bukan 30.
- Seluruh file asli dari 13 jurnal tambahan cocok SHA-256 dengan unggahan pengguna.
- 14/14 template terbaca oleh mesin JavaScript dalam pengujian programatis Node.js.
- 14/14 percobaan **AUTO dasar** menghasilkan file `.docx` yang mempertahankan XML footnote dokumen sumber persis sama.
- 14/14 percobaan **header/footer eksperimental** menghasilkan DOCX yang dapat dibaca `python-docx` setelah perbaikan relasi gambar dalam header/footer.
- Tiga hasil uji bisa diekspor menjadi PDF menggunakan LibreOffice. Hasil visual Karsa dan Perdikan telah diperiksa pada halaman pertama; **belum identik** dengan template asli karena elemen editorial yang berada di body dan tata letak rumit belum dimigrasikan.
- Pengujian browser end-to-end via Playwright tidak bisa dilakukan di lingkungan ini karena Chromium memblokir navigasi ke alamat lokal dengan `ERR_BLOCKED_BY_ADMINISTRATOR`. Pengujian engine independen dilakukan melalui DOM XML JS dan JSZip di Node.
- Tidak ada server Python, endpoint, kredensial, database, CDN atau token yang dibutuhkan.

## Keterbatasan

Pemformatan ini adalah baseline/beta: jangan menganggap output 100% sesuai template jurnal atau sudah layak publikasi tanpa diperiksa Word. User dapat mengunduh template asli untuk pemformatan manual yang presisi.
