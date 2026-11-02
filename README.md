# Perbaikan kompatibilitas Word 2010 — v6

**Peringatan:** Microsoft Word 2010 tidak tersedia untuk pengujian otomatis dalam lingkungan ini. Perbaikan mencegah satu kelas masalah XML yang terukur, tetapi belum menjamin Word 2010 tidak akan crash karena add-in, font, atau masalah Office lokal.

Pada v5 ditemukan 26 kelompok properti Word dengan urutan XML tidak sesuai schema pada hasil RABBANI; dokumen asli 0. Versi ini menormalisasi `w:pPr`, `w:rPr`, `w:tblPr`, `w:tcPr` sebelum DOCX diunduh. Perbaikan berlaku ke semua 15 jurnal.

Jika Word 2010 masih crash, gunakan dokumen pemulihan yang telah disimpan ulang lewat LibreOffice (tersedia terpisah) dan coba Word dalam safe mode (`winword /safe`).

# ZAIN.NET Multi Template Artikel — GitHub Pages Only v5.0

**15 jurnal, termasuk RABBANI NEW 2023. Tanpa Python backend, API key, Supabase, ataupun server pemroses.**

Paket ini menambahkan template DOCX **RABBANI** milik pengguna kepada versi v4.0 (14 jurnal) dengan menjaga salinan Word resmi tetap utuh pada `templates/rabbani/original.docx`. Total sekarang: **15 pilihan jurnal**, bukan 16 karena RABBANI belum ada dalam paket v4.

## Pasang / perbarui GitHub Pages

1. Ekstrak ZIP; **jangan** unggah ZIP sebagai satu file.
2. Buka repository: https://github.com/hanslakomesem-commits/Template-Artikel
3. Di branch `main` unggah **seluruh isi folder hasil ekstrak ke root**. Pastikan file `index.html`, `rules.json`, `templates.json`, folder `js`, `lib`, dan `templates` ada di root, bukan dalam `frontend/` atau folder ZIP tambahan.
4. File dengan nama yang sama harus **diganti** dengan versi terbaru (`index.html`, `rules.json`, `templates.json`, `js/app.js`, `js/docx-engine.js`). Folder `templates/rabbani/original.docx` juga harus ikut.
5. Pantau **Actions / Pages** sampai deploy berhasil; buka https://hanslakomesem-commits.github.io/Template-Artikel/ lalu lakukan `Ctrl+Shift+R` agar cache lama hilang.

## Cara pakai

1. Pilih jurnal RABBANI (atau jurnal lain).
2. Upload artikel `.docx` dengan bagian "Pendahuluan" / "Introduction" untuk deteksi badan artikel.
3. Pada panel metadata, isi nama penulis, judul, afiliasi, email, dan abstrak bila deteksi kurang tepat. Tersedia isian kontak WA khusus yang memerlukan nomor tersebut.
4. Isi DOI, volume/nomor/tahun/tanggal **hanya jika sudah disahkan redaksi**. Jika belum, biarkan kosong.
5. Proses dan unduh Word DOCX. Periksa halaman pertama, header genap/ganjil, footer, footnote dan kata kunci di Microsoft Word sebelum dipakai untuk pengajuan jurnal.

## Hasil dan batasan

- Uji dasar: file asli **15 jurnal** menghasilkan DOCX melalui engine browser dari satu artikel yang sama tanpa error.
- Uji khusus RABBANI dengan artikel Anggara: judul, identitas, halaman awal, header, logo, footnote, bibliografi dan dua abstrak diperiksa dalam render LibreOffice. **8 referensi catatan kaki asli tetap ada**.
- Template RABBANI memakai kotak teks dengan tinggi tetap. Bila isinya panjang, program mengalihkannya ke **tabel Word dua kolom** agar tidak terpotong. Perubahan ini berarti **tata letak tidak selalu 100% sama** dengan template asli.
- Ini bukan mesin validasi formal semua pedoman jurnal. Semua file masih harus diperiksa dalam Word 2010/Word terbaru sebelum cetak, publikasi atau submit.

Baca `ATURAN_RABBANI.md` untuk detail. Audit 14 jurnal terdahulu tersedia dalam `AUDIT_14_JURNAL.md`.
