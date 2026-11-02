# ZAIN.NET Multi Template Artikel — GitHub Pages Only v4.0 (Audit & Regenerasi)

**Status: BETA TERUJI STRUKTURAL, BUKAN JAMINAN 100% IDENTIK.**

Proyek ini mengolah artikel mahasiswa berformat `.docx` langsung di browser dan membuat `.docx` hasil berdasarkan **file Word asli jurnal**, bukan menggambar ulang header/logo dengan HTML. Tidak perlu API key, URL backend Python, Supabase, server, maupun login.

## 14 template dalam paket
KARSA, SOCIALE, PERDIKAN, PANYONARA, MABNY, MEYARSA, KIDDO, El Nubuwwah, GHANCARAN, Al-Huquq, SHAFIN, NUANSA, E-NASA, dan IQTISHADIA. Berkas `original.doc` lama juga disimpan untuk jurnal yang pada awalnya menggunakan `.doc` (NUANSA, E-NASA, IQTISHADIA). Versi konversinya `.docx` dipakai untuk pemrosesan browser.

## Isi folder yang diunggah ke GitHub

- `index.html`: antarmuka website responsif.
- `js/docx-engine.js`: pengolah OOXML template-first; rekonstruksi identitas dan penggabungan badan artikel.
- `js/app.js`: UI, unggah, unduh, dan peringatan.
- `lib/jszip.min.js`: pemroses ZIP/DOCX lokal (lisensi di `lib/JSZIP_LICENSE.md`).
- `rules.json`: konfigurasi aturan dan posisi field 14 jurnal.
- `templates.json`: daftar pilihan jurnal dan lokasi file asli.
- `templates/<jurnal>/`: berkas template Word yang pernah dikirim.
- `AUDIT_14_JURNAL.md`: ringkasan 112 tes dan batasan validasi.

## Memasang di GitHub Pages

1. Ekstrak file ZIP hasil download.
2. Masuk ke repository `https://github.com/hanslakomesem-commits/Template-Artikel`.
3. Unggah **seluruh isi** paket ke **root branch `main`**, bukan ke folder `frontend`, bukan pula ZIP-nya.
4. Di root harus tampak `index.html`, `js/`, `lib/`, `templates/`, `rules.json`, `templates.json`.
5. Bila ada file `index.html` lama, pastikan file versi baru menggantikannya. Tidak perlu menghapus seluruh repository atau memindahkan folder jurnal lain yang lama terlebih dahulu.
6. Buka `Settings → Pages`. Untuk branch deployment, pilih `main`, `/ (root)`, kecuali jika Actions sudah sengaja dipakai.
7. Tunggu deployment yang terbaru dinyatakan sukses di GitHub, lalu buka `https://hanslakomesem-commits.github.io/Template-Artikel/` dan tekan `Ctrl + Shift + R`.

## Cara memakai

1. Pilih salah satu jurnal.
2. Unggah artikel mahasiswa format DOCX. Untuk penggunaan otomatis, artikel harus mempunyai bagian `Pendahuluan` atau `Introduction`.
3. Bila judul, penulis, email, abstrak, dan keyword terdeteksi keliru, isi kolom metadata manual.
4. Klik `Proses dan Download Word (.docx)`.
5. **Wajib periksa semua halaman**, termasuk header ganjil/genap, DOI, kotak abstrak, footnote, tabel, dan daftar pustaka dalam Word sebelum diserahkan ke penerbit.

## Apa yang dilakukan mesin

- Menggunakan DOCX jurnal asli sebagai dokumen induk untuk mempertahankan logo, header/footer, tabel abstrak, style halaman dan aset.
- Mengisi slot judul, penulis, afiliasi, email, abstrak, kata kunci, dan tanggal apabila diberikan.
- Mengimpor teks, tabel, gambar, relasi media dan catatan kaki Word sumber.
- Menyesuaikan font/ukuran, perataan isi, heading serta spasi dasar sesuai konfigurasi jurnal.
- Menangani running header berbeda pada jurnal yang memilikinya, termasuk KIDDO dan KARSA.
- Tidak menciptakan tanggal editorial atau DOI mahasiswa.

## Batasan teknis yang belum dapat dijamin

- Nama penulis multi-afiliasi, rumus, sitasi otomatis/Zotero, chart OLE, komentar, objek terhubung, SmartArt, text box nonstandar, endnote kompleks, dan layout multi-section langka belum dijamin 100%.
- Judul dengan panjang melebihi batas jurnal tidak bisa dibuat patuh hanya dengan mengecilkan font; ikuti peringatan.
- Abstrak wajib berbentuk khusus (misalnya MABNY) tidak bisa diubah dari abstrak biasa tanpa campur tangan penulis.
- Headline volume/tahun, DOI dan tanggal contoh harus diperiksa per edisi. **Belum semua placeholder penerbit di template dapat diganti otomatis dengan data resmi**.
- Pemrosesan di browser menggunakan struktur OOXML Word, tetapi pengujian visual dijalankan memakai LibreOffice; tampilan Word 2010 pengguna mungkin berbeda.
- Versi browser tidak menyediakan konversi PDF berkualitas setara Word; gunakan Microsoft Word `Save As → PDF`.

## Pengujian

Lihat `AUDIT_14_JURNAL.md`. Pada pengujian lokal: **112 kombinasi file/jurnal** lolos pembuatan DOCX dan pemeriksaan kelengkapan elemen dokumen; 14 template asli, 14 dokumen uji standar, 14 judul panjang, dan 3 dokumen bergambar diperiksa melalui PDF dari LibreOffice. **Pengujian ini tidak memvalidasi kepatuhan akademik 100%.**
