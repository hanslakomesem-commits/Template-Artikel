# ZAIN.NET Multi Template Artikel v10 — hasil DOCX dari LibreOffice Server

**15 jurnal, GitHub Pages + backend Python/LibreOffice.** Ini adalah revisi dari versi browser-only yang pada Microsoft Word 2010 pengguna pernah gagal dibuka. **Uji lokal:** 15/15 jurnal berhasil melalui tombol browser → backend LibreOffice → download; 4/4 uji API berhasil. Ini bukan uji langsung di Word 2010 pengguna.

## Alur pemrosesan

1. Di GitHub Pages, pengguna upload artikel dan pilih salah satu dari 15 DOCX master jurnal.
2. JavaScript (JSZip) membuat **DOCX sementara** dari artikel dan master jurnal. Ini hanya tahap susun naskah, **tidak boleh diunduh sebagai hasil final**.
3. Browser mengirim DOCX sementara ke `POST /api/normalize` melalui HTTPS.
4. Backend Docker menjalankan **LibreOffice headless** untuk membuka dan **Save As DOCX** di direktori sementara.
5. Backend memeriksa struktur DOCX hasil dan jumlah referensi footnote; hasil LibreOffice dikirim ke browser.
6. Browser mengunduh **hanya file DOCX yang sudah disimpan ulang oleh LibreOffice**. Jika server gagal, tidak ada hasil mentah yang ditawarkan.

**Mengapa?** Pada Word 2010 pengguna, hasil browser RABBANI v8 terbukti crash sedangkan file yang disimpan ulang melalui LibreOffice terbukti terbuka. Ini bukan jaminan universal: setiap Word/OS dan dokumen berbeda, tetapi strategi server ini mengikuti metode yang sudah berhasil.

## Langkah A — upload ZIP ke GitHub

1. BACKUP repository yang sudah ada (misalnya unduh ZIP branch `main`).
2. Ekstrak paket v10 ini. Upload **isi** ZIP ke root repository `https://github.com/hanslakomesem-commits/Template-Artikel` branch `main`.
3. Pastikan file `index.html`, `config.js`, `rules.json`, `templates.json`, `render.yaml`, folder `js/`, `lib/`, `templates/` dan `backend/` berada langsung di root.
4. Commit changes. **Jangan gunakan website produksi sebelum backend berhasil deploy.**

## Langkah B — deploy backend LibreOffice di Render

1. Masuk `https://dashboard.render.com/` → `New` → `Web Service`.
2. Hubungkan GitHub dan pilih repository `Template-Artikel`, branch `main`.
3. Pilih **Docker** sebagai Language/Runtime; Dockerfile path: `./backend/Dockerfile`; Docker Context: `.` (root repo).
4. Pilih region dan instance sesuai sumber daya; Render memiliki opsi free untuk eksperimen dengan keterbatasan memori, penundaan ketika idle, dan kapasitas. LibreOffice dapat membutuhkan memori cukup besar; naikkan spesifikasi jika terjadi OOM.
5. Set environment variable: `ALLOWED_ORIGINS=https://hanslakomesem-commits.github.io`
   (jika website Anda menggunakan domain khusus, tambahkan origin tepatnya dengan koma, contoh `https://www.karyazainnet.net`).
6. Health Check Path: `/api/health`. Klik Create/Deploy dan pastikan deploy hijau.
7. Dapatkan URL HTTPS dari Render, misalnya `https://nama-service-anda.onrender.com`. Ini contoh; **jangan pakai URL contoh sebelum layanan Anda benar-benar ada**.
8. Uji lewat browser: `https://nama-service-anda.onrender.com/api/health`. Harus mengembalikan `"status":"ok"`.

**Alternatif Blueprint**: root memiliki `render.yaml`. Di Render pilih `New` → `Blueprint` dan pilih repo. Periksa pengaturan Docker dan paket sebelum deploy.

## Langkah C — hubungkan website GitHub Pages dengan backend

1. Di repository GitHub, buka `config.js` → klik ikon pensil Edit.
2. Ubah:
   ```js
   window.ZAINNET_CONFIG = { apiBaseUrl: 'https://URL-ASLI-BACKEND-ANDA.onrender.com' };
   ```
3. Commit changes ke branch `main`.
4. Tunggu GitHub Pages deploy. Buka `https://hanslakomesem-commits.github.io/Template-Artikel/` lalu `Ctrl + Shift + R`.
5. Upload artikel, pilih RABBANI, klik **Format → Server LibreOffice → Download DOCX**.
6. Nama unduhan akan diakhiri `_V10_LO_WORD2010.docx`.
7. Buka dan edit hasilnya menggunakan Microsoft Word 2010. Jika masih crash, jangan anggap semua jurnal lolos; catat file dan kesalahannya.

## Penting — perlindungan artikel

- File DOCX sementara dikirim ke server Render melalui HTTPS untuk diproses. Berbeda dari versi browser-only, isi artikel **tidak lagi sepenuhnya lokal**.
- Kode backend tidak menyimpan artikel secara permanen. File sementara dihapus ketika pemrosesan selesai. Provider hosting tetap memiliki kebijakan infrastrukturnya sendiri.
- Hindari mengirim naskah sensitif jika tidak ingin data diproses di server eksternal.
- Tidak ada akun, API key, atau `service_role` yang perlu disimpan di source website.
- CORS membatasi browser-origin, tetapi **bukan sistem autentikasi atau pembatasan pemakaian**. Untuk layanan publik dengan volume tinggi, tambahkan autentikasi, rate limiting di reverse proxy, pengawasan biaya, dan pembatasan sumber daya.
- DOCX maksimal 25 MB, proses server berseri (1 job setiap saat); server membutuhkan LibreOffice terpasang.

## Hasil pengujian dan risiko

- Uji server LibreOffice: **15/15 jurnal berhasil**.
- Uji UI Chromium lokal (request server disimulasikan melalui fungsi konverter yang sama): **15/15 pilihan jurnal menghasilkan unduhan hasil LibreOffice**, bukan DOCX mentah.
- Empat uji API otomatis berhasil.
- RABBANI hasil server telah dirender dan diperiksa sebagai dokumen 27 halaman; logo, kotak abstrak Indonesia, footer, dan footnote tampak.
- MABNY membutuhkan pembersihan *content control* kosong dari Google Docs sebelum konversi. Konten teks dan footnote tetap dipertahankan.
- LibreOffice dapat mengubah jumlah berkas media/objek gambar meski tampilan halaman tidak berubah jelas (contohnya SHAFIN). Buka hasil nyata di Word dan cek semua gambar, footnote, tabel, dan nomor halaman.

## Yang belum dijamin

- Kesesuaian visual tiap jurnal tidak otomatis 100%. Pembukaan berhasil tidak berarti seluruh format, header/footer, dan catatan kaki sepenuhnya identik. Periksa hasil sebelum dicetak/dikirim.
- LibreOffice bisa menata ulang halaman atau mengganti elemen Word kompleks (VML/textbox/chart/field). Jangan menebak volume, DOI, dan tanggal penerbitan.
- Backend tidak membuat PDF. Fitur final saat ini **DOCX yang disimpan ulang oleh LibreOffice**.
- Pada Render Free, kinerja LibreOffice bisa terganggu keterbatasan RAM, cold start, dan batas pemakaian; layanan produksi mungkin membutuhkan instance berbayar.

## Struktur kode

```text
index.html             — tampilan GitHub Pages
config.js              — URL HTTPS backend (atur setelah deploy)
rules.json             — aturan 15 jurnal
templates.json         — daftar 15 jurnal
js/docx-engine.js      — komposisi DOCX sementara berdasar master asli
js/app.js              — upload, susun, POST ke backend, download hasil LO
lib/jszip.min.js       — pustaka JSZip (bundled)
templates/*/           — seluruh master DOCX asli
backend/main.py        — API FastAPI, validasi DOCX, LibreOffice Save As
backend/Dockerfile     — image Docker dengan LibreOffice Writer
backend/requirements.txt
render.yaml            — pilihan deploy Render Blueprint
tests/                 — pengujian API
```

## Menjalankan backend lokal pada Linux (opsional)

```sh
sudo apt install libreoffice-writer fonts-liberation
python -m venv .venv && . .venv/bin/activate
pip install -r backend/requirements.txt
uvicorn backend.main:app --host 127.0.0.1 --port 10000
```

Uji `http://127.0.0.1:10000/api/health` — gunakan HTTP hanya saat pengujian localhost; production wajib HTTPS.
