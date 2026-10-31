# ZAIN.NET — Multi Template Artikel (Revisi Otomatis Tanpa API Key)

Sistem pemformatan dokumen Word untuk template **KARSA**, **RABBANI**, dan **AL-IHKAM**. Tidak memakai API OpenAI/Gemini, tidak meminta API key, dan tidak memakai Supabase.

**PENTING:** Program tetap memakai **server Python/FastAPI** untuk mengolah DOCX. GitHub Pages sendiri hanya dapat menampilkan halaman web; GitHub Pages tidak dapat menjalankan kode Python.

## Pilihan 1 — Paling mudah: satu hosting untuk web + mesin dokumen

Dengan pilihan ini semua fitur menggunakan **satu alamat website** dan tidak perlu mengatur alamat API satu kali pun.

1. Ekstrak ZIP dan unggah **semua isi ZIP** ke repository GitHub (termasuk `templates/`, `backend/`, dan `render.yaml`). Jangan mengunggah ZIP sebagai satu file.
2. Buka dashboard [Render](https://dashboard.render.com/), hubungkan akun GitHub, buat **New > Blueprint**, kemudian pilih repository tersebut. File `render.yaml` sudah berisi langkah deploy otomatis untuk layanan Python.
3. Deploy sesuai konfigurasi pada `render.yaml`. Jika tersedia paket gratis, paket itu dapat dipilih; cek batas penggunaan/ketentuan penyedia sebelum menggunakannya. **Jangan memasukkan API key.**
4. Setelah aktif, buka alamat website dari layanan Render. Web, `/health`, `/templates`, dan `/format` semuanya berada di alamat domain **yang sama**.
5. Pastikan indikator menunjukkan **“Mesin pemroses terhubung”**. Unggah DOCX, pilih template, tekan **Proses & Download Word**.

Perintah server jika diinstal secara manual (jalankan dari root proyek):

```bash
python -m pip install -r backend/requirements.txt
uvicorn backend.main:app --host 0.0.0.0 --port 8000
```

Kemudian buka http://localhost:8000 melalui browser di komputer yang menjalankan server. `index.html` tidak perlu dibuka terpisah. Di hosting yang menyediakan variabel `$PORT`, gunakan start command di `render.yaml`.

## Pilihan 2 — Tetap menggunakan GitHub Pages untuk tampilan web

Gunakan pilihan ini hanya jika ingin halaman web tetap berada di GitHub Pages. **Butuh dua layanan:** GitHub Pages untuk tampilan dan hosting Python untuk mesin dokumen. URL Python hanya perlu diatur **sekali oleh admin**, bukan diisi mahasiswa.

1. Deploy backend ke Render atau hosting Python lain menggunakan repository yang sama dan pastikan alamat seperti `https://alamat-backend-anda.onrender.com/health` mengembalikan JSON dengan `"ok": true`.
2. Di **file `config.js` pada root repository GitHub**, ubah konfigurasi:

   ```javascript
   window.ZAINNET_CONFIG = {
     apiBaseUrl: 'https://alamat-backend-anda.onrender.com'
   };
   ```

   Ganti URL contoh dengan **alamat backend sebenarnya** yang Anda terima setelah deploy. Jangan menambahkan `/format`, `/templates`, atau `/health`. URL ini **bukan API key** dan aman ditampilkan sebagai alamat layanan publik.
3. Pastikan `index.html`, `config.js`, dan `app.js` diunggah bersama pada folder yang menjadi sumber GitHub Pages. Jika GitHub Pages dikonfigurasi dari folder root, gunakan berkas root; folder `frontend/` adalah salinan untuk deploy terpisah.
4. Atur variabel `ALLOWED_ORIGINS` pada hosting Python sesuai alamat frontend, misalnya `https://USERNAME.github.io` atau domain kustom. Bila perlu mengizinkan lebih dari satu domain, pisahkan dengan koma. Variabel ini mengendalikan CORS, bukan autentikasi.
5. Buka website GitHub Pages. Indikator koneksi akan diperiksa otomatis dan pengguna tidak perlu memasukkan URL apa pun.

**Catatan:** Bila `config.js` masih kosong di GitHub Pages, akan muncul pesan konfigurasi untuk admin karena GitHub Pages tidak bisa menjalankan Python. Situs tidak dapat memproses DOCX sampai hosting backend benar-benar aktif. Untuk mode satu hosting, **biarkan `apiBaseUrl: ''`**.

## Fitur yang dipertahankan

- Pemilihan KARSA, RABBANI 2023/panduan web, dan AL-IHKAM 2026.
- Upload DOCX (maksimum 15 MB) dan hasil unduh Word DOCX.
- Metadata publikasi opsional, pengaturan jumlah paragraf judul, serta catatan validasi.
- Pemroses asli memakai sumber DOCX agar media, tabel, dan footnote Word tetap terjaga pada kasus yang didukung.
- Tautan unduh muncul setelah server memproses dokumen; verifikasi tata letak akhir tetap perlu dilakukan di Word.

## Fitur yang diperbaiki

- Kolom **Alamat mesin pemroses (API)** dihapus dari antarmuka pengguna.
- URL mesin pemroses menggunakan **alamat website sendiri** secara otomatis bila frontend dan backend digabung.
- Untuk GitHub Pages, URL disimpan hanya di **`config.js`**, tidak di localStorage/browser pengguna.
- Indikator koneksi **Terhubung / Belum terhubung** dan tombol **Coba Sambungkan Lagi**.
- Pemeriksaan layanan `/health` dan daftar `/templates` sebelum tombol pemrosesan diaktifkan.
- Pengecekan tipe file dan batas ukuran file sebelum mengunggah.
- Pesan gagal yang lebih jelas; pilihan template diambil dari backend.
- `render.yaml` untuk deploy satu hosting; endpoint root `/` dan aset JS dilayani oleh FastAPI.

## Pengujian

```bash
python -m pip install -r backend/requirements.txt
python -m pip install httpx
python -m unittest discover -s tests -v
```

Pengujian backend meliputi pemformatan dan detail layout ketiga jurnal, footnote/media, validasi file, CORS, serta akses halaman web/JS dari backend yang sama.

## Batasan yang belum berubah

Ini masih versi **beta**, bukan engine desain jurnal yang telah lulus verifikasi identik di semua jenis artikel. Header, pagination, tabel/gambar panjang, footnote, dan sitasi tertentu mungkin memerlukan pemeriksaan manual di Microsoft Word. Sistem belum menyiapkan PDF, login, antrean pengunjung, perlindungan terhadap penyalahgunaan layanan publik, ataupun penyimpanan arsip. Layanan hosting gratis dapat memiliki pembatasan sumber daya dan kondisi idle. Jika dipakai untuk mahasiswa dalam jumlah besar, gunakan hosting dengan kapasitas memadai dan tambahkan rate limit serta autentikasi sesuai kebutuhan.
