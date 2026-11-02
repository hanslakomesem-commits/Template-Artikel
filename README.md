# ZAIN.NET Multi Template Artikel — GitHub Pages Only

**Rilis:** 2 November 2026. **Tidak menggunakan backend Python, API, database, atau layanan eksternal.**

## Apa yang sudah tersedia

- Web mobile-friendly di `index.html`.
- **14 pilihan jurnal**: KARSA, SOCIALE, PERDIKAN, PANYONARA, MABNY, MEYARSA, KIDDO, El Nubuwwah, GHANCARAN, Al-Huquq, SHAFIN, NUANSA, E-NASA, IQTISHADIA.
- Semua template Word **disimpan dalam bentuk file yang sama persis** dengan berkas aslinya. Dokumen `.doc` juga memiliki versi `.docx` hasil konversi. Konversi belum divalidasi identik secara visual.
- Pencarian katalog dan download setiap template asli.
- Upload artikel mahasiswa dalam bentuk `.docx` dan pemrosesan di browser melalui `js/engine.js` + JSZip yang tersedia secara lokal pada folder `lib/`.
- **AUTO DASAR (beta)** untuk seluruh 14 jurnal: mengubah ukuran halaman, margin, font, ukuran huruf, spasi, dan beberapa gaya judul/abstrak/heading sesuai profil yang tersedia. Isi artikel Word tetap menjadi dokumen sumber sehingga native footnote, gambar, dan tabel tidak sengaja dibuang.
- Opsi eksperimental: menyalin sebagian header/footer resmi dari paket template, dan menambahkan nomor halaman otomatis. **Opsi ini dinonaktifkan secara default** karena belum selalu sesuai tata letak.
- Download DOCX hasil yang tetap dapat diedit.

### Penting: batasan format

Pemformatan AUTO ini **bukan pengganti layout Word asli** dan tidak boleh disebut "100% sesuai jurnal". Hal yang memerlukan pemeriksaan/manual: logo yang diletakkan di body (bukan header), tabel abstrak multi-kolom, running header penulis/judul pada halaman ganjil/genap, kotak tanggal redaksi, footer, footnote rumit, page breaks, kolom, daftar pustaka, dan pencetakan PDF. Tidak ada auto-PDF agar tidak menghasilkan PDF berbeda dari Microsoft Word. Cara paling aman untuk jurnal resmi adalah mengedit teks di file template asli.

**Template berisi contoh** (tanggal, volume, DOI, nama penulis) yang bukan data mahasiswa. Jika memilih impor header/footer secara eksperimental, pastikan semua placeholder/contoh redaksi tidak terbawa sebagai data nyata. Khusus kotak tanggal publikasi pada footer, mesin mengganti contoh tanggal dengan tanda kosong `—` jika pola tanggal diketahui.

## Cara instalasi di GitHub Pages

1. Download file ZIP dan **ekstrak**.
2. Buat repository GitHub, contohnya `zainnet-multi-template`.
3. Upload **isi folder ZIP**, bukan folder pembungkusnya; pastikan file `index.html` di ROOT repository dan folder `js/`, `lib/`, `templates/` ikut diunggah.
4. GitHub repository → **Settings → Pages → Build and deployment → Deploy from a branch**.
5. Pilih branch `main`, folder `/ (root)`, lalu **Save**.
6. Setelah GitHub Pages aktif, buka `https://USERNAME.github.io/zainnet-multi-template/` (ganti USERNAME).
7. Pilih artikel `.docx`, jurnal, lalu klik **Proses & Download DOCX**.

**Tidak ada kolom URL backend**. Tidak ada file backend di paket. Hanya file statis yang dihosting GitHub.

## Struktur

```
index.html                UI web responsif
js/app.js                 UI / upload / download / validasi
js/engine.js              Engine OpenXML di browser
lib/jszip.min.js          Dependensi lokal JSZip (MIT/GPLv3 dual licensed)
templates.json            Katalog 14 jurnal
profiles.json             Profil tata letak dasar (beta)
templates/<id>/...        Berkas asli dan, jika perlu, hasil konversi DOCX
```

## Cara kerja dan privasi

File input diproses menggunakan `File.arrayBuffer()` dalam memori browser. Browser mengambil file template Word dari repository Pages dan membuka keduanya lewat JSZip. Artikel pengguna **tidak dikirim ke server pihak lain** oleh kode ini; output dibuat sebagai `Blob` dan diunduh lokal. Pengguna tetap harus memercayai sumber website dan browser yang digunakan. Tidak ada tracking/upload jaringan untuk artikel di kode yang disertakan.

## Standar kualitas dan pengujian

Pengujian fungsional dilakukan dengan artikel DOCX Anggara Putra Pratama dari percakapan pengguna. Seluruh 14 profil berhasil menghasilkan DOCX secara programatis, dapat dibuka dengan `python-docx`, dan bagian `word/footnotes.xml` dari artikel tetap dipertahankan. Tiga output contoh juga berhasil dikonversi melalui LibreOffice ke PDF. **Ini bukan sertifikasi kesesuaian visual** dengan template jurnal; pengujian browser end-to-end tidak dapat dijalankan pada lingkungan uji karena browser menolak membuka alamat lokal.

Batas ukuran upload yang disarankan di UI adalah 35 MB. Pada perangkat bermemori rendah, DOCX besar dapat membuat browser lambat atau kehabisan memori. Berkas `.doc` lama untuk input artikel harus dikonversi ke `.docx` terlebih dahulu.

## Memperluas hingga 30 jurnal

Taruh file template jurnal baru di `templates/<slug>/original.docx`, tambah entri pada `templates.json` serta profil pada `profiles.json`. Jangan menganggap template baru sudah sesuai otomatis sebelum pengujian struktur dan layout masing-masing selesai. Saat ini ada **14 jurnal**, bukan 30 jurnal lengkap.
