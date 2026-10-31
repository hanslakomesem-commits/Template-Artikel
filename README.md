# ZAIN.NET Multi Template Artikel — 100% GitHub Pages

**Bisa langsung dipublikasikan di GitHub Pages. Tidak perlu API key, Render, Python, Node.js di hosting, Supabase, atau server pemrosesan.** Kode JavaScript dan JSZip dijalankan di browser pengguna. Dokumen artikel tidak diunggah ke layanan eksternal. Template DOCX disimpan di folder `templates/` dan dibaca dari situs statis melalui `fetch`.

## Instalasi GitHub Pages

1. Buat repository baru di GitHub, misalnya `zainnet-artikel`.
2. Ekstrak berkas ZIP ini. Unggah **isinya** ke root repository (`index.html`, `formatter.js`, `app.js`, `vendor/`, `templates/`), bukan ZIP-nya saja. Jangan melewatkan folder template atau vendor.
3. Buka **Settings → Pages → Build and deployment → Deploy from a branch**. Pilih branch `main`, folder `/(root)`, lalu Save.
4. Setelah GitHub Pages berhasil publish, kunjungi URL yang muncul di Settings → Pages, biasanya `https://USERNAME.github.io/zainnet-artikel/`. Nama repo disesuaikan.
5. Pilih jurnal, unggah artikel `.docx`, tekan **Proses & Download Word**. Unduhan diproses di browser. Jika tidak muncul otomatis, tekan tautan unduh yang muncul.

**Penting:** Membuka `index.html` melalui `file://` mungkin gagal memuat template karena keamanan browser. Gunakan GitHub Pages atau server file statis lokal seperti `python -m http.server 8000` untuk menguji. Perintah tersebut hanya untuk uji lokal; **tidak dibutuhkan** pada GitHub Pages.

## Struktur source code

- `index.html`: aplikasi web yang responsif untuk HP dan PC.
- `app.js`: tampilan, input, progres, dan unduh.
- `formatter.js`: mesin pemformat DOCX client-side (menyunting XML OOXML sambil mempertahankan struktur dokumen sumber).
- `vendor/jszip.min.js`: pustaka JSZip yang disertakan lokal, tidak memerlukan CDN atau akses ke layanan API.
- `templates/karsa/template.docx`: template asli KARSA.
- `templates/rabbani/template.docx`: template asli RABBANI.
- `templates/alihkam/template.docx`: template asli AL-IHKAM.
- `templates/*/rules.json`: deskripsi jurnal untuk pengembangan aturan selanjutnya.

## Fitur

- Memilih template KARSA, RABBANI, AL-IHKAM.
- Pemformatan judul, bagian artikel, abstrak, kata kunci, isi, dan referensi secara heuristik.
- Mengambil ukuran halaman dan margin serta bagian header/footer contoh dari file template jurnal.
- Nomor halaman `PAGE` pada footer yang memerlukan verifikasi akhir.
- Kotak abstrak/sidebar pada mode RABBANI 2023 (beta).
- Mempertahankan **paket DOCX sumber**, termasuk tabel, gambar, hyperlink, dan file footnote yang sudah ada, alih-alih membuat ulang artikel dari teks biasa.
- Bekerja secara statis tanpa API key dan tanpa database. File artikel diproses di memori browser.

## Batasan penting — jangan dilewatkan

Ini versi **BETA**. Pemformatan Word sempurna/identik dengan semua tata letak template **belum terjamin**. Mesin ini tidak menjalankan Microsoft Word, sehingga hasil bisa berbeda pada tabel kompleks, logo/artwork header, sitasi, penomoran, section break, layout multi-kolom, footnote, dan halaman sampul. Urutan elemen dan konten sumber dipertahankan semaksimal mungkin; isinya tidak diterjemahkan atau dibuat menggunakan AI. Pemeriksaan akhir dengan Word atau LibreOffice tetap wajib. File `.doc` lama, PDF, dan JPG tidak didukung. Jumlah paragraf awal untuk judul perlu dipilih jika judul terdiri atas 2–3 paragraf. Browser lama atau perangkat RAM rendah mungkin lambat untuk file yang sangat besar. Tampilan offline membutuhkan server lokal atau pengaturan cache dan tidak dijamin; setelah GitHub Pages diakses, pemformatannya tidak menggunakan backend.

Template AL-IHKAM contoh relatif besar (sekitar 5,6 MB), sehingga pengunjung perlu mengunduh template itu satu kali saat memproses. Seluruh aset berlisensi milik masing-masing jurnal/pemilik aslinya dan digunakan sebagai acuan tata letak.

## Diagnostik

- **Mesin gagal dimuat:** pastikan `vendor/jszip.min.js` dan `formatter.js` terunggah utuh.
- **Gagal mengambil template (404):** pastikan folder `templates/` dan nama direktori tepat, termasuk huruf kecil.
- **Word menyatakan ada masalah pada file:** gunakan fungsi repair Word lalu laporkan file sumber + template; dokumen kompleks tertentu memerlukan konversi backend yang lebih lengkap.
- **Font berbeda:** pasang font yang dibutuhkan template (Book Antiqua atau Times New Roman) di aplikasi pengolah Word.

## Lisensi

Kode ZAIN.NET disediakan untuk pengembangan sendiri. Pustaka `JSZip` MIT / GPLv3 (pilih salah satu lisensi sesuai proyek); lisensi sumber tersedia di https://github.com/Stuk/jszip . Hak template jurnal tetap pada pemilik masing-masing.
