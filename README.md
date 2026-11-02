# ZAIN.NET MULTI TEMPLATE ARTIKEL — GitHub Pages Only, v3.1 (2 November 2026)

**TIDAK PERLU URL BACKEND PYTHON, RENDER, SUPABASE, ATAU API KEY.**

## Yang benar-benar dibuat

1. Website HTML/CSS responsif, bisa digunakan dari HP/komputer lewat GitHub Pages.
2. 14 jurnal: 13 jurnal baru dari unggahan pengguna dan KARSA dari file asli yang pernah diunggah.
3. DOCX jurnal **asli sebagai dasar Word hasil**, bukan mengarang logo, meniru tampilan dengan CSS, atau sekadar mengambil ukuran margin. Identitas penerbit, logo, tabel abstrak, header/footer, watermark dan bentuk awal template disimpan.
4. Mesin JavaScript `js/docx-engine.js` membaca artikel mahasiswa, mendeteksi judul/penulis/abstrak/kata kunci dan mengganti teks contoh di posisi spesifik template (`rules.json`).
5. Isi artikel mulai Pendahuluan/Introduction diimpor ke dokumen jurnal; footnote Word asli, tabel, gambar dan relasi media ikut dipindahkan. Gambar lebar dikecilkan bila lebih lebar dari area teks.
6. Hasil diunduh sebagai **Word DOCX** yang bisa diedit. **Tidak ada ekspor PDF otomatis di browser**; untuk PDF buka DOCX di Word lalu `Save As → PDF`.
7. Metadata identitas artikel, abstrak, dan tanggal editorial bisa diisi manual jika otomatis salah mendeteksi. Tanggal contoh artikel terbit tidak diterapkan sebagai tanggal artikel pengguna.
8. Deteksi peringatan: judul terlalu panjang, abstrak keluar dari rentang jumlah kata, abstrak dua bahasa tidak ada, struktur abstrak MABNY belum lengkap.

## Cara mengganti aplikasi lama di GitHub

Repository yang sebelumnya digunakan adalah:
`https://github.com/hanslakomesem-commits/Template-Artikel`

1. **Ekstrak file ZIP** ini di komputer.
2. Buka repository GitHub lalu pastikan branch `main`.
3. Pilih `Add file` → `Upload files` atau gunakan Git di komputer.
4. **Unggah seluruh isi folder hasil ekstrak**, bukan ZIP dan bukan folder induknya. Di root repository harus tampak:

```
Template-Artikel/
├── index.html
├── templates.json
├── rules.json
├── js/
│   ├── app.js
│   └── docx-engine.js
├── lib/
│   ├── jszip.min.js
│   └── JSZIP_LICENSE.md
└── templates/
    ├── karsa/template.docx
    ├── perdikan/original.docx
    ├── sociale/original.docx
    └── ... 11 jurnal lain
```

5. Pilih **Commit changes**. Pastikan `index.html` lama benar-benar diganti.
6. Di `Settings → Pages`, gunakan branch `main`, folder `/(root)` untuk workflow "Deploy from a branch"; jika sudah memakai GitHub Actions, pastikan deploy berhasil.
7. Setelah indikator deploy hijau, buka `https://hanslakomesem-commits.github.io/Template-Artikel/` dan tekan `Ctrl + Shift + R`. Pada halaman baru seharusnya terlihat **14 jurnal**, tulisan **Template Asli**, dan **tidak ada kolom URL backend**.
8. Folder lawas `backend`, `frontend`, dan folder jurnal yang terlanjur berada di root dapat dibiarkan dahulu. Setelah versi baru berjalan, bersihkan bila perlu.

## Alur pemrosesan

- Artikel pengguna **tidak diunggah ke server**. Browser mengunduh template statis dari GitHub Pages, lalu membaca Word yang dipilih dan menyusun hasil menggunakan JavaScript/JSZip di perangkat pengguna.
- `rules.json` memuat peta slot asli untuk setiap jurnal, termasuk bidang dalam tabel. Jika penerbit mengubah template, peta indeks perlu diperbarui.
- Untuk memulai, pilih jurnal, pilih file `.docx`, kemudian klik **Proses dan Download Word**.
- Jika nama, abstrak, atau afiliasi terdeteksi keliru, buka menu **Perbaiki identitas artikel secara manual**.

## Batasan penting (WAJIB DIBACA)

- **Tidak ada jaminan format identik 100% atau otomatis lulus pemeriksaan editor.** Jurnal memakai banyak pengaturan DOCX rumit: text box, frame, section, header ganjil-genap, spasi persis, tabel asimetris, watermark, dan variasi Word.
- Dokumen keluaran **harus diperiksa di Microsoft Word**, terutama halaman pertama/kedua, logo, tabel abstrak, gambar, nomor halaman, footnote, daftar pustaka dan tanggal metadata.
- Template **MABNY** memakai kolom/bingkai kompleks dan membutuhkan **abstrak terstruktur** (Purpose, Methodology, Findings, Implications, Originality). Aplikasi tidak boleh mengarang elemen abstrak yang tidak ada di artikel asal. Perlu pengecekan layout manual.
- Template **NUANSA** menggunakan kotak judul dengan tinggi terbatas. Untuk judul panjang, mesin menurunkan ukuran font supaya judul tidak terpotong, sekaligus memberikan peringatan untuk memendekkan judul mengikuti batas jurnal.
- Sebagian template berasal dari tahun **2019–2026**; tanggal, volume, edisi, DOI, dan identitas contoh dalam template belum tentu sesuai edisi penerbit saat ini. Jangan menganggap tanggal contoh sebagai data mahasiswa.
- Kualitas ekspor PDF secara langsung dari browser tidak bisa disamakan dengan Microsoft Word. Gunakan `Save As PDF` di Word atau LibreOffice setelah mengedit.
- Aplikasi memerlukan koneksi untuk membuka GitHub Pages dan mengambil file template saat pertama kali dipilih; **tidak** memerlukan server yang memproses dokumen.

## Lisensi dan file asli

Kode kustom ZAIN.NET ditulis untuk proyek ini. `lib/jszip.min.js` adalah dependensi pihak ketiga dan lisensinya ada di `lib/JSZIP_LICENSE.md`. File template jurnal tetap berhak cipta penerbit/pembuat masing-masing dan sebaiknya dipakai sesuai izin mereka. Berkas DOC asli dan konversi DOCX yang tersedia tetap disertakan.

## Perubahan source code

- `index.html`: UI format artikel, manual metadata, katalog 14 jurnal.
- `js/app.js`: kontrol upload/pilih/unduh dan tampilan laporan.
- `js/docx-engine.js`: pengganti total mesin lama; template-first dan OOXML media/footnote importer.
- `rules.json`: pemetaan lokasi judul, abstrak, penulis, isi, style tiap jurnal.
- `templates.json`: manifest seluruh Word yang disimpan.
- `TEMPLATE_MANIFEST_SHA256.json`: checksums untuk melacak orisinalitas berkas.

Lihat `PENGUJIAN.md` untuk rincian hasil QA.
