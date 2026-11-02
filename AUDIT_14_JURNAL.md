# Audit ZAIN.NET — 14 Jurnal — Versi 4.0

## Ruang lingkup

- 14 file Word jurnal diperiksa dari `templates/<id>/original.docx` atau `template.docx` (KARSA).
- 14 template asli berhasil dikonversi ke PDF untuk baseline visual.
- **112 pengujian produksi DOCX**: 14 jurnal × 8 artikel uji (1 artikel pengguna yang nyata, 7 artikel sintetis dengan variasi judul panjang, paragraf panjang, tabel/gambar, multibahasa, abstrak Inggris tidak ada, dan artikel pendek).
- 112/112 proses selesai tanpa pengecualian. 112/112 hasil lolos pemeriksaan ZIP, XML, kelengkapan gambar asli, jumlah footnote sumber, keberadaan tabel sumber, daftar pustaka, dan placeholder running header utama.
- 14 hasil artikel standar berhasil dirender menjadi PDF. Halaman pertama dan kedua semua jurnal ditinjau melalui lembar kontak visual.
- 14 hasil judul panjang dirender menjadi PDF untuk memeriksa pemotongan judul; hasil tetap memerlukan validasi batas jumlah kata jurnal.
- 3 hasil bergambar/tabel (KARSA, KIDDO, SOCIALE) dirender dan diperiksa keberadaan gambar serta label tabel.

**Makna angka di atas:** lulus verifikasi proses dan kelengkapan struktur; *BUKAN* klaim identik pixel-per-pixel atau aturan publikasi 100% benar. Tidak dilakukan 100 kali per jurnal karena pengulangan identik bukan pengganti analisis variasi dokumen atau inspeksi editor Word.

## Hasil berdasarkan template

| Jurnal | Halaman contoh asli | Masalah yang ditemukan / diperbaiki | Sisa risiko |
|---|---:|---|---|
| KARSA | 8 | Header penulis/judul dinamis, tanggal contoh footer dihapus, logo dan kotak asli dipertahankan | Spasi halaman pertama; footnote dan tanggal editorial harus cek Word |
| SOCIALE | 6 | Header `Author Surnames`/`Article Tittle`, volume contoh **2099** dinetralkan | Posisi *corresponding author* dan tabel dua bahasa |
| PERDIKAN | 4 | Abstrak Indonesia/Inggris ke sel tabel asli, metadata judul dan penulis | Lebar kolom abstrak; info artikel setelah terbit |
| PANYONARA | 6 | Penulis di header lanjutan; aturan Cambria dan jarak tepat **17 pt** | Dua section, metadata penulis jamak, 4 kata kunci maksimum |
| MABNY | 5 | Judul lanjutan, area abstrak terstruktur dan tabel asli | **3 section** dan wajib Purpose/Methodology/Findings/Implications/Originality |
| MEYARSA | 7 | Header nama penulis/judul pendek dinamis | Footer tahun/volume contoh dan artikel dua bahasa |
| KIDDO | 5 | **Header genap nama penulis, ganjil judul pendek**; tabel abstrak + logo dipertahankan; rata kiri-kanan isi | Tabel halaman awal, Word 2010, dan nomor halaman |
| El Nubuwwah | 8 | Header penulis/judul, dua abstrak dan tanggal editorial contoh | Font Book Antiqua, footer penerbitan |
| GHANCARAN | 9 | Running header nama penulis, *first-page* dan tabel abstrak | **6 section** dari dokumen asli; interaksi section break dan ketentuan 12 kata judul |
| Al-Huquq | 3 | Header nama penulis/judul, font Book Antiqua | Footnote, panjang naskah, footer jurnal |
| SHAFIN | 4 | Header nama penulis, struktur tabel ARTICLE INFO/ABSTRACT | Konten abstrak, metadata penulis jamak, edisi penerbit |
| NUANSA | 7 | Judul dalam Word textbox/AlternateContent, tidak lagi memakai teks contoh header | **Judul panjang berisiko tidak muat**; versi DOC konversi dan Word 2010 |
| E-NASA | 6 | Header nama penulis/judul dan tabel abstrak asli | DOC lama hasil konversi, penomoran editorial |
| IQTISHADIA | 5 | Header nama penulis/judul pendek dan tabel ARTICLE INFO/ABSTRACT | DOC hasil konversi, edisi/DOI penerbitan |

## Risiko yang sengaja tidak ditutup-tutupi

1. **Jangan menyalin data penerbitan contoh sebagai fakta mahasiswa:** volume, DOI, issue, nomor, tanggal received/accepted/published harus disahkan penerbit. Perbaikan saat ini menetralkan contoh yang paling bermasalah (KARSA, SOCIALE), dan memberi peringatan untuk sisanya.
2. Teks artikel asli disalin secara konservatif, tetapi tidak semua elemen Word tingkat lanjut (OLE, persamaan, komentar, teks melayang) telah ditelusuri. Berkas yang mengandung objek ini perlu peninjauan manual.
3. Template yang memakai banyak section atau *text box* membutuhkan pemeriksaan halaman akhir per artikel: terutama **MABNY, GHANCARAN, PANYONARA, NUANSA**.
4. Tidak ada penerjemahan atau penulisan ulang abstrak otomatis; jika hanya ada satu bahasa, pengguna harus melengkapi bahasa kedua sesuai jurnal.
5. Pengujian visual memakai LibreOffice, bukan Microsoft Word 2010 (yang digunakan dalam screenshot Anda). Hasil dapat bergeser akibat font, kerning, compatMode dan layout engine.
6. **Tidak ada dasar untuk klaim 100% sama** pada artikel apa pun sebelum diuji penuh di Microsoft Word yang menjadi target pemakaian.

## Cara QA berikutnya untuk naskah nyata

Bandingkan halaman 1, 2, halaman bergambar/tabel, halaman dengan footnote, dan halaman terakhir dengan pedoman jurnal terbaru. Periksa isi semua halaman (jangan hanya halaman pertama); pastikan kata kunci, tanggal, DOI dan nomor halaman tepat, tidak ada judul terpotong, serta tidak ada referensi footnote hilang.
