# CHANGELOG

## v4.0 — Audit & regenerasi 14 jurnal

- Memulihkan audit 14 berkas template asli; render baseline 14/14.
- Mengganti teks running header dari judul/nama mahasiswa dan mempertahankan PAGE field; termasuk KIDDO, KARSA, MEYARSA, NUANSA, dan jurnal lain sesuai relasi Word masing-masing.
- Mengoreksi kapitalisasi judul dan nama penulis apabila sumber memakai HURUF KAPITAL semua.
- Memberlakukan perataan `justify` pada paragraf isi, tanpa memaksa heading/gambar menjadi justify.
- Menetralkan tanggal Received/Accepted/Published di *semua representasi* footer KARSA dan volume contoh SOCIALE 2099.
- Menambahkan 112 uji variasi dokumen dan laporan risiko 14 template.
- Menggunakan DOCX asli setiap jurnal sebagai dasar; tanpa backend, Python, URL API atau layanan eksternal pada website terpublikasi.

**Diketahui belum sempurna:** ketepatan sempurna Word 2010, kompleksitas multi-section, abstrak terstruktur MABNY, tanggal/DOI penerbit, dan objek Word nonstandar.
