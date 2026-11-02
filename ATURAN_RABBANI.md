# RABBANI — Integrasi DOCX asli ke ZAIN.NET v5.0

**Sumber:** `Template Jurnal Rabbani New 2023(1).docx` yang diunggah pengguna. Salinan tanpa perubahan tersimpan pada `templates/rabbani/original.docx`.

**Penting:** Meskipun nama berkas menyebut "New 2023", **contoh identitas penerbitan di dalam Word bertuliskan Volume 1, Nomor 1, Maret 2020**. Volume/tahun contoh, DOI artikel contoh, tanggal proses editorial, dan nama contoh tidak boleh digunakan sebagai metadata artikel mahasiswa. Mesin menggantinya dengan tanda `—` kecuali metadata resmi dimasukkan.

## Tata letak dan aturan dari Word asli

- A4, margin atas, bawah, kiri, kanan sekitar 3 cm.
- Kepala halaman pertama: nama jurnal, ISSN dan tautan jurnal di kiri, logo RABBANI di kanan.
- Judul artikel di tengah, Book Antiqua 13 pt **tebal**, dapat 2 baris.
- Nama penulis Book Antiqua 11 pt tebal, afiliasi dan email 10 pt tebal. Nama/afiliasi/nomor email tidak boleh membawa superskrip dan teks panduan berwarna biru dari template contoh.
- Ada kolom kontak WhatsApp; isian boleh kosong jika tidak diketahui.
- `ABSTRACT` bahasa Inggris dan `ABSTRAK` bahasa Indonesia, masing-masing sekitar **150–200 kata**, menggunakan Book Antiqua 11 pt dalam kotak kiri berwarna biru muda.
- Kotak kanan menampilkan ARTICLE HISTORY (Received, Revised, Accepted) dan KEYWORDS.
- Isi jurnal: Book Antiqua sekitar 12 pt; satu setengah spasi; alinea justify.
- Header halaman lanjut memiliki logo di sisi yang berganti menurut halaman genap/ganjil; teks identitas penulis dan judul pendek, serta DOI.
- Footnote tetap merupakan catatan kaki Word; referensi dan tabel tetap elemen DOCX.
- Daftar pustaka pada contoh menyebut **Chicago Manual Style edisi ke-17 / full footnote**, tetapi program tidak otomatis menulis ulang gaya sitasi sumber.

## Penanganan kotak abstrak

Template asli memakai gambar/shape Word dengan tinggi tetap: bila abstrak mahasiswa lebih panjang dari contoh, teks bisa terpotong. Mesin v5.0 mempertahankan shape asli untuk abstrak/kata kunci singkat, namun **beralih ke tabel Word dua kolom yang mengalir mengikuti halaman** jika total abstrak melewati 215 kata atau kata kunci panjang. Ini mempertahankan isi, warna, dan konsep dua kolom, tetapi **tidak identik piksel dengan kotak asli**. Pemeriksaan manual tetap wajib.

## Batasan

Belum divalidasi pada Microsoft Word 2010 di perangkat pengguna. RENDER uji menggunakan LibreOffice. Nomor volume, bulan terbit, DOI, tanggal editorial, dan halaman terbit harus diisi dari penerbit, bukan diduga dari artikel contoh.
