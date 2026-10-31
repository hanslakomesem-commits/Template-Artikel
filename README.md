# ZAIN.NET Multi Template Artikel — KARSA + RABBANI + AL-IHKAM

Paket revisi menambahkan template kedua RABBANI. Frontend ada di index.html dan frontend/index.html (identik). Pemroses tetap Python/FastAPI; GitHub Pages saja tidak bisa menjalankan pemroses Word.

Template ketiga: AL-IHKAM 2026, berdasarkan file docx.docx yang diunggah. Pilih AL-IHKAM di web; pemroses memakai Letter, margin 5/4,5 cm, abstrak tanpa panel, header/footer asli dan metadata editorial tersendiri. Rincian hasil pemeriksaan tujuh halaman ada di templates/alihkam/LAYOUT.md. Perbarui backend/alihkam.py dan seluruh folder templates/alihkam bersama frontend/backend utama.

## Cara memperbarui web
1. Ekstrak ZIP, lalu unggah seluruh isinya untuk menggantikan berkas proyek lama di GitHub.
2. GitHub Pages: gunakan branch main dan folder / (root); index.html sudah tersedia di root.
3. Perbarui backend juga, termasuk backend/rabbani.py dan folder templates/rabbani. Jangan hanya mengganti index.html.
4. Install: pip install -r backend/requirements.txt
5. Jalankan dari root proyek: uvicorn backend.main:app --host 0.0.0.0 --port 8000
6. Pada hosting yang menyediakan PORT, start command: uvicorn backend.main:app --host 0.0.0.0 --port $PORT
7. Masukkan alamat backend HTTPS di halaman web, pilih RABBANI, unggah artikel DOCX, lalu Proses & Download Word.
8. Atur ALLOWED_ORIGINS ke alamat frontend Anda, misalnya https://USERNAME.github.io.

## RABBANI
Sumber layout: templates/rabbani/template.docx, file asli yang diberikan pengguna (Template Jurnal Rabbani New 2023).
Panduan: https://ejournal.uinmadura.ac.id/index.php/rabbani/AuthorGuidelines

Dua pilihan disediakan karena file dan panduan berbeda:
- File template 2023 (default): judul 13 pt, abstrak 150–200 kata; abstrak dalam panel biru muda dan riwayat/kata kunci di sisi kanan.
- Panduan web: judul 12 pt dengan spasi tunggal; abstrak satu kolom sekitar 150 kata. Peringatan memakai rentang toleransi 130–170, bukan batas resmi jurnal.
- Keduanya: Book Antiqua, A4, margin 3 cm, isi 12 pt spasi 1,5, abstrak 11 pt spasi tunggal, daftar pustaka spasi tunggal.
- Halaman 1: identitas jurnal di kiri, logo di kanan; footer hak cipta, tanggal editorial, rentang halaman, dan nomor halaman di kanan.
- Halaman genap: penulis/judul pendek/DOI di kiri, logo kanan; footer diawali nomor halaman di kiri. Halaman ganjil lanjutan membalik posisi header dan menempatkan nomor halaman di akhir footer kanan.
- Header dan footer ganjil/genap menggunakan struktur serta posisi gambar dari template asli. Font Book Antiqua dan Garamond yang tertanam dalam template ikut dibawa agar metrik teks tidak berubah karena penggantian font.
- ABSTRACT bahasa Inggris lebih dahulu, lalu ABSTRAK Indonesia, dalam panel biru muda dengan bingkai. ARTICLE HISTORY dan KEYWORDS berada di kanan dengan bingkai tersendiri. Panel mengikuti aliran halaman; abstrak panjang dapat berlanjut ke halaman berikutnya. Isi tidak dipotong agar tampak sama dengan contoh pendek dalam template. Isi panel biru rata, bukan gradasi dekoratif template.
- Metadata dapat diisi di web. Kolom kosong ditandai —; tahun, DOI, volume, tanggal dan halaman contoh tidak dipakai. Judul pendek otomatis dipotong pada batas kata bila terlalu panjang; pengguna dapat menggantinya.

## Menyiapkan artikel
- Masukan adalah artikel DOCX yang sudah memiliki isi, bukan konversi skripsi menjadi artikel.
- Letakkan judul di awal. Pilih jumlah paragraf judul 1–3; judul yang membungkus menjadi beberapa baris tetapi tanpa Enter tetap satu paragraf.
- Berikutnya nama penulis, institusi, email, dan nomor kontak bila perlu.
- Gunakan judul bagian pada paragraf tersendiri: ABSTRACT, ABSTRAK, PENDAHULUAN, METODE PENELITIAN, HASIL DAN PEMBAHASAN, KESIMPULAN, DAFTAR PUSTAKA. Padanan bahasa Inggris dan awalan A./1. didukung.
- Kata kunci memakai label Keywords: / Kata kunci:, dipisahkan koma atau titik koma.
- Pendeteksian struktur bekerja pada paragraf utama. Abstrak yang berada di dalam text box atau tabel sumber perlu dipindah ke paragraf biasa terlebih dahulu; bila tidak ditemukan web menampilkan catatan.

## Batasan dan pemeriksaan
Isi ilmiah, nama, dan teks judul tidak ditulis ulang. Aplikasi tidak menerjemahkan abstrak, tidak memangkas judul, tidak mengarang daftar pustaka, dan tidak otomatis mengubah sitasi menjadi Chicago 17 Full Footnote. Catatan validasi ditampilkan setelah unduh.

Footnote asli, tabel, gambar, dan hubungan internal Word dipertahankan dengan memakai file sumber sebagai dasar. Periksa lebar gambar/tabel, pemenggalan halaman, gaya catatan kaki, serta detail sitasi pada Word. Gunakan font Book Antiqua pada komputer; aplikasi pembaca bisa menggantinya bila font tidak terpasang.

KARSA tetap memakai pemroses beta sebelumnya, dengan header sederhana dan keterbatasan layout yang sama. PDF dan pratinjau dalam web belum tersedia. Unggahan dikirim ke backend yang diisi pengguna, bukan diproses di GitHub Pages. Backend mengolah dalam memori dan belum menyediakan login, antrean, maupun pembatasan trafik.

## Pengujian
Dari root proyek: pip install httpx lalu python -m unittest discover -s tests -v
Pengujian mencakup daftar template, keluaran KARSA/RABBANI, kedua profil Rabbani, validasi permintaan, akses catatan melalui CORS, serta keberlangsungan byte footnote/media asli.
