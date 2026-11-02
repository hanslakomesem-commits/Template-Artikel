# Perubahan

## v5.0 — Tambah RABBANI New 2023

- Jumlah jurnal dari 14 menjadi **15**.
- Berkas DOCX asli RABBANI ditambahkan tanpa perubahan di `templates/rabbani/original.docx`.
- Mapping judul, penulis, afiliasi, email, nomor WA, abstrak dua bahasa, kata kunci dan riwayat artikel yang ditempatkan pada shape di halaman pertama.
- Judul contoh yang terbagi di dua paragraf kini diisi sesuai judul mahasiswa.
- Teks header dinamis diperbarui, logo tetap dari template asli, footer dan PAGE field dipertahankan.
- Volume, nomor, DOI dan tanggal terbit **contoh 2020** dibersihkan dan diganti dengan data resmi saja bila ada.
- Mekanisme dua kolom mengalir disediakan ketika abstrak atau kata kunci terlalu panjang untuk shape asli.
- Metadata tambahan: DOI, volume, nomor terbitan, bulan/tahun terbit, dan WhatsApp.
- Regresi satu artikel untuk 15/15 jurnal berhasil. Uji tiga skenario RABBANI (normal, judul panjang, metadata resmi) dan satu skenario abstrak singkat dilakukan; render visual pakai LibreOffice.

**Masih beta:** tidak ada klaim 100% identik dengan tata letak Word 2010 atau terverifikasi memenuhi seluruh aturan jurnal. Rujukan audit 14 jurnal sebelumnya tetap dalam `AUDIT_14_JURNAL.md`.
