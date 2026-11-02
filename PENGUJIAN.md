# HASIL PENGUJIAN — ZAIN.NET GitHub Pages Only v3.1

Tanggal: 2 November 2026. Sampel: artikel Word mahasiswa sepanjang 19 halaman yang memiliki 8 footnote Word asli, abstrak Indonesia/Inggris dan daftar pustaka.

## Pengujian otomatis

- **14 dari 14** template berhasil menghasilkan DOCX tanpa error pada pengujian JavaScript menggunakan pustaka DOM Word dan JSZip.
- Semua hasil dapat dibaca parser DOCX dan dikonversi ke PDF memakai LibreOffice headless (pengujian kompatibilitas, bukan jaminan hasil sama di semua versi Microsoft Word).
- 8 dari 8 referensi footnote dan definisi catatan kaki sampel masih ada pada output semua 14 jurnal.
- Judul, penulis, abstrak, kata kunci dan isi tersalin menurut konfigurasi 14 jurnal; teks daftar pustaka sampai entri terakhir tetap ada.
- Untuk **KARSA**, hasil visual halaman pertama berisi logo hijau/ISSN, judul/penulis, kotak Received–Accepted–Published dengan nilai kosong, nomor halaman di tengah bawah. Halaman kedua menampilkan nama penulis di header.
- Untuk **PERDIKAN** dan **SOCIALE**, logo resmi dan desain tabel informasi/abstrak tetap terlihat.
- **NUANSA**: kotak judul diuji dengan judul 16 kata; adaptasi font dilakukan agar judul tidak terpotong.
- **Gambar + tabel tambahan** di artikel uji terbawa ke keluaran KARSA. Gambar yang melampaui lebar halaman dikurangi ukurannya, dan tabel tetap dapat diedit.

## Keterbatasan yang masih harus diuji manual

1. Tidak ada pengujian akhir visual semua halaman untuk setiap kombinasi template + semua jenis artikel (khususnya objek OLE, SmartArt, track changes, dan sumber bermargin aneh).
2. **MABNY**: tata letak boks abstrak dan framing multikolom kompleks bisa berubah ketika abstrak panjang; periksa struktur abstrak dan halaman pertama secara manual.
3. Sebagian placeholder *volume, tahun edisi dan DOI* adalah contoh dari template; belum dapat menentukan volume/DOI resmi artikel yang belum diterbitkan.
4. Template penerbit mungkin mewajibkan perubahan isi/artikel (jumlah kata, struktur metode, bahasa, istilah). Mesin tidak mengarang data ilmiah atau menerjemahkan abstrak secara otomatis.
5. Kesesuaian di Word 2010/2016/365 dapat berbeda dari hasil rendering LibreOffice. Pemeriksaan langsung di versi Word pengguna tetap penting.
6. Mesin hanya menerima **artikel sumber DOCX**; tiga template DOC lawas telah disediakan juga versi DOCX hasil konversi.
7. Pengujian interaksi browser antarmuka dibatasi lingkungan eksekusi saat pembuatan; fungsi utama JavaScript diuji lewat runtime Node dengan antarmuka DOM yang kompatibel, bukan klaim uji semua browser.

**Status: BETA YANG DAPAT DIUJI; tidak boleh diberi label 100% sesuai sebelum pemeriksaan Word dan verifikasi tata letak setiap jurnal selesai.**
