# E-Pesantren: Sistem Manajemen Santri berbasis Google Sheets

Sistem manajemen santri terintegrasi yang memudahkan pengelolaan data dan keuangan santri pesantren menggunakan Google Sheets dan Apps Script.

![E-Pesantren Banner](https://epesantren.co.id/wp-content/uploads/2021/09/epesantren_hitm-1536x332.png)

## 📋 Fitur Utama

- **Sinkronisasi Data** - Tarik data santri dari API ke Google Sheets
- **Cek Saldo** - Periksa saldo santri berdasarkan NIS
- **Proses Transaksi** - Kelola pengeluaran saldo santri dengan validasi
- **Pencatatan Transaksi** - Catat semua transaksi dalam sheet terpisah
- **Laporan Otomatis** - Generate 4 jenis laporan berbeda:
  - Laporan Saldo Tertinggi
  - Laporan Saldo Terendah
  - Laporan Transaksi Harian
  - Laporan Transaksi Bulanan

## 🚀 Cara Penggunaan

### Persiapan Awal
1. Buat Google Spreadsheet baru
2. Buka Extensions > Apps Script
3. Copy-paste seluruh kode dari `e-pesantren.gs` ke editor
4. Simpan dan berikan nama proyek (contoh: "E-Pesantren")
5. Ganti `API_KEY` dan `BASE_URL` dengan kredensial API pesantren Anda

### Menjalankan Sistem
1. Refresh spreadsheet Anda
2. Menu "Sistem Santri" akan muncul di bagian atas
3. Pilih salah satu operasi dari menu tersebut:
   - Sinkronisasi Data Santri
   - Cek Saldo Santri
   - Proses Transaksi
   - Generate Laporan

## 📈 Struktur Sheet

Sistem ini akan membuat beberapa sheet di spreadsheet Anda:

- **Data Santri** - Berisi informasi semua santri (NIS, Nama, Saldo, dll.)
- **Transaksi** - Mencatat semua transaksi pengeluaran saldo
- **Laporan** - Berbagai sheet laporan akan dibuat sesuai kebutuhan

## 🔧 Konfigurasi API

Sistem ini menggunakan API dari epesantren.co.id untuk sinkronisasi data.
Edit baris berikut untuk menyesuaikan dengan API Anda:

```javascript
const API_KEY = "YOUR_API_KEY_HERE";
const BASE_URL = "YOUR_API_URL_HERE";
```

## 🤝 Kontribusi

Kontribusi selalu diterima! Jika Anda memiliki ide untuk perbaikan:

1. Fork repositori ini
2. Buat branch fitur baru (`git checkout -b fitur-baru`)
3. Commit perubahan Anda (`git commit -m 'Menambahkan fitur baru'`)
4. Push ke branch (`git push origin fitur-baru`)
5. Buat Pull Request

## 📄 Lisensi

Proyek ini dilisensikan di bawah [MIT License](LICENSE)

## 📞 Kontak

Jika Anda memiliki pertanyaan atau membutuhkan bantuan implementasi, silakan hubungi:
kontak@classy.id
