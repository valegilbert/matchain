# MatchaIn GC (Matcha Input Gak Culun)

Aplikasi otomatisasi untuk melakukan upload data Ground Check (GC) ke server BPS menggunakan kombinasi Selenium (untuk otentikasi) dan Requests (untuk pengiriman data cepat).

## Fitur Utama
- **Otomatisasi Login**: Login otomatis menggunakan akun SSO BPS.
- **Batch Processing**: Memproses semua file Excel (.xlsx/.xls) yang ada di folder `input/`.
- **Validasi Data**: Memastikan data Excel valid sebelum dikirim.
- **Smart Retry**: Otomatis memperbarui token jika sesi habis tanpa menghentikan proses.
- **Backup Otomatis**: Membuat backup data ke folder `backup/` sebelum proses dimulai.
- **Logging Lengkap**: Mencatat semua aktivitas ke file `app.log` dan konsol.
- **Resume Capability**: Melewati data yang sudah berstatus 'berhasil'.
- **Custom User Agent**: Bisa dikonfigurasi melalui file `.env`.

## Persyaratan Sistem
- Python 3.8 atau lebih baru.
- Google Chrome terinstall.
- Koneksi Internet.

## Instalasi & Penggunaan

### 1. Pertama Kali (Instalasi)
Double-click file **`install_and_run.bat`**.
Script ini akan otomatis:
- Membuat virtual environment (`.venv`).
- Menginstall semua library yang dibutuhkan.
- Menjalankan aplikasi.

### 2. Penggunaan Sehari-hari
Double-click file **`run.bat`**.
Script ini akan langsung menjalankan aplikasi tanpa melakukan instalasi ulang.

### 3. Konfigurasi Kredensial & User Agent
Buka file `.env` dan isi dengan username, password, dan User Agent yang diinginkan. Anda bisa mencontoh dari file `.env.example`.
```env
BPS_USERNAME=username_sso
BPS_PASSWORD=password_sso_anda
CUSTOM_USER_AGENT=Mozilla/5.0 (Linux; Android 11; Pixel 4 XL Build/RQ3A.210705.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.198 Mobile Safari/537.36
```

#### Daftar User Agent Android WebView (Referensi)
Anda bisa menggunakan salah satu dari daftar berikut:

**Android 11 – WebView berbasis Chrome 86**
`Mozilla/5.0 (Linux; Android 11; Pixel 4 XL Build/RQ3A.210705.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.198 Mobile Safari/537.36`

**Android 10 – WebView berbasis Chrome 80**
`Mozilla/5.0 (Linux; Android 10; SM‑G975F Build/QP1A.190711.020; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.162 Mobile Safari/537.36`

**Android 9 – WebView berbasis Chrome 69**
`Mozilla/5.0 (Linux; Android 9; Redmi Note 7 Build/PKQ1.180904.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/69.0.3497.100 Mobile Safari/537.36`

**Android 8 Oreo – WebView berbasis Chrome 66**
`Mozilla/5.0 (Linux; Android 8.1.0; Nexus 5X Build/OPM4.171019.021I; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/66.0.3359.158 Mobile Safari/537.36`

**Android 7 Nougat – WebView berbasis Chrome 51**
`Mozilla/5.0 (Linux; Android 7.1.1; Moto G (5) Build/NMF26F; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/51.0.2704.81 Mobile Safari/537.36`

**Android 6 Marshmallow – WebView berbasis Chrome 44**
`Mozilla/5.0 (Linux; Android 6.0; Nexus 6 Build/MRA58K; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/44.0.2403.119 Mobile Safari/537.36`

**Android 5 Lollipop – WebView berbasis Chrome 40**
`Mozilla/5.0 (Linux; Android 5.0; Nexus 5 Build/LRX21T; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/40.0.2214.89 Mobile Safari/537.36`

**Android 4.4 KitKat – WebView berbasis Chrome 30**
`Mozilla/5.0 (Linux; Android 4.4.4; Nexus 5 Build/KRT16S; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/30.0.1599.107 Mobile Safari/537.36`

### 4. Siapkan Data
1.  Buka folder **`input/`**.
2.  Letakkan semua file Excel yang ingin Anda proses di dalam folder ini.
3.  Pastikan format kolom di file Excel Anda sesuai dengan file **`contoh_excel.xlsx`** yang sudah ada di dalamnya.

Kolom wajib di setiap file Excel:
- `perusahaan_id` (Text/String)
- `latitude`
- `longitude`
- `hasilgc`
- `edit_nama`
- `edit_alamat`
- `nama_usaha`
- `alamat_usaha`

## Aturan Validasi Data (PENTING!)

Sebelum menjalankan aplikasi, pastikan data Excel Anda memenuhi aturan berikut:

1. **perusahaan_id**: Wajib terisi.
2. **hasilgc**: Harus salah satu dari angka berikut:
   - `1`: Aktif
   - `3`: Tutup Sementara
   - `4`: Tutup Permanen
   - `99`: Tidak Ditemukan
3. **edit_nama**:
   - `0`: Tidak ada perubahan nama.
   - `1`: Ada perubahan nama.
4. **edit_alamat**:
   - `0`: Tidak ada perubahan alamat.
   - `1`: Ada perubahan alamat.
5. **Konsistensi Data**:
   - Jika `nama_usaha` terisi, maka `edit_nama` **HARUS** `1`.
   - Jika `nama_usaha` kosong, maka `edit_nama` **HARUS** `0`.
   - Aturan yang sama berlaku untuk `alamat_usaha` dan `edit_alamat`.

## Troubleshooting
- **File Excel Terkunci**: Pastikan semua file Excel di folder `input/` **TERTUTUP** saat aplikasi berjalan.
- **Token Invalid**: Aplikasi akan mencoba login ulang otomatis. Jika gagal terus-menerus, cek koneksi internet atau kredensial di `.env`.
- **Log**: Cek file `app.log` untuk detail error.

---
*Daftar User Agent Android WebView bisa dilihat di revisi Git sebelumnya jika diperlukan.*
