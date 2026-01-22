# MatchaIn GC (Matcha Input Gak Culun)

Aplikasi otomatisasi (bot) berbasis Python untuk melakukan input dan konfirmasi data Ground Check (GC) ke sistem BPS (matchapro.web.bps.go.id). Aplikasi ini dirancang untuk efisiensi, ketangguhan (robustness), dan keamanan data.

## 🚀 Fitur Utama

*   **Otomatisasi Penuh**: Login SSO BPS otomatis (termasuk penanganan OTP via secret key).
*   **Input Cepat**: Menggunakan metode HTTP Request (bukan klik browser) untuk kecepatan maksimal.
*   **Validasi Cerdas**:
    *   Mengecek kelengkapan kolom wajib.
    *   **Validasi Lokasi (Geospasial)**: Memastikan koordinat (Latitude/Longitude) berada di dalam wilayah kabupaten yang sesuai (berdasarkan 2 digit kode kabupaten di kolom `kdkab`).
    *   Mencegah input data yang tidak konsisten.
*   **Ketangguhan (Robustness)**:
    *   **Auto-Retry**: Menangani gangguan koneksi internet dan timeout secara otomatis.
    *   **Rate Limit Handling**: Otomatis menunggu jika server sibuk (Error 429).
    *   **Auto-Refresh Token**: Memperbarui sesi secara otomatis jika token kedaluwarsa tanpa menghentikan proses.
*   **Keamanan Data**:
    *   **Backup Otomatis**: Membuat salinan file Excel sebelum diproses.
    *   **Real-time Saving**: Menyimpan status per 10 baris untuk mencegah kehilangan data jika crash.
    *   **Safe File Handling**: Mengecek apakah file sedang dibuka oleh user sebelum memproses.
*   **Manajemen File**:
    *   File yang selesai 100% otomatis dipindahkan ke folder `processed`.
*   **Pelaporan**: Menghasilkan laporan ringkasan (Summary Report) di akhir proses.

## 📋 Prasyarat

1.  **Python 3.8+** terinstal di komputer.
2.  **Google Chrome** browser terinstal.
3.  File `bounding_boxes.json` (untuk validasi lokasi).

## ⚙️ Instalasi

1.  Clone atau download repository ini.
2.  Buka terminal/command prompt di folder proyek.
3.  Instal library yang dibutuhkan:
    ```bash
    pip install -r requirements.txt
    ```

## 🛠️ Konfigurasi (.env)

Buat file bernama `.env` di folder root proyek dan isi dengan konfigurasi berikut:

```env
BPS_USERNAME=username_sso_anda
BPS_PASSWORD=password_sso_anda
BPS_OTP_SECRET=kode_rahasia_otp_anda  # Opsional, jika ingin OTP otomatis
USE_SESSION_CACHE=true                 # true/false (Simpan sesi login agar tidak login ulang terus)
HEADLESS=true                          # true/false (Jalankan browser di background)
```

> **Catatan**: `BPS_OTP_SECRET` adalah kode rahasia (biasanya string panjang) yang Anda gunakan di aplikasi Authenticator. Jika dikosongkan, aplikasi akan meminta input OTP manual di terminal.

### Referensi User Agent (Opsional)

Jika Anda ingin mengubah `CUSTOM_USER_AGENT` di `main.py` untuk mensimulasikan perangkat Android yang berbeda, berikut adalah beberapa referensi:

*   **Android 11 (Pixel 4 XL)**:
    `Mozilla/5.0 (Linux; Android 11; Pixel 4 XL Build/RQ3A.210705.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.198 Mobile Safari/537.36`
*   **Android 10 (Samsung S10+)**:
    `Mozilla/5.0 (Linux; Android 10; SM‑G975F Build/QP1A.190711.020; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.162 Mobile Safari/537.36`
*   **Android 9 (Redmi Note 7)**:
    `Mozilla/5.0 (Linux; Android 9; Redmi Note 7 Build/PKQ1.180904.001; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/69.0.3497.100 Mobile Safari/537.36`
*   **Android 8 Oreo (Nexus 5X)**:
    `Mozilla/5.0 (Linux; Android 8.1.0; Nexus 5X Build/OPM4.171019.021I; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/66.0.3359.158 Mobile Safari/537.36`
*   **Android 7 Nougat (Moto G5)**:
    `Mozilla/5.0 (Linux; Android 7.1.1; Moto G (5) Build/NMF26F; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/51.0.2704.81 Mobile Safari/537.36`
*   **Android 6 Marshmallow (Nexus 6)**:
    `Mozilla/5.0 (Linux; Android 6.0; Nexus 6 Build/MRA58K; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/44.0.2403.119 Mobile Safari/537.36`
*   **Android 5 Lollipop (Nexus 5)**:
    `Mozilla/5.0 (Linux; Android 5.0; Nexus 5 Build/LRX21T; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/40.0.2214.89 Mobile Safari/537.36`
*   **Android 4.4 KitKat (Nexus 5)**:
    `Mozilla/5.0 (Linux; Android 4.4.4; Nexus 5 Build/KRT16S; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/30.0.1599.107 Mobile Safari/537.36`

## 📂 Struktur Folder

*   `input/`: Letakkan file Excel (`.xlsx` / `.xls`) yang akan diproses di sini.
*   `backup/`: Aplikasi akan menyimpan backup file asli di sini sebelum memproses.
*   `processed/`: File yang sudah selesai 100% diproses akan dipindahkan ke sini.
*   `app.log`: File log detail untuk teknis/debugging.
*   `session.json`: File penyimpan sesi login (dibuat otomatis).

## 📝 Format Excel

File Excel di folder `input` **wajib** memiliki kolom-kolom berikut (nama kolom harus persis, huruf kecil):

| Nama Kolom | Keterangan |
| :--- | :--- |
| `perusahaan_id` | ID Perusahaan (Wajib) |
| `kdkab` | Kode Kabupaten (2 Digit, Wajib untuk validasi lokasi) |
| `latitude` | Koordinat Lintang |
| `longitude` | Koordinat Bujur |
| `hasilgc` | Kode Hasil GC (`1`, `3`, `4`, atau `99`) |
| `edit_nama` | Flag edit nama (`0` atau `1`) |
| `edit_alamat` | Flag edit alamat (`0` atau `1`) |
| `nama_usaha` | Nama Usaha (Wajib diisi jika `edit_nama` = 1) |
| `alamat_usaha` | Alamat Usaha (Wajib diisi jika `edit_alamat` = 1) |
| `status_upload` | (Opsional) Aplikasi akan mengisi kolom ini dengan status hasil upload. |

## ▶️ Cara Menjalankan

1.  Pastikan file Excel sudah ada di folder `input`.
2.  Jalankan aplikasi:
    ```bash
    python main.py
    ```
3.  Aplikasi akan menampilkan aturan validasi. Tekan **Enter** untuk memulai.
4.  Pantau progres di terminal.

## 📊 Output & Laporan

*   **Status di Excel**: Kolom `status_upload` di file Excel akan diupdate dengan:
    *   `berhasil`: Data sukses terkirim.
    *   Pesan Error (misal: `Invalid: kdkab kosong`, `gagal - HTTP 500`): Jika gagal.
*   **Laporan Akhir**: Setelah selesai, aplikasi akan membuat file `summary_report_YYYYMMDD_HHMMSS.txt` yang berisi statistik jumlah data sukses, gagal, dan dilewati.

## ⚠️ Catatan Penting

*   **Jangan membuka file Excel** yang sedang diproses di folder `input`. Aplikasi akan meminta Anda menutupnya jika terdeteksi.
*   Jika internet tidak stabil, aplikasi akan mencoba *reconnect* otomatis. Jika gagal total, ia akan meminta konfirmasi Anda.

---
**Disclaimer**: Aplikasi ini dibuat untuk membantu efisiensi kerja. Gunakan dengan bijak dan bertanggung jawab sesuai aturan BPS.
