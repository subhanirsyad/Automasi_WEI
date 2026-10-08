# Cari WA Creator (Kalodata) — Untuk Bang Zidane

Isi daftar username TikTok creator, klik satu tombol, dan nomor WhatsApp (plus email dan kontak lain yang diisi creator di Kalodata) terisi otomatis.

- **Input spreadsheet:** tiap creator selesai, barisnya langsung muncul di tab **`Hasil WA`** di spreadsheet yang sama. Tidak ada file Excel. Tab input tidak diubah.
- **Input Excel:** hasil disimpan ke file Excel baru.
- **Login Kalodata sendiri**, lewat jendela browser yang dibuka aplikasi (termasuk OTP). Tidak ada cookie atau cURL yang perlu di-copy.

## 1. Yang dibutuhkan

| Kebutuhan | Keterangan |
|---|---|
| Windows + **Python 3.10 atau lebih baru** | Saat memasang Python, centang **"Add python.exe to PATH"** |
| **Google Chrome atau Microsoft Edge** | Edge sudah ada di Windows 10/11 |
| Akun Kalodata | Yang punya akses penuh dan bisa menerima OTP |
| File **`credentials.json`** | Kunci service account Google, dikirim secara privat (lihat bagian 2) |
| Internet | Untuk Kalodata dan Google Sheets |

## 2. Pemasangan (sekali saja, ikuti urut dari atas)

Bagian ini ditulis selangkah demi selangkah. Anda **tidak perlu mengerti** isi perintahnya, cukup salin dan tempel.

### Langkah 1. Download dari GitHub

1. Buka link ini di browser: **`https://github.com/NAMA-AKUN/NAMA-REPO`** *(pemilik repo mengganti tulisan ini dengan link yang benar)*.
2. Klik tombol hijau **Code**, lalu klik **Download ZIP**.
3. Buka folder **Downloads**. Klik kanan file ZIP yang baru terunduh, pilih **Extract All...**, lalu klik **Extract**.
4. Buka folder hasil extract sampai ketemu folder bernama **"Untuk bang zidane"**. Di dalamnya ada file `app.py`, `requirements.txt`, dan lain-lain. Folder ini kita sebut **folder aplikasi**.
5. *(Opsional)* Pindahkan folder aplikasi ke **Documents** supaya mudah dicari.

### Langkah 2. Taruh file `credentials.json`

File ini **tidak ada di GitHub**. Pemilik aplikasi mengirimkannya terpisah (lewat chat pribadi). Taruh file itu di dalam folder aplikasi, **sejajar dengan `app.py`**.

> `credentials.json` memberi akses ke semua spreadsheet yang sudah di-share ke service account-nya. Jangan dikirim ke orang lain dan jangan di-upload ke tempat umum.

### Langkah 3. Buka terminal di folder aplikasi

**Terminal** itu jendela hitam tempat mengetik perintah. Kita hanya memakainya untuk menempel perintah yang sudah disediakan di bawah.

Cara membukanya (paling mudah):

1. Buka folder aplikasi di File Explorer (folder yang berisi `app.py`).
2. Klik **satu kali** di kotak alamat di bagian atas jendela (kotak yang menampilkan lokasi folder, misalnya `Documents > Untuk bang zidane`). Tulisannya akan terblok.
3. Ketik `cmd`, lalu tekan **Enter**.
4. Jendela hitam terbuka. Baris terakhirnya berakhiran `Untuk bang zidane>` dengan kursor berkedip. Itu tandanya terminal sudah berada di folder yang benar.

*(Cara lain di Windows 11: klik kanan di ruang kosong dalam folder aplikasi, lalu pilih **Open in Terminal**.)*

**Cara menempel perintah di terminal:** salin perintah dari kotak abu-abu di panduan ini (blok-blok di bawah), klik di jendela hitam, tekan **Ctrl + V** (atau klik kanan), lalu tekan **Enter** untuk menjalankannya.

### Langkah 4. Cek Python

Python adalah program yang menjalankan aplikasi ini. Tempel perintah ini di terminal, lalu Enter:

```
python --version
```

- Kalau muncul tulisan seperti **`Python 3.11.5`** (angka kedua minimal 10), lanjut ke Langkah 5.
- Kalau muncul **`'python' is not recognized`** atau Microsoft Store terbuka, Python belum terpasang:
  1. Buka https://www.python.org/downloads/ lalu klik tombol kuning **Download Python**.
  2. Buka file yang terunduh. Di jendela pertama, **centang kotak "Add python.exe to PATH"** (kotak kecil di bagian bawah, ini penting), lalu klik **Install Now**.
  3. Setelah selesai, **tutup jendela hitam** lalu ulangi Langkah 3 untuk membukanya lagi. Tempel `python --version` sekali lagi untuk memastikan.
- Kalau `python` tidak dikenali tapi `py --version` berhasil, pakai **`py`** sebagai pengganti **`python`** di semua perintah di bawah.

### Langkah 5. Install requirements (pasang library)

Requirements adalah program-program pendukung yang dibutuhkan aplikasi. Tempel perintah ini di terminal, lalu Enter:

```
python -m pip install -r requirements.txt
```

Tunggu 2 sampai 5 menit. Tulisan akan berjalan banyak dan itu normal. Sudah selesai kalau jendela kembali menampilkan baris `...Untuk bang zidane>` dan kursor berkedip lagi, biasanya dengan tulisan **`Successfully installed ...`** atau **`Requirement already satisfied`** di atasnya. Tulisan kuning **WARNING** atau **notice** boleh diabaikan. Kalau muncul tulisan merah **ERROR**, lihat bagian "Masalah umum" di bawah.

Langkah ini cukup **sekali** (atau lagi kalau pindah ke PC lain).

### Langkah 6. Jalankan aplikasi

Di terminal yang sama, tempel perintah ini, lalu Enter:

```
python app.py
```

Dalam beberapa detik jendela aplikasi **Cari WA Creator** terbuka. **Jangan tutup jendela hitamnya** selama aplikasi dipakai, karena menutupnya ikut menutup aplikasi (boleh di-minimize).

### Membuka aplikasi lagi di hari lain

Cukup **Langkah 3** (buka terminal di folder aplikasi), lalu tempel:

```
python app.py
```

Langkah 1, 2, 4, dan 5 tidak perlu diulang.

*(Jalan pintas tanpa terminal: klik dua kali **`Install.bat`** untuk Langkah 5, dan klik dua kali **`Jalankan.bat`** untuk Langkah 6. Hasilnya sama persis.)*

## 3. Cara pakai: input Google Spreadsheet

1. **Share spreadsheet** ke email service account ini sebagai **Editor**:

   ```
   video-bank-importer@automasi-507112.iam.gserviceaccount.com
   ```

   Email yang sama juga tampil otomatis di aplikasi, di bawah kolom link, dengan tombol **Salin email** (klik tombol itu supaya tidak salah ketik). Di Google Sheets: klik **Share**, **tempel email**, ubah ke **Editor**, lalu **Send**. Kalau muncul pertanyaan "kirim notifikasi?", boleh dimatikan. Cukup sekali per spreadsheet.
2. Di aplikasi, pilih **Google Spreadsheet**, lalu tempel **link** spreadsheet. Isi **nama tab** kalau daftar creator-nya bukan di tab pertama. Kalau link memuat `gid=...`, tab itu yang dipakai.
3. Klik **Cari Nomor WA**.
4. **Pertama kali:** jendela browser terbuka di Kalodata. Kotak **Log-in** terbuka otomatis (kalau tidak, klik tombol biru **Log-in / Sign-up** di pojok kanan atas jendela itu; kalau jendelanya sempit dan tombolnya terpotong, besarkan jendela dengan tombol kotak di pojok kanan atas). Isi akun Kalodata lalu selesaikan OTP. Halaman Kalodata bisa dibuka tanpa login tapi angkanya disembunyikan (`****`), jadi pastikan angkanya terbaca penuh. Aplikasi mendeteksi login sendiri dan langsung lanjut. Jangan tutup jendela itu selama proses jalan. Kalau muncul kotak verifikasi "Verify you are human", klik saja.
5. Lihat hasilnya di tab **`Hasil WA`**. Barisnya bertambah satu per satu selama proses.

Login tersimpan di folder `profil_chrome`, jadi lain kali biasanya tidak perlu login lagi.

## 4. Cara pakai: input Excel

1. Pilih **File Excel**, klik **Browse...**, pilih file `.xlsx`.
2. File hasil otomatis bernama `<nama file>_WA.xlsx` di folder yang sama. Bisa diganti di kotak **2. File hasil**.
3. Klik **Cari Nomor WA**. Login seperti di atas kalau diminta.

Excel disimpan tiap 25 creator dan di akhir, jadi kalau berhenti di tengah, hasilnya tetap ada.

## 5. Format daftar creator

Aplikasi mencari kolom bernama salah satu dari: `Creator name`, `Creator`, `Nama creator`, `Username`, `Handle`, `Creator ID`. Header boleh ada di baris mana pun dalam **30 baris pertama**. Kalau tidak ada header yang cocok, dipakai **kolom A**.

| Creator name |
|---|
| shellasaukia |
| @jastipbyvi |
| 6701563873919894530 |

- Isi boleh **username TikTok** (dengan atau tanpa `@`) atau **creator ID** numerik (15 digit atau lebih).
- Baris kosong, `Summary`, `-`, dan nama yang dobel dilewati (tiap nama hanya dicari sekali).

## 6. Hasil

| Kolom | Isi |
|---|---|
| Input | Teks asli dari daftar Anda |
| Handle, Nickname, Creator ID | Akun yang ditemukan |
| WhatsApp | Nomor seperti yang diisi creator (contoh `082167337702`) |
| Link WA | `https://wa.me/6282167337702`, klik untuk langsung chat |
| Email, Zalo, Line, Facebook, Instagram | Kontak lain, kosong kalau tidak diisi |
| MCN, Followers | Info profil |
| Status | Lihat di bawah |

| Status | Artinya |
|---|---|
| `OK` | Nomor WA ditemukan |
| `Tanpa kontak WA` | Creator ditemukan tapi tidak mengisi WhatsApp. Ini keterbatasan data, bukan error. |
| `Tidak ketemu` | Tidak ada creator dengan handle atau nickname **persis** sama |
| `Tidak ketemu (mirip: xxx)` | Hanya ada yang mirip. Periksa manual dan isi ulang dengan handle yang benar. |

Aplikasi **tidak menebak** creator yang hanya mirip. Nomor WA milik orang yang salah lebih berbahaya daripada kolom kosong.

## 7. Berhenti, lanjut, dan ulang

- **Stop** menghentikan proses setelah creator yang sedang diproses selesai. Yang sudah tertulis tetap ada.
- **Lanjut:** klik **Cari Nomor WA** lagi. Untuk spreadsheet, creator yang sudah ada di `Hasil WA` dilewati otomatis, jadi proses lanjut dari yang belum. (Untuk Excel, proses mulai dari awal. Buat file input baru berisi sisanya kalau mau lanjut.)
- **Mengulang creator tertentu:** hapus barisnya dari tab `Hasil WA`, lalu jalankan lagi.
- **Mengulang semuanya:** hapus semua baris di bawah header tab `Hasil WA` (atau hapus tab-nya).

## 8. Kecepatan

Tiap creator butuh dua permintaan dengan jeda 1,2 detik, jadi sekitar **2,4 detik per creator** (100 creator ≈ 4 menit, 1.000 creator ≈ 40 menit). Jeda sengaja tidak dipercepat supaya akun tidak kena pembatasan. Jangan mematikan atau men-sleep-kan PC selama proses.

## 9. Masalah umum

| Gejala | Penyebab dan solusi |
|---|---|
| `'python' is not recognized` / jendela hitam langsung menutup | Python belum terpasang atau tidak masuk PATH. Pasang ulang dan centang "Add python.exe to PATH" (Langkah 4), lalu buka terminal baru. |
| `ERROR` merah saat Langkah 5 | Internet mati atau diblokir. Coba lagi. Kalau tetap gagal, foto atau copy tulisan merahnya dan kirim ke pemilik aplikasi. |
| `No such file or directory: 'requirements.txt'` | Terminal tidak berada di folder aplikasi. Tutup jendela hitam dan ulangi Langkah 3. |
| `can't open file ... app.py` | Sama seperti di atas: terminal belum berada di folder aplikasi (Langkah 3). |
| `Chrome / Edge tidak bisa dibuka` | Pasang Google Chrome, lalu jalankan lagi. |
| Aplikasi terus menunggu login padahal terasa sudah login | **Cek angka di tabel jendela browser itu.** Kalau tampil **`****`** atau **`$****`**, Anda **belum login**: halaman Kalodata memang bisa dibuka tanpa login, tapi angkanya disembunyikan. Klik tombol login di pojok kanan atas jendela itu, lalu selesaikan OTP sampai angka terbaca penuh. Pastikan juga login dilakukan di **jendela yang dibuka aplikasi** (jendela terpisah dari Chrome harian Anda; cek di taskbar). Setelah login, aplikasi lanjut sendiri. Tombol **Sudah login ▶** hanya untuk kasus deteksi otomatis gagal padahal angka sudah terbaca penuh. |
| Ada kotak "Verify you are human" | Klik di jendela browser itu. Aplikasi lanjut sendiri. |
| `Belum login, atau ada verifikasi Cloudflare` di tengah proses | Sesi habis. Login ulang di jendela browser, lalu klik **Cari Nomor WA** lagi. Yang sudah selesai tidak diulang. |
| `Jendela browser tidak bisa dipakai` | Jendela browser tertutup di tengah proses. Jalankan lagi dan jangan tutup jendelanya. |
| `Spreadsheet tidak bisa diakses (HTTP 403/404)` | Belum di-share sebagai Editor ke email service account, atau link salah. |
| `Tab "..." tidak ada` | Nama tab salah. Pesan error menampilkan daftar tab yang ada. |
| `Tab input tidak boleh bernama "Hasil WA"` | Tab daftar creator memakai nama yang sama dengan tab hasil. Ganti nama tab input. |
| `File ini .xlsx yang di-upload ke Drive` | Itu bukan Google Sheets asli. Di Drive: buka file, **File → Save as Google Sheets**, lalu pakai link yang baru. |
| Banyak `Tidak ketemu` | Isi kolom bukan username TikTok (misalnya nama asli atau nama toko). |
| Nomor kolom WhatsApp kehilangan angka 0 | Tidak terjadi di `Hasil WA` (ditulis sebagai teks). Kalau Anda menyalin ke tempat lain, set kolomnya ke format teks. |

## 10. Data pribadi dan keamanan

- Hasilnya berisi nomor telepon pribadi creator. Pakai hanya untuk keperluan kerja, jangan dibagikan luas.
- Folder **`profil_chrome`** berisi sesi login Kalodata. Anggap seperti password: jangan dikirim atau di-upload.
- `credentials.json` dan `pengaturan.json` (menyimpan link terakhir) juga jangan dibagikan. Ketiganya sudah ada di `.gitignore` kalau folder ini dimasukkan ke git.

## 11. Isi folder

| File | Fungsi |
|---|---|
| `Install.bat` | Klik dua kali **sekali saja** untuk memasang semua library (requirements) |
| `Jalankan.bat` | Klik dua kali untuk membuka aplikasi (memasang library kalau belum ada) |
| `app.py` | Tampilan aplikasi |
| `inti.py` | Membaca daftar, memanggil pencarian, menulis ke Sheets/Excel. `python inti.py` menjalankan self-check. |
| `kalodata_browser.py` | Membuka browser, mendeteksi login, mengambil data dari Kalodata |
| `requirements.txt` | Daftar library |
| `credentials.json` | (dari pemilik) kunci service account Google |
| `profil_chrome/`, `pengaturan.json` | Dibuat otomatis saat dipakai |
