# Split Brand: Live Streaming & Short Video

Script Python untuk memecah laporan TAP (tab `Custom report`) di Google Sheets
menjadi tab per brand, dengan **Livestream room ID / Video ID yang unik** dan
angka GMV, orders, serta items sold yang **dijumlahkan** kalau ID-nya muncul
lebih dari sekali.

Semua script punya dua mode:

| Mode | Perintah | Efek |
|---|---|---|
| Dry run (default) | `python <script>.py` | Hanya membaca dan menampilkan ringkasan. Tidak ada yang ditulis. |
| Write | `python <script>.py --write` | Membuat tab baru dan menulis hasilnya. |

Jalankan dry run dulu, cek angkanya, baru `--write`.

## Script

| File | Sumber | Hasil |
|---|---|---|
| `split_brand_live_combined.py` | Live Streaming (TAP LS PC) | 1 tab per brand, GMV dan NON GMV digabung (kolom `Jenis`) |
| `split_brand_gmv_live.py` | Live Streaming (TAP LS PC) | 1 tab per brand **dan** per jenis, misal `PEPSODENT - GMV` dan `PEPSODENT - NON GMV` |
| `split_brand_sv_combined.py` | Short Video (TAP VID PC) | 1 tab per brand, GMV dan NON GMV digabung (kolom `Jenis`) |

`split_brand_live_combined.py` memakai fungsi dari `split_brand_gmv_live.py`,
jadi kedua file itu harus ada di folder yang sama.

## Persiapan

### 1. Python dan library

Python 3.10 atau lebih baru.

```bash
pip install google-api-python-client google-auth
```

### 2. Service account Google

1. Buka Google Cloud Console, buat project (atau pakai yang sudah ada).
2. Aktifkan **Google Sheets API**.
3. Buat **service account**, lalu buat key berformat **JSON**.
4. Simpan file key itu sebagai `credentials.json` di folder yang sama dengan script.
5. Buka spreadsheet yang mau diproses, klik **Share**, lalu tambahkan alamat
   `client_email` dari `credentials.json` (bentuknya
   `nama@project-id.iam.gserviceaccount.com`):
   - **Viewer** cukup untuk dry run.
   - **Editor** diperlukan untuk `--write`.

> `credentials.json` berisi kunci rahasia. Jangan pernah di-commit ke GitHub.
> File ini sudah dimasukkan ke `.gitignore`.

### 3. Atur konfigurasi di bagian atas script

```python
CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_ID = '...'      # ID dari URL: docs.google.com/spreadsheets/d/<ID>/edit
SOURCE_SHEET = 'Custom report'
```

Khusus `split_brand_sv_combined.py` dan `split_brand_live_combined.py`:

| Pengaturan | Fungsi |
|---|---|
| `TAB_SUFFIX` (SV) | Akhiran nama tab hasil, default `' - UNIK'`, supaya tidak bentrok dengan tab brand yang sudah ada. |
| `MERGE_BRANDS` | Menggabungkan beberapa brand ke satu tab, misal `{'DOVE DEO': 'DOVE', 'DOVE SCL': 'DOVE'}`. Default kosong, jadi tiap brand punya tab sendiri. |
| `WRITE_CHUNK_ROWS` (SV) | Jumlah baris per request tulis. Kecilkan kalau muncul error payload terlalu besar. |

## Cara pakai

```bash
# 1. Cek dulu (tidak menulis apa pun)
python split_brand_sv_combined.py

# 2. Kalau angkanya sudah benar, tulis ke spreadsheet
python split_brand_sv_combined.py --write
```

Contoh keluaran dry run:

```
Mode: DRY RUN (tidak ada yang ditulis)
134996 baris di "Custom report" (semua diambil, tanpa filter bulan).
38438 baris unik (brand + Video ID).

7 tab akan dibuat:
  PEPSODENT - UNIK        25372 baris  (GMV:  239, NON GMV:  25133)   GMV total: 20,792,074
  ...
  cek total Affiliate video-attributed GMV: sumber 76,433,359 | hasil 76,433,359 | OK
```

Baris `cek total ...` membandingkan jumlah di sumber dengan jumlah di hasil. Kalau
tertulis `BEDA!`, jangan lanjut `--write`.

## Aturan pengolahan

- **Semua baris diambil.** Tidak ada filter bulan atau tanggal.
- **Brand huruf besar.** `Pepsodent` dan `PEPSODENT` dianggap sama.
- **Unik per brand + ID.** ID yang muncul di dua brand menghasilkan satu baris di tiap brand.
- **Yang dijumlahkan** untuk baris dengan ID yang sama:

  | Script | Kolom yang dijumlahkan |
  |---|---|
  | Live Streaming | `Creator LIVE-attributed GMV`, `Creator LIVE-attributed orders`, `Creator-attributed items sold` |
  | Short Video | `Affiliate video-attributed GMV`, `Creator video-attributed orders`, `Creator-attributed items sold` |

- **Kolom lain** diambil dari baris pertama ID itu, kecuali `Date` yang diambil
  tanggal paling awal. Kolom yang tidak dijumlahkan (views, likes, produk, dan
  seterusnya) tidak mewakili total, hanya nilai dari baris pertama.
- **GMV atau NON GMV.** Total GMV lebih dari 0 berarti `GMV`, sama dengan 0
  berarti `NON GMV`.
- **Urutan.** Dari GMV terbesar, jadi baris `GMV` ada di atas.
- **Kolom tambahan Live Streaming.** `Start Live`, `End Live` (tanggal dan jam
  dari `LIVE time info`), dan `Date Live` (tanggal awal saja, tanpa jam).
- **Baris bergeser (Short Video).** Baris yang kolom `Post time`-nya bukan format
  `YYYY-MM-DD HH:MM:SS` dilewati, dan nomor barisnya ditampilkan di dry run.
  Itu tanda kolom di baris tersebut bergeser, dan angkanya tidak bisa dipercaya.

## Keamanan

- Tab yang sudah ada **tidak pernah ditimpa**. Kalau nama tab bentrok, script
  berhenti dan menyebut tab mana yang bentrok.
- Tab sumber (`Custom report`) tidak diubah.
- Script SV mengecek batas 10 juta sel per spreadsheet sebelum menulis.

## Masalah umum

| Pesan | Penyebab dan solusi |
|---|---|
| `Requested entity was not found` atau error 403/404 | Spreadsheet belum di-share ke `client_email` di `credentials.json`, atau `SPREADSHEET_ID` salah. |
| `The caller does not have permission` saat `--write` | Akses service account baru Viewer. Ubah ke Editor. |
| `Tab sudah ada, tidak ditimpa` | Hapus atau rename tab itu, atau ganti `TAB_SUFFIX`. |
| `Kolom "..." gak ada di header` | Nama kolom di laporan berubah. Sesuaikan nama kolom di bagian atas script. |
| `GAGAL di tengah jalan` | Sebagian tab mungkin sudah terbuat. Hapus tab hasil di spreadsheet, lalu jalankan ulang. |
| `Total sel melebihi batas Google Sheets` | Hapus tab lama yang tidak terpakai, atau pecah hasilnya ke spreadsheet lain. |

## Sebelum upload ke GitHub

- Pastikan `credentials.json` tidak ikut ter-commit (`git status` jangan menampilkannya).
- Folder kerja ini berisi banyak file data (`*.xlsx`, `*.csv`) yang memuat nama
  creator. Tambahkan polanya ke `.gitignore`, atau upload hanya file script dan
  README ini.
- `SPREADSHEET_ID` di script menunjuk ke spreadsheet internal. Gunakan repo
  **private**, atau ganti ID-nya dengan placeholder sebelum dipublikasikan.
