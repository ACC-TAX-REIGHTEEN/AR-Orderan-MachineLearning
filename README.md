# 🤖 AR Orderan — Machine Learning

> **Sinkronisasi AR ke Google Sheets order tracker dengan resolusi nama pelanggan berbasis Machine Learning — multi-produk, real-time, dan tanpa ketergantungan internet untuk data AR**

Sistem dua komponen: **pipeline data** (`Dapur/`) yang membaca AR dari `ARVIEWER.xlsm` dan menyuntikkan nilai + cell note ke Google Sheets secara berkala, didukung oleh **modul ML** (`ML/`) yang melatih model resolusi nama pelanggan menggunakan pendekatan hybrid empat lapis (local match → memori historis → RapidFuzz → TF-IDF + k-NN) — memungkinkan pencocokan akurat meski nama yang ditulis di order sheet tidak persis sama dengan nama di sistem AR.

---

## 📋 Daftar Isi

- [Gambaran Umum & Konteks](#-gambaran-umum--konteks)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
  - [Tahap 0 — Persiapan Data Training](#tahap-0--persiapan-data-training-lookupdata)
  - [Tahap 1 — Pelatihan Model ML](#tahap-1--pelatihan-model-ml-trainingmodel)
  - [Tahap 2 — Jalankan Pipeline Utama](#tahap-2--jalankan-pipeline-utama)
- [Alur Kerja Lengkap](#-alur-kerja-lengkap)
- [Detail Tiap Komponen](#-detail-tiap-komponen)
  - [`LookupData.py` — Persiapan data training](#lookupdata-py--persiapan-data-training)
  - [`TrainingModel.py` — Pelatihan model hybrid](#trainingmodel-py--pelatihan-model-hybrid)
  - [`1_CopyData.py` — Ekstraksi data dari ARVIEWER](#1_copydata-py--ekstraksi-data-dari-arviewer)
  - [`2_InjectDataToSS.py` — Loop sinkronisasi ke Google Sheets](#2_injectdatatos-py--loop-sinkronisasi-ke-google-sheets)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Format File Data Training](#-format-file-data-training)
- [Algoritma Resolusi Nama Pelanggan](#-algoritma-resolusi-nama-pelanggan)
- [Format Output: Cell Value & Cell Note](#-format-output-cell-value--cell-note)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum & Konteks

Proyek ini memecahkan satu masalah inti di sistem order berbasis Google Sheets: **nama pelanggan yang ditulis di kolom order tidak selalu sama persis dengan nama yang tercatat di sistem AR Accurate**. Misalnya, admin menulis `"Toko Maju Magelang IRC"` sementara sistem AR mencatat `"TOKO MAJU"`.

Alih-alih mengandalkan mapping manual atau download data AVG dari Sheets (seperti pada `Automasi-AR-Orderan`), proyek ini menggunakan pendekatan Machine Learning yang dilatih dari **data historis order nyata** untuk belajar secara otomatis bagaimana menerjemahkan variasi nama pelanggan ke nama kanonik di sistem AR.

Sistem ini juga mendukung **dua produk sekaligus** (IRC dan ZN) dari satu sumber ARVIEWER.xlsm, menginjeksi ke dua Google Sheets yang berbeda dalam satu sesi loop.

| Aspek | Automasi AR Orderan | AR Orderan ML (proyek ini) |
|---|---|---|
| Sumber data AR | `Piutang.xls` dari Accurate | `ARVIEWER.xlsm` (sheet `Source`) |
| Resolusi nama | AVG Spreadsheet + fuzzy | Model ML (TF-IDF + k-NN + RapidFuzz) |
| Data referensi | Download dari Google Sheets | File lokal `Hasil_Latihan.xlsx` |
| Jumlah produk | Satu per run | Multi-produk (IRC, ZN, ...) dalam satu run |
| Pelanggan cash | `FallbackCash` dari Sheets | `FBackCust.xlsx` dari ARVIEWER |
| Filter sales | `ar_key_filter` via Kontak | Filter "SR" hardcoded + `ar_key_filter` via Kontak |

---

## ✨ Fitur Utama

- **Model hybrid 4-lapis** — Resolusi nama menggunakan empat strategi secara berjenjang: local match dari data training, kamus memori historis (exact), RapidFuzz token matching, dan TF-IDF + k-NN cosine similarity — dengan threshold yang dapat dikonfigurasi.
- **Multi-produk dalam satu eksekusi** — Satu loop `2_InjectDataToSS.py` memproses semua produk yang dikonfigurasi (`irc`, `zn`, dst.) secara berurutan tanpa perlu menjalankan ulang.
- **RAM preloading** — Seluruh data (ML dict, FBackCust, ARClean) dimuat ke memori sebelum loop baris Google Sheets dimulai, sehingga pencocokan per baris berjalan dalam mikrodetik dengan bantuan cache in-memory.
- **Batch update 300 baris** — Permintaan update ke Google Sheets dikelompokkan maks 300 per panggilan API untuk efisiensi dan menghindari rate limit.
- **Pemisah bulan di cell note** — Baris faktur dalam ringkasan dikelompokkan berdasarkan bulan/tahun secara otomatis — dipisahkan baris kosong antar kelompok bulan yang berbeda.
- **Filter SR + filter produk** — Hanya baris AR dengan `Nama Penjual` mengandung "SR" dan `Nama Kontak` mengandung kode produk (`ar_key_filter`) yang diproses.
- **Cache resolver & AR lookup** — Nama pelanggan yang sudah pernah diselesaikan atau AR rows yang sudah pernah dicari di-cache agar tidak diproses ulang dalam iterasi yang sama.
- **Auto-detect header ARVIEWER** — Membaca sheet `Source` ARVIEWER.xlsm dengan deteksi header otomatis berdasarkan keyword kolom, bukan nomor baris yang tetap.
- **Loop real-time dengan interval** — `2_InjectDataToSS.py` berjalan terus dalam interval yang dikonfigurasi (bawaan 5 menit), menangkap baris order baru secara otomatis.
- **Toleransi error** — Error pada satu produk tidak menghentikan pemrosesan produk lainnya.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl scikit-learn rapidfuzz gspread google-auth
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Semua | Baca Excel, transformasi, filter |
| `openpyxl` | `1_CopyData.py` | Baca `.xlsm`, baca/tulis `.xlsx` |
| `scikit-learn` | `TrainingModel.py` | `TfidfVectorizer`, `NearestNeighbors` |
| `rapidfuzz` | `TrainingModel.py`, `2_InjectDataToSS.py` | Fuzzy string matching (token_set_ratio) |
| `gspread` | `2_InjectDataToSS.py` | Klien Google Sheets API |
| `google-auth` | `2_InjectDataToSS.py` | Autentikasi via Service Account |
| `re`, `os`, `shutil`, `configparser`, `datetime`, `time`, `collections`, `warnings`, `subprocess`, `sys` | Semua | Standard library |

### Dependensi eksternal
- **`ARVIEWER.xlsm`** — Workbook dashboard AR (dari proyek ARVIEWER). Harus sudah ada dan berisi data AR terkini di sheet `Source`. Path dikonfigurasi di `[DIR] arvi`.
- **`Hasil_Latihan.xlsx`** — Hasil pelatihan model ML. Harus ada di folder `ML/` dan disalin ke `Dapur/` oleh `1_CopyData.py`.
- **`TheTrainningData.xlsx`** — Data training historis. Harus ada di folder `ML/`.

---

## 📁 Struktur Folder & File

```
📦 AR-Orderan-MachineLearning/
│
├── 📄 Jalankan Automasi.py          ← Orkestrator pipeline utama. Jalankan ini
│
├── 📁 Dapur/                        ← Pipeline data (jangan diubah strukturnya)
│   ├── 📄 __init__.py
│   ├── 📄 1_CopyData.py             ← Ekstrak data dari ARVIEWER.xlsm + salin ML
│   ├── 📄 2_InjectDataToSS.py       ← Loop sinkronisasi AR → Google Sheets
│   ├── 📄 config.conf               ← Konfigurasi path, URL, dan semua flag
│   └── 📄 credentials.json          ← Kredensial Google Service Account (rahasia!)
│
└── 📁 ML/                           ← Modul Machine Learning (jalankan offline)
    ├── 📄 LookupData.py             ← Persiapan data training dari file historis
    ├── 📄 TrainingModel.py          ← Latih model hybrid → Hasil_Latihan.xlsx
    ├── 📄 TheTrainningData.xlsx     ← [INPUT] Data training historis (wajib ada)
    └── 📄 Hasil_Latihan.xlsx        ← [OUTPUT] Hasil pelatihan (wajib ada setelah training)
```

> **Alur ketergantungan file:**
> `TheTrainningData.xlsx` → `LookupData.py` (isi kolom E–H) → `TrainingModel.py` → `Hasil_Latihan.xlsx` → `1_CopyData.py` (salin ke Dapur/) → `2_InjectDataToSS.py`

---

## 🚀 Cara Penggunaan

Proyek ini memiliki **dua tahap terpisah**: pelatihan model (offline, dilakukan sekali atau periodik) dan eksekusi pipeline (online, berjalan terus-menerus).

### Tahap 0 — Persiapan Data Training (`LookupData.py`)

> Jalankan **hanya** jika `TheTrainningData.xlsx` belum memiliki kolom E–H terisi, atau jika ada data historis baru yang perlu ditambahkan.

1. Buka `ML/LookupData.py` dan ubah variabel `FOLDER_BASE` ke path folder yang berisi file-file order historis:
   ```python
   FOLDER_BASE = r"E:\ADM IRC AND ZN\2026"
   ```
2. Pastikan `ML/TheTrainningData.xlsx` sudah ada dengan kolom A–D terisi.
3. Jalankan:
   ```bash
   cd ML
   python LookupData.py
   ```
   Skrip akan mengisi kolom `Ekstrak_Sales (E)`, `Ekstrak_Customer_Detail (F)`, `Ekstrak_SR (G)`, dan `Ekstrak_Tgl_Nota (H)` di `TheTrainningData.xlsx`.

### Tahap 1 — Pelatihan Model ML (`TrainingModel.py`)

> Jalankan setiap kali data training diperbarui atau ada pelanggan baru yang perlu dikenali.

```bash
cd ML
python TrainingModel.py
```

Output: `ML/Hasil_Latihan.xlsx` berisi kolom tambahan `Hasil_Nama_Rekomendasi`, `Skor_Kemiripan_%`, `Status_Pencocokan`, `Sumber_Pencocokan`.

### Tahap 2 — Jalankan Pipeline Utama

1. Sesuaikan `Dapur/config.conf` — isi path ARVIEWER, URL Google Sheets, dan flag output.
2. Pastikan `ML/Hasil_Latihan.xlsx` sudah ada (hasil Tahap 1).
3. Jalankan:
   ```bash
   python "Jalankan Automasi.py"
   ```

Pipeline akan memvalidasi folder dan file, menjalankan `1_CopyData.py` (sekali), lalu masuk ke `2_InjectDataToSS.py` yang berjalan dalam loop tanpa henti. Tekan **`Ctrl+C`** untuk menghentikan.

---

## 🔄 Alur Kerja Lengkap

```
╔══════════════════════════════════════════════════════════════╗
║  OFFLINE — Persiapan & Pelatihan Model (ML/)                ║
╚══════════════════════════════════════════════════════════════╝

[File order historis]     [TheTrainningData.xlsx (kolom A–D)]
     (FOLDER_BASE)                      │
          │                             │
          └──────── LookupData.py ──────┘
                          │
                          │  Isi kolom E: Sales
                          │  Isi kolom F: Customer Detail  ← kunci referensi utama
                          │  Isi kolom G: SR Code
                          │  Isi kolom H: Tgl Nota
                          ↓
               TheTrainningData.xlsx (lengkap)
                          │
                          │
                   TrainingModel.py
                    (Hybrid 4-lapis)
                          │
                          ↓
               Hasil_Latihan.xlsx
                (kolom Hasil_Nama_Rekomendasi, Skor, Status, Sumber)


╔══════════════════════════════════════════════════════════════╗
║  ONLINE — Pipeline Sinkronisasi (Dapur/)                    ║
╚══════════════════════════════════════════════════════════════╝

[ARVIEWER.xlsm]                  [Hasil_Latihan.xlsx]
  sheet: Source    ──────────┐         │
  sheet: Nama Pelanggan SS ──┤   1_CopyData.py   ←───────────────────┘
                              │         │
                              │   ARClean_temp.xlsx (dari sheet Source)
                              │   FBackCust.xlsx (dari sheet Nama Pelanggan SS)
                              │   Hasil_Latihan.xlsx (disalin ke Dapur/)
                              │
                              ↓
                    2_InjectDataToSS.py
                     LOOP ∞ (tiap N menit)
                          │
                     Per produk (irc, zn):
                          │
                     [1] Preload ke RAM:
                          │   ml_dict (dari Hasil_Latihan)
                          │   fb_dict (dari FBackCust)
                          │   ar_memory (dari ARClean, filter SR + filter produk)
                          │
                     [2] Buka Google Sheets produk
                          │
                     [3] Per baris kosong di kolom target:
                          │   resolve_target_name_fast()
                          │     → Cek ml_dict (exact)
                          │     → Cek fb_dict (exact)
                          │     → Fuzzy RapidFuzz ke fb_list (threshold 60%)
                          │     → Fallback: gunakan raw input
                          │
                          │   get_ar_rows_fast()
                          │     → Cek ar_memory (exact)
                          │     → Fuzzy RapidFuzz ke ar_memory keys (threshold 80%)
                          │
                          │   Bangun cell note + nilai total piutang
                          │   Tambahkan ke batch requests
                          │
                     [4] Kirim batch (maks 300 per panggilan)
                          │
                     Tunggu interval → ulangi
```

---

## 🔍 Detail Tiap Komponen

### `LookupData.py` — Persiapan data training

Skrip utilitas **offline** yang membangun kamus `{nominal: (sales, customer, sr_code, tgl_nota)}` dari semua file Excel historis di `FOLDER_BASE`.

**Logika parsing nama file:**
File historis diharapkan memiliki nama dengan format: `Sales, CustomerDetail, SRCode, TglNota.xlsx`. Skrip memisahkan bagian-bagian ini dengan koma.

```
"Budi, Toko Makmur Magelang, IRC-001, 15012025.xlsx"
  ↓
  Sales    = "Budi"
  Customer = "Toko Makmur Magelang"    ← ini yang menjadi referensi (kolom F)
  SR Code  = "IRC-001"
  Tgl Nota = "15012025"
```

**Cara kerja lookup:**
Membaca nilai di kolom D setiap file historis sebagai kunci nominal. Jika nominal yang sama ditemukan di `TheTrainningData.xlsx` (kolom ke-4), baris tersebut diisi dengan data `(Sales, Customer, SR, Tgl)` dari file referensi yang cocok.

> ⚠️ **Wajib ubah sebelum dijalankan:** Ganti nilai `FOLDER_BASE` di baris 8 ke path folder historis Anda.

---

### `TrainingModel.py` — Pelatihan model hybrid

Melatih dan mengevaluasi model resolusi nama pelanggan menggunakan data di `TheTrainningData.xlsx`. Menggunakan **empat lapis pencocokan berjenjang** (lihat bagian [Algoritma](#-algoritma-resolusi-nama-pelanggan)).

**Input kolom yang dibutuhkan di `TheTrainningData.xlsx`:**

| Kolom | Nama | Keterangan |
|---|---|---|
| `Nama Customer dan Kota` | Input query | Nama seperti yang ditulis di order sheet |
| `Ekstrak_Customer_Detail (F)` | Referensi utama | Hasil LookupData — nama resmi dari file historis |
| `Ekstrak_Customer_Detail (F2)` | Referensi manual | Koreksi manual oleh admin (prioritas tertinggi) |
| `Ekstrak_Customer_Detail (F3)` | Referensi alternatif | Sumber ketiga jika F dan F2 kosong |

**Urutan prioritas sumber:** F2 → F → F3 (F2 dianggap paling akurat karena koreksi manual).

**Output kolom yang ditambahkan ke `TheTrainningData.xlsx`:**

| Kolom | Keterangan |
|---|---|
| `Hasil_Nama_Rekomendasi` | Nama kanonik yang direkomendasikan model |
| `Skor_Kemiripan_%` | Skor kemiripan (0–100) |
| `Status_Pencocokan` | Status detail: `SUCCESS (Match Lokal)`, `SUCCESS (RapidFuzz Match)`, dll. |
| `Sumber_Pencocokan` | Metode yang berhasil mencocokkan |

---

### `1_CopyData.py` — Ekstraksi data dari ARVIEWER

Mempersiapkan tiga file kerja di folder `Dapur/`:

| File output | Sumber | Keterangan |
|---|---|---|
| `ARClean_temp.xlsx` | ARVIEWER.xlsm sheet `arvi_ar_sheet` | Data AR baris-per-faktur untuk lookup |
| `FBackCust.xlsx` | ARVIEWER.xlsm sheet `arvi_name_out` | Daftar nama pelanggan resmi (feedback/master) |
| `Hasil_Latihan.xlsx` | File di path `ml_trainning` | Salinan lokal hasil pelatihan ML |

Header sheet `Source` di ARVIEWER dideteksi otomatis berdasarkan keyword: `NAMA PELANGGAN`, `NAMA CUSTOMER`, `NO. FAKTUR`, `NO FAKTUR`, `TGL FAKTUR`, `SISA PIUTANG`. Fallback ke baris dengan kepadatan kolom terbanyak jika keyword tidak ditemukan.

---

### `2_InjectDataToSS.py` — Loop sinkronisasi ke Google Sheets

Inti sistem. Memuat semua data ke RAM, lalu memindai baris kosong di setiap Google Sheets produk dan mengisinya dengan nilai piutang + cell note terstruktur.

**Filter data AR saat preloading:**

```
ARClean_temp.xlsx
  ├─ Filter: Nama Penjual mengandung "SR" (hardcoded, wajib)
  ├─ Filter: Nama Kontak mengandung ar_key_filter (misal "IRC" atau "ZN")
  └─ Filter: Nama Penjual TIDAK mengandung "FRAUD" (jika ar_data_fraud = No)
```

**Urutan resolusi nama (`resolve_target_name_fast`):**

```
Input: nama mentah dari kolom key_col Google Sheets
  ↓
1. Exact match di ml_dict (dari Hasil_Latihan.xlsx)
  ↓ jika gagal
2. Exact match di fb_dict (dari FBackCust.xlsx)
  ↓ jika gagal
3. Fuzzy RapidFuzz token_set_ratio terhadap fb_list (threshold 60%)
  ↓ jika gagal
4. Gunakan input mentah apa adanya
```

**Urutan lookup AR (`get_ar_rows_fast`):**

```
Input: nama kanonik (output resolusi)
  ↓
1. Cek ar_memory exact (key = bersihkan_teks(nama))
  ↓ jika tidak ada
2. Fuzzy RapidFuzz token_set_ratio terhadap semua key ar_memory (threshold 80%, limit 3)
   → gabungkan baris dari semua match (deduplikasi per No. Faktur)
```

---

## ⚙️ Konfigurasi `config.conf`

### `[DIR]` — Path file lokal

```ini
[DIR]
arvi = E:\ADM IRC AND ZN\ARVIEWER.xlsm
arvi_ar_sheet = Source
arvi_name_out = Nama Pelanggan SS
ml_trainning = E:\ADM IRC AND ZN\AR Pusat Machine Learning\ML\Hasil_Latihan.xlsx
```

| Key | Keterangan |
|---|---|
| `arvi` | Path absolut ke file `ARVIEWER.xlsm` (bisa relatif atau absolut) |
| `arvi_ar_sheet` | Nama sheet di ARVIEWER yang berisi data AR per faktur |
| `arvi_name_out` | Nama sheet di ARVIEWER yang berisi daftar nama pelanggan resmi |
| `ml_trainning` | Path ke `Hasil_Latihan.xlsx` yang akan disalin ke `Dapur/` |

---

### `[AR]` — Konfigurasi per produk & flag output

Setiap produk dikonfigurasi dengan suffix yang sama. Contoh untuk produk `irc` dan `zn`:

```ini
[AR]
ar_url_irc = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
ar_sheet_irc = Sheet1
ar_key_col_irc = Nama Customer dan Kota
ar_prod_key_col_irc =
ar_target_col_irc = Nominal Nota Belum Lunas
ar_key_filter_irc = IRC

ar_url_zn = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
ar_sheet_zn = Form Responses 1
ar_key_col_zn = Nama Customer dan Kota
ar_prod_key_col_zn =
ar_target_col_zn = Nominal Nota Belum Lunas
ar_key_filter_zn = ZN

ar_time_interval = 5
```

**Kunci konfigurasi per produk (ganti `xxx` dengan suffix produk: `irc`, `zn`, dst.):**

| Key | Keterangan |
|---|---|
| `ar_url_xxx` | URL Google Sheets target. Kosong → produk ini di-skip |
| `ar_sheet_xxx` | Nama sheet/tab di spreadsheet |
| `ar_key_col_xxx` | Kolom berisi nama pelanggan di Google Sheets |
| `ar_prod_key_col_xxx` | Kolom produk/divisi (opsional, kosong = pakai `ar_data_prod`) |
| `ar_target_col_xxx` | Kolom yang akan diisi nilai + note AR |
| `ar_key_filter_xxx` | Kata kunci filter kolom `Nama Kontak` di data AR |

**Menambahkan produk baru:** Tambahkan satu blok dengan suffix baru (misal `pcmo`) dan tambahkan suffix-nya ke variabel `daftar_produk` di `2_InjectDataToSS.py`:
```python
daftar_produk = ["irc", "zn", "pcmo"]
```

---

### Flag output cell note

| Key | Default | Keterangan |
|---|---|---|
| `ar_time_interval` | `5` | Interval loop sinkronisasi (menit) |
| `ar_data_fraud` | `No` | `No` = sembunyikan baris FRAUD; `Ya` = tampilkan + tandai `(FRAUD)` |
| `ar_data_codecus` | `Ya` | Tampilkan nama mentah (raw key) di header note |
| `ar_data_namecus` | `Ya` | Tampilkan nama kanonik (hasil resolusi ML) di header note |
| `ar_data_prod` | — | Nama produk fallback jika `ar_prod_key_col` kosong |
| `ar_data_dt_order` | `Ya` | Tampilkan tanggal order (kolom A Google Sheets) |
| `ar_data_calc` | `Ya` | Isi nilai sel dengan total Sisa Piutang |
| `ar_data_inv_numb` | `No` | Nomor faktur per baris |
| `ar_data_inv_dt` | `Ya` | Tanggal faktur per baris |
| `ar_data_inv_due` | `Ya` | Jatuh tempo per baris |
| `ar_data_inv_val` | `Ya` | Jumlah total faktur aktif |
| `ar_data_inv_orig` | `Ya` | Nilai faktur asli per baris |
| `ar_data_inv_ar` | `No` | Sisa piutang per baris |
| `ar_data_inv_pay` | `No` | Titip bayar per baris (`Nilai Faktur − Sisa Piutang`, jika > 0) |
| `ar_data_giro` | `Ya` | Tampilkan tanggal giro `(Tanggal JT)` per baris |
| `ar_data_age` | `Ya` | Umur piutang dalam hari |

---

## 📋 Format File Data Training

### `TheTrainningData.xlsx`

File Excel berisi data historis. Kolom yang dibutuhkan:

| Kolom | Nama | Deskripsi |
|---|---|---|
| A | *(bebas)* | Kolom data lain |
| B | *(bebas)* | |
| C | *(bebas)* | |
| D | *(nominal)* | Nilai nominal order — digunakan `LookupData.py` untuk lookup |
| Bebas | `Nama Customer dan Kota` | Nama pelanggan seperti ditulis di order sheet |
| Bebas | `Ekstrak_Customer_Detail (F)` | Nama resmi dari file historis (diisi `LookupData.py`) |
| Bebas | `Ekstrak_Customer_Detail (F2)` | Koreksi manual admin (prioritas tertinggi) |
| Bebas | `Ekstrak_Customer_Detail (F3)` | Sumber alternatif ketiga |

### `Hasil_Latihan.xlsx`

Output `TrainingModel.py`. Kolom penting yang dibaca oleh `2_InjectDataToSS.py`:

| Kolom | Keterangan |
|---|---|
| `Nama Customer dan Kota` | Input asli (kunci lookup) |
| `Hasil_Nama_Rekomendasi` | Nama kanonik hasil model |

---

## 🧠 Algoritma Resolusi Nama Pelanggan

### Saat Training (`TrainingModel.py`)

Model dilatih **sekali** dan hasilnya disimpan di `Hasil_Latihan.xlsx`. Empat lapis berjenjang:

```
Untuk setiap baris input di TheTrainningData.xlsx:

  LAPIS 1 — Local Match (prioritas tertinggi)
    Jika kolom F2, F, atau F3 tidak kosong → gunakan langsung (skor 100%)
    Status: "SUCCESS (Match Lokal)"

  LAPIS 2 — Exact Memori Historis
    Jika input_clean persis cocok di kamus memori → ambil nilai dari kamus
    Status: "SUCCESS (Match Exact Memori Historis)"

  LAPIS 3 — RapidFuzz Token Set Ratio
    Cari input paling mirip di kamus historis (threshold 75%)
    Jika ditemukan → ambil nama target dari kamus
    Status: "SUCCESS (RapidFuzz Match: '...')"

  LAPIS 4 — TF-IDF + k-NN Cosine Similarity
    Vectorize input dengan karakter n-gram (2–4)
    Temukan 1 tetangga terdekat via k-NN
    Jika similarity ≥ 0.65 → gunakan nama target tetangga
    Status: "SUCCESS (ML Prediction - Match: '...')"

  Jika semua gagal:
    Status: "GAGAL (Skor RapidFuzz & ML di bawah batas minimal)"
```

### Saat Runtime (`2_InjectDataToSS.py`)

Digunakan untuk mencocokkan nama dari Google Sheets ke nama kanonik:

```
resolve_target_name_fast(raw_key):
  1. Exact ke ml_dict (dari Hasil_Latihan)  → threshold: exact
  2. Exact ke fb_dict (dari FBackCust)      → threshold: exact
  3. Fuzzy ke fb_list via RapidFuzz         → threshold: 60%
  4. Fallback: gunakan raw_key apa adanya

get_ar_rows_fast(target_clean):
  1. Exact ke ar_memory                     → threshold: exact
  2. Fuzzy ke ar_memory keys via RapidFuzz  → threshold: 80%, limit 3
```

---

## 📤 Format Output: Cell Value & Cell Note

### Nilai Sel

Total `Sisa Piutang` seluruh faktur aktif dalam format IDR:
```
1.234.567
```

### Cell Note

```
Toko Makmur Mgl     Toko Makmur     IRC
Tanggal Order: 15/04/2026

========================================
RINGKASAN PERFORMA PIUTANG
========================================
Piutang                  :  1.234.567
Total Faktur Aktif (Inv) :  4

========================================
DAFTAR RINCIAN FAKTUR AKTIF
========================================

15 Jan 2026   15 Feb 2026   1.000.000   30 HR
22 Jan 2026   22 Feb 2026   500.000     37 HR (JT BG 15 Mar 2026)

01 Feb 2026   01 Mar 2026   400.000     13 HR
10 Feb 2026   10 Mar 2026   334.567     5 HR
```

Fitur khusus cell note versi ML Edition:
- **Header 3 kolom:** nama mentah (raw), nama kanonik (ML), nama produk
- **Baris faktur diurutkan** per tanggal faktur (ascending)
- **Pemisah bulan otomatis** — baris kosong disisipkan di antara kelompok bulan berbeda
- **Tanggal giro** `(Tanggal JT)` ditampilkan jika tersedia di data AR

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat/pilih project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Buka **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON Service Account Anda.

### 3. Berikan akses ke Google Sheets

Tambahkan `client_email` dari `credentials.json` sebagai **Editor** di setiap Google Sheets yang dikonfigurasi (`ar_url_irc`, `ar_url_zn`, dst.).

### 4. Struktur Google Sheets target

- Baris pertama = header kolom
- Kolom `ar_key_col_xxx` = berisi nama pelanggan per baris order
- Kolom `ar_target_col_xxx` = akan diisi otomatis (awalnya kosong)
- Kolom A = tanggal order (digunakan jika `ar_data_dt_order = Ya`)

---

## 🛠️ Troubleshooting

### ❌ `File 'Hasil_Latihan.xlsx' TIDAK ditemukan di dalam folder 'ML'`
Pelatihan model belum dijalankan. Jalankan `ML/TrainingModel.py` terlebih dahulu untuk menghasilkan file ini.

### ❌ `File ARVIEWER tidak ditemukan pada path: ...`
Path di `[DIR] arvi` di `config.conf` tidak valid. Pastikan path menunjuk ke file `ARVIEWER.xlsm` yang benar dan bisa diakses. Gunakan path absolut untuk menghindari ambiguitas.

### ❌ `Kolom target 'Nama Pelanggan' tidak ditemukan pada file 'ARClean_temp.xlsx'`
Sheet `Source` di ARVIEWER.xlsm tidak mengandung header yang dikenali. `1_CopyData.py` akan gagal mendeteksi header. Pastikan sheet berisi minimal salah satu kolom: `NAMA PELANGGAN`, `NO. FAKTUR`, `TGL FAKTUR`, atau `SISA PIUTANG`.

### ❌ `[MELEWATI PRODUK IRC] URL (ar_url_irc) tidak diisi di config.conf`
Kolom `ar_url_irc` kosong. Isi URL Google Sheets yang sesuai. Produk ini akan di-skip sampai URL diisi.

### ❌ `[ERROR] Kolom target di Google Sheets tidak ditemukan`
Nilai `ar_key_col_xxx` atau `ar_target_col_xxx` tidak ada di baris pertama (header) sheet Google Sheets. Periksa ejaan dan spasi — pencocokan bersifat case-insensitive tapi spasi ekstra dapat memengaruhi hasil.

### ❌ Banyak baris di Google Sheets tidak terisi meski ada data AR
Kemungkinan: (1) resolusi nama gagal → nama kanonik berbeda dari `Nama Pelanggan` di ARClean; (2) filter SR tidak terpenuhi — pastikan data AR di ARVIEWER sudah diperbarui dan `Nama Penjual` mengandung "SR"; (3) filter `ar_key_filter` terlalu ketat — periksa isi kolom `Nama Kontak` di ARClean.

### ❌ `Error: tidak ada data historis valid untuk melatih model`
`TheTrainningData.xlsx` tidak memiliki baris dengan kolom `Nama Customer dan Kota` dan setidaknya satu kolom F/F2/F3 terisi. Jalankan `LookupData.py` terlebih dahulu atau isi manual beberapa baris referensi.

### ❌ Nama yang diresolved salah padahal ada di ML dict
Cek file `Hasil_Latihan.xlsx` — cari baris dengan `Nama Customer dan Kota` yang relevan dan lihat kolom `Status_Pencocokan`. Jika `GAGAL`, pertimbangkan menambahkan koreksi manual di kolom `Ekstrak_Customer_Detail (F2)` dan latih ulang model.

### ❌ Error autentikasi Google
Periksa `Dapur/credentials.json` — pastikan `private_key` tersalin lengkap dengan `-----BEGIN PRIVATE KEY-----` dan `-----END PRIVATE KEY-----`.

---

## 📌 Catatan Penting

- **`LookupData.py` memiliki path hardcoded** — Ubah `FOLDER_BASE` di baris 8 sebelum menjalankan. Ini adalah skrip utilitas sekali pakai, bukan bagian dari pipeline otomatis.
- **Pelatihan ulang diperlukan saat ada pelanggan baru** — Jika ada nama pelanggan baru yang terus gagal dicocokkan, tambahkan ke `TheTrainningData.xlsx` (minimal isi `Nama Customer dan Kota` dan `Ekstrak_Customer_Detail (F2)`), lalu jalankan ulang `TrainingModel.py` dan `1_CopyData.py`.
- **Filter "SR" hardcoded** — Hanya baris AR dengan `Nama Penjual` mengandung "SR" yang masuk ke `ar_memory`. Ini tidak dapat diubah via `config.conf`. Pastikan data AR di ARVIEWER sudah benar.
- **ARVIEWER.xlsm harus diperbarui terlebih dahulu** — Proyek ini membaca dari ARVIEWER, bukan langsung dari Accurate. Jalankan pipeline ARVIEWER untuk memperbarui data sebelum menjalankan pipeline ini.
- **`credentials.json` bersifat rahasia** — Tambahkan ke `.gitignore`. Jangan commit ke repositori publik.
- **Hanya sel kosong yang diisi** — Baris yang kolom `ar_target_col` sudah berisi nilai apapun tidak akan ditimpa. Kosongkan manual jika ingin memperbarui.
- **Menambah produk baru** — Tambahkan blok konfigurasi `ar_url_newprod`, `ar_sheet_newprod`, dst. di `config.conf`, lalu tambahkan `"newprod"` ke `daftar_produk` di `2_InjectDataToSS.py`.
- **`Hasil_Latihan.xlsx` di `ML/` adalah master** — `1_CopyData.py` hanya menyalin, tidak mengubah file di `ML/`. Update selalu bermula dari menjalankan ulang `TrainingModel.py`.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
