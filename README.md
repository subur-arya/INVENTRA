# INVENTRA

<p align="center">
  <img src="logo_trsp.png" alt="INVENTRA Logo" width="420"/>
</p>

<p align="center">
  <b>INVENTORI REPORT ANALISIS</b><br/>
  Aplikasi desktop untuk analisis dan monitoring pengadaan material yang dikembangkan selama program magang di bidang Inventori PT PLN Nusantara Power UP Gresik.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white"/>
  <img src="https://img.shields.io/badge/GUI-Tkinter-orange"/>
  <img src="https://img.shields.io/badge/Excel-openpyxl%20%7C%20pandas-green?logo=microsoft-excel"/>
  <img src="https://img.shields.io/badge/Cloud-Google%20Sheets-yellow?logo=google-sheets"/>
  <img src="https://img.shields.io/badge/Build-PyInstaller-purple"/>
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows"/>
</p>

---

## Daftar Isi

- [Tentang INVENTRA](#tentang-inventra)
- [Fitur Utama](#fitur-utama)
- [Arsitektur Aplikasi](#arsitektur-aplikasi)
- [Struktur File Project](#struktur-file-project)
- [Penjelasan File Python](#penjelasan-file-python)
- [Penjelasan File Aset](#penjelasan-file-aset)
- [Alur Kerja Aplikasi](#alur-kerja-aplikasi)
- [Formula Kalkulasi Inventory](#formula-kalkulasi-inventori)
- [Cara Menjalankan](#cara-menjalankan)
- [Build ke EXE](#build-ke-exe)
- [Konfigurasi INVENTRA.json](#konfigurasi-inventrajson)
- [Dependensi](#dependensi)

---

## Tentang INVENTRA

**INVENTRA** adalah aplikasi manajemen inventori berbasis desktop yang dibangun menggunakan Python (Tkinter). Aplikasi ini dirancang untuk membantu tim inventori dalam:

1. **Menganalisis kebutuhan reorder** — menentukan item mana yang perlu dipesan ulang berdasarkan parameter ROP (Re-Order Point) dan ROQ (Re-Order Quantity).
2. **Memonitor status pengadaan** — melalui modul DRP (Dokumen Rencana Pengadaan) yang terhubung ke Google Sheets secara online.
3. **Memproses data dari sistem PLN** — dengan membaca file Excel dari berbagai sumber data (PLJM, SRD, SLN, IR, PO, LEVERING) dan menggabungkannya secara otomatis.

Aplikasi ini dapat dikompilasi menjadi file `.exe` menggunakan PyInstaller sehingga dapat dijalankan tanpa instalasi Python di komputer pengguna.

---

## Fitur Utama

| Fitur | Deskripsi |
|-------|-----------|
| 📂 **Upload Multi-File** | Upload beberapa file Excel sekaligus dengan auto-detect baris header |
| ⚙️ **Flexible Column Mapping** | Konfigurasi nama kolom Excel via dialog sehingga mendukung berbagai format file ERP |
| 📊 **Kalkulasi Otomatis** | Hitung ROP, ROQ, proyeksi stok, dan klasifikasi ORDER / TIDAK ORDER / PERLU REVIEW |
| 💾 **Export Excel** | Simpan hasil analisis ke file `.xlsx` dengan styling lengkap (header berwarna, freeze pane, autofilter) |
| 🔄 **Sync Google Sheets** | Kirim data DRP ke Google Sheets via Google Apps Script (GAS), dengan smart merge (data manual user tidak terhapus) |
| 📋 **Modul DRP** | Kelola data pengadaan berdasarkan nomor PRK dengan klasifikasi otomatis (PRK I, AO, Gabungan, Tanpa PRK) |
| ❓ **Tutorial In-App** | Panduan penggunaan interaktif langsung di dalam aplikasi |
| 💾 **Konfigurasi Persisten** | Semua pengaturan disimpan di `INVENTRA.json` — tidak perlu konfigurasi ulang setiap sesi |

---

## Arsitektur Aplikasi

```
┌─────────────────────────────────────────────────────────────┐
│                    INVENTRA Desktop App                      │
│                                                             │
│  ┌──────────┐    ┌──────────────────────────────────────┐  │
│  │  Splash  │───▶│         Module Chooser               │  │
│  │  Screen  │    │   [Settingan]      [Proses DRP]       │  │
│  └──────────┘    └──────┬───────────────────┬────────────┘  │
│                         │                   │               │
│               ┌─────────▼──────┐   ┌────────▼──────────┐   │
│               │   Main App     │   │     DRP App        │   │
│               │  (INVENTRA)    │   │   (drp_app.py)     │   │
│               │                │   │                    │   │
│               │ Upload Files   │   │ Upload AMP Excel   │   │
│               │ ↓              │   │ ↓                  │   │
│               │ MappingDialog  │   │ Proses DRP         │   │
│               │ ↓              │   │ ↓                  │   │
│               │ data.filter_   │   │ Sync Google Sheets │   │
│               │ all()          │   │                    │   │
│               │ ↓              │   └────────────────────┘   │
│               │ Preview Data   │                            │
│               │ ↓              │                            │
│               │ Save Excel     │                            │
│               └────────────────┘                            │
└─────────────────────────────────────────────────────────────┘
         │                              │
         ▼                              ▼
   File Excel (.xlsx)          Google Sheets (online)
   hasil analisis              via Google Apps Script
```

---

## Struktur File Project

```
INVENTRA 1.7/
│
├── 📄 main.py                  # Entry point & main window aplikasi
├── 📄 data.py                  # Proses perhitungan data dan filterisasi
├── 📄 drp_app.py               # Modul DRP dengan berisi UI dan logika pemrosesan filter file AMP
├── 📄 proses_GSheet.py         # Sinkronisasi data ke Google Sheets
├── 📄 tutorial.py              # Layar tutorial in-app
│
├── 📄 INVENTRA.json            # Konfigurasi persisten (mapping kolom + path file)
├── 📄 build_app.spec           # Konfigurasi build PyInstaller untuk menjadikannya ke `.exe`
├── 📄 build to apk.txt         # Catatan perintah build
│
├── 🖼️ logo_app.ico             # Ikon aplikasi (format Windows .ico)
├── 🖼️ logo_app.png             # Ikon aplikasi (format PNG)
├── 🖼️ logo_trsp.png            # Logo INVENTRA transparan (splash + header)
├── 🖼️ inventraabt.png          # Gambar halaman About di tutorial
├── 🖼️ icon_settingan.png       # Ikon tombol Settingan
├── 🖼️ icon_drp.png             # Ikon tombol Proses DRP
│
├── 🖼️ tutorial_upload.png      # Gambar langkah 1 tutorial (Upload File)
├── 🖼️ tutorial_mapping.png     # Gambar langkah 2 tutorial (Mapping)
├── 🖼️ tutorial_preview.png     # Gambar langkah 3 tutorial (Preview & Save)
├── 🖼️ tutorial_reset.png       # Gambar langkah 4 tutorial (Reset)
│
├── 📁 rumus_slide/             # Folder slide PNG penjelasan rumus dari proses seettingan
│   └── *.png                   # File-file gambar halaman penjelasan rumus
│
├── 📁 build/                   # Folder output sementara PyInstaller
├── 📁 .vscode/                 # Konfigurasi editor VS Code
└── 📁 __pycache__/             # Cache bytecode Python (auto-generated)
```

---

## Penjelasan File Python

### `main.py` — Entry Point & Pusat Kendali

File terbesar dan terpenting. Mengintegrasikan semua modul dan mengelola seluruh siklus hidup aplikasi.

**Dua kelas utama:**

#### `MappingDialog`
Dialog konfigurasi yang muncul saat user klik `⚙ Process`. Bertugas menghubungkan nama kolom di file Excel user dengan field yang dibutuhkan program.

```
Mengapa perlu dialog ini?
Setiap periode instansi mungkin punya nama kolom Excel yang berbeda.
Contoh: "STOCK_CODE" vs "Stock Code" vs "Material No."
MappingDialog memungkinkan konfigurasi tanpa mengubah kode program.
```

Fitur MappingDialog:
- 9 tab konfigurasi (PLJM01, PLJM08, SRD, SLN, IR, PO, LEVERING, ANALISIS SETTING, ANALISIS NON SETTING)
- Per tab: pilih Source File → Sheet Name → mapping kolom (dropdown dinamis berdasarkan isi sheet)
- Konfigurasi otomatis tersimpan ke `INVENTRA.json`
- Support drag dialog dan tombol close custom

#### `ModernRedINVENTRAManager`
Main window aplikasi dengan state machine navigasi berlapis.

**Alur state:**
```
Splash Screen → Module Chooser → [Main App] atau [DRP App]
```

**Mekanisme threading:**
- Operasi berat (baca Excel, simpan Excel, cek internet) dijalankan di background thread
- Hasil dikembalikan ke main thread via `root.after()` — karena Tkinter tidak thread-safe
- Loading overlay mencegah user berinteraksi saat proses berlangsung

**Dual sync mechanism (splash screen):**
```python
# Splash tidak hilang sampai DUA hal selesai:
self._splash_anim_done = False  # animasi progress bar 0→100%
self._splash_json_done = False  # pembuatan INVENTRA.json di background thread
# Keduanya harus True sebelum lanjut ke Module Chooser
```

---

### `data.py` — Engine Kalkulasi Inventory

File ini adalah "otak" pemrosesan data. Tidak memiliki UI, murni logika bisnis.

**Sumber data yang diproses:**

| Sumber | Keterangan | Filter Utama |
|--------|-----------|--------------|
| `PLJM01` | Tabel master parameter (ROP & ROQ per item) | (dipakai untuk VLOOKUP) |
| `PLJM08` | Tabel utama inventory (SOH, info item) | (tabel referensi utama) |
| `SRD` | Stock Receipt Document >> laporan penerimaan material dari supplier | `qty_rcv_uop > 0` |
| `SLN` | Service Level >> laporan pengeluaran material dari Gudang unit kerja | `qty_req > 0` AND `qty_issued > 0` |
| `IR` | Issue Requisition >> laporan data permintaan material yang belum terpenuhi | `qty_issued == 0` AND `req_by_date >= hari ini` |
| `PO` | Purchase Order >> laporan data pemesanan material stock yang sudah diterima gudang | `receipt_status == 2` AND `curr_qty != 0` |
| `LEVERING` | laporan data pemesanan material yang belum diterima gudang | `receipt_status == 0` AND `due_date >= (hari ini - 60 hari)` |

**Teknik penggabungan data (proses_pjm08):**

Alih-alih merge/join berulang yang lambat, digunakan teknik `groupby + map`:
```python
# 1. Buat lookup Series berindeks stock_code
map_qty_issued = df_sln.groupby(stock_code)[qty_issued].first()

# 2. Petakan ke setiap baris PLJM08
df_pljm08['Qty ISS'] = df_pljm08[stock_code].map(map_qty_issued)
```

**Formula klasifikasi:**

```
calc = SOH - Next_Req_Qty + Levering_Qty

Jika ROQ == 0 dan Next_Req == 0  → QTY_RO = 0       (tidak dianalisis)
Jika ROQ == 0 dan Next_Req != 0  → QTY_RO = -1      (PERLU REVIEW)
Jika Levering > ROP + ROQ        → QTY_RO = -1      (PERLU REVIEW)
Jika calc > ROP                  → QTY_RO = 0       (TIDAK ORDER)
Jika calc < ROP                  → QTY_RO = (ROP + ROQ) - calc  (ORDER)
Jika calc == ROP                 → QTY_RO = ROQ     (ORDER)
```

**Output akhir:**
- `ANALISIS SETTING` — item dengan ROQ yang sudah di-set
- `ANALISIS NON SETTING` — item tanpa ROQ tapi ada permintaan mendatang

---

### `drp_app.py` — Modul DRP

Modul Distribution Requirement Planning untuk memonitor status pengadaan dari PRK hingga pembayaran.

**Klasifikasi nomor PRK:**

| Kategori | Ciri | Contoh |
|----------|------|--------|
| **PRK I (AI)** | Huruf A setelah 3 digit, digit ke-3 = '4' | `GR254A0205` |
| **AO** | Ada huruf A tapi tidak memenuhi ciri PRK I | `GR253A0205` |
| **PRK GABUNGAN** | Lebih dari 1 nomor PRK dalam satu sel | `GR254A0205 GR253A0205` |
| **Tanpa PRK** | Angka murni atau teks deskriptif | `121046` / `PEKERJAAN SIPIL` |

**Regex PRK Pattern:**
```python
PRK_PATTERN = re.compile(
    r'(?:GR)?\d{2,}[A-Z]{1,2}\d{2,}'    # format huruf: GR254A0205
    r'|'
    r'(?<![.\d])(?:GR)?\d{6,}(?![.\d])' # format angka: 121046
)
```

**Alur proses AMP → DRP:**
```
File Excel AMP
  → Baca sheet (auto-detect header)
  → Extract PRK dari setiap sel (regex)
  → Normalisasi PRK (tambah prefix GR jika perlu)
  → Klasifikasikan PRK (AI / AO / Gabungan / Tanpa PRK)
  → groupby PRK → satu baris per PRK (multi-item digabung dengan "\n")
  → Tampilkan di Treeview preview
  → Sync ke Google Sheets
```

**Kelas DRPApp:**
- UI dengan tab preview: DRP | AMP | Tanpa PRK
- Threading untuk operasi proses dan sync
- Processing overlay mencegah double-click saat loading
- Dialog informatif saat sheet GSheet belum ada (lengkap dengan tombol "Salin Nama")

---

### `proses_GSheet.py` — Sinkronisasi Google Sheets

Menangani komunikasi antara INVENTRA (lokal) dan Google Sheets (cloud) melalui Google Apps Script Web App.

**Arsitektur komunikasi:**
```
INVENTRA App → HTTP POST (urllib) → Google Apps Script Web App → Google Sheets API → Spreadsheet
```

**Mengapa pakai Google Apps Script, bukan Google Sheets API langsung?**
- Tidak butuh OAuth / service account credential
- GAS sudah punya akses penuh ke spreadsheet milik akun Google yang deploy-nya
- Cukup satu URL endpoint, tanpa library google-auth yang kompleks

**Format data di Google Sheets:**

Sheet menggunakan marker rows untuk instruksi styling ke GAS:
```
##COVER_TITLE##   → judul besar (di-styling oleh GAS)
##SEKSI##AO       → label pemisah seksi (bold, berwarna)
##HEADER##        → baris header kolom
[data biasa]
##EMPTY##         → baris kosong visual
```

**Smart merge (merge_with_existing):**

Saat sync, kolom data DRP otomatis di-update, sedangkan kolom yang diisi manual di GSheet (DASPEN, TORRAB, METODE PENGADAAN, dll) **dipertahankan**:
```python
# Untuk setiap baris:
# - Kolom KOLOM_DRP    → ambil dari data lokal terbaru
# - Kolom KOLOM_MANUAL → pertahankan dari GSheet lama
# - PRK baru           → tambahkan di bawah
# - PRK lama yang hilang → tetap disimpan (tidak dihapus)
```

**Custom exceptions untuk UI informatif:**
```python
SheetTidakDitemukanError   # sheet belum dibuat → tampilkan instruksi + tombol salin nama
HeaderTidakDitemukanError  # header sheet rusak/berubah
GSheetResponseError        # GAS mengembalikan error
```

---

### `tutorial.py` — Layar Tutorial In-App

Layar panduan penggunaan yang bisa diakses kapan saja tanpa keluar dari aplikasi.

**Pola desain "Overlay Screen":**
```python
# TutorialScreen tidak membuat window baru (Toplevel)
# Melainkan frame besar yang menimpa konten main window via .place()
self.tutorial_frame = tk.Frame(self.root, bg=self.colors['bg'])
self.tutorial_frame.place(x=0, y=0, relwidth=1, relheight=1)
# Saat ditutup → frame di-destroy, main window kembali terlihat
```

**Tiga menu sidebar:**

| Menu | Konten |
|------|--------|
| **ⓘ About** | Gambar `inventraabt.png` di-resize proporsional ke lebar layar |
| **Tutorial** | 4 step panduan (Upload → Mapping → Preview → Reset), masing-masing dengan gambar dan tips |
| **Logika** | Slide PNG dari folder `rumus_slide/` |

**Scrollable canvas pattern (standar Tkinter):**
```python
Canvas → create_window(scrollable_frame) → Scrollbar
# Canvas = viewport, scrollable_frame = konten sebenarnya
# Saat frame membesar → update scrollregion canvas
```

**Gotcha ImageTk Tkinter:**
```python
# PENTING: referensi photo HARUS disimpan di variabel instance
# Jika tidak, Python garbage collector akan menghapusnya dan gambar hilang
img_label.image = photo          # simpan di widget
self.rumus_images.append(photo)  # simpan di list instance
```

---

## Penjelasan File Aset

| File | Ukuran | Keterangan |
|------|--------|-----------|
| `logo_app.ico` | 1.0 MB | Ikon aplikasi format Windows `.ico` dipakai oleh `root.iconbitmap()` dan PyInstaller |
| `logo_app.png` | 356 KB | Ikon aplikasi format PNG, fallback jika `.ico` gagal, juga dipakai di custom titlebar (20×20px) |
| `logo_trsp.png` | 98 KB | Logo INVENTRA dengan background transparan yang dipakai di splash screen, loading screen, dan header main app |
| `inventraabt.png` | 273 KB | Gambar halaman "About INVENTRA"  ditampilkan di sidebar tutorial |
| `icon_settingan.png` | 161 KB | Ikon tombol "Settingan" di Module Chooser |
| `icon_drp.png` | 173 KB | Ikon tombol "Proses DRP" di Module Chooser |
| `tutorial_upload.png` | 222 KB | Screenshot langkah Upload yang ditampilkan di panel Tutorial |
| `tutorial_mapping.png` | 149 KB | Screenshot langkah Mapping |
| `tutorial_preview.png` | 54 KB | Screenshot langkah Preview & Save |
| `tutorial_reset.png` | 198 KB | Screenshot langkah Reset |
| `rumus_slide/*.png` | — | Slide-slide penjelasan logika kalkulasi ROP/ROQ |
| `INVENTRA.json` | 4 KB | File konfigurasi persisten (mapping kolom + path file yang pernah diupload) |
| `build_app.spec` | 1.6 KB | Konfigurasi build PyInstaller (daftar aset yang di-bundle ke EXE) |
| `build to apk.txt` | 46 B | Catatan perintah build PyInstaller |

---

## Alur Kerja Aplikasi

### Modul INVENTRA (Settingan)

```
1. Buka aplikasi
   └── Splash Screen (animasi + buat INVENTRA.json)
       └── Module Chooser

2. Klik "Settingan"
   └── Loading Screen (baca INVENTRA.json + file Excel lama)
       └── Main App

3. Upload File Excel
   └── Pilih satu atau lebih file Excel (.xlsx/.xls)
       └── Auto-detect baris header (cari "Stock Code" + "District")
           └── Card file muncul di UI

4. Klik ⚙ Process
   └── MappingDialog terbuka
       └── Per tab: pilih Source File → Sheet → Kolom
           └── Klik "✓ Lanjut"
               └── data.filter_all() dipanggil:
                   ├── filter_srd() → filter_sln() → filter_ir()
                   ├── filter_po() → filter_levering()
                   ├── proses_pljm01() → VLOOKUP ROP & ROQ
                   ├── proses_pjm08() → gabungkan semua data + kalkulasi
                   ├── analisis_setting() → tabel ORDER/TIDAK ORDER
                   └── analisis_non_setting() → tabel item tanpa ROQ

5. Preview Data
   └── Pilih tab (PLJM08 / SRD / SLN / ... / ANALISIS SETTING)
       └── Treeview menampilkan data dengan auto-width kolom

6. Save Excel
   └── Dialog simpan file
       └── Background thread: tulis workbook dengan openpyxl
           └── Header merah, zebra stripes, freeze pane, autofilter, auto-fit kolom
```

### Modul DRP

```
1. Klik "Proses DRP" di Module Chooser
   └── Loading Screen (cek koneksi Google Sheets)
       └── DRP App

2. Upload File AMP
   └── Pilih file Excel AMP

3. Klik ⚙ Proses DRP
   └── proses_drp() → generate_drp_from_amp_df()
       ├── Baca sheet AMP (auto-detect header)
       ├── Extract PRK dari setiap baris (regex PRK_PATTERN)
       ├── Normalisasi PRK (tambah prefix GR)
       ├── Klasifikasikan: AI / AO / PRK GABUNGAN / Tanpa PRK
       └── groupby PRK → satu baris per PRK (multi-item digabung "\n")

4. Preview di Tab: DRP | AMP | Tanpa PRK

5. Sync ke Google Sheets
   └── Check sheet exists (HTTP GET ke GAS)
       └── Fetch data lama dari GSheet
           └── merge_with_existing() → smart merge
               └── susun_per_seksi() → tambah marker rows
                   └── push_sheet() → HTTP POST ke GAS
```

---

## Formula Kalkulasi Inventory

### Proyeksi Stok (calc)

```
calc = SOH_akhir - Next_Req_Qty + Levering_Qty
```

- **SOH_akhir** — Stock on Hand saat ini (dari PLJM08)
- **Next_Req_Qty** — Total qty permintaan mendatang yang belum dipenuhi (dari IR)
- **Levering_Qty** — Total qty barang yang akan tiba dari supplier (dari LEVERING)

### Quantity yang Perlu Di-Order (QTY_RO)

```
Kondisi                          → QTY_RO
─────────────────────────────────────────────────────────────
ROQ == 0  AND  Next_Req == 0    → 0          (tidak dianalisis)
ROQ == 0  AND  Next_Req != 0    → -1         (PERLU REVIEW — belum ada parameter)
Levering > ROP + ROQ            → -1         (PERLU REVIEW — levering sudah cukup)
calc > ROP                      → 0          (TIDAK ORDER — stok masih aman)
calc == ROP                     → ROQ        (ORDER sebesar satu ROQ)
calc < ROP                      → (ROP + ROQ) - calc   (ORDER — hitung kekurangan)
```

### Klasifikasi Teks

```python
QTY_RO == -1  →  "PERLU REVIEW"
QTY_RO ==  0  →  "TIDAK ORDER"
QTY_RO  >  0  →  "ORDER"
```

### Evaluasi Perubahan (Analisis Setting)

Membandingkan hasil analisis saat ini dengan data analisis periode sebelumnya:

| Kondisi | Label Evaluasi |
|---------|---------------|
| Item baru (tidak ada di data lama) | `Perlu di analisis` |
| Nama item berubah | `Item name order baru` |
| Supplier berubah | `Beda supplier` |
| Status klasifikasi berubah | `Status analisis berbeda` |
| Tidak ada perubahan | `telah di analisis` |

---

## Cara Menjalankan

### Prasyarat

```bash
Python 3.14.2 (ketika build sistem) atau yang lain yang masih bisa berjalan atau relevan
```

### Instalasi dependensi

```bash
pip install -r Requirements.txt
```

### Jalankan aplikasi

```bash
python main.py
```

### Persiapan file aset

Pastikan file-file berikut ada di folder yang sama dengan `main.py`:

```
logo_app.ico
logo_app.png
logo_trsp.png
inventraabt.png
icon_settingan.png
icon_drp.png
tutorial_upload.png
tutorial_mapping.png
tutorial_preview.png
tutorial_reset.png
rumus_slide/   (folder berisi file .png)
```

---

## Build ke EXE

### Menggunakan `build_app.spec`

```bash
pyinstaller build_app.spec --clean --noconfirm
```

Output `.exe` akan ada di folder `dist/` auto generate.

**Catatan `resource_path()`:** Semua akses file menggunakan fungsi ini agar path benar baik saat mode development (`.py`) maupun produksi (`.exe`):
```python
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # folder temp PyInstaller saat EXE berjalan
    except AttributeError:
        base_path = os.path.abspath(".")  # folder kerja saat development
    return os.path.join(base_path, relative_path)
```

---

## Konfigurasi INVENTRA.json

File ini dibuat otomatis saat pertama kali aplikasi dijalankan. Berisi dua bagian utama:

```json
{
  "uploaded_files": {
    "file_1": {
      "path": "D:/path/ke/file/PLJM08.xlsx",
      "size": 102400,
      "filename": "PLJM08.xlsx"
    }
  },
  "data_mappings": {
    "PLJM08": {
      "source_file_id": "file_1",
      "sheet_name": "Sheet1",
      "columns": {
        "stock_code": "STOCK_CODE",
        "item_name": "ITEM_NAME",
        "soh_akhir": "SOH Akhir",
        "distric": "Distric",
        "exp": "Exp"
      }
    },
    "SRD": { ... },
    "SLN": { ... },
    "IR":  { ... },
    "PO":  { ... },
    "LEVERING": { ... },
    "PLJM01": { ... },
    "ANALISIS SETTING": { ... },
    "ANALISIS NON SETTING": { ... }
  }
}
```

**`uploaded_files`** — Daftar file Excel yang pernah diupload. Path disimpan agar saat aplikasi dibuka kembali, file otomatis dimuat ulang.

**`data_mappings`** — Konfigurasi mapping kolom per jenis data. Nilai `columns` berisi pasangan `{field_internal: nama_kolom_di_excel}`.

---

## Dependensi

### Library Bawaan Python (tidak perlu diinstall)

| Library | Fungsi |
|---------|--------|
| `tkinter` | GUI framework utama — window, widget, event loop |
| `urllib` | HTTP GET/POST ke endpoint Google Apps Script |
| `json` | Baca/tulis konfigurasi `INVENTRA.json` |
| `threading` | Background thread untuk operasi berat (baca Excel, save, sync) |
| `re` | Regex untuk ekstraksi dan klasifikasi nomor PRK |
| `os` / `sys` | Path file dan deteksi mode EXE vs development |

### Library Pihak Ketiga (install via `pip install -r requirements.txt`)

| Library | Versi | Fungsi |
|---------|-------|--------|
| `pandas` | 3.0.1 | Pemrosesan data tabular — DataFrame, groupby, merge, map, filter |
| `openpyxl` | 3.1.5 | Baca/tulis file Excel `.xlsx` dengan styling (header, freeze pane, autofilter) |
| `xlrd` | 2.0.2 | Baca file Excel format `.xls` lama (Excel 97-2003) |
| `et_xmlfile` | 2.0.0 | Dependensi internal openpyxl |
| `numpy` | 2.4.2 | Operasi array dan `np.select` untuk klasifikasi ORDER/TIDAK ORDER |
| `Pillow` | 12.1.1 | Memuat dan memanipulasi gambar PNG/JPG (logo, ikon, tutorial) |
| `python-dateutil` | 2.9.0.post0 | Utilitas tanggal — dependensi internal pandas |
| `six` | 1.17.0 | Kompatibilitas Python — dependensi internal |
| `tzdata` | 2025.3 | Data timezone — dependensi internal pandas |
| `requests` | 2.32.5 | HTTP request tambahan ke Google Apps Script |
| `urllib3` | 2.6.3 | Dependensi internal requests |
| `certifi` | 2026.2.25 | Bundle sertifikat SSL — dependensi internal requests |
| `charset-normalizer` | 3.4.4 | Deteksi encoding — dependensi internal requests |
| `idna` | 3.11 | Internasionalisasi nama domain — dependensi internal requests |
| `pywinstyles` | 1.8 | *(Opsional)* Ubah warna title bar Windows sesuai tema merah INVENTRA |

### Library Build (hanya dibutuhkan saat kompilasi ke `.exe`)

| Library | Versi | Fungsi |
|---------|-------|--------|
| `pyinstaller` | 6.19.0 | Kompilasi aplikasi Python menjadi file `.exe` standalone |
| `pyinstaller-hooks-contrib` | 2026.1 | Koleksi hooks untuk library pihak ketiga di PyInstaller |
| `altgraph` | 0.17.5 | Analisis dependency graph — dependensi internal PyInstaller |
| `pefile` | 2024.8.26 | Parser file PE (format binary Windows) — dependensi PyInstaller |
| `pywin32-ctypes` | 0.2.3 | Windows API binding — dependensi internal PyInstaller |
| `packaging` | 26.0 | Parsing versi package — dependensi internal PyInstaller |

---

<p align="center">
  Dikembangkan selama program Magang — Semester 6<br/>
  Universitas Negeri Surabaya (UNESA)
</p>