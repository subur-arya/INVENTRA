# =============================================================================
# FILE: data.py
# PERAN DALAM PROJECT: "Otak" perhitungan inventory
#
# File ini adalah ENGINE utama pemrosesan data INVENTRA. Ia tidak punya UI,
# murni logika bisnis. Tugasnya adalah:
#   1. Membaca konfigurasi nama kolom dari JSON (fleksibel per instansi)
#   2. Memfilter 5 sumber data mentah (SRD, SLN, IR, PO, LEVERING)
#   3. Menggabungkan semua data ke tabel utama PLJM08 (mirip VLOOKUP Excel)
#   4. Menghitung apakah suatu item perlu di-ORDER atau tidak
#   5. Menghasilkan dua tabel analisis final (SETTING dan NON SETTING)
#
# KONTEKS BISNIS:
#   INVENTRA dipakai oleh tim pengadaan/gudang (kemungkinan PLN) untuk
#   memutuskan item mana yang harus di-reorder. Keputusan ini berdasarkan
#   rumus ROP/ROQ (Re-Order Point / Re-Order Quantity) — konsep standar
#   manajemen inventory.
#
# ALUR DATA (gambaran besar):
#   PLJM01 ──(ROP,ROQ)──┐
#   SRD ────(qty SRD)───┤
#   SLN ────(qty ISS)───┼──► PLJM08 (tabel gabungan) ──► KLASIFIKASI ──► ANALISIS
#   IR ─────(next req)──┤
#   PO ─────(supplier)──┤
#   LEVERING ──(qty)────┘
# =============================================================================

import pandas as pd
from datetime import datetime
import numpy as np
import json
import os


# =============================================================================
# BAGIAN 1: JSON LOADER
# Fungsi ini membaca file konfigurasi "rpo_settings.json" yang berisi
# mapping nama kolom Excel. Ini penting karena setiap instansi perusahaan
# mungkin punya nama kolom berbeda di file Excel mereka.
# Contoh isi JSON: {"SRD": {"columns": {"stock_code": "STOCK_CODE", ...}}}
# =============================================================================

def baca_json():
    # Cek apakah file konfigurasi sudah ada di direktori kerja
    if os.path.exists("rpo_settings.json"):
        with open("rpo_settings.json", "r") as f:
            return json.load(f)
    # Jika belum ada, kembalikan dict kosong (program tidak crash)
    return {}


# =============================================================================
# BAGIAN 2: HELPER col()
# Shortcut untuk mengambil nama kolom dari mapping JSON.
# Tanpa helper ini, kode akan penuh dengan:
#   mapping["SRD"]["columns"]["stock_code"]  ← panjang dan rawan typo
# Dengan helper: col(mapping, "SRD", "stock_code")  ← bersih
# =============================================================================

def col(mapping, section, key):
    # mapping  → seluruh isi JSON konfigurasi
    # section  → nama sheet/data (misal "SRD", "PLJM08")
    # key      → nama field internal (misal "stock_code", "qty_issued")
    return mapping[section]["columns"][key]


# =============================================================================
# BAGIAN 3: FILTER SRD (Stock Receipt Document)
# SRD = dokumen penerimaan barang dari supplier ke gudang.
# Filter ini membersihkan data SRD agar hanya berisi penerimaan yang valid.
#
# ATURAN FILTER:
#   - Kolom stock_code dipindah ke posisi ke-2 (untuk konsistensi tampilan)
#   - Tanggal diformat jadi 8 digit (YYYYMMDD), lalu diurutkan terbaru dulu
#   - Hanya baris dengan qty_rcv_uop > 0 yang dipakai (penerimaan nyata)
# =============================================================================

def filter_srd(df, mapping):

    # Ambil nama kolom dari konfigurasi (nama di Excel mungkin berbeda-beda)
    c_stock = col(mapping,"SRD","stock_code")
    c_qty   = col(mapping,"SRD","qty_rcv_uop")   # qty yang diterima (Unit of Purchase)
    c_date  = col(mapping,"SRD","creation_date")

    # Pindahkan kolom stock_code ke posisi ke-2 (index 1)
    # Agar mudah dibaca saat preview — stock code selalu ada di kolom kedua
    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    # Normalisasi tanggal ke format 8 digit string (YYYYMMDD)
    # .str.zfill(8) memastikan angka pendek seperti "20240101" tidak terpotong
    df[c_date] = df[c_date].astype(str).str.zfill(8)

    # Urutkan dari tanggal terbaru → terlama (descending)
    # Penting agar saat groupby nanti, .first() mengambil data TERBARU
    df = df.sort_values(by=c_date, ascending=False)

    # Konversi qty ke numerik (antisipasi jika ada teks/NaN dari Excel)
    df[c_qty] = pd.to_numeric(df[c_qty], errors='coerce')

    # Hanya ambil baris dengan qty > 0 (buang return/koreksi negatif)
    df = df[df[c_qty] > 0]

    df.reset_index(drop=True, inplace=True)
    return df


# =============================================================================
# BAGIAN 4: FILTER SLN (Stock Line / Pengeluaran Barang)
# SLN = dokumen pengeluaran barang dari gudang ke user/lapangan.
# Ini mencatat histori konsumsi barang.
#
# ATURAN FILTER:
#   - Tanggal last_acq_date dinormalisasi ke 8 digit
#   - Diurutkan dari tanggal terbaru (agar groupby ambil data terkini)
#   - Hanya baris dengan qty_req > 0 DAN qty_issued > 0
#     (permintaan yang benar-benar sudah dikeluarkan barangnya)
# =============================================================================

def filter_sln(df, mapping):

    c_stock        = col(mapping,"SLN","stock_code")
    c_qty_issued   = col(mapping,"SLN","qty_issued")    # qty yang sudah keluar
    c_qty_req      = col(mapping,"SLN","qty_req")       # qty yang diminta
    c_last_acq_date= col(mapping,"SLN","last_acq_date") # tanggal terakhir akuisisi

    # Normalisasi kolom tanggal (hanya last_acq_date yang dipakai sebagai sort key)
    date_cols = [c_last_acq_date]
    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)

    # Sort terbaru dulu agar saat groupby.first() kita dapat data paling baru
    df = df.sort_values(by=c_last_acq_date, ascending=False)

    # Reorder kolom — stock_code ke posisi ke-2
    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    # Konversi qty ke numerik
    df[c_qty_issued] = pd.to_numeric(df[c_qty_issued], errors='coerce')
    df[c_qty_req]    = pd.to_numeric(df[c_qty_req], errors='coerce')

    # Filter: hanya baris yang BENAR-BENAR dikeluarkan (req > 0 DAN issued > 0)
    # Baris dengan issued == 0 artinya permintaan ada tapi belum dilayani → bukan histori valid
    df = df[(df[c_qty_req] > 0) & (df[c_qty_issued] > 0)]

    df.reset_index(drop=True, inplace=True)
    return df


# =============================================================================
# BAGIAN 5: FILTER IR (Item Request / Permintaan Barang Mendatang)
# IR = dokumen permintaan barang yang BELUM dipenuhi di masa depan.
# Ini digunakan untuk memprediksi kebutuhan barang ke depan.
#
# ATURAN FILTER (lebih ketat dari SLN):
#   - Hanya req_by_date >= hari ini (permintaan yang AKAN datang, bukan masa lalu)
#   - qty_issued HARUS == 0 (belum dilayani sama sekali — masih open request)
#   - qty_req > 0
#   - Diurutkan dari tanggal paling awal (yang paling urgent duluan)
# =============================================================================

def filter_ir(df, mapping):

    c_stock      = col(mapping,"IR","stock_code")
    c_qty_req    = col(mapping,"IR","qty_req")
    c_qty_issued = col(mapping,"IR","qty_issued")
    c_req_date   = col(mapping,"IR","req_by_date")  # tanggal permintaan harus dipenuhi

    date_cols = [c_req_date]
    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)

    # Sort ASCENDING → yang paling cepat jatuh tempo di atas
    df = df.sort_values(by=c_req_date, ascending=True)

    df[c_qty_req]    = pd.to_numeric(df[c_qty_req], errors='coerce')
    df[c_qty_issued] = pd.to_numeric(df[c_qty_issued], errors='coerce')

    # Ambil tanggal hari ini dalam format YYYYMMDD (sama dengan format kolom)
    today = pd.Timestamp.today().strftime("%Y%m%d")

    # Filter tiga kondisi sekaligus:
    #   1. qty_req > 0              → ada permintaan nyata
    #   2. qty_issued == 0          → belum dipenuhi (masih open)
    #   3. req_by_date >= today     → tanggal jatuh tempo di masa depan
    df = df[
        (df[c_qty_req] > 0) &
        (df[c_qty_issued] == 0) &
        (df[c_req_date] >= today)
    ]

    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    df.reset_index(drop=True, inplace=True)
    return df


# =============================================================================
# BAGIAN 6: FILTER PO (Purchase Order)
# PO = order pembelian yang sudah dibuat ke supplier, tapi belum selesai diterima.
# Ini mencatat barang yang "sedang dalam perjalanan" (on-order).
#
# ATURAN FILTER:
#   - receipt_status == 2  → PO yang sudah partially received atau open
#     (nilai 2 = status kode dari sistem ERP, artinya masih aktif)
#   - qty_rcv_dir != 0     → ada quantity yang relevan
#   - curr_qty != 0        → sisa quantity di PO masih ada
# =============================================================================

def filter_po(df, mapping):

    c_stock          = col(mapping,"PO","stock_code")
    c_order_date     = col(mapping,"PO","order_date")
    c_receipt_status = col(mapping,"PO","receipt_status")
    c_qty_rcv_dir    = col(mapping,"PO","qty_rcv_dir")   # qty yang sudah diterima langsung
    c_curr_qty       = col(mapping,"PO","curr_qty")      # sisa qty yang belum diterima

    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    date_cols = [c_order_date]
    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)

    # Sort terbaru dulu (agar .first() pada groupby mendapat PO terbaru)
    df = df.sort_values(by=c_order_date, ascending=False)

    # Konversi kolom status dan qty ke numerik
    for c in [c_receipt_status, c_qty_rcv_dir, c_curr_qty]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # Filter: PO aktif (status=2) dengan qty valid
    df = df[
        (df[c_receipt_status] == 2) &  # masih open/partially received
        (df[c_qty_rcv_dir] != 0) &     # ada qty yang relevan
        (df[c_curr_qty] != 0)          # masih ada sisa qty
    ]

    df.reset_index(drop=True, inplace=True)
    return df


# =============================================================================
# BAGIAN 7: FILTER LEVERING (Pengiriman dari Supplier yang Akan Datang)
# LEVERING = barang yang sudah dipesan dan dijadwalkan akan dikirim supplier,
# tapi belum sampai ke gudang. Ini yang disebut "incoming stock".
#
# ATURAN FILTER (paling kompleks):
#   - receipt_status == 0  → belum diterima
#   - curr_qty_p != 0      → ada qty yang akan datang
#   - due_date >= (hari ini - 60 hari)
#     → ambil levering yang jatuh tempo dalam 60 hari ke belakang atau ke depan
#     → batas 60 hari ke belakang untuk mengakomodasi keterlambatan pengiriman
# =============================================================================

def filter_levering(df, mapping):

    c_stock          = col(mapping,"LEVERING","stock_code")
    c_due_date       = col(mapping,"LEVERING","levering_date")  # tanggal rencana tiba
    c_receipt_status = col(mapping,"LEVERING","receipt_status")
    c_curr_qty       = col(mapping,"LEVERING","curr_qty_p")     # qty yang akan datang

    df = df.copy()  # hindari SettingWithCopyWarning

    if c_stock in df.columns:
        cols = list(df.columns)
        cols.insert(1, cols.pop(cols.index(c_stock)))
        df = df[cols]

    for c in [c_receipt_status, c_curr_qty]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # Konversi due_date dari string YYYYMMDD ke Timestamp (untuk perbandingan tanggal)
    if c_due_date in df.columns:
        df[c_due_date] = pd.to_datetime(
            df[c_due_date].astype(str),
            format="%Y%m%d",
            errors="coerce"   # tanggal tidak valid → NaT (bukan crash)
        )

    # Batas bawah = 60 hari yang lalu (normalize() = hapus komponen waktu, jam=00:00)
    batas_due = pd.Timestamp.today().normalize() - pd.Timedelta(days=60)

    df = df[
        (df[c_receipt_status] == 0) &   # belum diterima
        (df[c_curr_qty] != 0) &          # ada qty
        (df[c_due_date] >= batas_due)   # tidak terlalu lama lewat
    ]

    # Sort dari due_date paling awal (yang paling urgent)
    df = df.sort_values(by=c_due_date, ascending=True)
    df.reset_index(drop=True, inplace=True)

    # Konversi kembali ke string YYYYMMDD untuk konsistensi output
    if c_due_date in df.columns:
        df[c_due_date] = df[c_due_date].dt.strftime("%Y%m%d")

    return df


# =============================================================================
# BAGIAN 8: FILTER ALL — Entry Point Pemrosesan
# Fungsi ini adalah gerbang masuk utama. Dipanggil dari main.py setelah
# semua file Excel berhasil dibaca dan data tersedia dalam dict `data`.
#
# `data` adalah dictionary berisi DataFrame untuk setiap sheet:
#   data["SRD"]      → DataFrame hasil baca Excel SRD
#   data["SLN"]      → DataFrame hasil baca Excel SLN
#   data["IR"]       → DataFrame hasil baca Excel IR
#   data["PO"]       → DataFrame hasil baca Excel PO
#   data["LEVERING"] → DataFrame hasil baca Excel LEVERING
#   data["PLJM01"]   → DataFrame tabel master (berisi ROP & ROQ per item)
#   data["PLJM08"]   → DataFrame tabel utama inventory (SOH, item info, dll)
#
# URUTAN PROSES:
#   1. Filter semua data sumber
#   2. VLOOKUP ROP/ROQ dari PLJM01 ke PLJM08
#   3. Gabungkan data dari semua sumber ke PLJM08 (proses_pjm08)
#   4. Hitung klasifikasi ORDER/TIDAK ORDER/PERLU REVIEW
#   5. Buat tabel ANALISIS SETTING dan NON SETTING
# =============================================================================

def filter_all(data, mapping):

    # Step 1: Filter semua data sumber secara berurutan
    data["SRD"]      = filter_srd(data["SRD"], mapping)
    data["SLN"]      = filter_sln(data["SLN"], mapping)
    data["IR"]       = filter_ir(data["IR"], mapping)
    data["PO"]       = filter_po(data["PO"], mapping)
    data["LEVERING"] = filter_levering(data["LEVERING"], mapping)

    # Step 2: Tambahkan kolom ROP dan ROQ ke PLJM08 dari PLJM01 (seperti VLOOKUP)
    data = proses_pljm01(data, mapping)

    # Bersihkan kolom PLJM08 yang seluruhnya kosong (NaN/None)
    # Ini penting agar kolom "hantu" dari Excel tidak mengacaukan proses
    data["PLJM08"] = data["PLJM08"].dropna(axis=1, how="all")

    # Step 3: Gabungkan semua data sumber ke PLJM08 (mapping qty, tanggal, supplier)
    data = proses_pjm08(data, mapping)

    # Step 4 & 5: Hitung analisis dan klasifikasi
    data = analisis_setting(data, mapping)
    data = analisis_non_setting(data, mapping)

    return data


# =============================================================================
# BAGIAN 9: VLOOKUP PLJM01 → PLJM08
# PLJM01 = tabel master parameter inventory (berisi ROP dan ROQ per item)
# PLJM08 = tabel utama yang berisi semua item inventory dengan SOH-nya
#
# Fungsi ini melakukan LEFT JOIN (mirip VLOOKUP di Excel):
#   PLJM08.stock_code → cari di PLJM01.stock_code → ambil ROP dan ROQ
#
# Stock code dinormalisasi ke 9 digit (zfill) agar "000123456" == "123456"
# tidak menyebabkan gagal join karena format berbeda.
# =============================================================================

def proses_pljm01(data, mapping):

    # Nama kolom di masing-masing sheet (dari konfigurasi JSON)
    c01_stock = col(mapping, "PLJM01", "stock_code")
    c01_rop   = col(mapping, "PLJM01", "rop")        # Re-Order Point
    c01_roq   = col(mapping, "PLJM01", "roq")        # Re-Order Quantity
    c08_stock = col(mapping, "PLJM08", "stock_code")

    df01 = data["PLJM01"].copy()
    df08 = data["PLJM08"].copy()

    # Normalisasi stock code ke 9 digit dengan zero-padding
    # Contoh: "12345" → "000012345" agar bisa match dengan "000012345" di sheet lain
    df01[c01_stock] = df01[c01_stock].astype(str).str.zfill(9)
    df08[c08_stock] = df08[c08_stock].astype(str).str.zfill(9)

    # Ambil hanya 3 kolom dari PLJM01 (tidak perlu kolom lain)
    # Rename ke "ROP" dan "ROQ" agar nama kolom seragam di seluruh program
    df01_slim = df01[[c01_stock, c01_rop, c01_roq]].copy()
    df01_slim = df01_slim.rename(columns={
        c01_rop: "ROP",
        c01_roq: "ROQ"
    })

    # LEFT JOIN: semua baris PLJM08 tetap ada
    # Item yang tidak ditemukan di PLJM01 akan punya ROP=NaN, ROQ=NaN
    df08 = df08.merge(
        df01_slim,
        left_on=c08_stock,    # kunci dari PLJM08
        right_on=c01_stock,   # kunci dari PLJM01
        how="left"
    )

    # Hapus kolom stock_code duplikat dari PLJM01 (jika nama kolomnya berbeda)
    # Setelah merge, ada dua kolom stock_code — kita hanya perlu milik PLJM08
    if c01_stock != c08_stock and c01_stock in df08.columns:
        df08 = df08.drop(columns=[c01_stock])

    data["PLJM08"] = df08
    return data


# =============================================================================
# BAGIAN 10: PROSES PLJM08 — Penggabungan Semua Data ke Tabel Utama
# Ini adalah inti dari seluruh pemrosesan. Fungsi ini menambahkan kolom-kolom
# baru ke PLJM08 berdasarkan data dari SLN, IR, SRD, LEVERING, dan PO.
#
# TEKNIK: Groupby + Map
#   Alih-alih merge (join) berulang yang lambat, digunakan teknik:
#   1. groupby(stock_code).agg() → hasilkan Series berindeks stock_code
#   2. df.map(series) → petakan nilai ke setiap baris PLJM08
#   Teknik ini jauh lebih cepat untuk data besar.
#
# KOLOM YANG DITAMBAHKAN KE PLJM08:
#   - Qty ISS       ← dari SLN (total qty yang pernah dikeluarkan)
#   - Last ISS      ← dari SLN (tanggal terakhir pengeluaran)
#   - Next Req Qty  ← dari IR  (total qty permintaan mendatang)
#   - Next Req Date ← dari IR  (tanggal permintaan paling awal)
#   - Qty SRD       ← dari SRD (qty penerimaan terakhir)
#   - Last SRD      ← dari SRD (tanggal penerimaan terakhir)
#   - levering qty  ← dari LEVERING (qty yang akan datang)
#   - Levering Date ← dari LEVERING (tanggal tiba)
#   - Suplier Name  ← dari PO  (nama supplier terakhir)
# =============================================================================

def proses_pjm08(data, mapping):

    # Ambil nama kolom stock_code dari semua sheet
    c_stock_pljm08  = col(mapping,"PLJM08","stock_code")
    c_stock_sln     = col(mapping,"SLN","stock_code")
    c_stock_ir      = col(mapping,"IR","stock_code")
    c_stock_srd     = col(mapping,"SRD","stock_code")
    c_stock_po      = col(mapping,"PO","stock_code")
    c_stock_levering= col(mapping,"LEVERING","stock_code")

    # Normalisasi semua stock_code ke 9 digit agar bisa di-map dengan benar
    data["PLJM08"][c_stock_pljm08]    = data["PLJM08"][c_stock_pljm08].astype(str).str.zfill(9)
    data["SLN"][c_stock_sln]          = data["SLN"][c_stock_sln].astype(str).str.zfill(9)
    data["IR"][c_stock_ir]            = data["IR"][c_stock_ir].astype(str).str.zfill(9)
    data["SRD"][c_stock_srd]          = data["SRD"][c_stock_srd].astype(str).str.zfill(9)
    data["PO"][c_stock_po]            = data["PO"][c_stock_po].astype(str).str.zfill(9)
    data["LEVERING"][c_stock_levering]= data["LEVERING"][c_stock_levering].astype(str).str.zfill(9)

    # ── Helper: groupby yang aman ────────────────────────────────────────────
    # safe_groupby_first: ambil nilai PERTAMA per stock_code (ingat: data sudah
    # disort terbaru dulu, jadi "first" = data paling baru)
    def safe_groupby_first(df, group_col, value_col):
        # Cek keamanan: jika DataFrame kosong atau kolom tidak ada, kembalikan Series kosong
        if df.empty or group_col not in df.columns or value_col not in df.columns:
            return pd.Series(dtype='object')
        return df.groupby(group_col)[value_col].first()

    # safe_groupby_sum: jumlahkan nilai per stock_code
    # Dipakai untuk IR (menjumlahkan semua next req qty dari satu item)
    # dan LEVERING (menjumlahkan semua qty levering dari satu item)
    def safe_groupby_sum(df, group_col, value_col):
        if df.empty or group_col not in df.columns or value_col not in df.columns:
            return pd.Series(dtype='float64')
        return df.groupby(group_col)[value_col].sum()

    # ── Ambil nama kolom yang akan diagregasi ────────────────────────────────
    qty_issued    = col(mapping, "SLN", "qty_issued")
    last_acq_date = col(mapping, "SLN", "last_acq_date")
    qty_req       = col(mapping, "IR", "qty_req")
    req_by_date   = col(mapping, "IR", "req_by_date")
    qty_rcv_uop   = col(mapping, "SRD", "qty_rcv_uop")
    creation_date = col(mapping, "SRD", "creation_date")
    curr_qty_p    = col(mapping, "LEVERING", "curr_qty_p")
    levering_date = col(mapping, "LEVERING", "levering_date")
    supplier_name = col(mapping, "PO", "supplier_name")

    # ── Buat lookup Series dari setiap sumber data ───────────────────────────
    # Hasilnya adalah Series dengan index = stock_code dan value = nilai yang dicari
    # Contoh: map_qty_issued["000012345"] → 50.0

    # Dari SLN: qty pengeluaran terakhir & tanggal terakhir
    map_qty_issued    = safe_groupby_first(data["SLN"], c_stock_sln, qty_issued)
    map_last_acq_date = safe_groupby_first(data["SLN"], c_stock_sln, last_acq_date)

    # Dari IR: TOTAL next request qty (sum semua baris) & tanggal paling awal
    map_qty_req    = safe_groupby_sum(data["IR"], c_stock_ir, qty_req)
    map_req_by_date= safe_groupby_first(data["IR"], c_stock_ir, req_by_date)

    # Dari SRD: qty penerimaan & tanggal penerimaan terakhir
    map_qty_rcv_uop  = safe_groupby_first(data["SRD"], c_stock_srd, qty_rcv_uop)
    map_creation_date= safe_groupby_first(data["SRD"], c_stock_srd, creation_date)

    # Dari LEVERING: TOTAL qty yang akan datang & tanggal levering paling awal
    map_curr_qty_p   = safe_groupby_sum(data["LEVERING"], c_stock_levering, curr_qty_p)
    map_levering_date= safe_groupby_first(data["LEVERING"], c_stock_levering, levering_date)

    # Dari PO: nama supplier terbaru
    map_supplier_name= safe_groupby_first(data["PO"], c_stock_po, supplier_name)

    # ── Petakan (map) lookup Series ke kolom baru di PLJM08 ─────────────────
    # .map(series) = untuk setiap stock_code di PLJM08, cari nilainya di series
    # Jika stock_code tidak ada di series → hasilnya NaN (akan diisi 0 di bawah)
    data["PLJM08"]['Qty ISS']      = data["PLJM08"][c_stock_pljm08].map(map_qty_issued)
    data["PLJM08"]['Last ISS']     = data["PLJM08"][c_stock_pljm08].map(map_last_acq_date)
    data["PLJM08"]['Next Req Qty'] = data["PLJM08"][c_stock_pljm08].map(map_qty_req)
    data["PLJM08"]['Next Req Date']= data["PLJM08"][c_stock_pljm08].map(map_req_by_date)
    data["PLJM08"]['Qty SRD']      = data["PLJM08"][c_stock_pljm08].map(map_qty_rcv_uop)
    data["PLJM08"]['Last SRD']     = data["PLJM08"][c_stock_pljm08].map(map_creation_date)
    data["PLJM08"]['levering qty'] = data["PLJM08"][c_stock_pljm08].map(map_curr_qty_p)
    data["PLJM08"]['Levering Date']= data["PLJM08"][c_stock_pljm08].map(map_levering_date)
    data["PLJM08"]['Suplier Name'] = data["PLJM08"][c_stock_pljm08].map(map_supplier_name)

    # ── Isi NaN dengan 0 untuk kolom numerik ────────────────────────────────
    # Item yang tidak ada di SLN/IR/SRD/LEVERING akan punya NaN → ganti 0
    num_cols = ['Qty ISS','Next Req Qty','Qty SRD','levering qty']
    data["PLJM08"][num_cols] = data["PLJM08"][num_cols].fillna(0)

    # ── Pastikan ROP dan ROQ numerik (hasil merge mungkin ada NaN) ───────────
    c_soh   = col(mapping, "PLJM08", "soh_akhir")  # Stock on Hand akhir
    rop_col = "ROP"
    roq_col = "ROQ"

    data["PLJM08"][rop_col] = pd.to_numeric(data["PLJM08"][rop_col], errors='coerce').fillna(0)
    data["PLJM08"][roq_col] = pd.to_numeric(data["PLJM08"][roq_col], errors='coerce').fillna(0)

    # ── KALKULASI UTAMA: Proyeksi Stock ──────────────────────────────────────
    # Rumus: calc = SOH - Next Req Qty + Levering Qty
    # Artinya: "Berapa stock yang tersisa setelah permintaan mendatang dipenuhi,
    #           dengan mempertimbangkan barang yang akan datang dari levering?"
    # Ini adalah proyeksi stock di masa dekat.
    calc = (
        data["PLJM08"][c_soh].fillna(0)
        - data["PLJM08"]['Next Req Qty'].fillna(0)
        + data["PLJM08"]['levering qty'].fillna(0)
    )

    # ── KONDISI KLASIFIKASI ──────────────────────────────────────────────────
    # Delapan kondisi (dibuat sebagai boolean Series) untuk np.select
    cond0  = (data["PLJM08"][roq_col] == 0) & (data["PLJM08"]['Next Req Qty'] == 0)
    # → ROQ tidak di-set DAN tidak ada permintaan → tidak perlu dianalisis
    
    cond01 = (data["PLJM08"][roq_col] == 0) & (data["PLJM08"]['Next Req Qty'] != 0)
    # → ROQ tidak di-set tapi ada permintaan mendatang → perlu review manual
    
    cond1  = data["PLJM08"]['levering qty'] > (data["PLJM08"][rop_col] + data["PLJM08"][roq_col])
    # → Levering yang akan datang sudah melebihi ROP+ROQ → tidak perlu order lagi
    
    cond2  = calc > data["PLJM08"][rop_col]
    # → Proyeksi stock masih di atas ROP → aman, tidak perlu order
    
    cond3  = calc < data["PLJM08"][rop_col]
    # → Proyeksi stock di bawah ROP → HARUS ORDER
    
    cond4  = calc == data["PLJM08"][rop_col]
    # → Proyeksi stock tepat di ROP → order sebesar ROQ (tepat restock)

    # ── HITUNG QTY_RO (Quantity yang harus di-order) ─────────────────────────
    # np.select = pilih nilai berdasarkan kondisi (seperti nested IF di Excel)
    # Urutan kondisi penting — kondisi pertama yang True yang dipakai
    data["PLJM08"]['QTY_RO'] = np.select(
        [cond0, cond01, cond1, cond2, cond3, cond4],  # daftar kondisi
        [
            0,      # cond0:  tidak ada yang perlu dilakukan → QTY_RO = 0
            -1,     # cond01: ROQ belum di-set → -1 sebagai flag "perlu review"
            -1,     # cond1:  levering sudah cukup → -1 = "perlu review" (bukan order)
            0,      # cond2:  stock aman → tidak perlu order
            # cond3:  KURANG STOCK → hitung berapa yang perlu di-order
            # Formula: (ROP + ROQ) - calc
            # = Kita ingin stock mencapai ROP+ROQ, sekarang proyeksinya "calc"
            # = Kita perlu beli selisihnya
            (data["PLJM08"][rop_col] + data["PLJM08"][roq_col]) - calc,
            data["PLJM08"][roq_col]  # cond4: tepat di ROP → order 1 ROQ
        ],
        default=0   # kondisi lain yang tidak terdefinisi → 0
    )

    # ── KLASIFIKASI TEKS ─────────────────────────────────────────────────────
    # Ubah angka QTY_RO menjadi label yang mudah dibaca manusia
    data["PLJM08"]['KLASIFIKASI'] = np.select(
        [
            data["PLJM08"]['QTY_RO'] == -1,  # flag review
            data["PLJM08"]['QTY_RO'] == 0,   # tidak perlu order
            data["PLJM08"]['QTY_RO'] > 0     # perlu order
        ],
        ['PERLU REVIEW', 'TIDAK ORDER', 'ORDER'],
        default='-'
    )

    return data


# =============================================================================
# BAGIAN 11: HELPER _proses_analisis
# Fungsi generik yang dipakai untuk membuat ANALISIS SETTING maupun NON SETTING.
# Menggunakan parameter `key` dan `roq_filter` agar tidak ada duplikasi kode.
#
# LOGIKA:
#   1. Filter dari PLJM08: hanya item yang 'ORDER' atau 'PERLU REVIEW'
#   2. Filter tambahan berdasarkan ROQ (beda untuk SETTING vs NON SETTING)
#   3. Bandingkan dengan data analisis SEBELUMNYA (dari file Excel lama)
#   4. Tandai item baru / item dengan perubahan dengan label "Evaluasi"
#
# MENGAPA MEMBANDINGKAN DENGAN DATA LAMA?
#   Karena user sudah mengisi keterangan manual di file Excel sebelumnya.
#   Kita tidak mau menghapus hasil analisis mereka. Kita hanya menandai
#   mana yang berubah supaya mereka tahu perlu re-review yang mana.
# =============================================================================

def _proses_analisis(data, mapping, key, roq_filter):
    """
    Helper untuk proses analisis.
    key        : "ANALISIS SETTING" atau "ANALISIS NON SETTING"
    roq_filter : fungsi lambda untuk filter ROQ, contoh:
                 lambda df: df["ROQ"] != 0
    """
    df_src = data["PLJM08"]

    # Step 1: Ambil hanya item yang perlu dianalisis (ORDER atau PERLU REVIEW)
    df_filter = df_src[
        df_src['KLASIFIKASI'].isin(['ORDER', 'PERLU REVIEW'])
    ].copy()

    # Step 2: Filter tambahan berdasarkan ROQ
    # SETTING: ROQ != 0 (item yang sudah punya parameter tetap)
    # NON SETTING: ROQ == 0 (item belum di-setting parameter-nya)
    df_filter = df_filter[roq_filter(df_filter)].copy()

    # Step 3: Susun tabel analisis baru dengan kolom yang diperlukan
    df_analisis = pd.DataFrame()
    df_analisis["Distric"]      = df_filter[col(mapping, "PLJM08", "distric")]
    df_analisis["Suplier Name"] = df_filter["Suplier Name"]   # dari hasil map PO
    df_analisis["Stock Code"]   = df_filter[col(mapping, "PLJM08", "stock_code")]
    df_analisis["Item Name"]    = df_filter[col(mapping, "PLJM08", "item_name")]
    df_analisis["EXP"]          = df_filter[col(mapping, "PLJM08", "exp")]
    df_analisis["QTY_RO"]       = df_filter["QTY_RO"]
    df_analisis["ROP"]          = df_filter["ROP"]
    df_analisis["ROQ"]          = df_filter["ROQ"]
    df_analisis["SOH akhir"]    = df_filter[col(mapping, "PLJM08", "soh_akhir")]
    df_analisis["Keterangan"]   = df_filter["KLASIFIKASI"]

    # Step 4: Ambil data analisis lama (dari file Excel periode sebelumnya)
    map_supplier  = col(mapping, key, "supplier_name")
    map_sc        = col(mapping, key, "stock_code")
    map_item_name = col(mapping, key, "item_name")
    map_analisis  = col(mapping, key, "analisis")  # kolom keterangan manual user

    required_cols = [map_supplier, map_sc, map_item_name, map_analisis]

    # Jika data lama kosong atau tidak punya kolom yang dibutuhkan → buat DataFrame kosong
    if data[key].empty or not all(c in data[key].columns for c in required_cols):
        df_lama = pd.DataFrame(columns=["Suplier Name lama", "Stock Code", "Item Name lama", "Keterangan lama"])
    else:
        df_lama = data[key][required_cols].copy()
        df_lama = df_lama.rename(columns={
            map_supplier:  "Suplier Name lama",
            map_sc:        "Stock Code",
            map_item_name: "Item Name lama",
            map_analisis:  "Keterangan lama"  # keterangan manual dari user sebelumnya
        })

    # Normalisasi stock code 9 digit sebelum join
    df_analisis["Stock Code"] = df_analisis["Stock Code"].astype(str).str.zfill(9)
    df_lama["Stock Code"]     = df_lama["Stock Code"].astype(str).str.zfill(9)

    # Step 5: Merge data baru dengan data lama (LEFT JOIN — data baru tetap lengkap)
    df_merge = df_analisis.merge(
        df_lama,
        on="Stock Code",
        how="left",
        suffixes=("", "_lama")
    )

    # Step 6: Evaluasi perubahan — tandai item yang perlu perhatian khusus
    # Ini adalah fitur "diff" — membandingkan kondisi sekarang vs. sebelumnya
    kondisi = [
        (df_merge["Item Name lama"].isna()) | (df_merge["Suplier Name lama"].isna()),
        # → Item baru (tidak ada di periode lama) → belum pernah dianalisis
        
        df_merge["Item Name"] != df_merge["Item Name lama"],
        # → Nama item berubah → mungkin salah mapping atau item diganti
        
        df_merge["Suplier Name"] != df_merge["Suplier Name lama"],
        # → Supplier berubah → perlu review ulang (harga/terms mungkin berbeda)
        
        df_merge["Keterangan"] != df_merge["Keterangan lama"]
        # → Status klasifikasi berubah (contoh: dulu TIDAK ORDER, sekarang ORDER)
    ]
    hasil = [
        "Perlu di analisis",
        "Item name order baru",
        "Beda supplier",
        "Status analisis berbeda"
    ]

    # np.select: pilih label pertama yang kondisinya True
    # default = "telah di analisis" → item yang tidak berubah sama sekali
    df_merge["Evaluasi"] = np.select(kondisi, hasil, default="telah di analisis")

    # Tampilkan keterangan lama sebagai referensi user
    df_merge["Review Sebelumnya"] = df_merge["Keterangan lama"].fillna("-")

    # Hapus kolom sementara yang tidak perlu ditampilkan ke user
    df_merge = df_merge.drop(
        columns=["Suplier Name lama", "Item Name lama", "Keterangan lama"],
        errors="ignore"
    )

    # Sort berdasarkan nama supplier agar mudah dikelompokkan saat review
    df_merge = df_merge.sort_values(by="Suplier Name", ascending=True, ignore_index=True)

    # Atur urutan kolom: Review Sebelumnya → Keterangan → Evaluasi
    cols = list(df_merge.columns)
    cols.insert(cols.index("Keterangan"), cols.pop(cols.index("Review Sebelumnya")))
    cols.insert(cols.index("Evaluasi"), cols.pop(cols.index("Keterangan")))
    df_merge = df_merge[cols]

    data[key] = df_merge
    return data


# =============================================================================
# BAGIAN 12: ANALISIS SETTING
# Item "SETTING" = item yang sudah punya ROP dan ROQ yang di-set (ROQ != 0).
# Ini adalah item yang sudah dikelola secara sistematis dengan parameter resmi.
#
# Langsung mendelegasikan ke _proses_analisis dengan filter ROQ != 0
# =============================================================================

def analisis_setting(data, mapping):
    return _proses_analisis(
        data, mapping,
        key="ANALISIS SETTING",
        roq_filter=lambda df: df["ROQ"] != 0  # hanya item yang punya ROQ
    )


# =============================================================================
# BAGIAN 13: ANALISIS NON SETTING
# Item "NON SETTING" = item yang belum punya ROQ (ROQ == 0), tapi ada
# permintaan mendatang (Next Req Date >= hari ini).
# Ini adalah item yang perlu diorder berdasarkan demand langsung, bukan parameter.
#
# Kolom tambahan khusus NON SETTING: IR Qty dan IR Date
# (karena item ini tidak punya ROQ, keputusan order murni dari IR)
# =============================================================================

def analisis_non_setting(data, mapping):

    df_src = data["PLJM08"]
    today_str = pd.Timestamp.today().strftime("%Y%m%d")

    # Filter: ORDER atau PERLU REVIEW + ROQ == 0 + ada Next Req Date yang valid
    df_filter = df_src[
        df_src['KLASIFIKASI'].isin(['ORDER', 'PERLU REVIEW'])
    ].copy()

    # Hanya item yang belum punya ROQ (Non Setting)
    df_filter = df_filter[df_filter["ROQ"] == 0].copy()

    # Hanya item yang punya permintaan mendatang (Next Req Date >= hari ini)
    # Kondisi ganda: notna() cek ada nilai, astype(str).strip() >= today cek tanggal
    df_filter = df_filter[
        df_filter["Next Req Date"].notna() &
        (df_filter["Next Req Date"].astype(str).str.strip() >= today_str)
    ].copy()

    # Susun tabel analisis — mirip ANALISIS SETTING tapi ada tambahan IR Qty & IR Date
    df_analisis = pd.DataFrame(index=df_filter.index)
    df_analisis["Distric"]      = df_filter[col(mapping, "PLJM08", "distric")]
    df_analisis["Suplier Name"] = df_filter["Suplier Name"]
    df_analisis["Stock Code"]   = df_filter[col(mapping, "PLJM08", "stock_code")]
    df_analisis["Item Name"]    = df_filter[col(mapping, "PLJM08", "item_name")]
    df_analisis["EXP"]          = df_filter[col(mapping, "PLJM08", "exp")]
    df_analisis["QTY_RO"]       = df_filter["QTY_RO"]
    df_analisis["ROP"]          = df_filter["ROP"]
    df_analisis["ROQ"]          = df_filter["ROQ"]
    df_analisis["SOH akhir"]    = df_filter[col(mapping, "PLJM08", "soh_akhir")]
    df_analisis["IR Qty"]       = df_filter["Next Req Qty"]   # qty dari IR
    df_analisis["IR Date"]      = df_filter["Next Req Date"]  # tanggal jatuh tempo dari IR

    # Ambil keterangan dari data analisis non-setting lama (review sebelumnya)
    key = "ANALISIS NON SETTING"
    map_sc       = col(mapping, key, "stock_code")
    map_analisis = col(mapping, key, "analisis")

    required_cols = [map_sc, map_analisis]
    if data[key].empty or not all(c in data[key].columns for c in required_cols):
        df_lama = pd.DataFrame(columns=["Stock Code", "Keterangan lama"])
    else:
        df_lama = data[key][required_cols].copy()
        df_lama = df_lama.rename(columns={
            map_sc:       "Stock Code",
            map_analisis: "Keterangan lama"
        })

    # Normalisasi dan merge dengan data lama
    df_analisis["Stock Code"] = df_analisis["Stock Code"].astype(str).str.zfill(9)
    df_lama["Stock Code"]     = df_lama["Stock Code"].astype(str).str.zfill(9)

    df_merge = df_analisis.merge(df_lama, on="Stock Code", how="left")

    # Kolom Keterangan diisi dari review sebelumnya (jika ada), default "-"
    # NON SETTING tidak punya kolom "Evaluasi" seperti SETTING
    df_merge["Keterangan"] = df_merge["Keterangan lama"].fillna("-")
    df_merge = df_merge.drop(columns=["Keterangan lama"], errors="ignore")

    df_merge = df_merge.sort_values(by="Suplier Name", ascending=True, ignore_index=True)

    data[key] = df_merge

    # Debug print: untuk verifikasi saat development
    print(f"ANALISIS NON SETTING: {len(df_merge)} baris")
    print(df_merge)

    return data