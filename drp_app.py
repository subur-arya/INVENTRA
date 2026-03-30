import re
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import os, sys, threading
from PIL import Image, ImageTk
from proses_GSheet import (
    sync_ke_gsheet, fetch_sheet, merge_with_existing, WEB_APP_URL,
    SheetTidakDitemukanError, HeaderTidakDitemukanError, GSheetResponseError,
)
PRK_PATTERN = re.compile(
    r'(?:GR)?\d{2,}[A-Z]{1,2}\d{2,}'    # angka + huruf + angka
    r'|'
    r'(?<![.\d])(?:GR)?\d{6,}(?![.\d])' # angka-only ≥ 6 digit
)


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


# =====================================================
# LOGIKA
# =====================================================

def _find_col(df, col_name):
    if not col_name:
        return None
    target = col_name.strip().lower()
    return next((c for c in df.columns if str(c).strip().lower() == target), None)


def _baca_sheet(xl, sheet, col_no_prk):
    try:
        temp = pd.read_excel(xl, sheet_name=sheet, header=None)
    except Exception:
        return None
    col_lower = col_no_prk.strip().lower()
    header_row = next(
        (i for i, row in temp.iterrows()
         if row.astype(str).str.strip().str.lower().str.contains(col_lower, na=False).any()),
        None
    )
    if header_row is None:
        return None
    df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
    df.columns = df.columns.str.strip().str.replace("\u00A0", " ", regex=False)
    return df


def _merge_by_no_prk(df, col_no_prk):
    actual = _find_col(df, col_no_prk)
    if actual is None:
        return df

    def agg_teks(series):
        vals = series.dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        unique_vals = vals.unique()
        if len(unique_vals) == 0:
            return ""
        return unique_vals[0] if len(unique_vals) == 1 else ", ".join(unique_vals)

    num_cols = df.select_dtypes(include="number").columns.tolist()
    agg = {c: "first" if c in num_cols else agg_teks for c in df.columns if c != actual}
    grouped = df.groupby(actual, as_index=False, sort=False).agg(agg)
    return grouped[[c for c in df.columns if c in grouped.columns]]


def _detect_header_row(temp, keywords=None, min_cols=5):
    if keywords:
        for i, row in temp.iterrows():
            vals = row.astype(str).str.strip().str.lower()
            if any(vals.str.contains(kw, na=False).any() for kw in keywords):
                if row.notna().sum() >= min_cols:
                    return i
    return next((i for i, row in temp.iterrows() if row.notna().sum() >= min_cols), None)


def _baca_excel_sheet(xl, sheet_name):
    try:
        temp = pd.read_excel(xl, sheet_name=sheet_name, header=None)
    except Exception:
        return None
    header_row = _detect_header_row(
        temp, keywords=["no requisisi", "peruntukan", "stockcode"]
    )
    if header_row is None:
        return None
    df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.str.strip().str.replace("\u00A0", " ", regex=False)
    return df


def _baca_amp(xl):
    sheet_upper = {s.strip().upper(): s for s in xl.sheet_names}
    sheet_amp = sheet_upper.get("AMP")
    return _baca_excel_sheet(xl, sheet_amp) if sheet_amp else None


def _baca_amp_by_name(xl, sheet_name):
    return _baca_excel_sheet(xl, sheet_name)


def _baca_drp_sheet(xl, sheet_name, auto_name="DRP"):
    if sheet_name is None:
        sheet_upper = {s.strip().upper(): s for s in xl.sheet_names}
        sheet_name  = sheet_upper.get(auto_name)
    if sheet_name is None:
        return None
    try:
        temp = pd.read_excel(xl, sheet_name=sheet_name, header=None)
    except Exception:
        return None
    header_row = _detect_header_row(temp, keywords=["nomor prk"])
    if header_row is None:
        return None
    df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row, dtype=str)
    df.columns = df.columns.str.strip().str.replace("\u00A0", " ", regex=False)
    prk_col = _find_col(df, "NOMOR PRK")
    if prk_col:
        df = df[
            df[prk_col].notna() &
            (df[prk_col].astype(str).str.strip() != "") &
            (df[prk_col].astype(str).str.strip().str.upper() != "NOMOR PRK")
        ]
    return df.fillna("")


def klasifikasi_prk(prk):
    """
    Klasifikasi PRK berdasarkan huruf di tengah nomor PRK.

    Aturan:
    - Lebih dari 1 nomor PRK dalam satu cell          → PRK GABUNGAN
    - Format GR + 2 digit bebas + 4 + A + sisa        → PRK I  (AI)
      (huruf pertama = A, 3 digit sebelumnya,
       digit terakhir sebelum A harus '4')
    - Tidak ada huruf sama sekali dalam nomor          → Tanpa PRK
    - Semua PRK berhuruf lainnya                      → AO

    Contoh:
        GR254A0205              → PRK I
        GR214A0101GR244A0109    → PRK GABUNGAN
        GR253A0205              → AO
        121046 / 10606          → Tanpa PRK (tidak ada huruf, tidak dikasih GR)
    """

    prk_str = str(prk).upper()

    semua = PRK_PATTERN.findall(prk_str)

    if len(semua) > 1:
        return "PRK GABUNGAN"

    if len(semua) == 1:
        p = semua[0]
        body = p[2:] if p.startswith("GR") else p
        match = re.search(r'([A-Z])', body)
        if not match:
            return "Tanpa PRK"
        digit_sebelum = body[:match.start()]
        huruf         = match.group(1)
        if huruf == "A" and len(digit_sebelum) == 3 and digit_sebelum[-1] == "4":
            return "PRK I"
        else:
            return "AO"

    # Tidak cocok PRK_PATTERN — cek ada huruf atau tidak
    if not re.search(r'[A-Z]', prk_str):
        return "Tanpa PRK"

    return "AO"

def generate_drp_from_amp_df(amp_df, col_mapping):
    if amp_df is None:
        return None

    amp_df = amp_df.copy()
    amp_df.columns = amp_df.columns.str.strip()

    def _cm(key):
        return col_mapping.get(key, {}).get("amp", "")

    col_prk           = _find_col(amp_df, _cm("col_no_prk"))
    col_item          = _find_col(amp_df, _cm("col_item"))
    col_tgl_ro        = _find_col(amp_df, _cm("col_tgl_ro"))
    col_no_ro         = _find_col(amp_df, _cm("col_no_ro"))
    col_hpe           = _find_col(amp_df, _cm("col_nilai_hpe"))
    col_kontrak       = _find_col(amp_df, _cm("col_kontrak"))
    col_levering      = _find_col(amp_df, _cm("col_levering"))
    col_stockcode     = _find_col(amp_df, _cm("col_stockcode"))
    col_tgl_kontrak   = _find_col(amp_df, _cm("col_tgl_kontrak"))
    col_tgl_selesai   = _find_col(amp_df, _cm("col_tgl_selesai"))
    col_no_requisisi  = _find_col(amp_df, _cm("col_no_requisisi"))
    col_satuan        = _find_col(amp_df, _cm("col_satuan"))
    col_Vol           = _find_col(amp_df, _cm("col_Vol"))
    col_nilai_kontrak = _find_col(amp_df, _cm("col_nilai_kontrak"))


    if col_prk is None:
        raise ValueError("Kolom PRK tidak ditemukan di sheet AMP")
    if col_item is None:
        raise ValueError("Kolom Item tidak ditemukan di sheet AMP")

    def extract_prk_item(text):
        """
        Ekstrak nomor PRK dan nama item dari satu sel teks.

        Aturan:
        - Pola PRK: (GR)? + digit(2+) + huruf(1-2) + digit(2+)
                    ATAU digit(6+) berdiri sendiri
        - Jika tidak ada pola PRK → seluruh teks dikembalikan sebagai PRK, item kosong
        - Semua PRK dinormalisasi dengan prefix GR
        - Sisa teks setelah PRK dihapus = item, separator dibersihkan

        Return: (prk_string, item_string)
        """
        if text is None or (isinstance(text, float) and pd.isna(text)):
            return "", ""

        text = str(text).replace("\u00A0", " ").replace("\t", " ").strip().upper()

        if not text or text == "NAN":
            return "", ""

        matches = PRK_PATTERN.findall(text)

        if not matches:
            # Tidak ada pola PRK → seluruh teks dianggap sebagai PRK
            return text, ""

        # Normalisasi: tambah prefix GR hanya jika ada huruf di dalamnya
        # Angka-only (seperti 121046, 10606) tidak perlu prefix GR
        def normalize(p):
            if p.startswith("GR"):
                return p
            if re.search(r'[A-Z]', p):
                return "GR" + p  # ada huruf → tambah GR
            return p              # angka-only → biarkan apa adanya

        prk = " ".join(normalize(p) for p in matches)

        # Ekstrak item: hapus semua PRK dari teks, bersihkan sisa separator
        item = text
        for p in matches:
            item = item.replace(p, "")
        item = re.sub(r'^[\s_\-/,:.]+|[\s_\-/,:.]+$', '', item)
        item = re.sub(r'\s{2,}', ' ', item).strip()

        return prk, item
 
    _prk_item_pairs = [extract_prk_item(x) for x in amp_df[col_prk]]
    amp_df["PRK"]  = [p[0] for p in _prk_item_pairs]
    amp_df["ITEM"] = [p[1] for p in _prk_item_pairs]

    # NORMALISASI PRK
    def normalize_prk(prk):

        if pd.isna(prk):
            return ""

        prk = str(prk).upper().strip()

        parts = prk.split()

        hasil = []

        for p in parts:

           if re.match(r"(?:GR[0-9A-Z]{7}|\d{3}[A-Z]\d{4}|\d{4}[A-Z]{2}\d{2}|\d{8}|\d{7})", p):
                # Tambah prefix GR hanya jika ada huruf (bukan angka-only)
                if not p.startswith("GR") and re.search(r'[A-Z]', p):
                    p = "GR" + p

                hasil.append(p)

        return " ".join(hasil)



    amp_df["PRK"] = [normalize_prk(v) for v in amp_df["PRK"]]

    amp_df = amp_df[
        (amp_df["PRK"] != "") &
        (amp_df["PRK"].str.upper() != "NAN") &
        (amp_df["PRK"].str.len() > 4)
    ]
    # buat kolom kategori dulu
    amp_df["Kategori PRK"] = [klasifikasi_prk(v) for v in amp_df["PRK"]]

    # baru diurutkan
    amp_df = amp_df.sort_values(["Kategori PRK", "PRK"])
    
    def _join(group, col):
        n = len(group)
        if not col:
            return "\n" * (n - 1)

        baris = []
        for v in group[col]:
            if pd.isna(v) or str(v).strip() == "" or str(v).strip().upper() == "NAN":
                baris.append("")
            else:
                s = str(v).strip()

                if s.endswith(".0"):
                    s = s[:-2]

                baris.append(s)

        return "\n".join(baris)
    def _join_hpe(group):
        # Iterasi langsung via kolom array — tanpa print debug, tanpa .loc[]
        baris = []
        tgl_vals = group[col_tgl_selesai].tolist() if col_tgl_selesai else [None] * len(group)
        hpe_vals = group[col_hpe].tolist()         if col_hpe         else [None] * len(group)
        for v_tgl, v_hpe in zip(tgl_vals, hpe_vals):
            tgl_str = str(v_tgl).strip() if v_tgl is not None else ""
            tgl_ok  = not (pd.isna(v_tgl) if v_tgl is not None else True) and tgl_str != "" and tgl_str.upper() != "NAN"
            if not tgl_ok or v_hpe is None:
                baris.append("")
            else:
                hpe_str = str(v_hpe).strip()
                baris.append("" if (pd.isna(v_hpe) or hpe_str == "" or hpe_str.upper() == "NAN") else hpe_str)
        return "\n".join(baris)

    def _join_nilai_dari_sumber(group):
        def _to_float(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return 0
            s = str(v).strip().replace(",", "").replace(" ", "")
            try:
                return float(s)
            except Exception:
                return 0

        nk_vals = group[col_nilai_kontrak].tolist() if col_nilai_kontrak else [None] * len(group)
        jm_vals = group[col_Vol].tolist()           if col_Vol           else [None] * len(group)
        baris = []
        for nk, jm in zip(nk_vals, jm_vals):
            nk_kosong = nk is None or (isinstance(nk, float) and pd.isna(nk)) or str(nk).strip().upper() in ("", "NAN")
            jm_kosong = jm is None or (isinstance(jm, float) and pd.isna(jm)) or str(jm).strip().upper() in ("", "NAN")
            if nk_kosong or jm_kosong:
                baris.append("")
            else:
                try:
                    baris.append(f"{_to_float(nk) * _to_float(jm):,.0f}")
                except Exception:
                    baris.append("")
        return "\n".join(baris)
    

    amp_df = amp_df.sort_values(["Kategori PRK","PRK"])
    hasil = []
    for prk, group in amp_df.groupby("PRK", sort=False):
        item_series = (
            group["ITEM"] if (not col_item or col_item == col_prk)
            else group[col_item]
        )
        hasil.append({
            "NOMOR PRK"             : prk,
            "ITEM PROSES PENGADAAN" : "\n".join(item_series.fillna("").astype(str)),
            "Stock Code"            : _join(group, col_stockcode),
            "No Requisisi"          : _join(group, col_no_requisisi),
            "Satuan"                : _join(group, col_satuan),
            "Vol"                   : _join(group, col_Vol),
            "Tanggal RO"            : _join(group, col_tgl_ro),
            "No RO"                 : _join(group, col_no_ro),
            "Nilai HPE"             : _join_hpe(group),
            "Tanggal Terkontrak"    : _join(group, col_tgl_kontrak),
            "Nomor Kontrak"         : _join(group, col_kontrak),
            "Nilai Kontrak"         : _join_nilai_dari_sumber(group),
            "Levering"              : _join(group, col_levering),
        })

    drp_df = pd.DataFrame(hasil)
    drp_df.insert(0, "NO", range(1, len(drp_df) + 1))

    # Kategori PRK: ambil dari amp_df yang sudah dihitung sebelumnya (tanpa apply ulang)
    kat_map = amp_df.groupby("PRK")["Kategori PRK"].first().to_dict()
    drp_df["Kategori PRK"] = [kat_map.get(p, klasifikasi_prk(p)) for p in drp_df["NOMOR PRK"]]

    return drp_df


def _extract_tanpa_prk(amp_df, col_mapping):
    """
    Kembalikan DataFrame TANPA PRK yang kolomnya sudah dinormalisasi
    ke nama-nama kanonik yang dipakai di preview/DRP.
    """
    if amp_df is None:
        return None

    def _cm(key):
        return col_mapping.get(key, {}).get("amp", "") if col_mapping else ""

    amp_df = amp_df.copy()
    amp_df.columns = amp_df.columns.str.strip()

    # cari kolom sumber sesuai mapping user
    col_prk           = _find_col(amp_df, _cm("col_no_prk"))
    col_item          = _find_col(amp_df, _cm("col_item"))
    col_stockcode     = _find_col(amp_df, _cm("col_stockcode"))
    col_no_requisisi  = _find_col(amp_df, _cm("col_no_requisisi"))
    col_satuan        = _find_col(amp_df, _cm("col_satuan"))
    col_Vol           = _find_col(amp_df, _cm("col_Vol"))
    col_tgl_ro        = _find_col(amp_df, _cm("col_tgl_ro"))
    col_no_ro         = _find_col(amp_df, _cm("col_no_ro"))
    col_hpe           = _find_col(amp_df, _cm("col_nilai_hpe"))
    col_tgl_kontrak   = _find_col(amp_df, _cm("col_tgl_kontrak"))
    col_kontrak       = _find_col(amp_df, _cm("col_kontrak"))
    col_nilai_kontrak = _find_col(amp_df, _cm("col_nilai_kontrak"))
    col_levering      = _find_col(amp_df, _cm("col_levering"))

    if col_prk is None:
        return None

    # keep only rows with non-empty source-peruntukan
    amp_df = amp_df[
        amp_df[col_prk].notna() &
        (amp_df[col_prk].astype(str).str.strip() != "")
    ].copy()

    # series bersih untuk cek PRK pattern
    prk_series = amp_df[col_prk].astype(str).str.replace("\u00A0", " ").str.strip()
    mask_prk_valid = prk_series.apply(lambda x: bool(PRK_PATTERN.search(str(x))))

    # baris tanpa PRK = tidak match pattern
    df_tanpa_src = amp_df[~mask_prk_valid].reset_index(drop=True)
    if df_tanpa_src.empty:
        # kembalikan df kosong dengan kolom kanonik agar preview stabil
        cols = ["NO", "Peruntukan", "Item", "Stock Code", "No Requisisi",
                "Vol", "Satuan", "Tanggal RO", "No RO",
                "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak",
                "Nilai Kontrak", "Levering"]
        return pd.DataFrame(columns=cols)

    # Susun output memakai nama kanonik (sesuai DRP)
    out = pd.DataFrame()
    out["Peruntukan"] = df_tanpa_src[col_prk].astype(str).str.replace("\u00A0", " ").str.strip() if col_prk else ""
    out["Item"] = df_tanpa_src[col_item].astype(str).str.strip() if (col_item and col_item != col_prk) else ""
    out["Stock Code"]   = df_tanpa_src[col_stockcode].astype(str).str.strip() if col_stockcode else ""
    out["No Requisisi"] = df_tanpa_src[col_no_requisisi].astype(str).str.strip() if col_no_requisisi else ""
    out["Vol"]          = df_tanpa_src[col_Vol].astype(str).str.strip() if col_Vol else ""
    out["Satuan"]       = df_tanpa_src[col_satuan].astype(str).str.strip() if col_satuan else ""
    out["Tanggal RO"]   = df_tanpa_src[col_tgl_ro].astype(str).str.strip() if col_tgl_ro else ""
    out["No RO"]        = df_tanpa_src[col_no_ro].astype(str).str.strip() if col_no_ro else ""
    out["Nilai HPE"]    = df_tanpa_src[col_hpe].astype(str).str.strip() if col_hpe else ""
    out["Tanggal Terkontrak"] = df_tanpa_src[col_tgl_kontrak].astype(str).str.strip() if col_tgl_kontrak else ""
    out["Nomor Kontrak"] = df_tanpa_src[col_kontrak].astype(str).str.strip() if col_kontrak else ""
    out["Nilai Kontrak"] = df_tanpa_src[col_nilai_kontrak].astype(str).str.strip() if col_nilai_kontrak else ""
    out["Levering"]      = df_tanpa_src[col_levering].astype(str).str.strip() if col_levering else ""

    # bersihkan literal "nan" dan spasi / nbsp
    out = out.fillna("").astype(str)
    for c in out.columns:
        out[c] = out[c].str.replace("\u00A0", " ").str.strip().replace("nan", "")

    # tambah kolom NO sebagai index 1..
    out.insert(0, "NO", range(1, len(out) + 1))

    return out
def proses_drp(file_path, col_mapping=None, sheet_amp=None):
    xl          = pd.ExcelFile(file_path)
    df_amp      = _baca_amp_by_name(xl, sheet_amp) if sheet_amp else _baca_amp(xl)
    df_drp      = generate_drp_from_amp_df(df_amp, col_mapping)
    df_tanpa    = _extract_tanpa_prk(df_amp, col_mapping)
    return {"DRP": df_drp, "AMP": df_amp, "TANPA_PRK": df_tanpa}


def merge_drp_dengan_gsheet(df_lokal, sheet_name):
    """
    Fetch data GSheet (SEMUA kolom) lalu merge dengan df_lokal.

    Aturan merge:
    - Primary key  : NOMOR PRK
    - Secondary key: No RO (per baris dalam satu PRK)
    - Hanya KOLOM_DRP yang di-update dari data lokal.
    - Kolom manual GSheet (di luar KOLOM_DRP) TIDAK disentuh.
    - PRK baru ditambahkan di bawah.
    - Pencocokan nama kolom bersifat case-insensitive.

    Return: (df_merged, df_lama)

    Exception yang bisa dilempar (diteruskan ke caller):
        SheetTidakDitemukanError  → sheet tidak ada di GSheet
        HeaderTidakDitemukanError → sheet ada tapi header NOMOR PRK tidak ditemukan
        ConnectionError           → gagal terhubung
    """
    from proses_GSheet import fetch_sheet, WEB_APP_URL, check_sheet_exists

    KOLOM_DRP = [
        "NOMOR PRK", "ITEM PROSES PENGADAAN", "Stock Code", "No Requisisi",
        "Vol", "Satuan", "Tanggal RO", "No RO",
        "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak", "Nilai Kontrak", "Levering",
    ]

    # ── Helper: cari kolom di df secara case-insensitive ─────────────────────
    def _ci(df, nama):
        """Kembalikan nama kolom asli di df yang cocok dengan nama
        (case-insensitive), atau None jika tidak ditemukan."""
        target = nama.strip().lower()
        return next((c for c in df.columns if str(c).strip().lower() == target), None)

    # ── Siapkan df_baru: rename kolom lokal ke nama kanonik KOLOM_DRP ────────
    rename_lokal = {}
    for k in KOLOM_DRP:
        actual = _ci(df_lokal, k)
        if actual and actual != k:
            rename_lokal[actual] = k
    df_baru = df_lokal.rename(columns=rename_lokal).copy()

    kolom_ada = [k for k in KOLOM_DRP if _ci(df_baru, k) is not None]
    df_baru   = df_baru[[_ci(df_baru, k) for k in kolom_ada]].copy()
    df_baru.columns = kolom_ada
    df_baru = df_baru.fillna("").astype(str)
    df_baru = df_baru[df_baru["NOMOR PRK"].str.strip() != ""].reset_index(drop=True)

    # ── Cek koneksi & keberadaan sheet terlebih dahulu ───────────────────────
    # Dilakukan sebelum fetch agar error bisa dibedakan sejak awal.
    exists = check_sheet_exists(WEB_APP_URL, sheet_name)
    if exists is None:
        raise ConnectionError(
            "Tidak dapat terhubung ke Google Sheets.\n"
            "Periksa koneksi internet Anda.")
    if not exists:
        raise SheetTidakDitemukanError(sheet_name)

    # ── Fetch GSheet — SEMUA kolom ───────────────────────────────────────────
    # fetch_sheet sekarang raise exception eksplisit:
    #   SheetTidakDitemukanError  → sheet tidak ada
    #   HeaderTidakDitemukanError → sheet ada tapi header rusak/tidak sesuai template
    #   return None               → sheet kosong (belum ada data sama sekali)
    df_lama = fetch_sheet(WEB_APP_URL, sheet_name)

    if df_lama is None:
        # Tidak ada data sama sekali — pakai df_baru apa adanya
        df_merged = df_baru.copy()
        df_merged["NOMOR PRK"] = df_merged["NOMOR PRK"].str.split("\n").str[0].str.strip()
        df_merged.insert(0, "NO", range(1, len(df_merged) + 1))
        return df_merged, df_lama

    if df_lama.empty:
        # Sheet baru/kosong dengan HEADER_DEFAULT — normalisasi kolom df_baru
        # agar cocok dengan urutan dan nama header default (case-insensitive).
        from proses_GSheet import _normalisasi_kolom
        all_cols  = list(df_lama.columns)
        df_baru   = _normalisasi_kolom(df_baru, all_cols)
        for c in df_baru.columns:
            if c not in all_cols:
                all_cols.append(c)
        df_merged = df_baru.reindex(columns=all_cols, fill_value="")
        df_merged["NOMOR PRK"] = df_merged["NOMOR PRK"].str.split("\n").str[0].str.strip()
        if "NO" in df_merged.columns:
            df_merged["NO"] = range(1, len(df_merged) + 1)
        else:
            df_merged.insert(0, "NO", range(1, len(df_merged) + 1))
        return df_merged, df_lama

    # ── Normalisasi kolom df_lama ─────────────────────────────────────────────
    df_lama.columns = df_lama.columns.str.strip().str.replace("\u00A0", " ", regex=False)

    # Rename kolom GSheet ke nama kanonik KOLOM_DRP (case-insensitive)
    rename_lama = {}
    for k in KOLOM_DRP:
        actual = _ci(df_lama, k)
        if actual and actual != k:
            rename_lama[actual] = k
    df_lama = df_lama.rename(columns=rename_lama)

    # Kolom KOLOM_DRP yang belum ada di GSheet sama sekali → tambahkan kosong
    for k in KOLOM_DRP:
        if k not in df_lama.columns:
            df_lama[k] = ""

    # Semua kolom GSheet (urutan asli dipertahankan, nama sudah ternormalisasi)
    all_gsheet_cols = list(df_lama.columns)

    # Bersihkan KOLOM_DRP di df_lama
    for k in KOLOM_DRP:
        df_lama[k] = df_lama[k].astype(str).str.lstrip("'").str.strip()

    df_lama["NOMOR PRK"] = df_lama["NOMOR PRK"].str.split("\n").str[0].str.strip()
    df_lama = df_lama[df_lama["NOMOR PRK"] != ""].reset_index(drop=True)
    df_lama = df_lama.drop_duplicates(subset=["NOMOR PRK"], keep="first").reset_index(drop=True)

    urutan_lama = {
        str(row["NOMOR PRK"]).strip().upper(): i
        for i, row in df_lama.iterrows()
    }

    # ── Helper functions ─────────────────────────────────────────────────────
    def _split_baris(val):
        return [b.strip() for b in str(val).split("\n")]

    def _norm_key(val):
        s = str(val).strip().lstrip("'").upper()
        return s if s and s != "NAN" else ""

    def _merge_nilai(v_lama, v_baru):
        lama_kosong = v_lama in ("", "nan")
        baru_kosong = v_baru in ("", "nan")
        if lama_kosong and baru_kosong:
            return ""
        elif lama_kosong:
            return v_baru
        elif baru_kosong:
            return v_lama   # lama dipertahankan jika baru kosong
        else:
            return v_baru   # baru menang jika keduanya ada

    def _sub_key(lama_cols, baru_cols, i, j):
        """
        Cek apakah sub-baris i (lama) cocok dengan sub-baris j (baru).

        Aturan:
        - No RO  : WAJIB sama jika keduanya tidak kosong; berbeda → tidak cocok.
                   Keduanya kosong → tidak bisa di-match.
        - SC     : kalau lama ada dan baru ada tapi beda → tidak cocok.
                   kalau salah satu kosong → abaikan sebagai key.
        - No Req : sama seperti SC.
        """
        def _get(cols, col, idx):
            lst = cols.get(col, [])
            return lst[idx] if idx < len(lst) else ""

        ro_l  = _norm_key(_get(lama_cols, "No RO",       i))
        ro_b  = _norm_key(_get(baru_cols, "No RO",       j))
        sc_l  = _norm_key(_get(lama_cols, "Stock Code",  i))
        sc_b  = _norm_key(_get(baru_cols, "Stock Code",  j))
        req_l = _norm_key(_get(lama_cols, "No Requisisi",i))
        req_b = _norm_key(_get(baru_cols, "No Requisisi",j))

        # No RO wajib cocok
        if ro_l and ro_b and ro_l != ro_b:
            return False
        # Keduanya kosong → tidak bisa di-match
        if not ro_l and not ro_b:
            return False

        # Stock Code: kalau lama ada, baru harus sama atau kosong
        if sc_l and sc_b and sc_l != sc_b:
            return False

        # No Requisisi: kalau lama ada, baru harus sama atau kosong
        if req_l and req_b and req_l != req_b:
            return False

        return True

    def _merge_satu_prk(row_lama, row_baru):
        """
        Merge satu PRK dengan secondary key kombinasi: No RO + Stock Code + No Requisisi.
        - No RO      : wajib cocok jika tidak kosong
        - Stock Code : fleksibel — kalau lama ada dan baru beda → tidak cocok
        - No Req     : fleksibel — sama seperti Stock Code
        Hanya kolom KOLOM_DRP yang di-merge dari baru ke lama.
        Kolom manual GSheet tidak disentuh.
        """
        KOLOM_NON_PRK = {"NOMOR PRK"}
        kolom_merge   = [c for c in KOLOM_DRP if c not in KOLOM_NON_PRK]

        lama_cols = {col: _split_baris(row_lama.get(col, "")) for col in kolom_merge}
        baru_cols = {col: _split_baris(row_baru.get(col, "")) for col in kolom_merge}

        n_lama = max((len(v) for v in lama_cols.values()), default=0)
        n_baru = max((len(v) for v in baru_cols.values()), default=0)

        # Pad agar semua kolom panjangnya sama
        for col in kolom_merge:
            while len(lama_cols[col]) < n_lama: lama_cols[col].append("")
            while len(baru_cols[col]) < n_baru: baru_cols[col].append("")

        hasil_cols   = {col: [] for col in kolom_merge}
        baru_dipakai = set()

        # ── Iterasi sub-baris LAMA ──────────────────────────────────────────
        for i in range(n_lama):
            # Cari sub-baris baru yang cocok (first match, belum dipakai)
            match_j = None
            for j in range(n_baru):
                if j in baru_dipakai:
                    continue
                if _sub_key(lama_cols, baru_cols, i, j):
                    match_j = j
                    break

            if match_j is not None:
                # Cocok → merge nilai per kolom
                baru_dipakai.add(match_j)
                for col in kolom_merge:
                    v_l = lama_cols[col][i]       if i       < len(lama_cols[col]) else ""
                    v_b = baru_cols[col][match_j] if match_j < len(baru_cols[col]) else ""
                    hasil_cols[col].append(_merge_nilai(v_l, v_b))
            else:
                # Tidak cocok → pertahankan sub-baris lama utuh
                for col in kolom_merge:
                    hasil_cols[col].append(
                        lama_cols[col][i] if i < len(lama_cols[col]) else "")

        # ── Sub-baris BARU yang belum dipakai → append ──────────────────────
        for i in range(n_baru):
            if i in baru_dipakai:
                continue
            for col in kolom_merge:
                hasil_cols[col].append(
                    baru_cols[col][i] if i < len(baru_cols[col]) else "")

        # Pad semua kolom agar panjangnya sama
        n_total = max((len(v) for v in hasil_cols.values()), default=0)
        for col in hasil_cols:
            while len(hasil_cols[col]) < n_total:
                hasil_cols[col].append("")

        # Susun hasil
        merged = {}
        merged["NOMOR PRK"] = str(row_lama.get("NOMOR PRK", "")).split("\n")[0].strip()
        for col in kolom_merge:
            merged[col] = "\n".join(hasil_cols[col])
        # Kolom manual: ambil langsung dari row_lama, tidak disentuh
        for col in all_gsheet_cols:
            if col not in merged:
                merged[col] = str(row_lama.get(col, "")).strip()
        return merged

    # ── Loop utama ───────────────────────────────────────────────────────────
    df_baru_dedup = df_baru.drop_duplicates(subset=["NOMOR PRK"], keep="first")
    baru_dict = {
        str(row["NOMOR PRK"]).strip().upper(): row.to_dict()
        for _, row in df_baru_dedup.iterrows()
    }

    hasil_rows = []

    for _, row_lama in df_lama.iterrows():
        prk_key = str(row_lama["NOMOR PRK"]).strip().upper()
        if prk_key in baru_dict:
            merged = _merge_satu_prk(row_lama.to_dict(), baru_dict[prk_key])
        else:
            # PRK tidak di data baru → pertahankan semua (termasuk kolom manual)
            merged = {col: str(row_lama.get(col, "")).strip()
                      for col in all_gsheet_cols}
        hasil_rows.append(merged)

    # PRK baru → tambahkan di bawah (kolom manual dikosongkan)
    for _, row_baru in df_baru_dedup.iterrows():
        prk_key = str(row_baru["NOMOR PRK"]).strip().upper()
        if prk_key not in urutan_lama:
            new_row = {col: "" for col in all_gsheet_cols}
            for col in KOLOM_DRP:
                if col in row_baru:
                    new_row[col] = str(row_baru.get(col, ""))
            hasil_rows.append(new_row)

    # Susun df_merged dengan urutan kolom GSheet asli
    df_merged = pd.DataFrame(hasil_rows, columns=all_gsheet_cols)
    df_merged["NOMOR PRK"] = df_merged["NOMOR PRK"].astype(str).str.split("\n").str[0].str.strip()

    # Update / insert kolom NO
    if "NO" in df_merged.columns:
        df_merged["NO"] = range(1, len(df_merged) + 1)
    else:
        df_merged.insert(0, "NO", range(1, len(df_merged) + 1))

    return df_merged, df_lama


# =====================================================
# UI
# =====================================================

class DRPApp:
    def __init__(self, root, colors, on_back, lift_titlebar=None):
        self.root            = root
        self.colors          = colors
        self.on_back         = on_back
        self.lift_titlebar   = lift_titlebar   # callback untuk angkat title bar
        self.file_path       = None
        self.hasil_data      = {}

        self.frame = tk.Frame(root, bg=colors['bg'])
        self.frame.place(x=0, y=32, relwidth=1, relheight=1)
        self.df_final = None
        self._build_ui()

        # Angkat title bar setelah frame DRP terpasang
        if callable(self.lift_titlebar):
            self.root.after(0, self.lift_titlebar)

    def _darken(self, color):
        c = color.lstrip('#')
        r, g, b = int(c[:2], 16), int(c[2:4], 16), int(c[4:], 16)
        return f"#{max(0,r-30):02x}{max(0,g-30):02x}{max(0,b-30):02x}"

    def _show_dialog(self, kind, title, message, detail=""):
        STYLES = {
            "info":    ("#2563eb", "#eff6ff", "#1e40af"),
            "warning": ("#d97706", "#fffbeb", "#92400e"),
            "error":   ("#b91c1c", "#fff1f2", "#7f1d1d"),
            "success": ("#059669", "#ecfdf5", "#065f46"),
        }
        ICONS = {"info": "ℹ", "warning": "⚠", "error": "✕", "success": "✓"}
        bg_hdr, bg_body, fg_body = STYLES.get(kind, STYLES["info"])

        W = 420   # lebar tetap dialog

        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.grab_set()
        dlg.configure(bg="#f8fafc")
        dlg.withdraw()   # sembunyikan dulu agar tidak flicker saat resize

        hdr = tk.Frame(dlg, bg=bg_hdr, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        inner = tk.Frame(hdr, bg=bg_hdr)
        inner.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(inner, text=ICONS.get(kind, "ℹ"), font=("Segoe UI", 16, "bold"),
                 bg=bg_hdr, fg="white").pack(side="left", padx=(0, 8))
        tk.Label(inner, text=title, font=("Segoe UI", 12, "bold"),
                 bg=bg_hdr, fg="white").pack(side="left")

        body = tk.Frame(dlg, bg="#f8fafc")
        body.pack(fill="x", padx=24, pady=16)
        msg_frame = tk.Frame(body, bg=bg_body, highlightthickness=1,
                             highlightbackground=bg_hdr)
        msg_frame.pack(fill="x")
        for txt, fnt in [(message, ("Segoe UI", 10)), (detail, ("Segoe UI", 9))]:
            if txt:
                tk.Label(msg_frame, text=txt, font=fnt, bg=bg_body, fg=fg_body,
                         wraplength=W - 60, justify="left", anchor="w"
                         ).pack(fill="x", padx=14, pady=(10 if txt == message else 2, 10))

        tk.Frame(dlg, bg="#e2e8f0", height=1).pack(fill="x")
        footer = tk.Frame(dlg, bg="#f1f5f9")
        footer.pack(fill="x", padx=24, pady=10)
        btn = tk.Button(footer, text="✕  Close", font=("Segoe UI", 10, "bold"),
                        bg=bg_hdr, fg="white", activebackground=fg_body,
                        activeforeground="white", relief="flat", cursor="hand2",
                        command=dlg.destroy, padx=32, pady=8)
        btn.pack(side="right")
        btn.bind("<Enter>", lambda e: btn.config(bg=fg_body))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_hdr))

        # Hitung tinggi sebenarnya setelah semua widget ter-render
        dlg.geometry(f"{W}x200")   # set lebar dulu agar wraplength bisa dihitung
        dlg.update_idletasks()
        dlg.update()
        H = dlg.winfo_reqheight()
        x = dlg.winfo_screenwidth()  // 2 - W // 2
        y = dlg.winfo_screenheight() // 2 - H // 2
        dlg.geometry(f"{W}x{H}+{x}+{y}")
        dlg.resizable(False, False)
        dlg.deiconify()   # tampilkan dengan ukuran yang sudah benar

    def _make_btn(self, parent, text, cmd, bg, bg_hover, bg_click=None,
                  font=("Segoe UI", 11, "bold"), **kwargs):
        bg_click = bg_click or self._darken(bg_hover)
        btn = tk.Button(parent, text=text, font=font, bg=bg, fg="white",
                        activebackground=bg_click, activeforeground="white",
                        relief="flat", bd=0, highlightthickness=0,
                        cursor="hand2", command=cmd, **kwargs)
        btn.bind("<Enter>",           lambda e: btn.config(bg=bg_hover))
        btn.bind("<Leave>",           lambda e: btn.config(bg=bg))
        btn.bind("<ButtonPress-1>",   lambda e: btn.config(bg=bg_click))
        btn.bind("<ButtonRelease-1>", lambda e: btn.config(bg=bg_hover))
        return btn

    def _build_ui(self):
        self._build_header()
        self._build_content()

    def _build_header(self):
        c = self.colors
        hf = tk.Frame(self.frame, bg=c['primary'], height=130)
        hf.pack(fill="x")
        hf.pack_propagate(False)
        hc = tk.Frame(hf, bg=c['primary'])
        hc.pack(expand=True, fill="both", padx=40, pady=20)

        left = tk.Frame(hc, bg=c['primary'])
        left.pack(side="left", fill="y")
        try:
            logo = Image.open(resource_path("logo_trsp.png"))
            logo.thumbnail((300, 120), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(logo)
            tk.Label(left, image=self._logo, bg=c['primary']).pack(anchor="w")
        except Exception:
            tk.Label(left, text="INVENTRA", font=("Segoe UI", 22, "bold"),
                     bg=c['primary'], fg="white").pack(anchor="w")

        right = tk.Frame(hc, bg=c['primary'])
        right.pack(side="right", anchor="e")
        self._make_btn(right, "← Kembali", self._go_back,
                       c['primary'], c['primary_dark'],
                       padx=25, pady=10).pack(side="left")

    def _build_content(self):
        c = self.colors
        cf = tk.Frame(self.frame, bg=c['bg'])
        cf.pack(fill="both", expand=True, padx=20, pady=20)

        style = ttk.Style()
        style.theme_use('default')
        style.configure('TNotebook', background=c['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', padding=[20, 10], font=('Segoe UI', 10, 'bold'))
        style.map('TNotebook.Tab',
                  background=[('selected', c['primary'])],
                  foreground=[('selected', 'white'), ('!selected', c['text'])])

        self.notebook = ttk.Notebook(cf)
        self.notebook.pack(fill="both", expand=True)

        for text, builder in [
            ("  Upload & Setting  ", self._build_tab_upload),
            ("  Preview Hasil  ",    self._build_tab_preview),
        ]:
            tab = tk.Frame(self.notebook, bg=c['card_bg'])
            self.notebook.add(tab, text=text)
            builder(tab)

    def _build_tab_upload(self, parent):
        c          = self.colors
        HDR_GREEN  = "#16a34a"
        HDR_ORANGE = "#ea580c"

        # Tombol Proses DRP FIXED di bawah — pack dulu sebelum canvas agar tidak tertutupi
        proc_frame = tk.Frame(parent, bg=c['card_bg'])
        proc_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 15))
        self.process_btn = self._make_btn(
            proc_frame, "⚙  Proses DRP", self._proses,
            c['primary'], c['primary_dark'],
            font=("Segoe UI", 12, "bold"), pady=13)
        self.process_btn.pack(fill="x")

        # Scroll area halaman
        page_canvas = tk.Canvas(parent, bg=c['card_bg'], highlightthickness=0)
        page_vsb    = ttk.Scrollbar(parent, orient="vertical", command=page_canvas.yview)
        page_canvas.configure(yscrollcommand=page_vsb.set)
        page_vsb.pack(side="right", fill="y")
        page_canvas.pack(side="left", fill="both", expand=True)

        page_inner = tk.Frame(page_canvas, bg=c['card_bg'])
        page_win   = page_canvas.create_window((0, 0), window=page_inner, anchor="nw")

        page_inner.bind("<Configure>",
            lambda e: page_canvas.configure(scrollregion=page_canvas.bbox("all")))
        page_canvas.bind("<Configure>",
            lambda e: page_canvas.itemconfig(page_win, width=e.width))

        def _page_scroll(e):
            page_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        page_canvas.bind("<Enter>", lambda e: page_canvas.bind_all("<MouseWheel>", _page_scroll))
        page_canvas.bind("<Leave>", lambda e: page_canvas.unbind_all("<MouseWheel>"))

        main = tk.Frame(page_inner, bg=c['card_bg'])
        main.pack(fill="both", expand=True, padx=20, pady=20)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        left = tk.Frame(main, bg=c['card_bg'])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        tk.Frame(left, bg=c['primary_light']).pack(fill="x", pady=(0, 15))
        tk.Label(left.winfo_children()[-1], text="📋  Konfigurasi File",
                 font=("Segoe UI", 13, "bold"), bg=c['primary_light'],
                 fg=c['text'], anchor="w").pack(fill="x", padx=15, pady=12)

        card_l  = tk.Frame(left, bg="white", relief="solid", bd=1)
        card_l.pack(fill="x")
        inner_l = tk.Frame(card_l, bg="white")
        inner_l.pack(fill="x", padx=20, pady=18)

        tk.Label(inner_l, text="Source File", font=("Segoe UI", 10, "bold"),
                 bg="white", fg=c['text'], anchor="w").pack(fill="x", pady=(0, 6))
        tk.Label(inner_l, text="File Excel (.xlsx / .xls) yang berisi data",
                 font=("Segoe UI", 8), bg="white", fg=c['text_light'],
                 anchor="w").pack(fill="x", pady=(0, 8))

        fr = tk.Frame(inner_l, bg="white")
        fr.pack(fill="x")

        # File label: ukuran FIXED agar tidak berubah saat nama file panjang
        file_container = tk.Frame(fr, bg=c['bg'], relief="solid", bd=1, height=34)
        file_container.pack(side="left", fill="x", expand=True)
        file_container.pack_propagate(False)
        self.file_label = tk.Label(
            file_container, text="Belum ada file dipilih",
            font=("Segoe UI", 9), bg=c['bg'], fg=c['text_light'],
            anchor="w", padx=8)
        self.file_label.place(relx=0, rely=0.5, anchor="w", relwidth=1)

        self._make_btn(fr, "+ Pilih File", self._pilih_file,
                       c['primary'], c['primary_dark'],
                       font=("Segoe UI", 9, "bold"),
                       padx=12, pady=7).pack(side="left", padx=(6, 0))

        tk.Frame(inner_l, bg=c['border'], height=1).pack(fill="x", pady=(16, 14))

        tk.Label(inner_l, text="Sheet AMP", font=("Segoe UI", 10, "bold"),
                 bg="white", fg=c['text'], anchor="w").pack(fill="x", pady=(0, 6))
        tk.Label(inner_l, text="Pilih sheet yang berisi data monitoring AMP",
                 font=("Segoe UI", 8), bg="white", fg=c['text_light'],
                 anchor="w").pack(fill="x", pady=(0, 8))

        row_s = tk.Frame(inner_l, bg="white")
        row_s.pack(fill="x")
        badge = tk.Frame(row_s, bg=HDR_ORANGE, padx=8, pady=4)
        badge.pack(side="left")
        tk.Label(badge, text="AMP", font=("Segoe UI", 9, "bold"),
                 bg=HDR_ORANGE, fg="white").pack()

        self._sheet_vars   = {}
        self._sheet_combos = {}
        var_amp_sheet = tk.StringVar(value="AMP")
        cb_amp_sheet  = ttk.Combobox(row_s, textvariable=var_amp_sheet,
                                     font=("Segoe UI", 10), state="readonly",
                                     height=12, values=[""])
        cb_amp_sheet.pack(side="left", fill="x", expand=True, ipady=4, padx=(10, 0))
        self._sheet_vars["sheet_amp"]   = var_amp_sheet
        self._sheet_combos["sheet_amp"] = cb_amp_sheet
        self._sheet_vars["sheet_drp"]   = tk.StringVar(value="DRP")

        right = tk.Frame(main, bg=c['card_bg'])
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        tk.Frame(right, bg=c['primary_light']).pack(fill="x", pady=(0, 15))
        tk.Label(right.winfo_children()[-1],
                 text="🗂  Column Headers Configuration",
                 font=("Segoe UI", 13, "bold"), bg=c['primary_light'],
                 fg=c['text'], anchor="w").pack(fill="x", padx=15, pady=12)

        card_r = tk.Frame(right, bg="white", relief="solid", bd=1)
        card_r.pack(fill="both", expand=True)

        # Semua field langsung di frame — scroll ditangani page_canvas level halaman
        inner_r = tk.Frame(card_r, bg="white")
        inner_r.pack(fill="both", expand=True, padx=20, pady=15)

        tk.Label(inner_r,
                 text="Sesuaikan kolom sheet AMP dengan field yang dibutuhkan",
                 font=("Segoe UI", 8), bg="white", fg=c['text_light'],
                 anchor="w").pack(fill="x", pady=(0, 10))

        FIELDS = [
            ("col_no_prk",        "NOMOR PRK",       "NOMOR PRK",             "Peruntukan"),
            ("col_item",          "Item Pengadaan",  "ITEM PROSES PENGADAAN", "Nama Barang"),
            ("col_stockcode",     "Stock Code",      "Stock Code",            "Stockcode"),
            ("col_no_requisisi",  "No Requisisi",    "No Requisisi",          "No Requisisi"),
            ("col_satuan",        "Satuan",          "Satuan",                "Satuan"),
            ("col_Vol",           "Vol",             "Vol",                   "Jml"),
            ("col_tgl_ro",        "Tanggal RO",      "Tanggal RO",            "Tgl But RO"),
            ("col_no_ro",         "No RO",           "No RO",                 "No RO"),
            ("col_tgl_selesai",   "Tgl Selesai",     "Tgl Selesai",           "Tgl Selesai.2"),
            ("col_nilai_hpe",     "Nilai HPE",       "Nilai HPE",             "Harga HPE"),
            ("col_tgl_kontrak",   "Tanggal Kontrak", "Tanggal Kontrak",       "Tgl Levering"),
            ("col_kontrak",       "No Kontrak",      "Nomor Kontrak",         "Tgl PO"),
            ("col_nilai_kontrak", "Nilai Kontrak",   "Nilai Kontrak",         "Supplier.1"),
            ("col_levering",      "Levering",        "Levering",              "Jenis Kontrak"),
        ]
        self._col_vars   = {}
        self._amp_combos = []

        ttk.Style().configure("Green.TCombobox",
                              fieldbackground="#f0fdf4", foreground="#15803d",
                              selectbackground=HDR_GREEN, selectforeground="white")

        hdr = tk.Frame(inner_r, bg="white")
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="Field", font=("Segoe UI", 8, "bold"),
                 bg="white", fg=c['text'], width=15, anchor="w").pack(side="left")
        hdr_amp = tk.Frame(hdr, bg=HDR_GREEN, padx=6, pady=3)
        hdr_amp.pack(side="left", expand=True, fill="x")
        tk.Label(hdr_amp, text="Kolom Sheet AMP", font=("Segoe UI", 8, "bold"),
                 bg=HDR_GREEN, fg="white").pack()
        tk.Frame(inner_r, bg=c['border'], height=1).pack(fill="x", pady=(0, 6))

        for attr, label, def_drp, def_amp in FIELDS:
            row_f = tk.Frame(inner_r, bg="white")
            row_f.pack(fill="x", pady=2)
            tk.Label(row_f, text=label, font=("Segoe UI", 8, "bold"),
                     bg="white", fg=c['text'], width=15, anchor="w").pack(side="left")
            var_amp = tk.StringVar(value=def_amp)
            cb = ttk.Combobox(row_f, textvariable=var_amp, font=("Segoe UI", 9),
                              state="readonly", style="Green.TCombobox", height=12)
            cb.pack(side="left", fill="x", expand=True, ipady=2)
            self._col_vars[attr] = {"drp": tk.StringVar(value=def_drp), "amp": var_amp}
            self._amp_combos.append((attr, cb, var_amp, def_amp))

        tk.Frame(inner_r, bg=c['border'], height=1).pack(fill="x", pady=(6, 4))
        tk.Label(inner_r, text="💡  Klik dropdown untuk memilih kolom dari sheet AMP",
                 font=("Segoe UI", 8), bg="white", fg=c['text_light'],
                 anchor="w").pack(fill="x")
        tk.Label(inner_r, text="⚠  Huruf besar/kecil tidak mempengaruhi pencocokan",
                 font=("Segoe UI", 8), bg="white", fg=c['warning'],
                 anchor="w").pack(fill="x", pady=(3, 0))

        cb_amp_sheet.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_amp_change())

    def _make_tree(self, parent):

        container = tk.Frame(parent, bg=self.colors['card_bg'])
        container.pack(fill="both", expand=True)

        style = ttk.Style()

        style.configure(
            "Custom.Treeview",
            rowheight=22,
            font=("Segoe UI", 9),
            borderwidth=1,
            relief="solid"
        )

        style.configure(
            "Custom.Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            borderwidth=1,
            relief="solid"
        )

        tree = ttk.Treeview(
            container,
            show="headings",
            style="Custom.Treeview"
        )

        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        # ── Smooth scroll: akumulasi delta trackpad 2 jari ────────────────────
        # Trackpad Windows kirim delta kecil (mis. 4, 8, 12) per event,
        # mouse wheel kirim 120 per klik. Akumulasi agar scroll konsisten.
        _delta_accum = [0]

        def _on_mousewheel(e):
            _delta_accum[0] += e.delta
            units = int(_delta_accum[0] / 120) * -1
            if units != 0:
                _delta_accum[0] %= 120
                tree.yview_scroll(units, "units")

        _delta_accum_x = [0]

        def _on_mousewheel_x(e):
            # Threshold 30 (bukan 120) agar trackpad 2 jari lebih responsif.
            # Setiap threshold terpenuhi, scroll 3 units sekaligus agar terasa ringan.
            _delta_accum_x[0] += e.delta
            units = int(_delta_accum_x[0] / 30) * -1
            if units != 0:
                _delta_accum_x[0] %= 30
                tree.xview_scroll(units * 3, "units")

        def _on_linux_scroll(e):
            tree.yview_scroll(-1 if e.num == 4 else 1, "units")

        def _on_linux_scroll_x(e):
            tree.xview_scroll(-3 if e.num == 6 else 3, "units")

        def _on_enter(e):
            tree.bind_all("<MouseWheel>",       _on_mousewheel)
            tree.bind_all("<Shift-MouseWheel>", _on_mousewheel_x)
            try:
                tree.bind_all("<Button-4>", _on_linux_scroll)
                tree.bind_all("<Button-5>", _on_linux_scroll)
            except Exception:
                pass
            try:
                tree.bind_all("<Button-6>", _on_linux_scroll_x)
                tree.bind_all("<Button-7>", _on_linux_scroll_x)
            except Exception:
                pass

        def _on_leave(e):
            tree.unbind_all("<MouseWheel>")
            tree.unbind_all("<Shift-MouseWheel>")
            try:
                tree.unbind_all("<Button-4>")
                tree.unbind_all("<Button-5>")
            except Exception:
                pass
            try:
                tree.unbind_all("<Button-6>")
                tree.unbind_all("<Button-7>")
            except Exception:
                pass

        tree.bind("<Enter>", _on_enter)
        tree.bind("<Leave>", _on_leave)

        return tree

    def _build_tab_preview(self, parent):
        c = self.colors
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=0)
        parent.rowconfigure(1, weight=0)
        parent.rowconfigure(2, weight=0)
        parent.rowconfigure(3, weight=1)
        parent.rowconfigure(4, weight=0)

        ih = tk.Frame(parent, bg=c['primary_light'])
        ih.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 6))
        self.preview_info = tk.Label(
            ih, text="📊  Belum ada data — klik ⚙ Proses DRP terlebih dahulu",
            font=("Segoe UI", 11, "bold"), bg=c['primary_light'],
            fg=c['text'], anchor="w")
        self.preview_info.pack(fill="x", padx=20, pady=12)

        PREV_FILTER_CFG = [
            ("  Semua  ",          "Semua",        c['primary']),
            ("  AO  ",             "AO",           "#0369a1"),
            ("  AI  ",             "AI",           "#7c3aed"),
            ("  PRK Gabungan  ",   "PRK Gabungan", "#d97706"),
            ("  Tanpa PRK  ",      "Tanpa PRK",    "#be185d"),
        ]
        self._preview_filter_cfg = PREV_FILTER_CFG
        self._preview_filter_infos = {}
        self._preview_active_tab   = [0]
        self._preview_tab_btn_list = []

        def _darken(color):
            c2 = color.lstrip('#')
            r, g, b = int(c2[:2],16), int(c2[2:4],16), int(c2[4:],16)
            return f"#{max(0,r-30):02x}{max(0,g-30):02x}{max(0,b-30):02x}"

        tab_bar = tk.Frame(parent, bg=c['card_bg'])
        tab_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 0))

        info_container = tk.Frame(parent, bg=c['card_bg'])
        info_container.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 4))

        def _select_preview_tab(idx):
            self._preview_active_tab[0] = idx
            for i, (_, _, hdr_color) in enumerate(PREV_FILTER_CFG):
                btn = self._preview_tab_btn_list[i]
                if i == idx:
                    btn.config(bg=hdr_color, fg="white")
                else:
                    btn.config(bg="#e2e8f0", fg=c['text'])
            for i, (_, kategori, _) in enumerate(PREV_FILTER_CFG):
                frm = self._preview_filter_infos[kategori]["frame"]
                if i == idx:
                    frm.pack(fill="x")
                else:
                    frm.pack_forget()
            _, kategori, _ = PREV_FILTER_CFG[idx]
            self._apply_preview_filter(kategori)

        for i, (tab_text, kategori, hdr_color) in enumerate(PREV_FILTER_CFG):
            dk = _darken(hdr_color)
            btn = tk.Button(
                tab_bar, text=tab_text,
                font=("Segoe UI", 9, "bold"),
                bg=hdr_color if i == 0 else "#e2e8f0",
                fg="white" if i == 0 else c['text'],
                relief="flat", cursor="hand2",
                activebackground=dk, activeforeground="white",
                padx=16, pady=7,
                command=lambda idx=i: _select_preview_tab(idx)
            )
            btn.pack(side="left", padx=(0, 2))
            btn.bind("<Enter>", lambda e, b=btn, h=hdr_color: b.config(bg=_darken(h), fg="white"))
            btn.bind("<Leave>", lambda e, b=btn, h=hdr_color, ii=i:
                b.config(bg=h if self._preview_active_tab[0]==ii else "#e2e8f0",
                         fg="white" if self._preview_active_tab[0]==ii else c['text']))
            self._preview_tab_btn_list.append(btn)

            frm = tk.Frame(info_container, bg=hdr_color)
            info_lbl = tk.Label(frm,
                text=f"🔍  {kategori} — belum ada data.",
                font=("Segoe UI", 9), bg=hdr_color, fg="white", anchor="w")
            info_lbl.pack(fill="x", padx=12, pady=5)
            self._preview_filter_infos[kategori] = {"frame": frm, "label": info_lbl}
            if i == 0:
                frm.pack(fill="x")

        tree_frame = tk.Frame(parent, bg=c['card_bg'])
        tree_frame.grid(row=3, column=0, sticky="nsew", padx=15)
        self.tree_drp = self._make_tree(tree_frame)

        save_frame = tk.Frame(parent, bg=c['card_bg'])
        save_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=(8, 15))
        save_frame.columnconfigure(0, weight=1)
        save_frame.columnconfigure(1, weight=1)

        self.send_btn = tk.Button(
            save_frame, text="☁  Kirim ke Google Sheets",
            font=("Segoe UI", 11, "bold"), bg="#dc2626", fg="white",
            relief="flat", cursor="hand2", activebackground="#b91c1c",
            activeforeground="white", command=self._kirim_ke_gsheet,
            pady=13, state="disabled")
        self.send_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.send_btn.bind("<Enter>", lambda e: self.send_btn.config(bg="#b91c1c") if self.send_btn["state"] == "normal" else None)
        self.send_btn.bind("<Leave>", lambda e: self.send_btn.config(bg="#dc2626") if self.send_btn["state"] == "normal" else None)

        self._make_btn(save_frame, "🔗  Open GSheet", self._buka_gsheet,
                       "#059669", "#047857",
                       font=("Segoe UI", 11, "bold"), pady=13
                       ).grid(row=0, column=1, sticky="ew")



    def _apply_preview_filter(self, kategori):

        # =========================
        # TANPA PRK
        # =========================
        if kategori == "Tanpa PRK":

            df = self.hasil_data.get("TANPA_PRK") if self.hasil_data else None

            if df is None or df.empty:
                self._fill_tree(self.tree_drp, None)
                return

            df = df.copy()

            # normalisasi supaya treeview stabil
            df = df.fillna("").astype(str)

            # batasi preview
            df = df.head(300)

            self._fill_tree(self.tree_drp, df)

            return

        # =========================
        # DRP
        # =========================
        df = self.hasil_data.get("DRP") if self.hasil_data else None

        if df is None or df.empty:
            self._fill_tree(self.tree_drp, None)
            return

        df = df.copy()

        if kategori != "Semua":

            kat_col = next(
                (col for col in df.columns if col.strip().lower() == "kategori prk"),
                None
            )

            if kat_col:

                val_upper = df[kat_col].astype(str).str.strip().str.upper()

                if kategori == "AO":
                    df = df[val_upper == "AO"]

                elif kategori == "AI":
                    df = df[val_upper == "PRK I"]

                elif kategori == "PRK Gabungan":
                    df = df[val_upper == "PRK GABUNGAN"]

                df = df.reset_index(drop=True)

                if "NO" in df.columns:
                    df["NO"] = range(1, len(df) + 1)

        # sembunyikan kolom kategori
        tampil_cols = [c for c in df.columns if c.strip().lower() != "kategori prk"]

        self._fill_tree(self.tree_drp, df[tampil_cols])

    def _pilih_file(self):
        path = filedialog.askopenfilename(
            title="Pilih File Excel DRP",
            filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            # ── FIX: teks dipersingkat agar tidak melar label ──
            nama = os.path.basename(path)
            ukuran = os.path.getsize(path) / 1024
            # Potong nama file jika terlalu panjang
            if len(nama) > 35:
                nama = nama[:32] + "..."
            self.file_label.config(
                text=f"📄  {nama}  ({ukuran:.1f} KB)",
                fg=self.colors['text'])
            self.hasil_data = {}
            self.send_btn.config(state="disabled")
            self._clear_preview()
            self._load_sheet_info(path)

    def _load_sheet_info(self, path):
        try:
            xl     = pd.ExcelFile(path)
            sheets = xl.sheet_names
        except Exception:
            sheets = []

        defaults = {"sheet_drp": "DRP", "sheet_amp": "AMP"}
        for attr, cb in self._sheet_combos.items():
            cb["values"] = sheets
            want  = defaults[attr]
            match = next((s for s in sheets if s.strip().upper() == want.upper()), None)
            cb.set(match or (sheets[0] if sheets else ""))

        self._refresh_amp_columns(xl)
        self._sheet_combos["sheet_amp"].bind(
            "<<ComboboxSelected>>", lambda e: self._on_sheet_amp_change())

    def _on_sheet_amp_change(self):
        try:
            self._refresh_amp_columns(pd.ExcelFile(self.file_path))
        except Exception:
            pass

    def _refresh_amp_columns(self, xl):
        sheet_amp = self._sheet_vars.get("sheet_amp", tk.StringVar()).get().strip()
        try:
            df   = (_baca_amp_by_name(xl, sheet_amp)
                    if sheet_amp and sheet_amp in xl.sheet_names
                    else _baca_amp(xl))
            cols = [c for c in (df.columns if df is not None else [])
                    if not str(c).startswith("Unnamed")]
        except Exception:
            cols = []

        for attr, cb, var, default in self._amp_combos:
            cb["values"] = cols
            cb.set(var.get() if var.get() in cols
                   else (default if default in cols
                         else (cols[0] if cols else "")))

    def _proses(self):
        if not self.file_path:
            self._show_dialog("warning", "File Belum Dipilih",
                              "Silakan pilih file Excel terlebih dahulu.")
            return

        col_mapping = {
            attr: {"drp": v["drp"].get().strip(), "amp": v["amp"].get().strip()}
            for attr, v in self._col_vars.items()
        }
        self._col_mapping = col_mapping
        samp = self._sheet_vars.get("sheet_amp", tk.StringVar()).get().strip() or None

        self.process_btn.config(state="disabled", text="⏳  Memproses...",
                                bg="#9ca3af", activebackground="#9ca3af")
        self.frame.update()

        def _t():
            try:
                h = proses_drp(self.file_path, self._col_mapping, samp)
                self.frame.after(0, lambda: self._finish(h))
            except Exception as e:
                self.frame.after(0, lambda err=str(e): self._err(err))

        threading.Thread(target=_t, daemon=True).start()

    def _finish(self, hasil):
        self.hasil_data = hasil
        c = self.colors
        self.process_btn.config(state="normal", text="⚙  Proses DRP",
                                bg=c['primary'], activebackground=c['primary_dark'])

        df_drp = self.hasil_data.get("DRP")
        df_tanpa = self.hasil_data.get("TANPA_PRK")

        # Simpan ke df_final: gabungkan DRP + Tanpa PRK
        # Tanpa PRK diberi kolom NOMOR PRK kosong dan Kategori PRK = "Tanpa PRK"
        if df_drp is not None and not df_drp.empty:
            df_drp_kirim = df_drp.copy()
            if df_tanpa is not None and not df_tanpa.empty:
                df_tanpa_kirim = df_tanpa.copy()
                # Sesuaikan kolom Tanpa PRK ke format DRP:
                # - "Item" → "ITEM PROSES PENGADAAN"
                # - "Peruntukan" → disimpan ke NOMOR PRK (tampil di GSheet)
                #   dikode sebagai "__TANPA_PRK__|<nilai peruntukan>" agar:
                #   (1) tidak dibuang filter NOMOR PRK != ""
                #   (2) tidak cocok pola PRK → tidak salah klasifikasi
                #   (3) push_sheet mengekstrak nilai asli untuk dikirim ke GSheet
                df_tanpa_kirim = df_tanpa_kirim.rename(columns={
                    "Item": "ITEM PROSES PENGADAAN",
                })
                # Ambil nilai Peruntukan, encode ke placeholder
                if "Peruntukan" in df_tanpa_kirim.columns:
                    def _encode_peruntukan(row):
                        val = str(row["Peruntukan"]).strip()
                        # Jika Peruntukan kosong, coba pakai ITEM PROSES PENGADAAN
                        if not val or val.lower() == "nan":
                            item_val = str(row.get("ITEM PROSES PENGADAAN", "")).strip()
                            val = item_val if (item_val and item_val.lower() != "nan") else f"TANPA_PERUNTUKAN_{row.name}"
                        return f"__TANPA_PRK__|{val}"
                    df_tanpa_kirim["NOMOR PRK"] = df_tanpa_kirim.apply(_encode_peruntukan, axis=1)
                    df_tanpa_kirim = df_tanpa_kirim.drop(columns=["Peruntukan"])
                else:
                    # Fallback jika kolom Peruntukan tidak ada
                    df_tanpa_kirim["NOMOR PRK"] = [
                        f"__TANPA_PRK__|{i+1}"
                        for i in range(len(df_tanpa_kirim))
                    ]
                df_tanpa_kirim["Kategori PRK"] = "Tanpa PRK"
                # Gabungkan, kolom yang tidak ada diisi kosong
                self.df_final = pd.concat(
                    [df_drp_kirim, df_tanpa_kirim],
                    ignore_index=True
                ).fillna("")
            else:
                self.df_final = df_drp_kirim
        else:
            self.df_final = df_drp

        if df_drp is not None and not df_drp.empty:
            self.send_btn.config(state="normal")
        else:
            self.send_btn.config(state="disabled")
        df_tanpa       = self.hasil_data.get("TANPA_PRK")
        total_drp      = len(df_drp) if df_drp is not None else 0
        total_tanpa    = len(df_tanpa) if df_tanpa is not None else 0
        sheet_amp_name = self._sheet_vars.get("sheet_amp", tk.StringVar()).get().strip() or "AMP"
        self.preview_info.config(
            text=f"📊  Sheet diproses: {sheet_amp_name}  →  {total_drp} baris DRP  |  {total_tanpa} baris Tanpa PRK")

        if df_drp is not None:
            kat_col = next(
                (col for col in df_drp.columns if col.strip().lower() == "kategori prk"), None)
            for _, kategori, _ in self._preview_filter_cfg:
                if kategori == "Tanpa PRK":
                    n = total_tanpa
                elif kategori == "Semua" or kat_col is None:
                    n = total_drp
                else:
                    val_upper = df_drp[kat_col].astype(str).str.strip().str.upper()
                    if kategori == "AO":
                        n = int((val_upper == "AO").sum())
                    elif kategori == "AI":
                        n = int((val_upper == "PRK I").sum())
                    elif kategori == "PRK Gabungan":
                        n = int((val_upper == "PRK GABUNGAN").sum())
                    else:
                        n = total_drp
                self._preview_filter_infos[kategori]["label"].config(
                    text=f"🔍  {kategori}  →  {n} baris  (dari total {total_drp} baris)")

        # Sembunyikan kolom Kategori PRK dari tampilan preview
        if df_drp is not None:
            tampil_cols = [col for col in df_drp.columns if col.strip().lower() != "kategori prk"]
            self._fill_tree(self.tree_drp, df_drp[tampil_cols])
        self.notebook.select(1)

    def _err(self, msg):
        c = self.colors
        self.process_btn.config(state="normal", text="⚙  Proses DRP",
                                bg=c['primary'], activebackground=c['primary_dark'])
        self._show_dialog("error", "Gagal Memproses", msg)

    def _export_excel_lokal(self):
        df_drp = self.hasil_data.get("DRP")
        if df_drp is None or df_drp.empty:
            self._show_dialog("warning", "Data Kosong",
                              "Proses DRP terlebih dahulu sebelum export.")
            return

        from datetime import datetime
        default_name = f"DRP_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Simpan Export Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")])

        if not path:
            return

        try:
            df_tampil = df_drp.drop(columns=["Kategori PRK"], errors="ignore")
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_tampil.to_excel(writer, sheet_name="DRP", index=False)

                df_amp = self.hasil_data.get("AMP")
                if df_amp is not None and not df_amp.empty:
                    df_amp.to_excel(writer, sheet_name="AMP_raw", index=False)

            self._show_dialog("success", "Export Berhasil",
                              f"File berhasil disimpan.",
                              os.path.basename(path))
        except Exception as e:
            self._show_dialog("error", "Gagal Export", str(e))

    # ─────────────────────────────────────────────────────────────────────────
    # Kirim ke Google Sheets
    # ─────────────────────────────────────────────────────────────────────────
    def _kirim_ke_gsheet(self):
        if self.df_final is None or self.df_final.empty:
            self._show_dialog("warning", "Data Kosong",
                              "Proses DRP terlebih dahulu sebelum mengirim.")
            return

        self.send_btn.config(state="disabled", text="⏳  Mengirim...",
                             bg="#9ca3af", activebackground="#9ca3af")
        self.frame.update()

        df_kirim   = self.df_final.copy()
        sheet_amp  = self._sheet_vars.get("sheet_amp", tk.StringVar()).get().strip() or "AMP"
        cover_info = {"sheet_amp": sheet_amp}

        def _t():
            import traceback
            try:
                from proses_GSheet import (
                    kirim_df_final_ke_gsheet,
                    SheetTidakDitemukanError,
                    GSheetResponseError,
                )
                print(f"[drp_app] Memulai kirim_df_final_ke_gsheet...")
                n_baru, n_lama, n_hasil = kirim_df_final_ke_gsheet(df_kirim, cover_info=cover_info)
                print(f"[drp_app] kirim_df_final_ke_gsheet selesai ✅")
                self.frame.after(0, lambda: self._finish_kirim(n_baru, n_lama, n_hasil))
            except SheetTidakDitemukanError as e:
                print(f"[drp_app] SheetTidakDitemukanError: {e.sheet_name}")
                self.frame.after(0, lambda sn=e.sheet_name: self._handle_sheet_tidak_ada(sn))
            except GSheetResponseError as e:
                print(f"[drp_app] GSheetResponseError: {e.pesan_gas}")
                self.frame.after(0, lambda err=e.pesan_gas: self._err_kirim(err))
            except ConnectionError as e:
                print(f"[drp_app] ConnectionError: {e}")
                self.frame.after(0, lambda err=str(e): self._err_kirim(err))
            except Exception as e:
                tb = traceback.format_exc()
                print(f"[drp_app] ❌ Exception tidak terduga:\n{tb}")
                self.frame.after(0, lambda err=f"{type(e).__name__}: {e}\n\n{tb}": self._err_kirim(err))

        threading.Thread(target=_t, daemon=True).start()

    def _finish_kirim(self, n_baru=0, n_lama=0, n_hasil=0):
        self.send_btn.config(state="normal", text="☁  Kirim ke Google Sheets",
                             bg="#dc2626", activebackground="#b91c1c")
        detail = (
            f"Data lokal   : {n_baru} baris\n"
            f"Data GSheet  : {n_lama} baris\n"
            f"Hasil merged : {n_hasil} baris"
        )
        self._show_dialog("success", "Berhasil Dikirim",
                          "Data berhasil dikirim ke Google Sheets.", detail)

    def _err_kirim(self, msg):
        self.send_btn.config(state="normal", text="☁  Kirim ke Google Sheets",
                             bg="#dc2626", activebackground="#b91c1c")
        self._show_dialog("error", "Gagal Mengirim", msg)

    def _err_kirim(self, msg):
        self.send_btn.config(state="normal", text="☁  Kirim ke Google Sheets",
                             bg="#dc2626", activebackground="#b91c1c")
        self._show_dialog("error", "Gagal Mengirim", msg)

    def _handle_sheet_tidak_ada(self, sheet_name):
        self.send_btn.config(state="normal", text="☁  Kirim ke Google Sheets",
                             bg="#dc2626", activebackground="#b91c1c")
        self._dialog_sheet_tidak_ada(sheet_name)

    def _dialog_sheet_tidak_ada(self, sheet_name):
        import webbrowser
        from proses_GSheet import GSHEET_URL

        TEMPLATE_KOLOM = [
            "NOMOR PRK", "ITEM PROSES PENGADAAN", "Stock Code", "No Requisisi",
            "Vol", "Satuan", "Tanggal RO", "No RO",
            "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak", "Nilai Kontrak", "Levering",
        ]

        dlg = tk.Toplevel(self.root)
        dlg.title("Sheet Tidak Ditemukan")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg="#f8fafc")
        dlg.update_idletasks()
        w, h = 520, 280
        x = dlg.winfo_screenwidth()  // 2 - w // 2
        y = dlg.winfo_screenheight() // 2 - h // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        hdr = tk.Frame(dlg, bg="#b91c1c", height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        hi = tk.Frame(hdr, bg="#b91c1c")
        hi.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(hi, text="⚠", font=("Segoe UI", 18),
                 bg="#b91c1c", fg="#fecaca").pack(side="left", padx=(0, 8))
        tk.Label(hi, text="Sheet Tidak Ditemukan", font=("Segoe UI", 13, "bold"),
                 bg="#b91c1c", fg="white").pack(side="left")

        body = tk.Frame(dlg, bg="#f8fafc")
        body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

        mf = tk.Frame(body, bg="#fff1f2", highlightthickness=1,
                      highlightbackground="#fecaca")
        mf.pack(fill="x")
        tk.Label(mf, text=f'Sheet "{sheet_name}" belum ada di Google Sheets.',
                 font=("Segoe UI", 10, "bold"), bg="#fff1f2", fg="#b91c1c",
                 anchor="w").pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(mf, text="Buat sheet baru dengan nama persis seperti di bawah, "
                          "lalu isi header kolom sesuai template.",
                 font=("Segoe UI", 9), bg="#fff1f2", fg="#7f1d1d",
                 wraplength=460, justify="left", anchor="w"
                 ).pack(fill="x", padx=14, pady=(0, 10))

        tk.Label(body, text="📋  Nama Sheet yang Dibutuhkan:",
                 font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#1f2937",
                 anchor="w").pack(fill="x", pady=(12, 4))

        name_frame = tk.Frame(body, bg="#f0f9ff", highlightthickness=1,
                              highlightbackground="#bae6fd")
        name_frame.pack(fill="x")
        name_inner = tk.Frame(name_frame, bg="#f0f9ff")
        name_inner.pack(fill="x", padx=12, pady=8)
        tk.Label(name_inner, text=sheet_name,
                 font=("Courier New", 11, "bold"), bg="#f0f9ff", fg="#0369a1",
                 anchor="w").pack(side="left", fill="x", expand=True)

        def _salin_nama():
            dlg.clipboard_clear()
            dlg.clipboard_append(sheet_name)
            salin_btn.config(text="✓  Tersalin!", bg="#059669")
            dlg.after(1500, lambda: salin_btn.config(
                text="📋  Salin Nama", bg="#0369a1"))

        salin_btn = tk.Button(name_inner, text="📋  Salin Nama",
                              font=("Segoe UI", 8, "bold"),
                              bg="#0369a1", fg="white", relief="flat",
                              cursor="hand2", command=_salin_nama,
                              padx=10, pady=4)
        salin_btn.pack(side="right")

        tk.Frame(dlg, bg="#e2e8f0", height=1).pack(fill="x", pady=(12, 0))
        bf = tk.Frame(dlg, bg="#f1f5f9")
        bf.pack(fill="x", padx=24, pady=12)
        bf.columnconfigure(0, weight=3)
        bf.columnconfigure(1, weight=1)

        self._make_btn(bf, "🔗  Buka Google Sheets",
                       lambda: webbrowser.open(GSHEET_URL),
                       "#059669", "#047857",
                       font=("Segoe UI", 10, "bold"), pady=10
                       ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._make_btn(bf, "✕  Tutup", dlg.destroy,
                       "#64748b", "#475569",
                       font=("Segoe UI", 10, "bold"), pady=10
                       ).grid(row=0, column=1, sticky="ew")

    def _dialog_header_tidak_sesuai(self, sheet_name, kolom_hilang):
        import webbrowser
        from proses_GSheet import GSHEET_URL

        TEMPLATE_KOLOM = [
            "NOMOR PRK", "ITEM PROSES PENGADAAN", "Stock Code", "No Requisisi",
            "Vol", "Satuan", "Tanggal RO", "No RO",
            "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak", "Nilai Kontrak", "Levering",
        ]

        dlg = tk.Toplevel(self.root)
        dlg.title("Header Tidak Sesuai")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg="#f8fafc")
        dlg.update_idletasks()
        w, h = 520, 280
        x = dlg.winfo_screenwidth()  // 2 - w // 2
        y = dlg.winfo_screenheight() // 2 - h // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        hdr = tk.Frame(dlg, bg="#d97706", height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        hi = tk.Frame(hdr, bg="#d97706")
        hi.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(hi, text="⚠", font=("Segoe UI", 18),
                 bg="#d97706", fg="#fef3c7").pack(side="left", padx=(0, 8))
        tk.Label(hi, text="Header Kolom Tidak Sesuai", font=("Segoe UI", 13, "bold"),
                 bg="#d97706", fg="white").pack(side="left")

        body = tk.Frame(dlg, bg="#f8fafc")
        body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

        mf = tk.Frame(body, bg="#fffbeb", highlightthickness=1,
                      highlightbackground="#fcd34d")
        mf.pack(fill="x")
        tk.Label(mf, text=f'Sheet "{sheet_name}" ditemukan, namun header kolom tidak lengkap.',
                 font=("Segoe UI", 10, "bold"), bg="#fffbeb", fg="#92400e",
                 anchor="w").pack(fill="x", padx=14, pady=(10, 2))
        tk.Label(mf, text="Pastikan baris header sheet sesuai template berikut, "
                          "lalu coba sinkronisasi ulang.",
                 font=("Segoe UI", 9), bg="#fffbeb", fg="#92400e",
                 wraplength=460, justify="left", anchor="w"
                 ).pack(fill="x", padx=14, pady=(0, 10))

        tk.Label(body, text="✅  Template header yang dibutuhkan:",
                 font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#15803d",
                 anchor="w").pack(fill="x", pady=(12, 4))

        tpl_frame = tk.Frame(body, bg="#f0fdf4", highlightthickness=1,
                             highlightbackground="#86efac")
        tpl_frame.pack(fill="x")
        tk.Label(tpl_frame, text="  |  ".join(TEMPLATE_KOLOM),
                 font=("Courier New", 8), bg="#f0fdf4", fg="#15803d",
                 wraplength=460, justify="left", anchor="w"
                 ).pack(fill="x", padx=14, pady=8)

        tk.Frame(dlg, bg="#e2e8f0", height=1).pack(fill="x", pady=(12, 0))
        bf = tk.Frame(dlg, bg="#f1f5f9")
        bf.pack(fill="x", padx=24, pady=12)
        bf.columnconfigure(0, weight=3)
        bf.columnconfigure(1, weight=1)

        self._make_btn(bf, "🔗  Buka Google Sheets",
                       lambda: webbrowser.open(GSHEET_URL),
                       "#059669", "#047857",
                       font=("Segoe UI", 10, "bold"), pady=10
                       ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._make_btn(bf, "✕  Tutup", dlg.destroy,
                       "#64748b", "#475569",
                       font=("Segoe UI", 10, "bold"), pady=10
                       ).grid(row=0, column=1, sticky="ew")

    def _buka_gsheet(self):
        import webbrowser
        from proses_GSheet import get_gsheet_url, check_sheet_exists, WEB_APP_URL, _sheet_gid_cache
        from datetime import datetime
        sheet_name = f"MONITORING DRP {datetime.now().year}"
        try:
            # Pastikan gid sudah ada di cache — fetch jika belum
            if sheet_name not in _sheet_gid_cache:
                check_sheet_exists(WEB_APP_URL, sheet_name)
            url = get_gsheet_url(sheet_name)
            webbrowser.open(url)
        except Exception as e:
            self._show_dialog("error", "Gagal Membuka",
                              "Tidak dapat membuka link Google Sheets.", str(e))

    def _clear_preview(self):

        for item in self.tree_drp.get_children():
            self.tree_drp.delete(item)

        self.preview_info.config(
            text="📊  Belum ada data — klik ⚙ Proses DRP terlebih dahulu"
        )
        
    def _fill_tree(self, tree, df):

        tree.delete(*tree.get_children())

        if df is None or df.empty:
            return

        # batasi preview supaya cepat
        df = df.head(300)

        cols = list(df.columns)

       # style zebra row
        tree.tag_configure("even", background="#f1f5f9")
        tree.tag_configure("odd", background="white")

        # set kolom hanya jika berubah
        if tree["columns"] != tuple(cols):

            tree["columns"] = cols

            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=150, anchor="w", stretch=True)

        for i, row in enumerate(df.values):

            split_cols = [str(v).split("\n") for v in row]
            max_lines = max(len(col) for col in split_cols)

            for line in range(max_lines):

                new_row = []

                for col in split_cols:
                    new_row.append(col[line] if line < len(col) else "")

                tag = "even" if i % 2 == 0 else "odd"

                tree.insert("", "end", values=new_row, tags=(tag,))

    def _go_back(self):
        self.frame.destroy()
        self.on_back()