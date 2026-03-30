from datetime import datetime
import pandas as pd
import json
import urllib.request
import urllib.parse


WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzqUyf2wpSUFdVPtV_qZ1DZsGxpmdBcEfUsNuvws1bKvAKqARolJYQUGLbyfg4Z-4ZU/exec"
GSHEET_URL  = "https://docs.google.com/spreadsheets/d/1q65xJgzOg9rSD2L6IDVfOUQUkoAsTKCGUr1aq5oOkJA/edit?usp=sharing"

# Kolom yang nilainya harus diperlakukan sebagai teks murni
_KOLOM_TEKS = {"NOMOR PRK", "No RO", "Stock Code", "STOCK CODE"}

# Urutan seksi kategori di GSheet
SEKSI_KATEGORI = [
    ("AO",           "AO"),
    ("AI",           "PRK I"),
    ("PRK GABUNGAN", "PRK GABUNGAN"),
    ("TANPA PRK",    "Tanpa PRK"),
]

# Penanda baris label seksi (nilai di kolom NOMOR PRK)
_LABEL_PREFIX = "##SEKSI##"

# ============================================================
# Header default GSheet — dipakai saat sheet baru/kosong tidak punya header.
# Urutan kolom ini adalah urutan resmi yang akan ditulis ke GSheet.
# ============================================================
HEADER_DEFAULT = [
    "NO",
    "NOMOR PRK",
    "ITEM PROSES PENGADAAN",
    "STOCK CODE",
    "No Requisisi",
    "VOL",
    "SATUAN",
    "KATEGORI (SPM)",
    "REVIEW VFM",
    "NILAI ANGGARAN (Rp x 1000)",
    "PEMILIK ANGGARAN",
    "DISBURSE ANGGARAN",
    "TAHUN DISBURSE",
    "METODE PENGADAAN",
    "PELAKSANA PENGADAAN",
    "PELAKSANA PENGADAAN SESUAI KEWENANGAN BARU",
    "JENIS PROGRAM",
    "POTENSI PENYEDIA PLN GROUP/PLN NP GROUP",
    "DASPEN",
    "TORRAB",
    "INFO HARGA >1",
    "ANALISA RISIKO",
    "KKP/ KT",
    "USULAN KE SSCM/ IKG",
    "KELENGKAPAN",
    "METODE PENGADAAN (AKHIR)",
    "Tanggal RO",
    "No RO",
    "Tanggal RKS (Jika Ada)",
    "Nilai HPE",
    "Kirim Pengadaan",
    "Tanggal Terkontrak",
    "Nomor Kontrak",
    "Nilai Kontrak",
    "Levering",
    "Kedatangan",
    "Proses BA",
    "Pembayaran",
]


# ============================================================
# Exception khusus
# ============================================================
class SheetTidakDitemukanError(Exception):
    """Dilempar ketika sheet target tidak ada di Google Sheets."""
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name
        super().__init__(sheet_name)


class HeaderTidakDitemukanError(Exception):
    """
    Dilempar ketika sheet ditemukan dan berisi data,
    namun tidak ada baris header yang mengandung kolom 'NOMOR PRK'.
    Ini menandakan struktur sheet rusak atau tidak sesuai template.
    """
    def __init__(self, sheet_name, kolom_hilang=None):
        self.sheet_name   = sheet_name
        self.kolom_hilang = kolom_hilang or []
        super().__init__(sheet_name)


class GSheetResponseError(Exception):
    """
    Dilempar ketika push_sheet mendapat respons error dari Google Apps Script.
    """
    def __init__(self, pesan_gas):
        self.pesan_gas = pesan_gas
        super().__init__(pesan_gas)


# ============================================================
# Helper: paksa nilai sebagai teks (anti scientific notation)
# ============================================================
def _force_text(val):
    s = str(val).strip() if val is not None else ""
    if s == "" or s == "nan":
        return ""
    if not s.startswith("'"):
        return "'" + s
    return s


# ============================================================
# Helper: normalisasi kolom df_sumber agar cocok dengan acuan
#
# Kolom df_sumber yang namanya sama (case-insensitive) dengan
# kolom di df_acuan / list acuan akan di-rename ke nama versi
# acuan. Kolom yang benar-benar baru dibiarkan apa adanya.
#
# Parameter:
#   df       — DataFrame yang kolomnya akan dinormalisasi
#   acuan    — list nama kolom acuan (urutan dipertahankan)
#
# Return: DataFrame dengan kolom yang sudah di-rename
# ============================================================
def _normalisasi_kolom(df, acuan):
    """
    Rename kolom df agar cocok (case-insensitive) dengan nama di acuan.
    Kolom yang tidak ada di acuan dibiarkan apa adanya.
    """
    acuan_lower = [c.lower() for c in acuan]
    rename_map  = {}
    for c in df.columns:
        if c.lower() in acuan_lower:
            canonical = acuan[acuan_lower.index(c.lower())]
            if c != canonical:
                rename_map[c] = canonical
    return df.rename(columns=rename_map) if rename_map else df


# ============================================================
# Helper: susun df menjadi seksi-seksi berdasarkan kategori
# Setiap seksi diawali baris label (NOMOR PRK = ##SEKSI##<nama>)
# ============================================================
def susun_per_seksi(df, kolom_kategori="Kategori PRK", cover_info=None):
    """
    Susun df dengan struktur per seksi:
        [##COVER_TITLE##]   ← baris judul cover (opsional)
        [##COVER_DATE##]    ← baris tanggal generate
        [##COVER_AMP##]     ← baris nama sheet AMP
        [##COVER_END##]     ← baris penutup cover
        [##SEKSI##<NAMA>]   ← baris label kategori
        [##HEADER##]        ← baris header kolom
        ... data ...
        [##EMPTY##]         ← baris kosong pemisah

    cover_info: dict opsional {"sheet_amp": str, "tanggal": str}
    """
    from drp_app import klasifikasi_prk
    from datetime import datetime as _dt

    df = df.copy()

    if kolom_kategori not in df.columns:
        prk_col = next((c for c in df.columns
                        if c.strip().upper() == "NOMOR PRK"), None)
        if prk_col:
            df[kolom_kategori] = df[prk_col].apply(klasifikasi_prk)
        else:
            return df

    all_cols = list(df.columns)
    bagian   = []

    def _marker_row(prk_val, extra_col=None, extra_val=""):
        row = {col: "" for col in all_cols}
        row["NOMOR PRK"] = prk_val
        if extra_col and extra_col in row:
            row[extra_col] = extra_val
        return pd.DataFrame([row], columns=all_cols)

    # ── Baris Cover di paling atas ───────────────────────────────────────────
    if cover_info:
        now       = _dt.now()
        tgl_str   = cover_info.get("tanggal") or now.strftime("%d %B %Y  |  %H:%M WIB")
        amp_str   = cover_info.get("sheet_amp") or "AMP"
        tahun     = now.year

        bagian.append(_marker_row("##COVER_TITLE##",
                                  "ITEM PROSES PENGADAAN",
                                  f"MONITORING DRP {tahun}"))
        bagian.append(_marker_row("##COVER_DATE##",
                                  "ITEM PROSES PENGADAAN",
                                  f"Generated : {tgl_str}"))
        bagian.append(_marker_row("##COVER_AMP##",
                                  "ITEM PROSES PENGADAAN",
                                  f"Sheet AMP : {amp_str}"))
        bagian.append(_marker_row("##COVER_END##"))

    # ── Seksi data ───────────────────────────────────────────────────────────
    for nama_seksi, nilai_kategori in SEKSI_KATEGORI:
        if nilai_kategori == "Tanpa PRK":
            mask = df[kolom_kategori].astype(str).str.strip() == nilai_kategori
        else:
            mask = df[kolom_kategori].astype(str).str.strip().str.upper() == nilai_kategori.upper()
        subset = df.loc[mask, all_cols]

        bagian.append(_marker_row(f"{_LABEL_PREFIX}{nama_seksi}"))
        bagian.append(_marker_row("##HEADER##"))
        if not subset.empty:
            # Untuk seksi TANPA PRK: decode placeholder __TANPA_PRK__| di NOMOR PRK
            # agar nilai Peruntukan bersih yang dikirim ke GSheet (bukan placeholder)
            if nilai_kategori == "Tanpa PRK" and "NOMOR PRK" in subset.columns:
                subset = subset.copy()
                def _decode_prk(val):
                    s = str(val).strip()
                    if s.startswith("__TANPA_PRK__|"):
                        return s[len("__TANPA_PRK__|"):]
                    return s
                subset["NOMOR PRK"] = subset["NOMOR PRK"].apply(_decode_prk)
            bagian.append(subset)
        bagian.append(_marker_row("##EMPTY##"))

    if not bagian:
        return df

    return pd.concat(bagian, ignore_index=True)


# ============================================================
# Cek apakah sheet sudah ada
# ============================================================
# Cache gid sheet — diisi saat check_sheet_exists berhasil
_sheet_gid_cache = {}


def check_sheet_exists(url, sheet_name):
    """
    Kirim GET request ke GAS untuk mengecek keberadaan sheet.

    Return:
        True  → sheet ada (GAS tidak mengembalikan key 'error')
        False → sheet tidak ada (GAS mengembalikan key 'error')
        None  → gagal terhubung ke internet / timeout
    """
    params   = urllib.parse.urlencode({"sheet_name": sheet_name})
    full_url = f"{url}?{params}"
    try:
        req = urllib.request.Request(full_url)
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read().decode("utf-8")
        result = json.loads(raw)
        if "error" not in result:
            # Simpan gid ke cache jika tersedia
            if "gid" in result:
                _sheet_gid_cache[sheet_name] = result["gid"]
            return True
        return False
    except Exception as e:
        print(f"[check_sheet] error: {e}")
        return None


def get_gsheet_url(sheet_name=None):
    """
    Return URL Google Sheets langsung ke sheet tertentu.
    Jika gid tersedia di cache, tambahkan #gid=... agar browser
    langsung buka tab sheet yang sesuai.
    """
    from datetime import datetime as _dt
    if sheet_name is None:
        sheet_name = f"MONITORING DRP {_dt.now().year}"
    gid = _sheet_gid_cache.get(sheet_name)
    if gid is not None:
        # Format: https://docs.google.com/spreadsheets/d/ID/edit#gid=GID
        base = GSHEET_URL.split("#")[0].replace("/view", "/edit")
        return f"{base}#gid={gid}"
    return GSHEET_URL


# ============================================================
# Fetch data lama dari sheet — kembalikan SEMUA kolom GSheet
# Baris label seksi (##SEKSI##) dilewati saat merge
#
# Kondisi yang ditangani:
#   1. Sheet tidak ada                          → raise SheetTidakDitemukanError
#   2. Sheet ada, kosong (< 2 baris)            → return None (sheet kosong biasa)
#   3. Sheet ada, ada data, NOMOR PRK tidak ada → return DataFrame kosong
#                                                  dengan HEADER_DEFAULT
#                                                  (header hilang → reset ke default)
#   4. Normal                                   → return DataFrame
# ============================================================
def fetch_sheet(url, sheet_name):
    params   = urllib.parse.urlencode({"sheet_name": sheet_name})
    full_url = f"{url}?{params}"
    try:
        req = urllib.request.Request(full_url)
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
        result = json.loads(raw)

        # ── GAS mengembalikan error → sheet tidak ada ────────────────────────
        if "error" in result:
            raise SheetTidakDitemukanError(sheet_name)

        data = result.get("data", [])

        # ── Sheet ada tapi benar-benar kosong (0 atau 1 baris) ───────────────
        # Kembalikan DataFrame kosong dengan HEADER_DEFAULT agar merge dan push
        # tetap bisa berjalan dengan struktur kolom yang benar.
        if not data or len(data) < 2:
            print(f"[fetch_sheet] Sheet '{sheet_name}' kosong, "
                  f"menggunakan HEADER_DEFAULT ({len(HEADER_DEFAULT)} kolom).")
            return pd.DataFrame(columns=HEADER_DEFAULT)

        # ── Cari baris header yang mengandung "NOMOR PRK" ────────────────────
        headers    = None
        data_start = 1
        for idx, row in enumerate(data):
            row_clean = [str(c).replace("'", "").strip().upper() for c in row]
            if "NOMOR PRK" in row_clean:
                headers    = [str(h) for h in row]
                data_start = idx + 1
                break

        # ── Header tidak ditemukan → gunakan HEADER_DEFAULT ────────────────
        # Sheet ada dan berisi data tapi tidak punya baris header NOMOR PRK.
        # Kemungkinan: sheet baru dibuat tapi belum ada header, atau header
        # terhapus tidak sengaja.
        # Solusi: anggap sheet kosong berstruktur default — kembalikan DataFrame
        # kosong dengan HEADER_DEFAULT sehingga merge bisa berjalan normal
        # dan push akan menulis header yang benar ke GSheet.
        if headers is None:
            print(f"[fetch_sheet] Header tidak ditemukan di '{sheet_name}', "
                  f"menggunakan HEADER_DEFAULT ({len(HEADER_DEFAULT)} kolom).")
            return pd.DataFrame(columns=HEADER_DEFAULT)

        n_cols        = len(headers)
        headers_upper = [h.strip().upper() for h in headers]

        # Cari indeks kolom NOMOR PRK
        prk_idx = next((i for i, h in enumerate(headers)
                        if h.strip().upper() == "NOMOR PRK"), None)

        _SKIP_MARKERS = (_LABEL_PREFIX, "##HEADER##", "##EMPTY##")

        # ── Kumpulkan semua sub-baris per PRK ────────────────────────────────
        # Karena di GSheet cell NOMOR PRK di-merge secara vertikal,
        # sub-baris ke-2 dst punya NOMOR PRK kosong.
        # Solusi: sub-baris kosong di-group ke PRK terakhir yang ditemukan,
        # lalu semua sub-baris digabung dengan \n per kolom.
        #
        # PENGECUALIAN: baris di seksi TANPA PRK selalu punya NOMOR PRK kosong
        # dan harus diperlakukan sebagai baris mandiri (bukan sub-baris).
        grouped_prk = []   # [(prk_val, [row, row, ...]), ...]
        in_tanpa_prk_seksi = False   # flag: sedang di seksi TANPA PRK

        for row in data[data_start:]:
            row_padded = list(row) + [""] * (n_cols - len(row))
            row_padded = row_padded[:n_cols]

            # Skip baris header berulang
            row_upper = [str(c).strip().lstrip("'").upper() for c in row_padded]
            if row_upper == headers_upper:
                continue

            # Cek marker di kolom NOMOR PRK
            if prk_idx is not None:
                val_prk = str(row_padded[prk_idx]).strip().lstrip("'")

                # Deteksi marker seksi — update flag in_tanpa_prk_seksi
                if val_prk.startswith(_LABEL_PREFIX):
                    seksi_nama = val_prk[len(_LABEL_PREFIX):].strip().upper()
                    in_tanpa_prk_seksi = (seksi_nama == "TANPA PRK")
                    continue   # skip baris marker

                # Skip marker header / empty
                if any(val_prk.startswith(m) for m in ("##HEADER##", "##EMPTY##")):
                    continue

                if val_prk != "":
                    # Baris PRK baru → buat group baru
                    grouped_prk.append((val_prk, [row_padded]))
                else:
                    if in_tanpa_prk_seksi:
                        # Di seksi TANPA PRK → setiap baris adalah baris mandiri
                        # meskipun NOMOR PRK-nya kosong
                        grouped_prk.append(("", [row_padded]))
                    else:
                        # NOMOR PRK kosong di luar seksi TANPA PRK →
                        # sub-baris dari PRK sebelumnya (efek merge cell vertikal)
                        if grouped_prk:
                            grouped_prk[-1][1].append(row_padded)
            else:
                # Tidak ada kolom NOMOR PRK → kumpulkan apa adanya
                grouped_prk.append(("", [row_padded]))

        if not grouped_prk:
            return None

        # ── Gabungkan sub-baris per PRK dengan \n ────────────────────────────
        # Kolom key (No RO, Stock Code, No Requisisi) di sub-baris yang kosong
        # di-forward-fill dari nilai baris sebelumnya dalam grup yang sama.
        _KEY_COLS_LOWER = {"no ro", "stock code", "no requisisi"}

        key_col_idxs = set()
        for ci, h in enumerate(headers):
            if h.strip().lower() in _KEY_COLS_LOWER:
                key_col_idxs.add(ci)

        rows_bersih = []
        for prk_val, sub_rows in grouped_prk:
            if len(sub_rows) == 1:
                rows_bersih.append(sub_rows[0])
            else:
                # Forward fill kolom key pada sub-baris yang kosong
                filled_rows = []
                last_val = {}   # col_idx → nilai terakhir yang tidak kosong
                for row in sub_rows:
                    filled = list(row)
                    for ci in key_col_idxs:
                        v = str(filled[ci]).strip().lstrip("'")
                        if v and v.upper() != "NAN":
                            last_val[ci] = v
                        elif ci in last_val:
                            filled[ci] = last_val[ci]
                    filled_rows.append(filled)

                # Gabung tiap kolom dengan \n
                merged_row = []
                for col_i in range(n_cols):
                    if col_i == prk_idx:
                        merged_row.append(prk_val)
                    else:
                        vals = [str(r[col_i]).strip().lstrip("'") for r in filled_rows]
                        merged_row.append("\n".join(vals))
                rows_bersih.append(merged_row)

        df = pd.DataFrame(rows_bersih, columns=headers)

        # Deduplikasi nama kolom (jika GSheet punya kolom ganda)
        seen = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols

        for col in df.columns:
            df[col] = df[col].astype(str).str.lstrip("'").str.strip()

        return df

    # Teruskan exception khusus agar bisa ditangkap di layer atas
    except (SheetTidakDitemukanError, HeaderTidakDitemukanError):
        raise
    except Exception as e:
        print(f"[fetch_sheet] error tidak terduga: {e}")
        raise ConnectionError(
            f"Gagal membaca data dari Google Sheets.\nDetail: {e}"
        )


# ============================================================
# Merge data baru dengan data lama
# ============================================================
def merge_with_existing(df_baru, df_lama, key_col="NOMOR PRK",
                        kolom_diisi=None):
    """
    Merge df_baru (dari lokal/AMP) ke df_lama (dari GSheet).

    - Kolom yang ada di kolom_diisi → di-update dari df_baru, per sub-baris.
    - Kolom yang TIDAK ada di kolom_diisi (kolom manual GSheet) → dipertahankan penuh.
    - PRK baru → ditambahkan di bawah.
    - Pencocokan nama kolom bersifat case-insensitive.

    Logika merge per sub-baris (secondary key kombinasi fleksibel):
        - No RO          : WAJIB cocok jika tidak kosong; berbeda → data baru
        - Stock Code     : fleksibel, kosong di salah satu sisi = tidak menghalangi match
        - No Requisisi   : fleksibel, kosong di salah satu sisi = tidak menghalangi match

    Aturan merge nilai per kolom:
        lama kosong + baru ada     → pakai baru
        lama ada    + baru kosong  → pakai lama  (lama dipertahankan)
        lama ada    + baru ada     → pakai baru  (update)
        keduanya kosong            → ""
    """
    # ── Helper ────────────────────────────────────────────────────────────────
    def _ci(df, nama):
        """Cari nama kolom asli secara case-insensitive."""
        target = nama.strip().lower()
        return next((c for c in df.columns if str(c).strip().lower() == target), None)

    def _split(val):
        """Split nilai multiline menjadi list, strip tiap baris."""
        return [b.strip() for b in str(val).split("\n")]

    def _merge_nilai(v_lama, v_baru):
        """Aturan merge satu nilai skalar."""
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

    def _norm_key(val):
        """Normalisasi nilai key: upper, strip, kosong jika nan."""
        s = str(val).strip().lstrip("'").upper()
        return s if s and s != "NAN" else ""

    def _sub_key(lama_cols, baru_cols, i, j,
                 no_ro_col, sc_col, req_col):
        """
        Cek apakah sub-baris i (lama) cocok dengan sub-baris j (baru).

        Aturan:
        - No RO WAJIB sama (jika keduanya tidak kosong).
          Jika No RO berbeda → tidak cocok (return False).
        - Stock Code & No Requisisi bersifat fleksibel:
          kosong di salah satu sisi → dianggap cocok (tidak menghalangi match).
          keduanya ada tapi berbeda → tidak cocok.
        """
        def _get(cols, col, idx):
            if col is None:
                return ""
            lst = cols.get(col, [])
            return lst[idx] if idx < len(lst) else ""

        ro_l  = _norm_key(_get(lama_cols, no_ro_col, i))
        ro_b  = _norm_key(_get(baru_cols, no_ro_col, j))
        sc_l  = _norm_key(_get(lama_cols, sc_col,    i))
        sc_b  = _norm_key(_get(baru_cols, sc_col,    j))
        req_l = _norm_key(_get(lama_cols, req_col,   i))
        req_b = _norm_key(_get(baru_cols, req_col,   j))

        # No RO harus cocok (kalau keduanya tidak kosong)
        if ro_l and ro_b and ro_l != ro_b:
            return False
        # Kalau keduanya kosong No RO-nya → tidak bisa di-match
        if not ro_l and not ro_b:
            return False

        # Stock Code:
        #   - lama ada, baru ada, berbeda → tidak cocok
        #   - lama ada, baru kosong       → cocok (baru tidak menimpa)
        #   - lama kosong                 → abaikan sebagai key
        if sc_l and sc_b != sc_l and sc_b != "":
            return False

        # No Requisisi: aturan sama seperti Stock Code
        if req_l and req_b != req_l and req_b != "":
            return False

        return True

    def _merge_satu_prk(row_lama, row_baru, kolom_merge, kolom_manual):
        """
        Merge satu PRK secara per sub-baris.
        Matching sub-baris menggunakan dict index No RO → O(1) lookup
        menggantikan nested loop O(n×m).
        """
        no_ro_col = next((c for c in kolom_merge if c.strip().lower() == "no ro"),        None)
        sc_col    = next((c for c in kolom_merge if c.strip().lower() == "stock code"),   None)
        req_col   = next((c for c in kolom_merge if c.strip().lower() == "no requisisi"), None)

        lama_cols = {col: _split(row_lama.get(col, "")) for col in kolom_merge}
        baru_cols = {col: _split(row_baru.get(col, "")) for col in kolom_merge}

        _key_cols = [c for c in [no_ro_col, sc_col, req_col] if c is not None]
        if _key_cols:
            n_lama = max((len(lama_cols[c]) for c in _key_cols), default=1)
            n_baru = max((len(baru_cols[c]) for c in _key_cols), default=1)
        else:
            n_lama = 1
            n_baru = 1

        for col in kolom_merge:
            while len(lama_cols[col]) < n_lama:
                lama_cols[col].append("")
            while len(baru_cols[col]) < n_baru:
                baru_cols[col].append("")

        # ── Bangun index No RO dari baru → O(1) lookup ───────────────────────
        # Key: no_ro_norm → list of j (satu No RO bisa muncul >1 kali)
        if no_ro_col:
            baru_ro_index = {}
            for j in range(n_baru):
                ro_b = _norm_key(baru_cols[no_ro_col][j])
                if ro_b:
                    baru_ro_index.setdefault(ro_b, []).append(j)
        else:
            baru_ro_index = {}

        hasil_cols   = {col: [] for col in kolom_merge}
        baru_dipakai = set()

        def _find_match(i):
            """Cari j yang cocok dengan sub-baris i dari lama. Return j atau None."""
            ro_l = _norm_key(lama_cols[no_ro_col][i]) if no_ro_col else ""
            sc_l = _norm_key(lama_cols[sc_col][i])    if sc_col    else ""
            req_l= _norm_key(lama_cols[req_col][i])   if req_col   else ""

            # Kandidat j: ambil dari index No RO (O(1)) jika ro_l ada,
            # fallback ke semua j jika ro_l kosong (jarang terjadi)
            candidates = baru_ro_index.get(ro_l, []) if ro_l else list(range(n_baru))

            for j in candidates:
                if j in baru_dipakai:
                    continue
                if _sub_key(lama_cols, baru_cols, i, j, no_ro_col, sc_col, req_col):
                    return j
            return None

        # ── Iterasi sub-baris LAMA ──────────────────────────────────────────
        for i in range(n_lama):
            match_j = _find_match(i)
            if match_j is not None:
                baru_dipakai.add(match_j)
                for col in kolom_merge:
                    v_l = lama_cols[col][i]          if i       < len(lama_cols[col]) else ""
                    v_b = baru_cols[col][match_j]    if match_j < len(baru_cols[col]) else ""
                    hasil_cols[col].append(_merge_nilai(v_l, v_b))
            else:
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

        # Pad semua kolom hasil agar panjang sama
        n_total = max((len(v) for v in hasil_cols.values()), default=0)
        for col in hasil_cols:
            while len(hasil_cols[col]) < n_total:
                hasil_cols[col].append("")

        merged = {}
        merged[key_col] = str(row_lama.get(key_col, "")).split("\n")[0].strip()
        for col in kolom_merge:
            merged[col] = "\n".join(hasil_cols[col])
        for col in kolom_manual:
            merged[col] = str(row_lama.get(col, "")).strip()

        return merged

    # ── Setup awal ────────────────────────────────────────────────────────────
    if df_lama is None:
        return df_baru.copy()

    df_baru = df_baru.copy()
    df_lama = df_lama.copy()

    # Jika df_lama kosong (empty DataFrame berkolom HEADER_DEFAULT):
    # Gunakan kolomnya sebagai acuan urutan agar header GSheet
    # mengikuti HEADER_DEFAULT, bukan hanya kolom dari data AMP.
    if df_lama.empty:
        # Normalisasi kolom df_baru agar cocok dengan HEADER_DEFAULT (case-insensitive).
        # Kolom baru yang tidak ada di header default tetap ditambahkan di akhir.
        all_cols = list(df_lama.columns)
        df_baru  = _normalisasi_kolom(df_baru, all_cols)
        for c in df_baru.columns:
            if c not in all_cols:
                all_cols.append(c)
        df_result = df_baru.reindex(columns=all_cols, fill_value="")
        if "NO" in df_result.columns:
            df_result["NO"] = range(1, len(df_result) + 1)
        return df_result
    df_lama = df_lama.copy()

    # Normalisasi nama kolom ke kanonik (case-insensitive)
    semua_kanon = list(kolom_diisi) if kolom_diisi else []
    if key_col not in semua_kanon:
        semua_kanon = [key_col] + semua_kanon

    for kanon in semua_kanon:
        for df in (df_baru, df_lama):
            actual = _ci(df, kanon)
            if actual and actual != kanon:
                df.rename(columns={actual: kanon}, inplace=True)

    # Normalisasi key value
    for df in (df_baru, df_lama):
        if key_col in df.columns:
            df[key_col] = df[key_col].astype(str).str.strip().str.upper()

    # Susun all_cols: urutan lama dipertahankan, kolom baru ditambahkan di akhir.
    # Normalisasi case-insensitive: "Stock Code" di df_baru → "STOCK CODE" dari df_lama.
    all_cols = list(df_lama.columns)
    df_baru  = _normalisasi_kolom(df_baru, all_cols)
    for c in df_baru.columns:
        if c not in all_cols:
            all_cols.append(c)

    df_lama = df_lama.reindex(columns=all_cols, fill_value="")
    df_baru = df_baru.reindex(columns=all_cols, fill_value="")

    # Pisahkan kolom_merge vs kolom_manual
    kolom_update = set(kolom_diisi) if kolom_diisi else set(df_baru.columns)
    kolom_merge  = [c for c in all_cols if c != key_col and c in kolom_update]
    kolom_manual = [c for c in all_cols if c != key_col and c not in kolom_update]

    # Buat dict baru: PRK_KEY → row dict
    baru_dict = {}
    for _, row in df_baru.iterrows():
        k = str(row.get(key_col, "")).strip().upper()
        if k:
            baru_dict[k] = row.to_dict()

    # ── Loop utama ────────────────────────────────────────────────────────────
    hasil     = []
    prk_sudah = set()

    for _, row_lama in df_lama.iterrows():
        prk_key = str(row_lama.get(key_col, "")).strip().upper()
        if not prk_key:
            continue
        prk_sudah.add(prk_key)

        if prk_key not in baru_dict:
            hasil.append(row_lama.to_dict())
        else:
            merged = _merge_satu_prk(
                row_lama.to_dict(),
                baru_dict[prk_key],
                kolom_merge,
                kolom_manual,
            )
            hasil.append(merged)

    # PRK baru (tidak ada di lama) → append di bawah
    for _, row_baru in df_baru.iterrows():
        prk_key = str(row_baru.get(key_col, "")).strip().upper()
        if prk_key and prk_key not in prk_sudah:
            hasil.append(row_baru.to_dict())

    df_hasil = pd.DataFrame(hasil, columns=all_cols)

    if "NO" in df_hasil.columns:
        df_hasil["NO"] = range(1, len(df_hasil) + 1)

    return df_hasil


# ============================================================
# Push data ke Google Sheets
# ============================================================
def push_sheet(url, sheet_name, df, kolom_diisi=None):
    """
    Kirim df ke GSheet dengan struktur berseksi per kategori PRK.
    Baris label (##SEKSI##) dikirim sebagai penanda pemisah.
    GAS menggunakan action replace_all untuk menulis ulang seluruh sheet.
    """
    kolom_update = list(kolom_diisi) if kolom_diisi else list(df.columns)

    print(f"[push_sheet] Mulai — sheet: '{sheet_name}', "
          f"baris: {len(df)}, kolom_update: {len(kolom_update)}")

    rows_as_dict = []
    for _, row in df.iterrows():
        row_dict = {}
        for col in df.columns:
            val = row[col]
            s   = "" if (not isinstance(val, str) and pd.isna(val)) else str(val)
            if col == "NOMOR PRK":
                prk_val = s.split("\n")[0].strip()
                if prk_val.startswith(_LABEL_PREFIX):
                    row_dict[col] = prk_val
                    continue
                # Placeholder baris Tanpa PRK → ekstrak nilai Peruntukan asli
                # format: "__TANPA_PRK__|<nilai peruntukan>"
                if prk_val.startswith("__TANPA_PRK__|"):
                    row_dict[col] = prk_val[len("__TANPA_PRK__|"):]
                    continue
                s = prk_val
            if col in _KOLOM_TEKS and not s.startswith(_LABEL_PREFIX):
                s = _force_text(s)
            row_dict[col] = s
        rows_as_dict.append(row_dict)

    payload = {
        "action":       "replace_all",
        "sheet_name":   sheet_name,
        "data":         rows_as_dict,
        "kolom_update": kolom_update,
    }

    payload_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    print(f"[push_sheet] Payload size: {len(payload_bytes) / 1024:.1f} KB")

    req = urllib.request.Request(
        url, data=payload_bytes,
        headers={"Content-Type": "application/json"}
    )
    try:
        print(f"[push_sheet] Mengirim POST ke GAS (timeout=60s)...")
        with urllib.request.urlopen(req, timeout=60) as resp:
            http_status = resp.getcode()
            raw_resp    = resp.read().decode("utf-8")

        print(f"[push_sheet] HTTP status : {http_status}")
        print(f"[push_sheet] Raw response: {raw_resp[:300]!r}")

        # ── Validasi respons GAS ─────────────────────────────────────────────
        # GAS bisa return plain text ("Replace All Success", "Success", dll)
        # atau JSON {"error": "..."} jika ada masalah di sisi GAS.
        raw_stripped = raw_resp.strip()
        try:
            result = json.loads(raw_stripped)
            # Response adalah JSON
            if "error" in result:
                print(f"[push_sheet] GAS error: {result['error']}")
                raise GSheetResponseError(
                    f"Google Apps Script mengembalikan error:\n{result['error']}"
                )
            # JSON tanpa key 'error' → sukses
            print(f"[push_sheet] ✅ Sukses (JSON response)")
        except json.JSONDecodeError:
            # Plain text response — cek apakah mengandung kata "error"
            if "error" in raw_stripped.lower():
                print(f"[push_sheet] GAS plain-text error: {raw_stripped}")
                raise GSheetResponseError(
                    f"Google Apps Script mengembalikan error:\n{raw_stripped}"
                )
            # Plain text sukses ("Replace All Success", "Success", dll)
            print(f"[push_sheet] ✅ Sukses (plain text response: {raw_stripped!r})")

    except (GSheetResponseError, ConnectionError):
        raise
    except Exception as e:
        print(f"[push_sheet] ❌ Exception: {type(e).__name__}: {e}")
        raise ConnectionError(
            f"Gagal mengirim data ke Google Sheets.\nDetail: {e}"
        )


# ============================================================
# MAIN: Sync ke Google Sheets
# ============================================================

# Kolom yang diisi otomatis dari kode (data AMP)
KOLOM_DRP = [
    "NOMOR PRK", "ITEM PROSES PENGADAAN", "Stock Code", "No Requisisi",
    "Vol", "Satuan", "Tanggal RO", "No RO",
    "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak", "Nilai Kontrak", "Levering",
]

# ============================================================
# Kolom manual = HEADER_DEFAULT dikurangi KOLOM_DRP
# Nilainya dipertahankan dari GSheet saat update.
# ============================================================
_KOLOM_DRP_LOWER = {k.strip().lower() for k in [
    "NOMOR PRK", "ITEM PROSES PENGADAAN", "Stock Code", "No Requisisi",
    "Vol", "Satuan", "Tanggal RO", "No RO",
    "Nilai HPE", "Tanggal Terkontrak", "Nomor Kontrak", "Nilai Kontrak", "Levering",
]}

KOLOM_MANUAL = [
    k for k in HEADER_DEFAULT
    if k.strip().lower() not in _KOLOM_DRP_LOWER
]  # 25 kolom manual


def kirim_df_final_ke_gsheet(df_final, cover_info=None):
    """
    Kirim df_final (hasil proses DRP) ke Google Sheets.

    Alur:
        1. Cek koneksi & keberadaan sheet
        2. Fetch data lama dari GSheet
           - Kosong / sheet baru  → buat struktur HEADER_DEFAULT kosong
           - Ada data             → merge: KOLOM_DRP ditimpa, 25 kolom manual dipertahankan
        3. Reindex ke HEADER_DEFAULT (38 kolom)
        4. Susun per seksi
        5. Push (replace_all)

    Return:
        (n_baru, n_lama, n_hasil)  — jumlah baris untuk info di UI

    Exception:
        ConnectionError           → tidak bisa konek internet
        SheetTidakDitemukanError  → sheet belum dibuat di GSheet
        GSheetResponseError       → GAS menolak data
    """
    tahun      = datetime.now().year
    sheet_name = f"MONITORING DRP {tahun}"

    print(f"\n{'='*60}")
    print(f"[kirim_gsheet] MULAI — sheet target: '{sheet_name}'")
    print(f"[kirim_gsheet] df_final shape: {df_final.shape}")
    print(f"[kirim_gsheet] df_final columns: {list(df_final.columns)}")

    # ── 1. Cek koneksi & sheet ───────────────────────────────────────────────
    print(f"\n[kirim_gsheet] STEP 1: Cek koneksi & keberadaan sheet...")
    exists = check_sheet_exists(WEB_APP_URL, sheet_name)
    print(f"[kirim_gsheet] check_sheet_exists → {exists}")
    if exists is None:
        raise ConnectionError(
            "Tidak dapat terhubung ke Google Sheets.\n"
            "Periksa koneksi internet Anda.")
    if not exists:
        raise SheetTidakDitemukanError(sheet_name)
    print(f"[kirim_gsheet] Sheet '{sheet_name}' ditemukan ✅")

    # ── Siapkan df_baru ──────────────────────────────────────────────────────
    print(f"\n[kirim_gsheet] STEP 1b: Siapkan df_baru dari df_final...")
    df_baru = df_final.copy()
    df_baru = df_baru.drop(columns=["NO"], errors="ignore")

    # Simpan kolom Kategori PRK terpisah untuk keperluan susun_per_seksi nanti
    kategori_map = {}
    if "Kategori PRK" in df_baru.columns:
        df_baru["NOMOR PRK"] = df_baru["NOMOR PRK"].astype(str).str.strip()
        for prk_raw, kat in zip(df_baru["NOMOR PRK"], df_baru["Kategori PRK"].astype(str)):
            key = prk_raw.strip().upper()
            # Decode placeholder Tanpa PRK agar key = nilai Peruntukan bersih
            if key.startswith("__TANPA_PRK__|"):
                key = key[len("__TANPA_PRK__|"):]
            if key:
                kategori_map[key] = kat
    df_baru = df_baru.drop(columns=["Kategori PRK"], errors="ignore")

    rename_map = {}
    for k in KOLOM_DRP:
        actual = next((c for c in df_baru.columns
                       if c.strip().lower() == k.strip().lower()), None)
        if actual and actual != k:
            rename_map[actual] = k
    if rename_map:
        print(f"[kirim_gsheet] Rename kolom: {rename_map}")
        df_baru = df_baru.rename(columns=rename_map)

    kolom_ada = [k for k in KOLOM_DRP if k in df_baru.columns]
    kolom_hilang = [k for k in KOLOM_DRP if k not in df_baru.columns]
    print(f"[kirim_gsheet] KOLOM_DRP ditemukan ({len(kolom_ada)}): {kolom_ada}")
    if kolom_hilang:
        print(f"[kirim_gsheet] KOLOM_DRP tidak ada di df_final: {kolom_hilang}")

    df_baru = df_baru[kolom_ada].copy()
    df_baru = df_baru[
        df_baru["NOMOR PRK"].notna() &
        (df_baru["NOMOR PRK"].astype(str).str.strip() != "")
    ].fillna("").astype(str).reset_index(drop=True)

    n_baru = len(df_baru)
    print(f"[kirim_gsheet] df_baru siap: {n_baru} baris")
    print(f"[kirim_gsheet] kategori_map: {len(kategori_map)} entri")

    # ── 2. Fetch data lama ───────────────────────────────────────────────────
    print(f"\n[kirim_gsheet] STEP 2: Fetch data lama dari GSheet...")
    try:
        df_lama = fetch_sheet(WEB_APP_URL, sheet_name)
        if df_lama is None:
            print(f"[kirim_gsheet] fetch_sheet → None (sheet kosong < 2 baris)")
        elif df_lama.empty:
            print(f"[kirim_gsheet] fetch_sheet → DataFrame kosong (header default dipakai)")
        else:
            print(f"[kirim_gsheet] fetch_sheet → {len(df_lama)} baris, "
                  f"{len(df_lama.columns)} kolom")
            print(f"[kirim_gsheet] Kolom GSheet: {list(df_lama.columns)}")
    except Exception as e:
        print(f"[kirim_gsheet] ❌ fetch_sheet error: {type(e).__name__}: {e}")
        raise

    n_lama = 0

    # ── 3. Gabungkan: kolom DRP dari df_baru, kolom manual dari df_lama ───────
    print(f"\n[kirim_gsheet] STEP 3: Gabungkan data...")
    try:
        if df_lama is None or df_lama.empty:
            # Sheet kosong — langsung pakai df_baru, reindex ke HEADER_DEFAULT
            print(f"[kirim_gsheet] Mode: SHEET KOSONG → langsung pakai df_baru")
            df_merged = df_baru.copy()

        else:
            n_lama = len(df_lama)
            print(f"[kirim_gsheet] Mode: ADA DATA → replace DRP, pertahankan manual")

            # Normalisasi df_lama ke HEADER_DEFAULT (case-insensitive)
            df_lama = _normalisasi_kolom(df_lama, HEADER_DEFAULT)
            for col in HEADER_DEFAULT:
                if col not in df_lama.columns:
                    df_lama[col] = ""
            df_lama = df_lama.reindex(columns=HEADER_DEFAULT, fill_value="")

            # Normalisasi df_baru ke HEADER_DEFAULT juga
            df_baru = _normalisasi_kolom(df_baru, HEADER_DEFAULT)

            # Key: NOMOR PRK uppercase untuk matching
            df_lama["_key"] = df_lama["NOMOR PRK"].astype(str).str.strip().str.upper()
            df_baru["_key"] = df_baru["NOMOR PRK"].astype(str).str.strip().str.upper()

            # Kolom DRP yang ada di HEADER_DEFAULT (nama kanonik)
            kolom_drp_final = [c for c in HEADER_DEFAULT
                               if c.strip().lower() in _KOLOM_DRP_LOWER]
            # Kolom manual = HEADER_DEFAULT dikurangi kolom DRP dan key
            kolom_manual_final = [c for c in HEADER_DEFAULT
                                  if c not in kolom_drp_final and c != "NO"]

            hasil_rows = []

            # ── Pisahkan df_baru: PRK normal vs Tanpa PRK (placeholder) ──────
            # Placeholder format baru: "__TANPA_PRK__|<nilai peruntukan>"
            _TANPA_PRK_MASK = df_baru["_key"].str.startswith("__TANPA_PRK__|")
            df_baru_normal   = df_baru[~_TANPA_PRK_MASK].copy()
            df_baru_tanpa    = df_baru[_TANPA_PRK_MASK].copy()

            # ── Buat lookup baris Tanpa PRK baru berdasarkan Peruntukan ──────
            # Key: nilai peruntukan (diekstrak dari placeholder) uppercase
            # Dipakai untuk matching dengan baris tanpa PRK di df_lama
            tanpa_prk_baru_dict = {}
            for _, rb in df_baru_tanpa.iterrows():
                # Ekstrak nilai Peruntukan dari placeholder key
                raw_key = str(rb.get("NOMOR PRK", rb.get("_key", ""))).strip()
                if raw_key.startswith("__TANPA_PRK__|"):
                    peruntukan_key = raw_key[len("__TANPA_PRK__|"):].strip().upper()
                else:
                    peruntukan_key = raw_key.upper()
                if peruntukan_key and peruntukan_key not in tanpa_prk_baru_dict:
                    tanpa_prk_baru_dict[peruntukan_key] = rb.to_dict()

            tanpa_prk_baru_dipakai = set()  # peruntukan_key yang sudah di-match

            # ── Helper: cek apakah prk_key adalah nilai Peruntukan (bukan nomor PRK) ──
            # Baris tanpa PRK di GSheet punya NOMOR PRK berisi teks Peruntukan
            # yang tidak cocok pola PRK_PATTERN → tidak ada di df_baru_normal
            def _is_tanpa_prk_row(prk_key_upper):
                """True jika prk_key tidak cocok pola PRK dan tidak ada di df_baru_normal."""
                if not prk_key_upper:
                    return True
                # Cek di df_baru_normal — kalau tidak ada, kemungkinan baris tanpa PRK
                return df_baru_normal[df_baru_normal["_key"] == prk_key_upper].empty and \
                       prk_key_upper not in {r.upper() for r in tanpa_prk_baru_dict.keys()} and \
                       not any(prk_key_upper == k for k in
                               df_baru_normal["_key"].str.upper().tolist())

            # Iterasi PRK dari df_lama — pertahankan urutan GSheet
            for _, row_lama in df_lama.iterrows():
                prk_key = row_lama["_key"]

                # ── Baris TANPA PRK dari GSheet ───────────────────────────────
                # Dikenali dari: (1) prk_key kosong, atau
                # (2) prk_key berisi teks Peruntukan yang tidak cocok PRK normal
                match_normal = df_baru_normal[df_baru_normal["_key"] == prk_key]
                is_tanpa = (not prk_key) or (match_normal.empty and prk_key in tanpa_prk_baru_dict)

                if is_tanpa:
                    # ── Tentukan peruntukan_lama ──────────────────────────────────────
                    # Prioritas 1: prk_key langsung (NOMOR PRK di GSheet = nilai Peruntukan)
                    # Prioritas 2: jika prk_key kosong (baris sub-merge di GSheet),
                    #   coba pakai nilai ITEM PROSES PENGADAAN sebagai fallback key
                    #   karena kolom itu tidak di-merge dan bisa jadi identifier
                    peruntukan_lama = prk_key
                    if not peruntukan_lama:
                        item_val = str(row_lama.get("ITEM PROSES PENGADAAN", "")).strip().upper()
                        if item_val and item_val in tanpa_prk_baru_dict:
                            peruntukan_lama = item_val

                    # Coba match dengan baris tanpa PRK baru via Peruntukan
                    if peruntukan_lama and peruntukan_lama in tanpa_prk_baru_dict:
                        # Ada di data baru → update kolom DRP, pertahankan manual
                        rb = tanpa_prk_baru_dict[peruntukan_lama]
                        tanpa_prk_baru_dipakai.add(peruntukan_lama)
                        row_hasil = {}
                        for col in HEADER_DEFAULT:
                            if col == "NOMOR PRK":
                                # Pertahankan nilai Peruntukan di NOMOR PRK
                                row_hasil[col] = row_lama.get(col, "")
                            elif col in kolom_drp_final:
                                row_hasil[col] = rb.get(col, "")
                            else:
                                row_hasil[col] = row_lama.get(col, "")
                    else:
                        # Tidak ada di data baru → pertahankan semua (manual + Peruntukan)
                        row_hasil = {}
                        for col in HEADER_DEFAULT:
                            if col in kolom_drp_final and col != "NOMOR PRK":
                                row_hasil[col] = ""   # kolom DRP dikosongkan
                            else:
                                row_hasil[col] = row_lama.get(col, "")  # manual + Peruntukan dipertahankan
                        # Jika NOMOR PRK masih kosong (sub-baris merge di GSheet),
                        # isi dengan nilai ITEM PROSES PENGADAAN sebagai Peruntukan
                        if not str(row_hasil.get("NOMOR PRK", "")).strip():
                            item_fallback = str(row_lama.get("ITEM PROSES PENGADAAN", "")).strip()
                            if item_fallback:
                                row_hasil["NOMOR PRK"] = item_fallback
                    hasil_rows.append(row_hasil)
                    continue

                # ── Baris PRK normal ──────────────────────────────────────────
                if match_normal.empty:
                    # PRK tidak ada di data baru → kosongkan kolom DRP, pertahankan manual
                    row_hasil = {}
                    for col in HEADER_DEFAULT:
                        if col in kolom_drp_final and col != "NOMOR PRK":
                            row_hasil[col] = ""
                        else:
                            row_hasil[col] = row_lama.get(col, "")
                else:
                    # PRK ada di data baru → replace kolom DRP, pertahankan manual
                    row_baru = match_normal.iloc[0]
                    row_hasil = {}
                    for col in HEADER_DEFAULT:
                        if col == "NOMOR PRK":
                            row_hasil[col] = row_lama[col]
                        elif col in kolom_drp_final:
                            row_hasil[col] = row_baru.get(col, "")
                        else:
                            row_hasil[col] = row_lama.get(col, "")

                hasil_rows.append(row_hasil)

            # ── PRK normal baru yang belum ada di GSheet → tambahkan di bawah ─
            prk_lama_keys = set(df_lama["_key"].tolist())
            for _, row_baru in df_baru_normal.iterrows():
                if row_baru["_key"] not in prk_lama_keys:
                    row_hasil = {col: row_baru.get(col, "") for col in HEADER_DEFAULT}
                    hasil_rows.append(row_hasil)
                    print(f"[kirim_gsheet] PRK baru ditambahkan: {row_baru['_key']}")

            # ── Baris tanpa PRK baru yang belum di-match → tambahkan di bawah ─
            for _, rb in df_baru_tanpa.iterrows():
                raw_key = str(rb.get("NOMOR PRK", rb.get("_key", ""))).strip()
                if raw_key.startswith("__TANPA_PRK__|"):
                    peruntukan_key = raw_key[len("__TANPA_PRK__|"):].strip().upper()
                    peruntukan_val = raw_key[len("__TANPA_PRK__|"):].strip()
                else:
                    peruntukan_key = raw_key.upper()
                    peruntukan_val = raw_key
                if peruntukan_key not in tanpa_prk_baru_dipakai:
                    row_hasil = {col: rb.get(col, "") for col in HEADER_DEFAULT}
                    # Isi NOMOR PRK dengan nilai Peruntukan asli (bukan placeholder)
                    row_hasil["NOMOR PRK"] = peruntukan_val
                    hasil_rows.append(row_hasil)
                    print(f"[kirim_gsheet] Tanpa PRK baru ditambahkan: {peruntukan_val[:60]}")

            df_merged = pd.DataFrame(hasil_rows, columns=HEADER_DEFAULT)
            print(f"[kirim_gsheet] df_merged: {len(df_merged)} baris")

    except Exception as e:
        print(f"[kirim_gsheet] ❌ error: {type(e).__name__}: {e}")
        raise

    n_hasil = len(df_merged)

    # ── 4. Reindex ke 38 kolom HEADER_DEFAULT + nomor urut ───────────────────
    print(f"\n[kirim_gsheet] STEP 4: Reindex ke 38 kolom HEADER_DEFAULT...")
    df_merged = _normalisasi_kolom(df_merged, HEADER_DEFAULT)
    for col in HEADER_DEFAULT:
        if col not in df_merged.columns:
            df_merged[col] = ""
    df_merged = df_merged.reindex(columns=HEADER_DEFAULT, fill_value="")
    df_merged["NO"] = range(1, len(df_merged) + 1)
    df_merged = df_merged.fillna("").astype(str)

    # Tambahkan kembali kolom Kategori PRK dari kategori_map
    # agar susun_per_seksi bisa memfilter seksi TANPA PRK dengan benar
    def _resolve_kategori(prk_val):
        key = str(prk_val).strip().upper()
        # NOMOR PRK kosong → pasti Tanpa PRK
        if not key or key == "NAN":
            return "Tanpa PRK"
        # Placeholder baris Tanpa PRK (format: __TANPA_PRK__|<peruntukan>)
        if key.startswith("__TANPA_PRK__|"):
            return "Tanpa PRK"
        # Cari di kategori_map (key sudah uppercase)
        if key in kategori_map:
            return kategori_map[key]
        # Fallback: klasifikasi ulang via pola PRK
        from drp_app import klasifikasi_prk
        hasil = klasifikasi_prk(prk_val)
        # Jika klasifikasi_prk tidak mengenali pola PRK → ini baris Tanpa PRK
        # (nilai Peruntukan yang tidak mengandung nomor PRK valid)
        if hasil == "AO":
            from drp_app import PRK_PATTERN
            import re as _re
            if not PRK_PATTERN.search(key):
                return "Tanpa PRK"
        return hasil

    df_merged["Kategori PRK"] = df_merged["NOMOR PRK"].apply(_resolve_kategori)
    print(f"[kirim_gsheet] df_merged final: {df_merged.shape}")
    print(f"[kirim_gsheet] Distribusi Kategori PRK:\n"
          f"{df_merged['Kategori PRK'].value_counts().to_dict()}")

    # ── 5. Susun per seksi ───────────────────────────────────────────────────
    print(f"\n[kirim_gsheet] STEP 5: Susun per seksi...")
    print(f"[kirim_gsheet] cover_info: {cover_info}")
    try:
        df_kirim = susun_per_seksi(df_merged, cover_info=cover_info)
        print(f"[kirim_gsheet] df_kirim (inkl label seksi + cover): {len(df_kirim)} baris")
    except Exception as e:
        print(f"[kirim_gsheet] ❌ susun_per_seksi error: {type(e).__name__}: {e}")
        raise

    # ── 6. Push ──────────────────────────────────────────────────────────────
    print(f"\n[kirim_gsheet] STEP 6: Push ke GSheet...")
    try:
        # kolom_diisi=None → push semua 38 kolom HEADER_DEFAULT
        # sehingga GAS replace_all menulis ulang seluruh sheet dengan benar
        push_sheet(WEB_APP_URL, sheet_name, df_kirim, kolom_diisi=None)
        print(f"[kirim_gsheet] ✅ SELESAI — "
              f"lokal={n_baru}, lama={n_lama}, hasil={n_hasil}")
        print(f"{'='*60}\n")
    except Exception as e:
        print(f"[kirim_gsheet] ❌ push_sheet error: {type(e).__name__}: {e}")
        raise

    return n_baru, n_lama, n_hasil


def sync_ke_gsheet(sections, KOLOM_DIPAKAI=None):
    """
    Alur:
        0. Cek sheet ada
        1. Siapkan data baru
        2. Fetch data lama (SEMUA kolom, label seksi dilewati)
        3. Merge
        4. Susun per seksi kategori
        5. Push (replace_all)

    Exception yang mungkin dilempar:
        ConnectionError           → gagal terhubung internet
        SheetTidakDitemukanError  → sheet tidak ada di GSheet
        HeaderTidakDitemukanError → sheet ada tapi header NOMOR PRK tidak ditemukan
        GSheetResponseError       → GAS menolak data saat push
    """
    tahun      = datetime.now().year
    sheet_name = f"MONITORING DRP {tahun}"
    kolom_diisi = KOLOM_DIPAKAI if KOLOM_DIPAKAI else KOLOM_DRP

    # ── 0. Validasi koneksi & keberadaan sheet ───────────────────────────────
    print(f"Memeriksa sheet '{sheet_name}'...")
    exists = check_sheet_exists(WEB_APP_URL, sheet_name)
    if exists is None:
        raise ConnectionError(
            "Tidak dapat terhubung ke Google Sheets.\n"
            "Periksa koneksi internet Anda.")
    if not exists:
        raise SheetTidakDitemukanError(sheet_name)

    # ── 1. Siapkan data baru ─────────────────────────────────────────────────
    df_all = pd.concat(sections, ignore_index=True)
    rename_map = {}
    for k in kolom_diisi:
        target = k.strip().lower()
        actual = next((c for c in df_all.columns
                       if str(c).strip().lower() == target), None)
        if actual and actual != k:
            rename_map[actual] = k
    if rename_map:
        df_all = df_all.rename(columns=rename_map)

    kolom_ada = [k for k in kolom_diisi if k in df_all.columns]
    df_baru   = df_all[kolom_ada].copy()
    df_baru   = df_baru[df_baru["NOMOR PRK"].notna()]
    df_baru   = df_baru[df_baru["NOMOR PRK"].astype(str).str.strip() != ""]
    df_baru   = df_baru.fillna("").astype(str).reset_index(drop=True)
    print(f"Data baru siap: {len(df_baru)} baris")

    # ── 2. Fetch data lama ───────────────────────────────────────────────────
    # fetch_sheet sekarang raise exception eksplisit, tidak silent return None
    print(f"Mengambil data lama dari sheet '{sheet_name}'...")
    df_lama = fetch_sheet(WEB_APP_URL, sheet_name)
    if df_lama is not None and not df_lama.empty:
        print(f"Data lama: {len(df_lama)} baris, {len(df_lama.columns)} kolom")
    else:
        print(f"Sheet kosong/baru, menggunakan header default ({len(HEADER_DEFAULT)} kolom).")

    # ── 3. Merge ─────────────────────────────────────────────────────────────
    print("Menggabungkan data...")
    df_merged = merge_with_existing(
        df_baru, df_lama,
        key_col="NOMOR PRK",
        kolom_diisi=kolom_ada,
    )
    print(f"Hasil merge: {len(df_merged)} baris")

    # ── 4. Susun per seksi ───────────────────────────────────────────────────
    print("Menyusun per seksi kategori...")
    df_final = susun_per_seksi(df_merged)
    print(f"Siap kirim: {len(df_final)} baris (termasuk baris label seksi)")

    # ── 5. Push ──────────────────────────────────────────────────────────────
    print(f"Mengirim ke sheet '{sheet_name}'...")
    push_sheet(WEB_APP_URL, sheet_name, df_final, kolom_diisi=kolom_ada)
    print(f"✅ Sync berhasil → {len(df_merged)} baris data di sheet '{sheet_name}'")
    print(f"   Lama: {len(df_lama) if (df_lama is not None and not df_lama.empty) else 0} | "
          f"Baru: {len(df_baru)} | Hasil: {len(df_merged)}")