import pandas as pd
from datetime import datetime
import numpy as np
import json
import os


# =====================================================
# JSON LOADER
# =====================================================

def baca_json():
    if os.path.exists("rpo_settings.json"):
        with open("rpo_settings.json", "r") as f:
            return json.load(f)
    return {}


# =====================================================
# HELPER UNTUK AMBIL NAMA KOLOM DARI JSON
# =====================================================

def col(mapping, section, key):
    return mapping[section]["columns"][key]


# =====================================================
# FILTER SRD
# =====================================================

def filter_srd(df, mapping):

    c_stock = col(mapping,"SRD","stock_code")
    c_qty   = col(mapping,"SRD","qty_rcv_uop")
    c_date  = col(mapping,"SRD","creation_date")

    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    df[c_date] = df[c_date].astype(str).str.zfill(8)
    # df[c_date] = df[c_date].str[:4] + '/' + df[c_date].str[4:6] + '/' + df[c_date].str[6:8]

    df = df.sort_values(by=c_date, ascending=False)

    df[c_qty] = pd.to_numeric(df[c_qty], errors='coerce')
    df = df[df[c_qty] > 0]

    df.reset_index(drop=True, inplace=True)
    return df


# =====================================================
# FILTER SLN
# =====================================================

def filter_sln(df, mapping):

    c_stock = col(mapping,"SLN","stock_code")
    c_qty_issued = col(mapping,"SLN","qty_issued")
    c_qty_req = col(mapping,"SLN","qty_req")
    c_last_acq_date = col(mapping, "SLN", "last_acq_date")

    # date_cols = ['Creation Date', "req_by_date", 'Last Acq Date']
    date_cols = [c_last_acq_date]

    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)
            # df[colx] = df[colx].str[:4] + '/' + df[colx].str[4:6] + '/' + df[colx].str[6:8]

    df = df.sort_values(by=c_last_acq_date, ascending=False)

    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    df[c_qty_issued] = pd.to_numeric(df[c_qty_issued], errors='coerce')
    df[c_qty_req] = pd.to_numeric(df[c_qty_req], errors='coerce')

    df = df[(df[c_qty_req] > 0) & (df[c_qty_issued] > 0)]

    df.reset_index(drop=True, inplace=True)
    return df


# =====================================================
# FILTER IR
# =====================================================

def filter_ir(df, mapping):

    c_stock = col(mapping,"IR","stock_code")
    c_qty_req = col(mapping,"IR","qty_req")
    c_qty_issued = col(mapping,"IR","qty_issued")
    c_req_date = col(mapping,"IR","req_by_date")

    # date_cols = ['Creation Date', c_req_date, 'Last Acq Date']
    date_cols = [c_req_date]

    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)

    df = df.sort_values(by=c_req_date, ascending=True)

    df[c_qty_req] = pd.to_numeric(df[c_qty_req], errors='coerce')
    df[c_qty_issued] = pd.to_numeric(df[c_qty_issued], errors='coerce')

    today = pd.Timestamp.today().strftime("%Y%m%d")

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


# =====================================================
# FILTER PO
# =====================================================

def filter_po(df, mapping):

    c_stock = col(mapping,"PO","stock_code")
    c_order_date = col(mapping,"PO","order_date")
    c_receipt_status = col(mapping,"PO","receipt_status")
    c_qty_rcv_dir = col(mapping,"PO","qty_rcv_dir")
    c_curr_qty = col(mapping,"PO","curr_qty")

    cols = list(df.columns)
    cols.insert(1, cols.pop(cols.index(c_stock)))
    df = df[cols]

    # date_cols = ['ORDER_DATE','DUE_SITE_DATE']
    date_cols = [c_order_date]
    for colx in date_cols:
        if colx in df.columns:
            df[colx] = df[colx].astype(str).str.zfill(8)

    df = df.sort_values(by=c_order_date, ascending=False)

    for c in [c_receipt_status,c_qty_rcv_dir,c_curr_qty]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    df = df[
        (df[c_receipt_status] == 2) &
        (df[c_qty_rcv_dir] != 0) &
        (df[c_curr_qty] != 0)
    ]

    df.reset_index(drop=True, inplace=True)
    return df


# =====================================================
# FILTER LEVERING
# =====================================================

def filter_levering(df, mapping):

    c_stock = col(mapping,"LEVERING","stock_code")
    c_due_date = col(mapping,"LEVERING","levering_date")
    c_receipt_status = col(mapping,"LEVERING","receipt_status")
    c_curr_qty = col(mapping,"LEVERING","curr_qty_p")

    df = df.copy()

    # Reorder kolom
    if c_stock in df.columns:
        cols = list(df.columns)
        cols.insert(1, cols.pop(cols.index(c_stock)))
        df = df[cols]

    # Convert numeric
    for c in [c_receipt_status, c_curr_qty]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # Convert due date dari YYYYMMDD
    if c_due_date in df.columns:
        df[c_due_date] = pd.to_datetime(
            df[c_due_date].astype(str),
            format="%Y%m%d",
            errors="coerce"
        )

    batas_due = pd.Timestamp.today().normalize() - pd.Timedelta(days=60)

    df = df[
        (df[c_receipt_status] == 0) &
        (df[c_curr_qty] != 0) &
        (df[c_due_date] >= batas_due)
    ]

    df = df.sort_values(by=c_due_date, ascending=True)
    df.reset_index(drop=True, inplace=True)

    if c_due_date in df.columns:
        df[c_due_date] = df[c_due_date].dt.strftime("%Y%m%d")

    return df

# =====================================================
# FILTER ALL (ENTRY POINT)
# =====================================================

def filter_all(data, mapping):

    data["SRD"] = filter_srd(data["SRD"], mapping)
    data["SLN"] = filter_sln(data["SLN"], mapping)
    data["IR"] = filter_ir(data["IR"], mapping)
    data["PO"] = filter_po(data["PO"], mapping)
    data["LEVERING"] = filter_levering(data["LEVERING"], mapping)

    data = proses_pljm01(data, mapping)

    # Hapus kolom di PLJM08 yang seluruh isinya kosong (NaN / None)
    data["PLJM08"] = data["PLJM08"].dropna(axis=1, how="all")

    data = proses_pjm08(data, mapping)
    data = analisis_setting(data, mapping)
    data = analisis_non_setting(data, mapping)

    return data


# =====================================================
# VLOOKUP PLJM01 -> PLJM08
# =====================================================

def proses_pljm01(data, mapping):

    c01_stock = col(mapping, "PLJM01", "stock_code")
    c01_rop   = col(mapping, "PLJM01", "rop")
    c01_roq   = col(mapping, "PLJM01", "roq")
    c08_stock = col(mapping, "PLJM08", "stock_code")

    df01 = data["PLJM01"].copy()
    df08 = data["PLJM08"].copy()

    # Normalisasi stock code ke 9 digit
    df01[c01_stock] = df01[c01_stock].astype(str).str.zfill(9)
    df08[c08_stock] = df08[c08_stock].astype(str).str.zfill(9)

    # Ambil hanya 3 kolom yang dibutuhkan dari PLJM01
    df01_slim = df01[[c01_stock, c01_rop, c01_roq]].copy()
    df01_slim = df01_slim.rename(columns={
        c01_rop: "ROP",
        c01_roq: "ROQ"
    })

    # Merge (vlookup) ke PLJM08 berdasarkan stock_code
    df08 = df08.merge(
        df01_slim,
        left_on=c08_stock,
        right_on=c01_stock,
        how="left"
    )

    # Buang kolom duplikat stock_code dari PLJM01 jika nama kolomnya berbeda
    if c01_stock != c08_stock and c01_stock in df08.columns:
        df08 = df08.drop(columns=[c01_stock])

    data["PLJM08"] = df08
    return data


# =====================================================
# PROSES PJM08 (DYNAMIC)
# =====================================================

def proses_pjm08(data, mapping):

    c_stock_pljm08 = col(mapping,"PLJM08","stock_code")
    c_stock_sln = col(mapping,"SLN","stock_code")
    c_stock_ir  = col(mapping,"IR","stock_code")
    c_stock_srd = col(mapping,"SRD","stock_code")
    c_stock_po  = col(mapping,"PO","stock_code")
    c_stock_levering  = col(mapping,"LEVERING","stock_code")

    data["PLJM08"][c_stock_pljm08] = data["PLJM08"][c_stock_pljm08].astype(str).str.zfill(9)
    data["SLN"][c_stock_sln] = data["SLN"][c_stock_sln].astype(str).str.zfill(9)
    data["IR"][c_stock_ir] = data["IR"][c_stock_ir].astype(str).str.zfill(9)
    data["SRD"][c_stock_srd] = data["SRD"][c_stock_srd].astype(str).str.zfill(9)
    data["PO"][c_stock_po] = data["PO"][c_stock_po].astype(str).str.zfill(9)
    data["LEVERING"][c_stock_levering] = data["LEVERING"][c_stock_levering].astype(str).str.zfill(9)

    # ================= GROUPBY =================
    def safe_groupby_first(df, group_col, value_col):
        if df.empty or group_col not in df.columns or value_col not in df.columns:
            return pd.Series(dtype='object')
        return df.groupby(group_col)[value_col].first()


    def safe_groupby_sum(df, group_col, value_col):
        if df.empty or group_col not in df.columns or value_col not in df.columns:
            return pd.Series(dtype='float64')
        return df.groupby(group_col)[value_col].sum()
    
    qty_issued = col(mapping, "SLN", "qty_issued")
    last_acq_date = col(mapping, "SLN", "last_acq_date")
    qty_req = col(mapping, "IR", "qty_req")
    req_by_date = col(mapping, "IR", "req_by_date")
    qty_rcv_uop = col(mapping, "SRD", "qty_rcv_uop")
    creation_date = col(mapping, "SRD", "creation_date")
    curr_qty_p = col(mapping, "LEVERING", "curr_qty_p")
    levering_date = col(mapping, "LEVERING", "levering_date")
    supplier_name = col(mapping, "PO", "supplier_name")

    map_qty_issued = safe_groupby_first(data["SLN"], c_stock_sln, qty_issued)
    map_last_acq_date = safe_groupby_first(data["SLN"], c_stock_sln, last_acq_date)

    map_qty_req = safe_groupby_sum(data["IR"], c_stock_ir, qty_req)
    map_req_by_date = safe_groupby_first(data["IR"], c_stock_ir, req_by_date)

    map_qty_rcv_uop = safe_groupby_first(data["SRD"], c_stock_srd, qty_rcv_uop)
    map_creation_date = safe_groupby_first(data["SRD"], c_stock_srd, creation_date)

    map_curr_qty_p = safe_groupby_sum(data["LEVERING"], c_stock_levering, curr_qty_p)
    map_levering_date = safe_groupby_first(data["LEVERING"], c_stock_levering, levering_date)

    map_supplier_name = safe_groupby_first(data["PO"], c_stock_po, supplier_name)


    data["PLJM08"]['Qty ISS'] = data["PLJM08"][c_stock_pljm08].map(map_qty_issued)
    data["PLJM08"]['Last ISS'] = data["PLJM08"][c_stock_pljm08].map(map_last_acq_date)
    data["PLJM08"]['Next Req Qty'] = data["PLJM08"][c_stock_pljm08].map(map_qty_req)
    data["PLJM08"]['Next Req Date'] = data["PLJM08"][c_stock_pljm08].map(map_req_by_date)
    data["PLJM08"]['Qty SRD'] = data["PLJM08"][c_stock_pljm08].map(map_qty_rcv_uop)
    data["PLJM08"]['Last SRD'] = data["PLJM08"][c_stock_pljm08].map(map_creation_date)
    data["PLJM08"]['levering qty'] = data["PLJM08"][c_stock_pljm08].map(map_curr_qty_p)
    data["PLJM08"]['Levering Date'] = data["PLJM08"][c_stock_pljm08].map(map_levering_date)
    data["PLJM08"]['Suplier Name'] = data["PLJM08"][c_stock_pljm08].map(map_supplier_name)

    num_cols = ['Qty ISS','Next Req Qty','Qty SRD','levering qty']
    data["PLJM08"][num_cols] = data["PLJM08"][num_cols].fillna(0)

    # ROP & ROQ diambil dari kolom hasil vlookup PLJM01
    c_soh   = col(mapping, "PLJM08", "soh_akhir")
    rop_col = "ROP"
    roq_col = "ROQ"

    data["PLJM08"][rop_col] = pd.to_numeric(data["PLJM08"][rop_col], errors='coerce').fillna(0)
    data["PLJM08"][roq_col] = pd.to_numeric(data["PLJM08"][roq_col], errors='coerce').fillna(0)

    calc = (
        data["PLJM08"][c_soh].fillna(0)
        - data["PLJM08"]['Next Req Qty'].fillna(0)
        + data["PLJM08"]['levering qty'].fillna(0)
    )

    cond0  = (data["PLJM08"][roq_col] == 0) & (data["PLJM08"]['Next Req Qty'] == 0)
    # cond01 = (data["PLJM08"][roq_col] == 0) & (data["PLJM08"]['Next Req Qty'] != 0)
    cond1  = data["PLJM08"]['levering qty'] > (data["PLJM08"][rop_col] + data["PLJM08"][roq_col])
    cond2  = calc > data["PLJM08"][rop_col]
    cond3  = calc < data["PLJM08"][rop_col]
    cond4  = calc == data["PLJM08"][rop_col]

    data["PLJM08"]['QTY_RO'] = np.select(
        [cond0, cond1, cond2, cond3, cond4],
        [
            -1,
            -1,
            0,
            (data["PLJM08"][rop_col] + data["PLJM08"][roq_col]) - calc,
            data["PLJM08"][roq_col]
        ],
        default=0
    )

    data["PLJM08"]['KLASIFIKASI'] = np.select(
        [
            data["PLJM08"]['QTY_RO']==-1,
            data["PLJM08"]['QTY_RO']==0,
            data["PLJM08"]['QTY_RO']>0
        ],
        ['PERLU REVIEW','TIDAK ORDER','ORDER'],
        default='-'
    )

    return data


# =====================================================
# HELPER ANALISIS
# =====================================================

def _proses_analisis(data, mapping, key, roq_filter):
    """
    Helper untuk proses analisis.
    key        : "ANALISIS SETTING" atau "ANALISIS NON SETTING"
    roq_filter : fungsi lambda untuk filter ROQ, contoh:
                 lambda df: df["ROQ_PLJM01"] != 0
    """
    df_src = data["PLJM08"]

    # Filter KLASIFIKASI ORDER/PERLU REVIEW + filter ROQ
    df_filter = df_src[
        df_src['KLASIFIKASI'].isin(['ORDER', 'PERLU REVIEW'])
    ].copy()
    df_filter = df_filter[roq_filter(df_filter)].copy()

    df_analisis = pd.DataFrame()
    df_analisis["Suplier Name"]   = df_filter["Suplier Name"]
    df_analisis["Stock Code"]     = df_filter[col(mapping, "PLJM08", "stock_code")]
    df_analisis["Item Name"]      = df_filter[col(mapping, "PLJM08", "item_name")]
    df_analisis["EXP"]            = df_filter[col(mapping, "PLJM08", "exp")]
    df_analisis["QTY_RO"]         = df_filter["QTY_RO"]
    df_analisis["ROP"]            = df_filter["ROP"]
    df_analisis["ROQ"]            = df_filter["ROQ"]
    df_analisis["SOH akhir"]      = df_filter[col(mapping, "PLJM08", "soh_akhir")]
    df_analisis["Keterangan"]     = df_filter["KLASIFIKASI"]

    map_supplier  = col(mapping, key, "supplier_name")
    map_sc        = col(mapping, key, "stock_code")
    map_item_name = col(mapping, key, "item_name")
    map_analisis  = col(mapping, key, "analisis")

    df_lama = data[key][[map_supplier, map_sc, map_item_name, map_analisis]].copy()
    df_lama = df_lama.rename(columns={
        map_supplier:  "Suplier Name lama",
        map_sc:        "Stock Code",
        map_item_name: "Item Name lama",
        map_analisis:  "Keterangan lama"
    })

    df_analisis["Stock Code"] = df_analisis["Stock Code"].astype(str).str.zfill(9)
    df_lama["Stock Code"]     = df_lama["Stock Code"].astype(str).str.zfill(9)

    df_merge = df_analisis.merge(
        df_lama,
        on="Stock Code",
        how="left",
        suffixes=("", "_lama")
    )

    kondisi = [
        (df_merge["Item Name lama"].isna()) | (df_merge["Suplier Name lama"].isna()),
        df_merge["Item Name"] != df_merge["Item Name lama"],
        df_merge["Suplier Name"] != df_merge["Suplier Name lama"],
        df_merge["Keterangan"] != df_merge["Keterangan lama"]
    ]
    hasil = [
        "Perlu di analisis",
        "Item name order baru",
        "Beda supplier",
        "Status analisis berbeda"
    ]

    df_merge["Evaluasi"] = np.select(kondisi, hasil, default="telah di analisis")
    df_merge["Review Sebelumnya"]  = df_merge["Keterangan lama"].fillna("-")

    df_merge = df_merge.drop(
        columns=["Suplier Name lama", "Item Name lama", "Keterangan lama"],
        errors="ignore"
    )

    df_merge = df_merge.sort_values(by="Suplier Name", ascending=True, ignore_index=True)

    cols = list(df_merge.columns)
    cols.insert(cols.index("Keterangan"), cols.pop(cols.index("Review Sebelumnya")))
    cols.insert(cols.index("Evaluasi"), cols.pop(cols.index("Keterangan")))
    df_merge = df_merge[cols]

    data[key] = df_merge
    return data


# =====================================================
# ANALISIS SETTING (ROQ != 0)
# =====================================================

def analisis_setting(data, mapping):
    return _proses_analisis(
        data, mapping,
        key="ANALISIS SETTING",
        roq_filter=lambda df: df["ROQ"] != 0
    )


# =====================================================
# ANALISIS NON SETTING (ROQ == 0 + Next Req Date >= hari ini)
# Header: Suplier Name | Stock Code | Item Name | EXP |
#         QTY_RO | ROP | ROQ | SOH akhir |
#         IR Qty | IR Date | Keterangan
# =====================================================

def analisis_non_setting(data, mapping):

    df_src = data["PLJM08"]
    today_str = pd.Timestamp.today().strftime("%Y%m%d")

    # Filter: KLASIFIKASI ORDER/PERLU REVIEW + ROQ == 0 + Next Req Date >= hari ini
    df_filter = df_src[
        df_src['KLASIFIKASI'].isin(['ORDER', 'PERLU REVIEW'])
    ].copy()

    df_filter = df_filter[df_filter["ROQ"] == 0].copy()

    df_filter = df_filter[
        df_filter["Next Req Date"].notna() &
        (df_filter["Next Req Date"].astype(str).str.strip() >= today_str)
    ].copy()

    # Susun kolom sesuai header yang diminta
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
    df_analisis["IR Qty"]       = df_filter["Next Req Qty"]
    df_analisis["IR Date"]      = df_filter["Next Req Date"]

    # Ambil Keterangan dari data ANALISIS NON SETTING lama (review sebelumnya)
    key = "ANALISIS NON SETTING"
    map_sc       = col(mapping, key, "stock_code")
    map_analisis = col(mapping, key, "analisis")

    df_lama = data[key][[map_sc, map_analisis]].copy()
    df_lama = df_lama.rename(columns={
        map_sc:       "Stock Code",
        map_analisis: "Keterangan lama"
    })

    df_analisis["Stock Code"] = df_analisis["Stock Code"].astype(str).str.zfill(9)
    df_lama["Stock Code"]     = df_lama["Stock Code"].astype(str).str.zfill(9)

    df_merge = df_analisis.merge(
        df_lama,
        on="Stock Code",
        how="left"
    )

    # Kolom Keterangan diisi dari review sebelumnya (lama), default "-"
    df_merge["Keterangan"] = df_merge["Keterangan lama"].fillna("-")
    df_merge = df_merge.drop(columns=["Keterangan lama"], errors="ignore")

    df_merge = df_merge.sort_values(by="Suplier Name", ascending=True, ignore_index=True)

    data[key] = df_merge

    print(f"ANALISIS NON SETTING: {len(df_merge)} baris")
    print(df_merge)

    return data