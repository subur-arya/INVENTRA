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
    c_due_date = col(mapping,"LEVERING","due_site_date")
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

    data = proses_pjm08(data, mapping)
    data = suplier_order(data, mapping)
    data = perlu_review(data, mapping)
    data = analisis(data, mapping)

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
    due_site_date = col(mapping, "LEVERING", "due_site_date")
    supplier_name = col(mapping, "PO", "supplier_name")

    map_qty_issued = safe_groupby_first(data["SLN"], c_stock_sln, qty_issued)
    map_last_acq_date = safe_groupby_first(data["SLN"], c_stock_sln, last_acq_date)

    map_qty_req = safe_groupby_sum(data["IR"], c_stock_ir, qty_req)
    map_req_by_date = safe_groupby_first(data["IR"], c_stock_ir, req_by_date)

    map_qty_rcv_uop = safe_groupby_first(data["SRD"], c_stock_srd, qty_rcv_uop)
    map_creation_date = safe_groupby_first(data["SRD"], c_stock_srd, creation_date)

    map_curr_qty_p = safe_groupby_sum(data["LEVERING"], c_stock_po, curr_qty_p)
    map_due_site_date = safe_groupby_first(data["LEVERING"], c_stock_po, due_site_date)

    map_supplier_name = safe_groupby_first(data["PO"], c_stock_po, supplier_name)


    data["PLJM08"]['Qty ISS'] = data["PLJM08"][c_stock_pljm08].map(map_qty_issued)
    data["PLJM08"]['Last ISS'] = data["PLJM08"][c_stock_pljm08].map(map_last_acq_date)
    data["PLJM08"]['Next Req Qty'] = data["PLJM08"][c_stock_pljm08].map(map_qty_req)
    data["PLJM08"]['Next Req Date'] = data["PLJM08"][c_stock_pljm08].map(map_req_by_date)
    data["PLJM08"]['Qty SRD'] = data["PLJM08"][c_stock_pljm08].map(map_qty_rcv_uop)
    data["PLJM08"]['Last SRD'] = data["PLJM08"][c_stock_pljm08].map(map_creation_date)
    data["PLJM08"]['levering qty'] = data["PLJM08"][c_stock_pljm08].map(map_curr_qty_p)
    data["PLJM08"]['Due Site Date'] = data["PLJM08"][c_stock_pljm08].map(map_due_site_date)
    data["PLJM08"]['Suplier Name'] = data["PLJM08"][c_stock_pljm08].map(map_supplier_name)

    num_cols = ['Qty ISS','Next Req Qty','Qty SRD','levering qty']
    data["PLJM08"][num_cols] = data["PLJM08"][num_cols].fillna(0)

    calc = (
        data["PLJM08"]['SOH Akhir'].fillna(0)
        - data["PLJM08"]['Next Req Qty'].fillna(0)
        + data["PLJM08"]['levering qty'].fillna(0)
    )

    cond0 = (data["PLJM08"]['ROQ'] == 0) & (data["PLJM08"]['Next Req Qty'] == 0)
    cond01 = (data["PLJM08"]['ROQ'] == 0) & (data["PLJM08"]['Next Req Qty'] != 0)
    cond1 = data["PLJM08"]['levering qty'] > (data["PLJM08"]['ROP'] + data["PLJM08"]['ROQ'])
    cond2 = calc > data["PLJM08"]['ROP']
    cond3 = calc < data["PLJM08"]['ROP']
    cond4 = calc == data["PLJM08"]['ROP']

    data["PLJM08"]['QTY_RO'] = np.select(
        [cond0,cond01,cond1,cond2,cond3,cond4],
        [
            0,
            -1,
            -1,
            0,
            (data["PLJM08"]['ROP']+data["PLJM08"]['ROQ'])-calc,
            data["PLJM08"]['ROQ']
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
# SUPPLIER ORDER
# =====================================================

def suplier_order(data, mapping):

    df_src = data["PLJM08"]

    df_order = df_src.loc[df_src['KLASIFIKASI'] == 'ORDER'].copy()

    kolom_order = [
        'Suplier Name',
        col(mapping,"PLJM08","stock_code"),
        col(mapping,"PLJM08","item_name"),
        'KLASIFIKASI',
        'QTY_RO'
    ]

    kolom_ada = [c for c in kolom_order if c in df_order.columns]

    df_baru = df_order.loc[:, kolom_ada].copy()

    df_baru = df_baru.sort_values(by='Suplier Name', ignore_index=True)

    data["SupplierOrder"] = df_baru

    return data


# =====================================================
# PERLU REVIEW
# =====================================================

def perlu_review(data, mapping):

    df_src = data["PLJM08"]

    df_filter = df_src.loc[df_src['KLASIFIKASI'] == 'PERLU REVIEW'].copy()

    kolom_review = list(df_filter.columns)

    df_perlu_review = df_filter.loc[:, kolom_review].copy()

    df_perlu_review = df_perlu_review.sort_values(by=col(mapping,"PLJM08","stock_code"), ignore_index=True)

    data["PerluReview"] = df_perlu_review

    return data



# =====================================================
# ANALISIS
# =====================================================

def analisis(data, mapping):
    df_src = data["PLJM08"]

    # Ambil hanya ORDER dan PERLU REVIEW
    df_filter = df_src[
        df_src['KLASIFIKASI'].isin(['ORDER', 'PERLU REVIEW'])
    ].copy()

    df_analisis = pd.DataFrame()

    df_analisis["Suplier Name"] = df_filter["Suplier Name"]
    df_analisis["Stock Code"] = df_filter[col(mapping, "PLJM08", "stock_code")]
    df_analisis["Item Name"] = df_filter[col(mapping, "PLJM08", "item_name")]
    df_analisis["EXP"] = df_filter[col(mapping, "PLJM08", "exp")]
    df_analisis["QTY_RO"] = df_filter["QTY_RO"]
    df_analisis["ROP"] = df_filter[col(mapping, "PLJM08", "rop")]
    df_analisis["ROQ"] = df_filter[col(mapping, "PLJM08", "roq")]
    df_analisis["SOH akhir"] = df_filter[col(mapping, "PLJM08", "soh_akhir")]
    df_analisis["Last SRD"] = df_filter["Last SRD"]
    df_analisis["Keterangan Baru"] = df_filter["KLASIFIKASI"]

    map_supplier = col(mapping, "ANALISIS", "supplier_name")
    map_sc = col(mapping, "ANALISIS", "stock_code")
    map_item_name = col(mapping, "ANALISIS", "item_name")
    map_analisis = col(mapping, "ANALISIS", "analisis")

    # Ambil hanya kolom yang diperlukan saja
    df_lama = data["ANALISIS"][
        [map_supplier, map_sc, map_item_name, map_analisis]
    ].copy()

    df_lama = df_lama.rename(columns={
        map_supplier: "Suplier Name lama",
        map_sc: "Stock Code",
        map_item_name: "Item Name lama",
        map_analisis: "Keterangan lama"
    })

    # Samakan tipe dulu
    df_analisis["Stock Code"] = df_analisis["Stock Code"].astype(str).str.zfill(9)
    df_lama[map_sc] = df_lama[map_sc].astype(str).str.zfill(9)

    # Merge berdasarkan Stock Code
    df_merge = df_analisis.merge(
        df_lama,
        left_on="Stock Code",
        right_on=map_sc,
        how="left",
        suffixes=("", "_lama")
    )

    kondisi = [
        (df_merge["Item Name lama"].isna()) | (df_merge["Suplier Name lama"].isna()),
        df_merge["Item Name"] != df_merge["Item Name lama"],
        df_merge["Suplier Name"] != df_merge["Suplier Name lama"],
        df_merge["Keterangan Baru"] != df_merge["Keterangan lama"]
    ]

    hasil = [
        "Perlu di analisis",
        "Item name order baru",
        "Beda supplier",
        "Status analisis berbeda"
    ]

    df_merge["Hasil Perbandingan"] = np.select(
        kondisi,
        hasil,
        default="telah di analisis"
    )

    # df_merge["Keterangan"] = df_merge[map_analisis]

    df_merge["Review Sebelumnya"] = df_merge["Keterangan lama"].fillna("-")

    df_merge = df_merge.drop(
        columns=["Suplier Name lama", "Item Name lama", "Keterangan lama"],
        errors="ignore"
    )

    df_merge = df_merge.sort_values(
        by="Suplier Name",
        ascending=True,
        ignore_index=True
    )

    cols = list(df_merge.columns)
    cols.insert(cols.index("Keterangan Baru"), cols.pop(cols.index("Review Sebelumnya")))
    cols.insert(cols.index("Hasil Perbandingan"), cols.pop(cols.index("Keterangan Baru")))
    df_merge = df_merge[cols]

    data["ANALISIS"] = df_merge

    print("df_lama")
    print(df_lama)
    print("df_merge")
    print(df_merge)

    # ===============================================
    # SPLIT ANALISIS BERDASARKAN KOLOM TERTENTU
    # ===============================================

    kolom_split = "EXP"   # GANTI dengan nama kolom yang kamu mau
    # kolom_split = col(mapping, "PLJM08", "exp")
    # print(kolom_split)

    if kolom_split in df_merge.columns:
        # print("ada")

        unique_values = (
            df_merge[kolom_split]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        for val in unique_values:

            safe_name = (
                val.upper()
                .replace(" ", "_")
                .replace("/", "_")
                .replace("-", "_")
            )

            
            key_name = f"ANALISIS {safe_name}"

            data[key_name] = df_merge[
                df_merge[kolom_split].astype(str).str.strip() == val
            ].copy()

            print(f"Sheet dibuat: {key_name}")

    return data