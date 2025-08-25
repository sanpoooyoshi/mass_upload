import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re
import unicodedata
import numpy as np

st.title("📦 Shopee Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロード後に保護解除→保存し直してからアップロードしてください")
st.image("images/unlock_tip.png", width=600)

# STEPごとのアップローダー
basic_info_path   = st.file_uploader("STEP1: basic_info", type=["xlsx"], key="basic")
sales_info_path   = st.file_uploader("STEP2: sales_info", type=["xlsx"], key="sales")
media_info_path   = st.file_uploader("STEP3: media_info", type=["xlsx"], key="media")
shipment_info_path= st.file_uploader("STEP4: shipment_info", type=["xlsx"], key="shipment")
template_path     = st.file_uploader("STEP5: Shopee公式テンプレート", type=["xlsx"], key="template")

# ── 共通: 列名正規化（Shopeeテンプレ特有の |n|n サフィックス除去） ──
def normalize_columns(cols):
    return [re.sub(r"\|\d+\|\d+$", "", str(c)).strip() for c in cols]

# ── 文字列正規化（NFKC・全半角空白を単一半角に・両端strip・改行タブ除去） ──
def clean_name(x):
    if pd.isna(x):
        return None
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return s.strip() or None

# ── product_id を文字列化（Excelの 123456.0 問題を吸収） ──
def to_str_id(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    # 先頭ゼロを保持したい場合があるためint化はしない
    return s

# ── media_info の Option Name/Image の列ペアを 1..30 で厳密対応 ──
def extract_option_pairs(columns):
    name_map = {}
    img_map  = {}
    for c in columns:
        cs = str(c)
        m1 = re.search(r"Option\s*Name\s*(\d+)$", cs, flags=re.I)
        m2 = re.search(r"Option\s*(\d+)\s*Name$", cs, flags=re.I)
        if m1 or m2:
            idx = int((m1 or m2).group(1))
            name_map[idx] = cs
        m3 = re.search(r"Option\s*Image\s*(\d+)$", cs, flags=re.I)
        m4 = re.search(r"Option\s*(\d+)\s*Image$", cs, flags=re.I)
        if m3 or m4:
            idx = int((m3 or m4).group(1))
            img_map[idx] = cs
    pairs = []
    for i in sorted(set(name_map.keys()) & set(img_map.keys())):
        pairs.append((i, name_map[i], img_map[i]))
    return pairs

if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # ====== データ読み込み ======
    basic_df    = pd.read_excel(basic_info_path,    sheet_name="Sheet1")
    sales_df    = pd.read_excel(sales_info_path,    sheet_name="Sheet1")
    media_df    = pd.read_excel(media_info_path,    sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path,      sheet_name="Template")

    # 公式列名（順序保持用）
    original_columns = template_df.columns
    template_df_norm = template_df.copy()
    template_df_norm.columns = normalize_columns(template_df.columns)

    # 不足列を補完（欠けがあれば空列を追加）
    for col in normalize_columns(original_columns):
        if col not in template_df_norm.columns:
            template_df_norm[col] = None
    template_df_norm = template_df_norm[normalize_columns(original_columns)]

    # ====== 取り込み元の基本列（開始行以降で埋める） ======
    start_row = 5
    product_ids       = sales_df["et_title_product_id"].astype(object)
    variation_ids     = sales_df["et_title_variation_id"].astype(object)
    variation_names   = sales_df["et_title_variation_name"].astype(object)
    skus              = sales_df["et_title_variation_sku"].astype(object)
    variation_prices  = sales_df["et_title_variation_price"].astype(object)
    variation_stocks  = sales_df["et_title_variation_stock"].astype(object)
    product_names     = sales_df["et_title_product_name"].astype(object)
    weight_num        = shipment_df["et_title_product_weight"].astype(object)

    # SGD→MYR換算（必要なら調整）
    sgd_to_myr_rate = 3.4

    # 行数計算（sales_dfのデータ部に合わせる）
    n = len(product_ids) - start_row
    if n < 0:
        n = 0
    rows_needed = start_row + n
    if len(template_df_norm) < rows_needed:
        template_df_norm = pd.concat(
            [template_df_norm, pd.DataFrame([{}] * (rows_needed - len(template_df_norm)))],
            ignore_index=True
        )

    # ====== 値の埋め込み（開始行以降） ======
    sl = slice(start_row, start_row + n)

    template_df_norm.loc[sl, "et_title_variation_integration_no"] = product_ids.iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_variation_id"]             = variation_ids.iloc[start_row:].values
    template_df_norm.loc[sl, "ps_product_name"]                   = product_names.iloc[start_row:].values
    template_df_norm.loc[sl, "ps_sku_short"]                      = skus.iloc[start_row:].values
    template_df_norm.loc[sl, "ps_price"]                          = variation_prices.iloc[start_row:].values
    template_df_norm.loc[sl, "ps_stock"]                          = variation_stocks.iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_option_for_variation_1"]   = variation_names.iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_variation_1"]              = "type"
    template_df_norm.loc[sl, "ps_weight"]                         = weight_num.iloc[start_row:].values
    template_df_norm.loc[sl, "channel_id.28057"]                  = "On"

    # 価格換算
    template_df_norm.loc[sl, "ps_price"] = (
        pd.to_numeric(template_df_norm.loc[sl, "ps_price"], errors="coerce") * sgd_to_myr_rate
    ).round(2)

    # ====== media_info: Option 1〜30 Name/Image を縦持ち化 ======
    pairs = extract_option_pairs(media_df.columns)
    media_long_list = []
    for idx, name_col, img_col in pairs:
        tmp = media_df[["et_title_product_id", name_col, img_col]].copy()
        tmp.rename(columns={name_col:"variation_name_raw", img_col:"variation_image"}, inplace=True)
        tmp["product_id"]    = tmp["et_title_product_id"].map(to_str_id)
        tmp["variation_name"]= tmp["variation_name_raw"].map(clean_name)
        tmp["variation_image"]= tmp["variation_image"].astype(str).str.strip()
        tmp = tmp[(tmp["variation_name"].notna()) & (tmp["variation_image"] != "") & (tmp["product_id"].notna())]
        media_long_list.append(tmp[["product_id","variation_name","variation_image"]])
    if media_long_list:
        media_long = pd.concat(media_long_list, ignore_index=True)
    else:
        media_long = pd.DataFrame(columns=["product_id","variation_name","variation_image"])

    # 同一キーで重複があれば最初のURLを採用
    if not media_long.empty:
        media_long = (media_long
                      .sort_values(["product_id","variation_name"])
                      .drop_duplicates(["product_id","variation_name"], keep="first"))

    # ====== sales側の variation_map を正規化（マージキー統一） ======
    variation_map = sales_df[["et_title_product_id","et_title_variation_name"]].copy()
    variation_map["product_id"]     = variation_map["et_title_product_id"].map(to_str_id)
    variation_map["variation_name"] = variation_map["et_title_variation_name"].map(clean_name)
    variation_map = variation_map[["product_id","variation_name"]].dropna().drop_duplicates()

    # ====== 画像URLを variation_map に合流 ======
    img_map = pd.merge(variation_map, media_long, on=["product_id","variation_name"], how="left")
    # 画像が見つからないものは NaN のまま（後段でそのまま入る）

    # ====== Templateにマージして et_title_image_per_variation を埋める ======
    template_df_norm["product_id"]    = template_df_norm["et_title_variation_integration_no"].map(to_str_id)
    template_df_norm["variation_name"]= template_df_norm["et_title_option_for_variation_1"].map(clean_name)

    template_df_norm = template_df_norm.merge(
        img_map[["product_id","variation_name","variation_image"]],
        on=["product_id","variation_name"],
        how="left"
    )

    # ここが肝: 開始行以降のみ画像を書き込む
    template_df_norm.loc[sl, "et_title_image_per_variation"] = template_df_norm.loc[sl, "variation_image"].values

    # ====== 商品説明を統合 ======
    product_description_df = basic_df[["et_title_product_id","et_title_product_description"]].copy()
    product_description_df.rename(columns={"et_title_product_id":"product_id"}, inplace=True)
    product_description_df["product_id"] = product_description_df["product_id"].map(to_str_id)

    template_df_norm = template_df_norm.merge(product_description_df, on="product_id", how="left")
    template_df_norm.loc[sl, "ps_product_description"] = template_df_norm.loc[sl, "et_title_product_description"].values

    # ====== 一時列の後始末 ======
    template_df_norm.drop(columns=[
        "variation_image","et_title_product_description","product_id","variation_name"
    ], inplace=True, errors="ignore")

    # ====== 列順を公式に揃える（余計な列はない状態） ======
    template_df_norm = template_df_norm[normalize_columns(original_columns)]
    # 元のラベル（|n|n付きの表示文字列）に戻す
    template_df_norm.columns = original_columns

    # ====== Excel 書き込み ======
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    # 値だけ上書き（検証/ドロップダウン等の定義は維持）
    for r_idx, row in enumerate(template_df_norm.itertuples(index=False, name=None), start=1):
        for c_idx, value in enumerate(row, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=None if (isinstance(value, float) and np.isnan(value)) else value)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("処理完了：画像URL・説明文がマージされました。")
    st.download_button(
        label="📥 Excelをダウンロード",
        data=buf,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
