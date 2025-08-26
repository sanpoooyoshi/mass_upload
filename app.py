import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re, unicodedata, numpy as np

st.title("📦 Shopee Mass Upload Excel作成アプリ")
st.markdown("### ⚠️ STEP1~4のExcelは保護解除→保存し直してからアップロードしてください")
st.image("images/unlock_tip.png", width=600)

# ── アップローダ ──
basic_info_path    = st.file_uploader("STEP1: basic_info", type=["xlsx"], key="basic")
sales_info_path    = st.file_uploader("STEP2: sales_info", type=["xlsx"], key="sales")
media_info_path    = st.file_uploader("STEP3: media_info", type=["xlsx"], key="media")
shipment_info_path = st.file_uploader("STEP4: shipment_info", type=["xlsx"], key="shipment")
template_path      = st.file_uploader("STEP5: Shopee公式テンプレート", type=["xlsx"], key="template")

# ── ユーティリティ ──
def normalize_columns(cols):
    return [re.sub(r"\|\d+\|\d+$", "", str(c)).strip() for c in cols]

def clean_name(x):
    if pd.isna(x): return None
    s = unicodedata.normalize("NFKC", str(x))
    s = s.replace("\u3000"," ")
    s = re.sub(r"\s+"," ", s).replace("\r"," ").replace("\n"," ").replace("\t"," ")
    return s.strip() or None

def to_str_id(x):
    if pd.isna(x): return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return s

def extract_option_pairs(columns):
    name_map, img_map = {}, {}
    for c in columns:
        cs = str(c)
        m1 = re.search(r"Option\s*Name\s*(\d+)$", cs, flags=re.I) or re.search(r"Option\s*(\d+)\s*Name$", cs, flags=re.I)
        m2 = re.search(r"Option\s*Image\s*(\d+)$", cs, flags=re.I) or re.search(r"Option\s*(\d+)\s*Image$", cs, flags=re.I)
        if m1: name_map[int(m1.group(1))] = cs
        if m2: img_map[int(m2.group(1))]  = cs
    return [(i, name_map[i], img_map[i]) for i in sorted(set(name_map)&set(img_map))]

def find_image_per_variation_col(norm_cols):
    # 公式テンプレの実在列（正規化後）をパターンで探索
    pats = [r"(?i)image\s*per\s*variation", r"(?i)et_title_image_per_variation", r"(?i)ps.*image.*per.*variation"]
    for p in pats:
        for c in norm_cols:
            if re.search(p, c): return c
    return None

if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # ====== 読み込み ======
    basic_df    = pd.read_excel(basic_info_path,    sheet_name="Sheet1")
    sales_df    = pd.read_excel(sales_info_path,    sheet_name="Sheet1")
    media_df    = pd.read_excel(media_info_path,    sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path,      sheet_name="Template")

    # 列準備
    original_columns     = template_df.columns
    original_cols_normal = normalize_columns(original_columns)
    template_df_norm     = template_df.copy()
    template_df_norm.columns = original_cols_normal

    # 欠損公式列の補完
    for col in original_cols_normal:
        if col not in template_df_norm.columns:
            template_df_norm[col] = None
    template_df_norm = template_df_norm[original_cols_normal]

    # ====== データ抽出（開始行5以降） ======
    start_row = 5
    n = max(0, len(sales_df) - start_row)
    sl = slice(start_row, start_row + n)

    template_df_norm.loc[sl, "et_title_variation_integration_no"] = sales_df["et_title_product_id"].iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_variation_id"]             = sales_df["et_title_variation_id"].iloc[start_row:].values
    template_df_norm.loc[sl, "ps_product_name"]                   = sales_df["et_title_product_name"].iloc[start_row:].values
    template_df_norm.loc[sl, "ps_sku_short"]                      = sales_df["et_title_variation_sku"].iloc[start_row:].values
    template_df_norm.loc[sl, "ps_price"]                          = sales_df["et_title_variation_price"].iloc[start_row:].values
    template_df_norm.loc[sl, "ps_stock"]                          = sales_df["et_title_variation_stock"].iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_option_for_variation_1"]   = sales_df["et_title_variation_name"].iloc[start_row:].values
    template_df_norm.loc[sl, "et_title_variation_1"]              = "type"
    template_df_norm.loc[sl, "ps_weight"]                         = shipment_df["et_title_product_weight"].iloc[start_row:].values
    template_df_norm.loc[sl, "channel_id.28057"]                  = "On"

    # 価格換算
    template_df_norm.loc[sl, "ps_price"] = pd.to_numeric(template_df_norm.loc[sl, "ps_price"], errors="coerce").mul(3.4).round(2)

    # ====== media_info: Option1～30 Name/Imageを縦持ち ======
    pairs = extract_option_pairs(media_df.columns)
    media_long_list = []
    for _, name_col, img_col in pairs:
        tmp = media_df[["et_title_product_id", name_col, img_col]].copy()
        tmp.rename(columns={name_col:"variation_name_raw", img_col:"variation_image"}, inplace=True)
        tmp["product_id"]     = tmp["et_title_product_id"].map(to_str_id)
        tmp["variation_name"] = tmp["variation_name_raw"].map(clean_name)
        tmp["variation_image"]= tmp["variation_image"].astype(str).str.strip()
        tmp = tmp[(tmp["product_id"].notna()) & (tmp["variation_name"].notna()) & (tmp["variation_image"]!="")]
        media_long_list.append(tmp[["product_id","variation_name","variation_image"]])
    media_long = pd.concat(media_long_list, ignore_index=True) if media_long_list else pd.DataFrame(columns=["product_id","variation_name","variation_image"])
    if not media_long.empty:
        media_long = media_long.sort_values(["product_id","variation_name"]).drop_duplicates(["product_id","variation_name"], keep="first")

    # ====== キーを統一してマージ ======
    variation_map = sales_df[["et_title_product_id","et_title_variation_name"]].copy()
    variation_map["product_id"]     = variation_map["et_title_product_id"].map(to_str_id)
    variation_map["variation_name"] = variation_map["et_title_variation_name"].map(clean_name)
    variation_map = variation_map[["product_id","variation_name"]].dropna().drop_duplicates()

    img_map = variation_map.merge(media_long, on=["product_id","variation_name"], how="left")

    # Template側キーを用意
    template_df_norm["product_id"]     = template_df_norm["et_title_variation_integration_no"].map(to_str_id)
    template_df_norm["variation_name"] = template_df_norm["et_title_option_for_variation_1"].map(clean_name)

    # 実在する「Image per Variation」列名（正規化後）を特定
    image_per_var_col = find_image_per_variation_col(original_cols_normal)
    if image_per_var_col is None:
        st.error("テンプレートに『Image per Variation』列が見つかりません。テンプレの列名をご確認ください。")
        st.stop()

    # 画像URLを合流して、公式の Image per Variation 列へ書き込み
    template_df_norm = template_df_norm.merge(img_map, on=["product_id","variation_name"], how="left")
    template_df_norm.loc[sl, image_per_var_col] = template_df_norm.loc[sl, "variation_image"].values

    # ====== 商品説明を結合 ======
    desc_df = basic_df[["et_title_product_id","et_title_product_description"]].copy()
    desc_df.rename(columns={"et_title_product_id":"product_id"}, inplace=True)
    desc_df["product_id"] = desc_df["product_id"].map(to_str_id)
    template_df_norm = template_df_norm.merge(desc_df, on="product_id", how="left")
    template_df_norm.loc[sl, "ps_product_description"] = template_df_norm.loc[sl, "et_title_product_description"].values

    # ====== 一時列の後始末と列順復元 ======
    template_df_norm.drop(columns=["variation_image","et_title_product_description","product_id","variation_name"], inplace=True, errors="ignore")
    template_df_norm = template_df_norm[original_cols_normal]
    template_df_norm.columns = original_columns

    # ====== 出力 ======
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]
    for r_idx, row in enumerate(template_df_norm.itertuples(index=False, name=None), start=1):
        for c_idx, value in enumerate(row, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=None if (isinstance(value, float) and np.isnan(value)) else value)

    buf = BytesIO(); wb.save(buf); buf.seek(0)
    st.success("処理完了：『Image per Variation』にURLを投入しました。")
    st.download_button("📥 Excelをダウンロード", data=buf, file_name="output_file.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 参考: マッチ率の簡易診断
    matched = template_df_norm.iloc[sl][image_per_var_col].notna().sum()
    total   = n
    st.info(f"画像URLマッチ: {matched}/{total}")