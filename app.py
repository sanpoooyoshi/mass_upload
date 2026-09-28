# -*- coding: utf-8 -*-
# 📦 Shopee Mass Upload Excel作成アプリ（フル版・自動流し込み対応）
# - media_info: Option 1〜30 Name/Image を縦持ち化（実列名に対応）
# - product_id × variation_name（正規化）で突合 → Image per Variation へ自動流し込み
# - 商品説明を結合
# - 未マッチ一覧CSV / media候補カタログCSVを出力
# - ヘッダー破損防止のため、Excelは「2行目から」書き込み

import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Shopee Mass Upload Builder", layout="wide")

# === Googleログインゲート ===
GOOGLE_LOGO_B64 = (
    "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0OCA0OCIgd2lkdGg9IjIwIiBoZWlnaHQ9IjIwIj4"
    "KPHBhdGggZmlsbD0iI0ZGQzEwNyIgZD0iTTQzLjYxMSwyMC4wODNINDJWMjBIMjR2OGgxMS4zMDNjLTEuNjQ5LDQuNjU3LTYuMDgsOC0xMS4zMDMs"
    "OGMtNi42MjcsMC0xMi01LjM3My0xMi0xMmMwLTYuNjI3LDUuMzczLTEyLDEyLTEyYzMuMDU5LDAsNS44NDIsMS4xNTQsNy45NjEsMy4wMzlsNS42"
    "NTctNS42NTdDMzQuMDQ2LDYuMDUzLDI5LjI2OCw0LDI0LDRDMTIuOTU1LDQsNCwxMi45NTUsNCwyNGMwLDExLjA0NSw4Ljk1NSwyMCwyMCwyMGMx"
    "MS4wNDUsMCwyMC04Ljk1NSwyMC0yMEM0NCwyMi42NTksNDMuODYyLDIxLjM1LDQzLjYxMSwyMC4wODN6Ii8+CjxwYXRoIGZpbGw9IiNGRjNEMDAi"
    "IGQ9Ik02LjMwNiwxNC42OTFsNi41NzEsNC44MTlDMTQuNjU1LDE1LjEwOCwxOC45NjEsMTIsMjQsMTJjMy4wNTksMCw1Ljg0MiwxLjE1NCw3Ljk2"
    "MSwzLjAzOWw1LjY1Ny01LjY1N0MzNC4wNDYsNi4wNTMsMjkuMjY4LDQsMjQsNEMxNi4zMTgsNCw5LjY1Niw4LjMzNyw2LjMwNiwxNC42OTF6Ii8+"
    "CjxwYXRoIGZpbGw9IiM0Q0FGNTAiIGQ9Ik0yNCw0NGM1LjE2NiwwLDkuODYtMS45NzcsMTMuNDA5LTUuMTkybC02LjE5LTUuMjM4QzI5LjIxMSwz"
    "NS4wOTEsMjYuNzE1LDM2LDI0LDM2Yy01LjIwMiwwLTkuNjE5LTMuMzE3LTExLjI4My03Ljk0NmwtNi41MjIsNS4wMjVDOS41MDUsMzkuNTU2LDE2"
    "LjIyNyw0NCwyNCw0NHoiLz4KPHBhdGggZmlsbD0iIzE5NzZEMiIgZD0iTTQzLjYxMSwyMC4wODNINDJWMjBIMjR2OGgxMS4zMDNjLTAuNzkyLDIu"
    "MjM3LTIuMjMxLDQuMTY2LTQuMDg3LDUuNTcxYzAuMDAxLTAuMDAxLDAuMDAyLTAuMDAxLDAuMDAzLTAuMDAybDYuMTksNS4yMzhDMzYuOTcxLDM5"
    "LjIwNSw0NCwzNCw0NCwyNEM0NCwyMi42NTksNDMuODYyLDIxLjM1LDQzLjYxMSwyMC4wODN6Ii8+Cjwvc3ZnPg=="
)

try:
    is_logged_in = st.user.is_logged_in
    auth_configured = True
except Exception:
    is_logged_in = False
    auth_configured = False

if not is_logged_in:
    st.markdown(
        f"""
        <style>
        div[data-testid="stAppViewContainer"] {{ background-color: #0e0e0e; }}
        .login-card {{
            max-width: 420px;
            margin: 12vh auto 0 auto;
            text-align: center;
        }}
        .login-card h1 {{
            color: #f5f5f5;
            font-size: 1.6rem;
            margin-bottom: 0.4rem;
        }}
        .login-card p {{
            color: #9a9a9a;
            margin-bottom: 2rem;
        }}
        .st-key-google_login button {{
            background-color: #1e1e1e !important;
            color: #f5f5f5 !important;
            border: 1px solid #3c3c3c !important;
            border-radius: 999px !important;
            padding: 0.6rem 1.5rem !important;
            font-weight: 500 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 10px !important;
        }}
        .st-key-google_login button::before {{
            content: "";
            background-image: url("data:image/svg+xml;base64,{GOOGLE_LOGO_B64}");
            background-size: contain;
            background-repeat: no-repeat;
            width: 20px;
            height: 20px;
            display: inline-block;
        }}
        .st-key-google_login button:hover {{
            border-color: #5f5f5f !important;
            color: #ffffff !important;
        }}
        </style>
        <div class="login-card">
            <h1>📦 Shopee Mass Upload Builder</h1>
            <p>続けるにはGoogleアカウントでログインしてください</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        if auth_configured:
            st.button("Googleで登録", key="google_login", on_click=st.login, use_container_width=True)
        else:
            st.button("Googleで登録", key="google_login", use_container_width=True, disabled=True)
            st.caption(
                "⚠️ Google認証が未設定です。`.streamlit/secrets.toml.example` を参考に "
                "`.streamlit/secrets.toml` を作成し、Google OAuthクライアント情報を設定してください。"
            )
    st.stop()

with st.sidebar:
    user_label = st.user.get("email") or st.user.get("name") or "ユーザー"
    st.markdown(f"👤 **{user_label}**")
    if st.button("ログアウト"):
        st.logout()

st.title("📦 Shopee Mass Upload Excel作成アプリ（自動流し込み対応）")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4のExcelは、**保護解除→保存し直し**のうえアップロードしてください。")
st.image("images/unlock_tip.png", width=640)

# === アップローダ ===
basic_info_path    = st.file_uploader("STEP1: basic_info", type=["xlsx"], key="basic")
sales_info_path    = st.file_uploader("STEP2: sales_info", type=["xlsx"], key="sales")
media_info_path    = st.file_uploader("STEP3: media_info", type=["xlsx"], key="media")
shipment_info_path = st.file_uploader("STEP4: shipment_info", type=["xlsx"], key="shipment")
template_path      = st.file_uploader("STEP5: Shopee公式テンプレート", type=["xlsx"], key="template")

# === ユーティリティ ===
INSTRUCTION_ROWS = 5  # mass_update_* の上部説明行数

def normalize_columns(cols):
    """Shopeeテンプレ特有の |n|n サフィックスを除去"""
    return [re.sub(r"\|\d+\|\d+$", "", str(c)).strip() for c in cols]

def clean_name(x):
    """バリエーション名の正規化"""
    if pd.isna(x): return None
    s = unicodedata.normalize("NFKC", str(x)).replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).replace("\r"," ").replace("\n"," ").replace("\t"," ")
    return s.strip() or None

def to_str_id(x):
    """Excel数値→文字列化（12345.0 → 12345）"""
    if pd.isna(x): return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return s

def find_image_per_variation_col(norm_cols):
    """テンプレ内の『Image per Variation』列（正規化名）を自動検出"""
    for p in [r"(?i)image\s*per\s*variation", r"(?i)et_title_image_per_variation"]:
        for c in norm_cols:
            if re.search(p, c):
                return c
    return None

def detect_pairs_media_columns(media_df):
    """media_info の実列名に合わせた Name/Image の30ペアを抽出"""
    pairs = []
    for i in range(1, 30 + 1):
        name_col = f"et_title_option_{i}_for_variation_1"
        img_col  = f"et_title_option_image_{i}_for_variation_1"
        if name_col in media_df.columns and img_col in media_df.columns:
            pairs.append((i, name_col, img_col))
    return pairs

def df_to_csv_bytes(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO(); df.to_csv(buf, index=False, encoding="utf-8-sig"); buf.seek(0); return buf

# === メイン処理 ===
if all([basic_info_path, sales_info_path, media_info_path, shipment_info_path, template_path]):

    # ===== 取り込み =====
    try:
        basic_df    = pd.read_excel(basic_info_path,    sheet_name="Sheet1")
        sales_df    = pd.read_excel(sales_info_path,    sheet_name="Sheet1")
        media_df    = pd.read_excel(media_info_path,    sheet_name="Sheet1")
        shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
        template_df = pd.read_excel(template_path,      sheet_name="Template")
    except Exception as e:
        st.error(f"Excel読み込みでエラー: {e}")
        st.stop()

    # 公式列名（順序保持）
    original_columns     = template_df.columns
    original_cols_normal = normalize_columns(original_columns)

    # テンプレを処理用に正規化
    template_df_norm = template_df.copy()
    template_df_norm.columns = original_cols_normal

    # === 行拡張（核心） ===
    start = INSTRUCTION_ROWS                 # データ開始位置（説明行の直後）
    n = max(0, len(sales_df) - start)        # データ行数
    rows_needed = start + n
    if len(template_df_norm) < rows_needed:
        template_df_norm = pd.concat(
            [template_df_norm, pd.DataFrame([{}] * (rows_needed - len(template_df_norm)))],
            ignore_index=True
        )

    # 欠けている公式列があれば空で補完
    for col in original_cols_normal:
        if col not in template_df_norm.columns:
            template_df_norm[col] = None
    template_df_norm = template_df_norm[original_cols_normal]

    sl = slice(start, start + n)  # 以降「開始行以降のみ」上書き

    # ===== sales / shipment の値を埋め込み =====
    template_df_norm.loc[sl, "et_title_variation_integration_no"] = sales_df["et_title_product_id"].iloc[start:].values
    template_df_norm.loc[sl, "et_title_variation_id"]             = sales_df["et_title_variation_id"].iloc[start:].values
    template_df_norm.loc[sl, "ps_product_name"]                   = sales_df["et_title_product_name"].iloc[start:].values
    template_df_norm.loc[sl, "ps_sku_short"]                      = sales_df["et_title_variation_sku"].iloc[start:].values
    template_df_norm.loc[sl, "ps_price"]                          = sales_df["et_title_variation_price"].iloc[start:].values
    template_df_norm.loc[sl, "ps_stock"]                          = sales_df["et_title_variation_stock"].iloc[start:].values
    template_df_norm.loc[sl, "et_title_option_for_variation_1"]   = sales_df["et_title_variation_name"].iloc[start:].values
    if "et_title_variation_1" in template_df_norm.columns:
        template_df_norm.loc[sl, "et_title_variation_1"]          = "type"
    if "ps_weight" in template_df_norm.columns and "et_title_product_weight" in shipment_df.columns:
        template_df_norm.loc[sl, "ps_weight"] = shipment_df["et_title_product_weight"].iloc[start:].values
    if "channel_id.28057" in template_df_norm.columns:
        template_df_norm.loc[sl, "channel_id.28057"] = "On"

    # 価格換算オプション
    with st.expander("💱 価格換算（SGD→MYR）"):
        do_fx   = st.checkbox("SGD→MYRを掛ける", value=True)
        fx_rate = st.number_input("レート", value=3.4, step=0.1, format="%.2f")
    if do_fx and "ps_price" in template_df_norm.columns:
        template_df_norm.loc[sl, "ps_price"] = pd.to_numeric(
            template_df_norm.loc[sl, "ps_price"], errors="coerce"
        ).mul(fx_rate).round(2)

    # ===== media_info: Option 1〜30 Name/Image を縦持ち化 =====
    pairs = detect_pairs_media_columns(media_df)
    media_long_list = []
    for _, name_col, img_col in pairs:
        tmp = media_df[["et_title_product_id", name_col, img_col]].copy()
        tmp.rename(columns={name_col: "variation_name_raw", img_col: "variation_image"}, inplace=True)
        tmp["product_id"]      = tmp["et_title_product_id"].map(to_str_id)
        tmp["variation_name"]  = tmp["variation_name_raw"].map(clean_name)
        tmp["variation_image"] = tmp["variation_image"].astype(str).str.strip()
        tmp = tmp[(tmp["product_id"].notna()) & (tmp["variation_name"].notna()) & (tmp["variation_image"] != "")]
        media_long_list.append(tmp[["product_id", "variation_name", "variation_image"]])
    media_long = pd.concat(media_long_list, ignore_index=True) if media_long_list else pd.DataFrame(columns=["product_id","variation_name","variation_image"])
    # 数値以外の product_id（ヘッダー混入など）を排除 & 重複キーは最初を採用
    media_long = media_long[media_long["product_id"].str.fullmatch(r"\d+")]
    media_long = (media_long
                  .sort_values(["product_id", "variation_name"])
                  .drop_duplicates(["product_id", "variation_name"], keep="first"))

    # ===== sales側キーを正規化（突合用） =====
    variation_map = sales_df[["et_title_product_id", "et_title_variation_name"]].copy()
    variation_map["product_id"]     = variation_map["et_title_product_id"].map(to_str_id)
    variation_map["variation_name"] = variation_map["et_title_variation_name"].map(clean_name)
    variation_map = variation_map[["product_id", "variation_name"]].dropna().drop_duplicates()

    # ======== ★ 自動流し込みロジック ★ ========
    # 1) (product_id × variation_name) で media_long と結合し URL を得る
    img_map = variation_map.merge(media_long, on=["product_id", "variation_name"], how="left")

    # 2) テンプレ側にも同じキー列を作り、img_map を合流
    template_df_norm["product_id"]     = template_df_norm["et_title_variation_integration_no"].map(to_str_id)
    template_df_norm["variation_name"] = template_df_norm["et_title_option_for_variation_1"].map(clean_name)
    template_df_norm = template_df_norm.merge(
        img_map[["product_id", "variation_name", "variation_image"]],
        on=["product_id", "variation_name"],
        how="left"
    )

    # 3) 実在の「Image per Variation」列を特定し、開始行以降へURLを書き込み
    image_per_var_col = find_image_per_variation_col(original_cols_normal)
    if not image_per_var_col:
        st.error("テンプレートに『Image per Variation』列が見つかりません。")
        st.stop()
    template_df_norm.loc[sl, image_per_var_col] = template_df_norm.loc[sl, "variation_image"].values

    # ===== 商品説明を結合 =====
    desc_df = basic_df[["et_title_product_id", "et_title_product_description"]].copy()
    desc_df.rename(columns={"et_title_product_id": "product_id"}, inplace=True)
    desc_df["product_id"] = desc_df["product_id"].map(to_str_id)
    template_df_norm = template_df_norm.merge(desc_df, on="product_id", how="left")
    if "ps_product_description" in template_df_norm.columns:
        template_df_norm.loc[sl, "ps_product_description"] = template_df_norm.loc[sl, "et_title_product_description"].values

    # 一時列の後始末 & 列順復元
    template_df_norm.drop(columns=["variation_image", "et_title_product_description", "product_id", "variation_name"],
                          inplace=True, errors="ignore")
    template_df_norm = template_df_norm[original_cols_normal]
    template_df_norm.columns = original_columns

    # ===== 未マッチCSV / media候補カタログCSV =====
    variation_map_dbg = sales_df[["et_title_product_id", "et_title_product_name", "et_title_variation_name"]].copy()
    variation_map_dbg["product_id"] = variation_map_dbg["et_title_product_id"].map(to_str_id)
    variation_map_dbg["variation_name_clean"] = variation_map_dbg["et_title_variation_name"].map(clean_name)
    variation_map_dbg = (variation_map_dbg[["product_id", "et_title_product_name", "et_title_variation_name", "variation_name_clean"]]
                         .dropna(subset=["product_id", "variation_name_clean"])
                         .drop_duplicates()
                         .rename(columns={"et_title_product_name": "product_name",
                                          "et_title_variation_name": "variation_name_sales_raw"}))

    img_map_dbg = variation_map.merge(media_long, on=["product_id", "variation_name"], how="left")
    img_map_dbg["has_url"] = img_map_dbg["variation_image"].notna() & (img_map_dbg["variation_image"] != "")

    media_candidates = (media_long.groupby("product_id")["variation_name"]
                        .agg(lambda s: " | ".join(sorted(set([x for x in s if isinstance(x, str)]))))
                        .reset_index()
                        .rename(columns={"variation_name": "media_candidates"}))

    unmatched = img_map_dbg[~img_map_dbg["has_url"]][["product_id", "variation_name"]].copy()
    unmatched = unmatched.rename(columns={"variation_name": "variation_name_clean"})
    unmatched = unmatched.merge(variation_map_dbg, on=["product_id", "variation_name_clean"], how="left")
    unmatched = unmatched.merge(media_candidates, on="product_id", how="left")
    unmatched = unmatched[["product_id", "product_name", "variation_name_sales_raw", "variation_name_clean", "media_candidates"]]

    # ===== Excelへ安全に書き戻し（2行目から書く） =====
    wb = load_workbook(template_path, data_only=True)
    sh = wb["Template"]
    for r_idx, row in enumerate(template_df_norm.itertuples(index=False, name=None), start=2):
        for c_idx, val in enumerate(row, start=1):
            if isinstance(val, float) and np.isnan(val): val = None
            sh.cell(row=r_idx, column=c_idx, value=val)

    out_xlsx = BytesIO(); wb.save(out_xlsx); out_xlsx.seek(0)

    # CSVバッファ
    unmatched_csv = df_to_csv_bytes(unmatched)
    catalog_csv   = df_to_csv_bytes(media_long[["product_id", "variation_name", "variation_image"]])

    # ===== 画面出力 =====
    matched_cnt = int(img_map_dbg["has_url"].sum())
    total_keys  = len(img_map_dbg)
    st.success(f"✅ 完了：画像URLマッチ {matched_cnt}/{total_keys}（{matched_cnt/total_keys:.2%}）")
    st.download_button("📥 Excelをダウンロード", data=out_xlsx,
                       file_name="shopee_mass_upload_output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ 未マッチ一覧CSVをダウンロード", data=unmatched_csv,
                       file_name="unmatched_variations.csv", mime="text/csv")
    st.download_button("⬇️ media候補カタログCSVをダウンロード", data=catalog_csv,
                       file_name="media_variation_catalog.csv", mime="text/csv")

else:
    st.info("上の5つのExcelファイルをすべて選択すると処理を実行します。")