import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.title("📦 Shopee Mass Upload Excel作成アプリ")

st.markdown("### ⚠️ STEP1~5のExcelはダウンロード後に保護解除＆保存し直してからアップロードしてください")
st.image("images/unlock_tip.png", width=600)

# アップローダー
basic_info_path = st.file_uploader("STEP1: basic_info", type=["xlsx"], key="basic")
sales_info_path = st.file_uploader("STEP2: sales_info", type=["xlsx"], key="sales")
media_info_path = st.file_uploader("STEP3: media_info", type=["xlsx"], key="media")
shipment_info_path = st.file_uploader("STEP4: shipment_info", type=["xlsx"], key="shipment")
template_path = st.file_uploader("STEP5: Shopee公式テンプレート", type=["xlsx"], key="template")

# 列名正規化
def normalize_columns(cols):
    return [re.sub(r"\|\d+\|\d+$", "", str(c)) for c in cols]

# Option Image マッピング
def build_option_image_map(media_df):
    image_map = {}
    for _, row in media_df.iterrows():
        pid = row["et_title_product_id"]
        for i in range(1, 31):
            opt_name_col = f"Option {i} Name"
            opt_img_col  = f"Option {i} Image"
            if opt_name_col in media_df.columns and opt_img_col in media_df.columns:
                opt_name = str(row[opt_name_col]).strip()
                opt_img  = row[opt_img_col]
                if opt_name and opt_name != "nan":
                    image_map[(pid, opt_name)] = opt_img
    return image_map

if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # データ読み込み
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # 公式列名と正規化列名
    original_columns = template_df.columns
    template_df_norm = template_df.copy()
    template_df_norm.columns = normalize_columns(template_df.columns)

    # データ抽出
    start_row = 5
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    parent_skus = basic_df[["et_title_product_id", "et_title_parent_sku"]].copy()
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)[5:]
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)[5:]
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)[5:]
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)[5:]
    weight_num = shipment_df["et_title_product_weight"].reset_index(drop=True)[5:]

    sgd_to_myr_rate = 3.4
    num_ids = len(product_ids)

    rows_needed = start_row + num_ids
    if len(template_df_norm) < rows_needed:
        template_df_norm = pd.concat(
            [template_df_norm, pd.DataFrame([{}] * (rows_needed - len(template_df_norm)))],
            ignore_index=True
        )

    # 値の埋め込み
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_1"] = "type"
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_weight"] = weight_num.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "channel_id.28057"] = "On"
    template_df_norm["ps_price"].iloc[start_row:] = (
        template_df_norm["ps_price"].iloc[start_row:].astype(float) * sgd_to_myr_rate
    ).round(2)

    # Option Image マッピング
    image_map = build_option_image_map(media_df)
    for idx in range(start_row, start_row+num_ids):
        pid = template_df_norm.loc[idx, "et_title_variation_integration_no"]
        vname = template_df_norm.loc[idx, "et_title_option_for_variation_1"]
        if (pid, vname) in image_map:
            template_df_norm.loc[idx, "et_title_image_per_variation"] = image_map[(pid, vname)]

    # 親SKU統合
    parent_skus.rename(columns={"et_title_product_id": "product_id"}, inplace=True)
    template_df_norm["product_id"] = template_df_norm["et_title_variation_integration_no"]
    merged = pd.merge(template_df_norm, parent_skus, on="product_id", how="left")
    merged.loc[start_row:start_row+num_ids-1, "ps_sku_parent_short"] = merged["et_title_parent_sku"].iloc[start_row:]

    # 商品説明統合
    product_description_df = basic_df[["et_title_product_id", "et_title_product_description"]].copy()
    product_description_df.rename(columns={"et_title_product_id": "product_id"}, inplace=True)
    merged = pd.merge(merged, product_description_df, on="product_id", how="left")
    merged["ps_product_description"].iloc[start_row:] = merged["et_title_product_description"].iloc[start_row:]

    # 不要列削除
    merged.drop(columns=[
        "et_title_product_description",
        "et_title_variation_id",
        "product_id",
        "et_title_parent_sku"
    ], inplace=True)

    # 列名を公式に戻す
    merged.columns = original_columns

    # Excel 出力
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged.to_excel(writer, index=False, sheet_name="Template")

    output.seek(0)

    st.download_button(
        label="📥 完了しました。Excelをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    output.close()
