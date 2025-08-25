import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.title("📦 Shopee Mass Upload Excel作成アプリ")

# 注意事項
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロードした後に保護を解除して保存し直してからアップロードしてください")
st.image("images/unlock_tip.png", width=600)

# STEPごとのアップローダー
basic_info_path = st.file_uploader("STEP1: basic_info", type=["xlsx"], key="basic")
sales_info_path = st.file_uploader("STEP2: sales_info", type=["xlsx"], key="sales")
media_info_path = st.file_uploader("STEP3: media_info", type=["xlsx"], key="media")
shipment_info_path = st.file_uploader("STEP4: shipment_info", type=["xlsx"], key="shipment")
template_path = st.file_uploader("STEP5: Shopee公式テンプレート", type=["xlsx"], key="template")

# 列名正規化関数
def normalize_columns(cols):
    return [re.sub(r"\|\d+\|\d+$", "", str(c)) for c in cols]

if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # データ読み込み
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # 公式列名を保存
    original_columns = template_df.columns

    # 正規化して処理用にコピー
    template_df_norm = template_df.copy()
    template_df_norm.columns = normalize_columns(template_df.columns)

    # 不足列を補完（et_title_parent_sku もここで保証される）
    for col in normalize_columns(original_columns):
        if col not in template_df_norm.columns:
            template_df_norm[col] = None
    template_df_norm = template_df_norm[normalize_columns(original_columns)]

    # ===== 各データの抽出 =====
    start_row = 5
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    parent_skus = basic_df["et_title_parent_sku"].reset_index(drop=True)[5:]
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

    # ===== 値を埋め込み =====
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_parent_sku"] = parent_skus.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_1"] = "type"
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_weight"] = weight_num.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "channel_id.28057"] = "On"

    template_df_norm["ps_price"].iloc[start_row:] = (
        template_df_norm["ps_price"].iloc[start_row:].astype(float) * sgd_to_myr_rate
    ).round(2)

    # ===== 画像情報統合 =====
    top_image_df = media_df[[
        "et_title_product_id", "ps_item_cover_image",
        "ps_item_image.1", "ps_item_image.2", "ps_item_image.3",
        "ps_item_image.4", "ps_item_image.5", "ps_item_image.6",
        "ps_item_image.7", "ps_item_image.8"
    ]].copy()
    top_image_df.rename(columns={
        "et_title_product_id": "product_id",
        "ps_item_cover_image": "ps_item_cover_image_"
    }, inplace=True)

    template_df_norm["product_id"] = template_df_norm["et_title_variation_integration_no"]
    merged = pd.merge(template_df_norm, top_image_df, on="product_id", how="left")

    merged["ps_item_cover_image"].iloc[start_row:] = merged["ps_item_cover_image_"].iloc[start_row:]
    for i in range(1, 9):
        merged[f"ps_item_image_{i}"].iloc[start_row:] = merged[f"ps_item_image.{i}"].iloc[start_row:]
        merged.drop(columns=[f"ps_item_image.{i}"], inplace=True)

    # ===== 商品説明統合 =====
    product_description_df = basic_df[["et_title_product_id", "et_title_product_description"]].copy()
    product_description_df.rename(columns={"et_title_product_id": "product_id"}, inplace=True)
    merged = pd.merge(merged, product_description_df, on="product_id", how="left")
    merged["ps_product_description"].iloc[start_row:] = merged["et_title_product_description"].iloc[start_row:]

    # ===== 不要列削除 =====
    merged.drop(columns=[
        "et_title_product_description",
        "ps_item_cover_image_",
        "et_title_variation_id",
        "product_id"
    ], inplace=True, errors="ignore")

    # ===== 列名を公式に戻す =====
    merged.columns = original_columns

    # ===== Excel 出力 =====
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    for row_idx, row_data in enumerate(merged.values, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理が完了しました。ここをクリックしてExcelをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    output.close()
