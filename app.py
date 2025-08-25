import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.title("Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロードした後に保護を解除して、保存し直してから、アップロードしてください")
st.image("images/unlock_tip.png", width=600)  # 注意画像


# STEP1
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step1.png", width=120)
with col2:
    st.markdown("### 📄 STEP1-1: mass_upload_basic_info*****.xlsx をダウンロード")
    st.markdown("### 📄 STEP1-2: mass_upload_basic_info*****.xlsx をアップロード")
    basic_info_path = st.file_uploader(label="", type=["xlsx"], key="basic")

# STEP2
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step2.png", width=120)
with col2:
    st.markdown("### 📄 STEP2-1: mass_upload_sales_info*****.xlsx をダウンロード")
    st.markdown("### 📄 STEP2-2: mass_upload_sales_info*****.xlsx をアップロード")
    sales_info_path = st.file_uploader(label="", type=["xlsx"], key="sales")

# STEP3
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step3.png", width=120)
with col2:
    st.markdown("### 📄 STEP3-1: mass_upload_media_info*****.xlsx をダウンロード")
    st.markdown("### 📄 STEP3-2: mass_upload_media_info*****.xlsx をアップロード")
    media_info_path = st.file_uploader(label="", type=["xlsx"], key="media")

# STEP4
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=120)
with col2:
    st.markdown("### 📄 STEP4-1: mass_upload_shipment_info*****.xlsx をダウンロード")
    st.markdown("### 📄 STEP4-2: mass_upload_shipment_info*****.xlsx をアップロード")
    shipment_info_path = st.file_uploader(label="", type=["xlsx"], key="shipment")

# STEP5
col1, col2 = st.columns([1, 5])
with col1:
    st.image("images/step4.png", width=120)
with col2:
    st.markdown("### 📄 STEP5-1: mass_upload_***_basic_template.xlsx をダウンロード")
    st.markdown("### 📄 STEP5-2: mass_upload_***_basic_template.xlsx をアップロード")
    template_path = st.file_uploader(label="", type=["xlsx"], key="template")


# ====== メイン処理 ======
if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # --- DataFrame 読み込み ---
    media_df = pd.read_excel(media_info_path, sheet_name=0)
    sales_df = pd.read_excel(sales_info_path, sheet_name=0)
    basic_df = pd.read_excel(basic_info_path, sheet_name=0)
    shipment_df = pd.read_excel(shipment_info_path, sheet_name=0)
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # 元の列名（Shopee公式）
    original_columns = template_df.columns.copy()

    # --- 各列を取得 ---
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)[5:]
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)[5:]
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)[5:]
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)[5:]
    weight_num = shipment_df["et_title_product_weight"].reset_index(drop=True)[5:]

    # --- 貼り付け ---
    start_row = 5
    num_ids = len(product_ids)

    # 足りない行を追加
    rows_needed = start_row + num_ids
    if len(template_df) < rows_needed:
        extra_rows = rows_needed - len(template_df)
        empty_rows = pd.DataFrame([{}] * extra_rows)
        template_df = pd.concat([template_df, empty_rows], ignore_index=True)

    sgd_to_myr_rate = 3.4

    template_df.loc[start_row:start_row+num_ids-1, "et_title_variation_integration_no"] = product_ids.values
    template_df.loc[start_row:start_row+num_ids-1, "et_title_variation_id"] = variation_ids.values
    template_df.loc[start_row:start_row+num_ids-1, "ps_product_name"] = product_names.values
    template_df.loc[start_row:start_row+num_ids-1, "ps_sku_short"] = skus.values
    template_df.loc[start_row:start_row+num_ids-1, "ps_price"] = variation_prices.values
    template_df.loc[start_row:start_row+num_ids-1, "ps_stock"] = variation_stocks.values
    template_df.loc[start_row:start_row+num_ids-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df.loc[start_row:start_row+num_ids-1, "et_title_variation_1"] = "type"
    template_df.loc[start_row:start_row+num_ids-1, "ps_weight"] = weight_num.values

    # SGD → MYR 変換
    template_df.loc[start_row:, "ps_price"] = (
        template_df.loc[start_row:, "ps_price"].astype(float) * sgd_to_myr_rate
    ).round(2)

    # --- 画像結合 ---
    def return_image_list(media_df):
        option_cols = [col for col in media_df.columns if col.startswith("et_title_option_") and col.endswith("_for_variation_1")]
        matched_data = []
        for pid, vname in zip(product_ids, variation_names):
            matching_rows = media_df[media_df["et_title_product_id"] == pid]
            for _, row in matching_rows.iterrows():
                for col in option_cols:
                    if row[col] == vname:
                        col_number = col.replace("et_title_option_", "").replace("_for_variation_1", "")
                        image_col = f"et_title_option_image_{col_number}_for_variation_1"
                        matched_data.append({
                            "product_id": pid,
                            "variation_name": vname,
                            "image_value": row.get(image_col, None)
                        })
        return pd.DataFrame(matched_data)

    images_df = return_image_list(media_df)
    template_df["product_id"] = template_df["et_title_variation_integration_no"]
    template_df["variation_name"] = template_df["et_title_option_for_variation_1"]
    merged = pd.merge(template_df, images_df, on=["product_id", "variation_name"], how="left")
    merged.loc[start_row:, "et_title_image_per_variation"] = merged.loc[start_row:, "image_value"]

    # --- 商品説明を統合 ---
    product_description_df = basic_df[["et_title_product_id", "et_title_product_description"]].rename(
        columns={"et_title_product_id": "product_id"}
    )
    merged = pd.merge(merged, product_description_df, on="product_id", how="left")
    merged.loc[start_row:, "ps_product_description"] = merged.loc[start_row:, "et_title_product_description"]

    # --- 不要カラム削除 ---
    merged = merged.drop(columns=["product_id", "variation_name", "image_value", "et_title_product_description"], errors="ignore")

    # 列名をShopee公式のまま戻す
    merged = merged.reindex(columns=original_columns)

    # --- Excel 出力 ---
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    for row_idx, row_data in enumerate(merged.values, start=2):  # ヘッダーはテンプレートそのまま
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理が完了しました。ここをクリックしてExcelファイルをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    output.close()
