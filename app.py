import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.title("Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロードした後に保護を解除して、保存し直してから、アップロードしてください")
st.image("images/unlock_tip.png", width=600)

# STEP1
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step1.png", width=120)
with col2:
    st.markdown("### 📄 STEP1: mass_upload_basic_info*****.xlsx をアップロード")
    basic_info_path = st.file_uploader(label="", type=["xlsx"], key="basic")

# STEP2
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step2.png", width=120)
with col2:
    st.markdown("### 📄 STEP2: mass_upload_sales_info*****.xlsx をアップロード")
    sales_info_path = st.file_uploader(label="", type=["xlsx"], key="sales")

# STEP3
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step3.png", width=120)
with col2:
    st.markdown("### 📄 STEP3: mass_upload_media_info*****.xlsx をアップロード")
    media_info_path = st.file_uploader(label="", type=["xlsx"], key="media")

# STEP4
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=120)
with col2:
    st.markdown("### 📄 STEP4: mass_upload_shipment_info*****.xlsx をアップロード")
    shipment_info_path = st.file_uploader(label="", type=["xlsx"], key="shipment")

# STEP5
col1, col2 = st.columns([1, 5])
with col1:
    st.image("images/step4.png", width=120)
with col2:
    st.markdown("### 📄 STEP5: 出品したい国の mass_upload_***_basic_template.xlsx をアップロード")
    template_path = st.file_uploader(label="", type=["xlsx"], key="template")

# ===== 処理開始 =====
if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # 各シートを読み込み
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # 元の列名を保存（公式フォーマット）
    original_columns = template_df.columns

    # 列名正規化関数
    def normalize_columns(df):
        df = df.copy()
        df.columns = df.columns.str.replace(r"\|\d+\|\d+$", "", regex=True)
        return df

    # 正規化版で処理
    template_df_norm = normalize_columns(template_df)

    # 貼り付け開始行
    start_row = 5
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)[5:]
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)[5:]
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)[5:]
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)[5:]
    weight_num = shipment_df["et_title_product_weight"].reset_index(drop=True)[5:]

    # 値を転記
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_variation_1"] = "type"
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_weight"] = weight_num.values

    # ===== 画像や説明文の統合処理（省略：現行の処理をそのまま使う） =====

    # 最後に公式列名を復元
    template_df_norm.columns = original_columns

    # Excel出力
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    # DataFrameの内容を再書き込み
    for row_idx, row_data in enumerate(template_df_norm.values, start=1):
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
