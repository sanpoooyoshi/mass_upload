import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.title("Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~5に必要なExcelシートは、ダウンロードした後に保護を解除して、保存し直してから、アップロードしてください")
st.image("images/unlock_tip.png", width=600)

# ファイルアップロードUI
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step1.png", width=100)
with col2:
    st.markdown("### 📄 STEP1: mass_upload_basic_info*****.xlsx をアップロード")
    basic_info_path = st.file_uploader(label="", type=["xlsx"], key="basic")

col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step2.png", width=100)
with col2:
    st.markdown("### 📄 STEP2: mass_upload_sales_info*****.xlsx をアップロード")
    sales_info_path = st.file_uploader(label="", type=["xlsx"], key="sales")

col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step3.png", width=100)
with col2:
    st.markdown("### 📄 STEP3: mass_upload_media_info*****.xlsx をアップロード")
    media_info_path = st.file_uploader(label="", type=["xlsx"], key="media")

col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP4: mass_upload_shipment_info*****.xlsx をアップロード")
    shipment_info_path = st.file_uploader(label="", type=["xlsx"], key="shipment")

col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP5: mass_upload_***_basic_template.xlsx をアップロード")
    template_path = st.file_uploader(label="", type=["xlsx"], key="template")

# 列名正規化用関数
def normalize_colname(name):
    if name is None:
        return None
    return re.sub(r"\|\d+\|\d+$", "", str(name))

# すべてアップロードされたら処理開始
if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # ===== 1. データ読み込み =====
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # オリジナル列名を保存（Shopee公式のまま）
    original_columns = list(template_df.columns)

    # 正規化したコピーを作成（内部処理用）
    template_df_norm = template_df.copy()
    template_df_norm.columns = [normalize_colname(c) for c in template_df_norm.columns]
    sales_df.columns = [normalize_colname(c) for c in sales_df.columns]
    media_df.columns = [normalize_colname(c) for c in media_df.columns]
    shipment_df.columns = [normalize_colname(c) for c in shipment_df.columns]

    # ===== 2. 値の抽出 =====
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)[5:]
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)[5:]
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)[5:]
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)[5:]
    weight_num = shipment_df["et_title_product_weight"].reset_index(drop=True)[5:]

    # ===== 3. テンプレートに転記 =====
    sgd_to_myr_rate = 3.4
    start_row = 5
    num_ids = len(product_ids)

    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_weight"] = weight_num.values
    template_df_norm["ps_price"].iloc[5:] = (template_df_norm["ps_price"].iloc[5:].astype(float) * sgd_to_myr_rate).round(2)

    # ===== 4. 出力前に列名を戻す =====
    merged = template_df_norm.copy()
    merged.columns = original_columns  # Shopee公式の列名に復元

    # ===== 5. Excelに保存 =====
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    for row_idx, row_data in enumerate(merged.values, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理が完了しました。Excelファイルをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    output.close()
