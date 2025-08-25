import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.title("📦 Shopee Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロードした後に保護を解除して、保存し直してから、アップロードしてください")
st.image("images/unlock_tip.png", width=600)

# STEP1
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step1.png", width=100)
with col2:
    st.markdown("### 📄 STEP1: mass_upload_basic_info*****.xlsx をアップロード")
    basic_info_path = st.file_uploader(label="", type=["xlsx"], key="basic")

# STEP2
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step2.png", width=100)
with col2:
    st.markdown("### 📄 STEP2: mass_upload_sales_info*****.xlsx をアップロード")
    sales_info_path = st.file_uploader(label="", type=["xlsx"], key="sales")

# STEP3
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step3.png", width=100)
with col2:
    st.markdown("### 📄 STEP3: mass_upload_media_info*****.xlsx をアップロード")
    media_info_path = st.file_uploader(label="", type=["xlsx"], key="media")

# STEP4
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP4: mass_upload_shipment_info*****.xlsx をアップロード")
    shipment_info_path = st.file_uploader(label="", type=["xlsx"], key="shipment")

# STEP5
col1, col2 = st.columns([1, 4])
with col1:
    st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP5: 出品したい国の mass_upload_***_basic_template.xlsx をアップロード")
    template_path = st.file_uploader(label="", type=["xlsx"], key="template")


# 列名正規化関数（|0|0 や |1|1 を削除）
def normalize_columns(cols):
    return [re.sub(r"\|\d+\|\d+$", "", str(c)) for c in cols]


# =============================
# 実行処理
# =============================
if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # データ読み込み
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # 公式の列名を保存
    original_columns = template_df.columns  

    # 正規化した列名に置換して処理用コピーを作成
    template_df_norm = template_df.copy()
    template_df_norm.columns = normalize_columns(template_df.columns)

    # 不足列を追加して original_columns と列数を揃える
    for col in normalize_columns(original_columns):
        if col not in template_df_norm.columns:
            template_df_norm[col] = None

    # 列順も合わせる
    template_df_norm = template_df_norm[normalize_columns(original_columns)]

    # ===== データ埋め込み処理 =====
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)[5:]
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)[5:]
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)[5:]
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)[5:]
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)[5:]
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)[5:]
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)[5:]
    weight_num = shipment_df["et_title_product_weight"].reset_index(drop=True)[5:]

    sgd_to_myr_rate = 3.4
    start_row = 5
    num_ids = len(product_ids)

    rows_needed = start_row + num_ids
    if len(template_df_norm) < rows_needed:
        extra_rows = rows_needed - len(template_df_norm)
        empty_rows = pd.DataFrame([{}] * extra_rows)
        template_df_norm = pd.concat([template_df_norm, empty_rows], ignore_index=True)

    # 値を埋め込む
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_integration_no"] = product_ids.values
    #template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row + num_ids - 1, "et_title_variation_1"] = "type"
    template_df_norm.loc[start_row:start_row + num_ids - 1, "ps_weight"] = weight_num.values
    template_df_norm["ps_price"].iloc[start_row:] = (
        template_df_norm["ps_price"].iloc[start_row:].astype(float) * sgd_to_myr_rate
    ).round(2)

    # === 列名を公式の Shopee テンプレートに戻す ===
    st.write(original_columns)
    st.write("template")
    st.write(template_df_norm.columns)

    
    template_df_norm.columns = original_columns

    # =============================
    # Excel 出力処理
    # =============================
    wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]

    # データをシートに書き込み
    for row_idx, row_data in enumerate(template_df_norm.values, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)

    # メモリに保存
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # ダウンロードボタン
    st.download_button(
        label="📥 処理が完了しました。ここをクリックしてExcelファイルをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    output.close()
