import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.title("Mass Upload Excel作成アプリ")

# 注意メッセージ
st.markdown("### ⚠️ STEP1~5に必要なExcelシートは、ダウンロード後に保護を解除して保存し直してからアップロードしてください")
st.image("images/unlock_tip.png", width=600)

# ファイルアップロード
basic_info_path = st.file_uploader("STEP1: mass_upload_basic_info*****.xlsx", type=["xlsx"], key="basic")
sales_info_path = st.file_uploader("STEP2: mass_upload_sales_info*****.xlsx", type=["xlsx"], key="sales")
media_info_path = st.file_uploader("STEP3: mass_upload_media_info*****.xlsx", type=["xlsx"], key="media")
shipment_info_path = st.file_uploader("STEP4: mass_upload_shipment_info*****.xlsx", type=["xlsx"], key="shipment")
template_path = st.file_uploader("STEP5: mass_upload_***_basic_template.xlsx", type=["xlsx"], key="template")

# 列名正規化
def normalize_colname(name):
    if name is None:
        return None
    return re.sub(r"\|\d+\|\d+$", "", str(name))

if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:

    # ===== 1. 読み込み =====
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")

    # オリジナル列名保存
    original_columns = list(template_df.columns)

    # 正規化して内部処理用にコピー
    template_df_norm = template_df.copy()
    template_df_norm.columns = [normalize_colname(c) for c in template_df_norm.columns]
    sales_df.columns = [normalize_colname(c) for c in sales_df.columns]
    media_df.columns = [normalize_colname(c) for c in media_df.columns]
    shipment_df.columns = [normalize_colname(c) for c in shipment_df.columns]

    # ===== 2. データ抽出 =====
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
    rows_needed = start_row + len(product_ids)

    # ===== 3. 空行補充 =====
    if len(template_df_norm) < rows_needed:
        extra_rows = rows_needed - len(template_df_norm)
        empty_rows = pd.DataFrame([{}] * extra_rows)
        template_df_norm = pd.concat([template_df_norm, empty_rows], ignore_index=True)

    # ===== 4. データ転記 =====
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_product_name"] = product_names.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_price"] = variation_prices.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_stock"] = variation_stocks.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row+len(product_ids)-1, "ps_weight"] = weight_num.values

    # 通貨換算
    template_df_norm["ps_price"].iloc[5:] = (
        template_df_norm["ps_price"].iloc[5:].astype(float) * sgd_to_myr_rate
    ).round(2)

    # ===== 5. 出力用に列名を戻す =====
    merged = template_df_norm.copy()
    merged.columns = original_columns

    # ===== 6. Excel保存 =====
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
