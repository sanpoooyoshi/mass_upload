import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import re

st.title("Mass Upload Excel作成アプリ")

# 🟡 注意コメント + 補助画像
st.markdown("### ⚠️ STEP1~4に必要なExcelシートは、ダウンロードした後に保護を解除して、保存し直してから、アップロードしてください")
st.image("images/unlock_tip.png", width=400)  # 注意画像


# ============ ファイルアップロード ============ #
col1, col2 = st.columns([1, 4])
with col1: st.image("images/step1.png", width=100)
with col2:
    st.markdown("### 📄 STEP1: mass_upload_basic_info*****.xlsx をアップロード")
    basic_info_path = st.file_uploader("", type=["xlsx"], key="basic")

col1, col2 = st.columns([1, 4])
with col1: st.image("images/step2.png", width=100)
with col2:
    st.markdown("### 📄 STEP2: mass_upload_sales_info*****.xlsx をアップロード")
    sales_info_path = st.file_uploader("", type=["xlsx"], key="sales")

col1, col2 = st.columns([1, 4])
with col1: st.image("images/step3.png", width=100)
with col2:
    st.markdown("### 📄 STEP3: mass_upload_media_info*****.xlsx をアップロード")
    media_info_path = st.file_uploader("", type=["xlsx"], key="media")

col1, col2 = st.columns([1, 4])
with col1: st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP4: mass_upload_shipment_info*****.xlsx をアップロード")
    shipment_info_path = st.file_uploader("", type=["xlsx"], key="shipment")

col1, col2 = st.columns([1, 4])
with col1: st.image("images/step4.png", width=100)
with col2:
    st.markdown("### 📄 STEP5: 出品したい国の mass_upload_***_basic_template.xlsx をアップロード")
    template_path = st.file_uploader("", type=["xlsx"], key="template")


# ============ 正規化関数 ============ #
def normalize_columns(cols):
    """Shopee公式テンプレ列名から |0|0, |1|1 を削除して比較用に正規化"""
    return [re.sub(r"\|\d+\|\d+$", "", str(c)) for c in cols]


# ============ 処理開始 ============ #
if basic_info_path and sales_info_path and media_info_path and shipment_info_path and template_path:
    # 元のテンプレートを読み込み
    template_df = pd.read_excel(template_path, sheet_name="Template", header=2)  # ←オレンジ行をヘッダーに
    original_columns = template_df.columns  # 元の列名（|0|0付き）を保持
    template_df_norm = template_df.copy()
    template_df_norm.columns = normalize_columns(original_columns)

    # 各種データ読み込み
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    shipment_df = pd.read_excel(shipment_info_path, sheet_name="Sheet1")

    # 必要カラム取得
    product_ids = sales_df["et_title_product_id"].iloc[5:].reset_index(drop=True)
    variation_ids = sales_df["et_title_variation_id"].iloc[5:].reset_index(drop=True)
    variation_names = sales_df["et_title_variation_name"].iloc[5:].reset_index(drop=True)
    skus = sales_df["et_title_variation_sku"].iloc[5:].reset_index(drop=True)
    prices = sales_df["et_title_variation_price"].iloc[5:].reset_index(drop=True)
    stocks = sales_df["et_title_variation_stock"].iloc[5:].reset_index(drop=True)
    names = sales_df["et_title_product_name"].iloc[5:].reset_index(drop=True)
    weights = shipment_df["et_title_product_weight"].iloc[5:].reset_index(drop=True)

    start_row = 5
    num_ids = len(product_ids)

    # 空行を追加
    if len(template_df_norm) < start_row + num_ids:
        extra = (start_row + num_ids) - len(template_df_norm)
        template_df_norm = pd.concat([template_df_norm, pd.DataFrame([{}]*extra)], ignore_index=True)

    # 値を転記
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_integration_no"] = product_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_id"] = variation_ids.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_product_name"] = names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_sku_short"] = skus.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_price"] = prices.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_stock"] = stocks.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_option_for_variation_1"] = variation_names.values
    template_df_norm.loc[start_row:start_row+num_ids-1, "et_title_variation_1"] = "type"
    template_df_norm.loc[start_row:start_row+num_ids-1, "ps_weight"] = weights.values

    # --- ここに media_df / basic_df のマージ処理を追加しても良い ---

    # 列名を公式テンプレートのまま戻す
    template_df_norm.columns = original_columns

    # 保存
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template_df_norm.to_excel(writer, index=False, sheet_name="Template")
    output.seek(0)

    # ダウンロード
    st.download_button(
        label="📥 処理が完了しました。Excelをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
