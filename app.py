import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📦 Shopee Mass Upload Excel作成アプリ")

# 1. ファイルアップロード
basic_info_path = st.file_uploader("STEP1: mass_upload_basic_info*****.xlsx をアップロード", type=["xlsx"], key="basic")
sales_info_path = st.file_uploader("STEP2: mass_upload_sales_info*****.xlsx をアップロード", type=["xlsx"], key="sales")
media_info_path = st.file_uploader("STEP3: mass_upload_media_info*****.xlsx をアップロード", type=["xlsx"], key="media")
template_path = st.file_uploader("STEP4: 出品したい国の mass_upload_***_basic_template.xlsx をアップロード", type=["xlsx"], key="template")


# すべてのファイルがアップロードされたときだけ処理を実行
if basic_info_path and sales_info_path and media_info_path:
    # 2. 各ファイルの読み込み
    #basic_df = pd.read_excel(basic_file, sheet_name="Template")
    #sales_df = pd.read_excel(sales_file, sheet_name="Sheet1")
    #media_df = pd.read_excel(media_file)

    from openpyxl import load_workbook
    import openai,os
    from openai import OpenAI
    from dotenv import load_dotenv
    import pandas as pd


    def return_image_list(media_df):
        
        # オプション列のうち、"et_title_option_〇_for_variation_1" のみ抽出
        option_cols = [col for col in media_df.columns if col.startswith("et_title_option_") and col.endswith("_for_variation_1")]
        
        # 結果を格納するリスト
        matched_data = []
        
        # 各 product_id と variation_name を走査
        for pid, vname in zip(product_ids, variation_names):
            # product_id に一致する行のみ抽出
            matching_rows = media_df[media_df["et_title_product_id"] == pid]
        
            for _, row in matching_rows.iterrows():
                for col in option_cols:
                    if row[col] == vname:
                        # カラム名から番号部分（〇）を抽出
                        col_number = col.replace("et_title_option_", "").replace("_for_variation_1", "")
                        matched_data.append({
                            "product_id": pid,
                            "variation_name": vname,
                            "option_column": col,
                            "option_number": col_number
                            
                        })
        
        # 結果をDataFrameに変換
        matched_df = pd.DataFrame(matched_data)
        
        # 各一致データに対して画像列を特定し、画像URLなどを取得
        for entry in matched_data:
            option_number = entry["option_number"]
            product_id = entry["product_id"]
            
            # 対応する画像列名を構築
            image_col = f"et_title_option_image_{option_number}_for_variation_1"
            
            # その product_id に一致する行を取得
            row = media_df[media_df["et_title_product_id"] == product_id]
            
            # 値が取得できる場合は格納
            if not row.empty and image_col in row.columns:
                entry["option_image_column"] = image_col
                entry["image_value"] = row.iloc[0][image_col]
            else:
                entry["option_image_column"] = None
                entry["image_value"] = None
        
        # 再度DataFrame化
        matched_with_images_df = pd.DataFrame(matched_data)
        
        
        return matched_with_images_df


    # ファイルパス
    #basic_info_path = 'sg_basic_info_1101994425_20250410221946.xlsx'
    #sales_info_path = 'sg_sales_info_1101994425_20250410191659.xlsx'
    #shipping_info_path = 'sg_shipping_info_1101994425_20250410191659.xlsx'
    #dts_info_path = 'sg_dts_info_1101994425_20250410191659.xlsx'
    #media_info_path = 'sg_media_info_1101994425_20250410213356.xlsx'
    #template_path = 'my_basic_template_.xlsx'



    # Excelファイルを読み込む 保護ビューになっていると開けないので、シートの保護を解除する
    if template_path is not None:
        wb_output_file = load_workbook(template_path, data_only=True)
    sheet_output_file = wb_output_file["Template"]  

    # 結合セル情報の取得
    merged_cells = sheet_output_file.merged_cells.ranges


    from collections import defaultdict
    # ヘッダー行を取得
    headers = [cell.value for cell in sheet_output_file[1]]

    # `ps_item_image_1` ~ `ps_item_image_8` の列インデックスを取得
    image_columns = [idx + 1 for idx, col in enumerate(headers) if col and col.startswith("ps_item_image")]

    # `ps_sku_short` の列インデックスを取得
    if "ps_sku_short" in headers:
        sku_column_idx = headers.index("ps_sku_short") + 1
    else:
        raise ValueError("ps_sku_short列が見つかりません。")


    # sales_info Excelファイルを読み込む
    wb_sales_info = load_workbook(sales_info_path, data_only=True)
    sheet_sales_info = wb_sales_info.active 





    # 必要なシートの読み込み
    media_df = pd.read_excel(media_info_path, sheet_name="Sheet1")
    sales_df = pd.read_excel(sales_info_path, sheet_name="Sheet1")
    basic_df = pd.read_excel(basic_info_path, sheet_name="Sheet1")
    template_df = pd.read_excel(template_path, sheet_name="Template")
    template_df["et_title_variation_id"]=1
    # カラム名の一覧を確認
    columns_sales_df=sales_df.columns
    columns_template_df=template_df.columns
    columns_media_df=media_df.columns

    # Excelファイルの読み込み（通常1シート目に入っている想定）
    df = pd.read_excel(sales_info_path, sheet_name="Sheet1")  # または sheet_name=0
    # et_title_product_id 列のすべての行を取得
    product_ids = df["et_title_product_id"]
    # 結果を表示（上位5件だけ表示する例）
    print(product_ids.iloc[9])


    # 貼り付ける値
    product_ids = sales_df["et_title_product_id"].reset_index(drop=True)
    product_ids =product_ids[5:]
    # 貼り付ける値
    variation_ids = sales_df["et_title_variation_id"].reset_index(drop=True)
    variation_ids =variation_ids[5:]
    # 貼り付ける値
    variation_names = sales_df["et_title_variation_name"].reset_index(drop=True)
    variation_names =variation_names[5:]
    # 貼り付ける値
    skus = sales_df["et_title_variation_sku"].reset_index(drop=True)
    skus = skus[5:]
    # 貼り付ける値
    variation_prices = sales_df["et_title_variation_price"].reset_index(drop=True)
    variation_prices = variation_prices[5:]
    # 貼り付ける値
    variation_stocks = sales_df["et_title_variation_stock"].reset_index(drop=True)
    variation_stocks = variation_stocks[5:]
    # 貼り付ける値
    product_names = sales_df["et_title_product_name"].reset_index(drop=True)
    product_names = product_names[5:]
    # typeは、どのdfにもなかった
    #variation_1_titles = sales_df["et_title_variation_1"].reset_index(drop=True)
    #variation_1_titles = variation_1_titles[5:]



    sgd_to_myr_rate = 3.4
    # 貼り付けを開始する行番号（0始まりで5行目 → index=4）
    start_row = 5
    num_ids = len(product_ids)

    # 空行を追加（必要なら）
    rows_needed = start_row + num_ids
    if len(template_df) < rows_needed:
        extra_rows = rows_needed - len(template_df)
        empty_rows = pd.DataFrame([{}] * extra_rows)
        template_df = pd.concat([template_df, empty_rows], ignore_index=True)

    # 値の貼り付け
    template_df.loc[start_row:start_row + num_ids - 1, 'et_title_variation_integration_no'] = product_ids.values
    template_df.loc[start_row:start_row + num_ids - 1, 'et_title_variation_id'] = variation_ids.values
    template_df.loc[start_row:start_row + num_ids - 1, 'ps_product_name'] = product_names.values
    template_df.loc[start_row:start_row + num_ids - 1, 'ps_sku_short'] = skus.values
    template_df.loc[start_row:start_row + num_ids - 1, 'ps_price'] = variation_prices.values
    template_df.loc[start_row:start_row + num_ids - 1, 'ps_stock'] = variation_stocks.values
    #template_df.loc[start_row:start_row + num_ids - 1, 'et_title_variation_1'] = variation_1_titles.values
    template_df.loc[start_row:start_row + num_ids - 1, 'et_title_option_for_variation_1'] = variation_names.values
    template_df.loc[start_row:start_row + num_ids - 1, 'et_title_variation_1'] = "type"
    template_df.loc[start_row:start_row + num_ids - 1, 'ps_weight'] = 1
    template_df['ps_price'].iloc[5:] = (template_df['ps_price'].iloc[5:].astype(float)* sgd_to_myr_rate).round(2) 


    images_df=return_image_list(media_df)
    # マージのため、比較対象の列を揃える
    template_df["product_id"] = template_df["et_title_variation_integration_no"]
    template_df["variation_name"] = template_df["et_title_option_for_variation_1"]

    # matched_with_images_df から必要な列だけ抽出
    match_df = images_df[["product_id", "variation_name", "image_value"]]

    # 両方のデータを product_id & variation_name でマージ
    merged = pd.merge(template_df, match_df, on=["product_id", "variation_name"], how="left")

    # image_value を et_title_image_per_variation に格納
    merged["et_title_image_per_variation"].iloc[5:] = merged["image_value"].iloc[5:]





    top_image_df = media_df[["et_title_product_id","ps_item_cover_image", "ps_item_image.1", "ps_item_image.2","ps_item_image.3",
                             "ps_item_image.4","ps_item_image.5","ps_item_image.6","ps_item_image.7","ps_item_image.8"]]
    top_image_df["product_id"]=top_image_df["et_title_product_id"]
    top_image_df["ps_item_cover_image_"]=top_image_df["ps_item_cover_image"]
    top_image_df = top_image_df.drop(columns=["ps_item_cover_image","et_title_product_id"])

    # 両方のデータを product_id & variation_name でマージ
    merged = pd.merge(merged, top_image_df, on=["product_id"], how="left")
    merged["ps_item_cover_image"].iloc[5:] = merged["ps_item_cover_image_"].iloc[5:]
    merged["ps_item_image_1"].iloc[5:] = merged["ps_item_image.1"].iloc[5:]
    merged["ps_item_image_2"].iloc[5:] = merged["ps_item_image.2"].iloc[5:]
    merged["ps_item_image_3"].iloc[5:] = merged["ps_item_image.3"].iloc[5:]
    merged["ps_item_image_4"].iloc[5:] = merged["ps_item_image.4"].iloc[5:]
    merged["ps_item_image_5"].iloc[5:] = merged["ps_item_image.5"].iloc[5:]
    merged["ps_item_image_6"].iloc[5:] = merged["ps_item_image.6"].iloc[5:]
    merged["ps_item_image_7"].iloc[5:] = merged["ps_item_image.7"].iloc[5:]
    merged["ps_item_image_8"].iloc[5:] = merged["ps_item_image.8"].iloc[5:]
    for i in range(8): 
        merged = merged.drop(columns=["ps_item_image."+str(i+1)])
    merged = merged.iloc[2:,:]


    product_description_df = basic_df[["et_title_product_id","et_title_product_description"]]
    product_description_df["product_id"]=product_description_df["et_title_product_id"]
    product_description_df = product_description_df.drop(columns=["et_title_product_id"])
    merged = pd.merge(merged, product_description_df, on=["product_id"], how="left")
    merged["ps_product_description"].iloc[5:] = merged["et_title_product_description"].iloc[5:]
    a=merged

    # 不要な作業用カラムを削除
    merged = merged.drop(columns=["product_id", "variation_name", "image_value","et_title_variation_id",
                                  "et_title_product_description","ps_item_cover_image_"])
    #merged = merged.iloc[2:,:]





    # pd.concatでデータを結合
    updated_data = merged.iloc[2:,:]
    # カラム名を最初の行に挿入
    columns_as_first_row = pd.DataFrame([template_df.columns], columns=template_df.columns)
    updated_data_with_columns = pd.concat([columns_as_first_row, updated_data], ignore_index=True)
            
    if template_path is not None:
        wb = load_workbook(template_path, data_only=True)
    sheet = wb["Template"]  
    # Step 3: openpyxlで更新データを再挿入
    for row_idx, row_data in enumerate(updated_data_with_columns.values, start=1):
        for col_idx, value in enumerate(row_data, start=1):
            sheet.cell(row=row_idx, column=col_idx, value=value)


    # メモリ上に保存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Streamlit ダウンロードボタン
    st.download_button(
        label="📥 処理が完了ました。ここをクリックしてExcelファイルをダウンロード",
        data=output,
        file_name="output_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    
    # 🔚 メモリ解放（使い終わったら閉じる）
    output.close()
    


