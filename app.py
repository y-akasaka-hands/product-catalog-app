import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="商品カタログ 店舗別売上集計", layout="wide")
st.title("📊 商品カタログ 店舗別売上集計ツール")

st.markdown("""
**使い方:**
1. 下記のエリアに「**⑩店コード表.xlsx**」と「**商品カタログ_〜.xlsx**」を2つ同時にドラッグ＆ドロップしてください。
2. 自動的に「売単価 × 売上数」で計算された店舗別売上高が集計されます。
""")

# ファイルアップローダー（複数ファイルの受け入れ）
uploaded_files = st.file_uploader(
    "「商品カタログ」と「店コード表」をアップロード（複数可）", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    catalog_file = None
    store_code_file = None

    # ファイルの判別
    for f in uploaded_files:
        if "店コード表" in f.name:
            store_code_file = f
        elif "商品カタログ" in f.name:
            catalog_file = f

    if not store_code_file:
        st.warning("⚠️ 「店コード表」が選択されていません。")
    if not catalog_file:
        st.warning("⚠️ 「商品カタログ」が選択されていません。")

    if store_code_file and catalog_file:
        try:
            # 1. 店コード表の読み込み
            df_store_raw = pd.read_excel(store_code_file, header=None)
            store_map = {}
            for idx, row in df_store_raw.iterrows():
                if pd.notna(row[1]) and pd.notna(row[2]):
                    code = str(int(row[1])).zfill(3) if str(row[1]).isdigit() else str(row[1]).zfill(3)
                    name = str(row[2]).strip()
                    store_map[name] = code

            # 2. 商品カタログの読み込み
            df_cat = pd.read_excel(catalog_file, header=None)
            row11 = df_cat.iloc[11].values # 12行目
            row12 = df_cat.iloc[12].values # 13行目

            # 列名のマッピング
            current_group = ""
            combined_cols = []
            for g, c in zip(row11, row12):
                if pd.notna(g) and str(g).strip() != "":
                    current_group = str(g).strip()
                c_str = "" if pd.isna(c) else str(c).strip()

                if current_group and c_str in ["売上数", "売上高", "在庫数"]:
                    combined_cols.append(f"{current_group}_{c_str}")
                else:
                    combined_cols.append(c_str)

            # 実データの抽出
            df_clean = df_cat.iloc[13:].copy()
            df_clean.columns = combined_cols
            df_clean['売単価'] = pd.to_numeric(df_clean['売単価'], errors='coerce').fillna(0)

            # 店舗ごとの計算
            store_results = []
            for col in df_clean.columns:
                if col.endswith('_売上数'):
                    store_name = col.replace('_売上数', '')
                    qty = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
                    sales_amount = (qty * df_clean['売単価']).sum()

                    code = store_map.get(store_name, "999")
                    store_results.append({
                        "店コード": code,
                        "店舗名": store_name,
                        "売上高（売単価×数量）": sales_amount
                    })

            result_df = pd.DataFrame(store_results)
            result_df = result_df.sort_values(by="店コード").reset_index(drop=True)

            # 結果の表示
            st.success("✅ 集計が完了しました！")

            col1, col2 = st.columns([2, 1])
            with col1:
                st.subheader("📋 店舗別売上高一覧")
                st.dataframe(
                    result_df.style.format({"売上高（売単価×数量）": "¥{:,.0f}"}),
                    use_container_width=True
                )

            with col2:
                st.subheader("📊 概況")
                total_sales = result_df["売上高（売単価×数量）"].sum()
                st.metric("総合計 売上高", f"¥{int(total_sales):,}")

                # Excelダウンロード用データ生成
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='店舗別売上')

                st.download_button(
                    label="📥 集計結果（Excel）をダウンロード",
                    data=output.getvalue(),
                    file_name=f"店舗別売上集計_{catalog_file.name}",
                    mime="application/vnd.ms-excel"
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
