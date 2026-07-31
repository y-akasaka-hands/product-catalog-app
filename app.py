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

# リセット処理関数
def reset_app():
    # セッション状態をクリアして画面をリロード
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# ファイルアップローダー（複数ファイルの受け入れ）
uploaded_files = st.file_uploader(
    "「商品カタログ」と「店コード表」をアップロード（複数可）", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True,
    key="file_uploader"
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
            # 1. 店コード表の読み込み (B列:コード, C列:店舗名)
            df_store_raw = pd.read_excel(store_code_file, header=None)
            store_map = {}
            for idx, row in df_store_raw.iterrows():
                if pd.notna(row[1]) and pd.notna(row[2]):
                    raw_code = str(row[1]).strip()
                    if raw_code.replace('.0', '').isdigit():
                        code = str(int(float(raw_code))).zfill(3)
                    else:
                        code = raw_code.zfill(3)
                    name = str(row[2]).strip()
                    store_map[name] = code

            # 2. 商品カタログの読み込み
            df_cat = pd.read_excel(catalog_file, header=None)
            row11 = df_cat.iloc[11].values  # 12行目 (店舗名)
            row12 = df_cat.iloc[12].values  # 13行目 (売上数・売単価など)

            # 3. 「売単価」列の位置を判定
            unit_price_col_idx = -1
            for idx, val in enumerate(row12):
                if pd.notna(val) and str(val).strip() == "売単価":
                    unit_price_col_idx = idx
                    break

            if unit_price_col_idx == -1:
                st.error("❌ 商品カタログ内に「売単価」の列が見つかりませんでした。")
                st.stop()

            # 売単価のデータを数値に変換（14行目以降）
            unit_prices = pd.to_numeric(df_cat.iloc[13:, unit_price_col_idx], errors='coerce').fillna(0)

            # 4. 各店舗の「売上数」列の位置と店舗名を正確に特定
            store_columns = []

            for col_idx in range(len(row11)):
                top_val = str(row11[col_idx]).strip() if pd.notna(row11[col_idx]) else ""
                
                # 12行目に店舗名がしっかり入っている列のみ処理
                if top_val != "" and top_val not in ["品番", "枝番", "取引先"]:
                    for offset in range(3):
                        check_idx = col_idx + offset
                        if check_idx < len(row12):
                            sub_val = str(row12[check_idx]).strip() if pd.notna(row12[check_idx]) else ""
                            if sub_val == "売上数":
                                store_columns.append((top_val, check_idx))
                                break

            # 5. 店舗ごとに「売単価 × 売上数」を算出して合計
            store_results = []
            for store_name, col_idx in store_columns:
                qty_series = pd.to_numeric(df_cat.iloc[13:, col_idx], errors='coerce').fillna(0)
                sales_amount = (qty_series * unit_prices).sum()
                code = store_map.get(store_name, "999")

                store_results.append({
                    "店コード": code,
                    "店舗名": store_name,
                    "売上高（売単価×数量）": sales_amount
                })

            # データフレームの作成・並び替え
            result_df = pd.DataFrame(store_results)
            result_df = result_df.sort_values(by="店コード").reset_index(drop=True)

            # 6. 結果の表示
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
                st.metric("データ合計 売上高", f"¥{int(total_sales):,}")
                
                # Excelダウンロード用データ生成
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='店舗別売上')
                
                st.download_button(
                    label="📥 集計結果（Excel）をダウンロード",
                    data=output.getvalue(),
                    file_name=f"店舗別売上集計_{catalog_file.name}",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

                st.divider()

                # 初期画面（クリア）に戻るボタン
                st.button(
                    "🔄 最初の画面に戻る（選択解除）", 
                    on_click=reset_app, 
                    type="secondary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
