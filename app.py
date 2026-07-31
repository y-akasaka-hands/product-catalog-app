import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="商品カタログ 店舗別売上集計", layout="wide")
st.title("📊 商品カタログ 店舗別売上集計ツール")

st.markdown("""
**使い方:**
1. 下記のエリアに「**⑩店コード表.xlsx**」と「**商品カタログ_〜.xlsx**」を2つ同時にドラッグ＆ドロップしてください。
2. 「**🚀 集計を開始する**」ボタンを押すと、店舗別・取引先別・品番別の売上高が集計されます。
""")

# セッション状態の初期化
if "run_calc" not in st.session_state:
    st.session_state.run_calc = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# 完全リセット処理関数
def reset_app():
    st.session_state.run_calc = False
    st.session_state.uploader_key += 1  # uploaderのキーを変更してファイル選択を完全クリア
    st.rerun()


# ファイルアップローダー（キー可変でクリア可能に）
uploaded_files = st.file_uploader(
    "「商品カタログ」と「店コード表」をアップロード（複数可）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.uploader_key}",
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

    # ファイルが2つ揃っている場合
    if store_code_file and catalog_file:
        # 集計未実行の場合は「集計開始ボタン」を表示
        if not st.session_state.run_calc:
            st.info("💡 ファイルの準備ができました。下のボタンを押して集計を開始してください。")
            if st.button("🚀 集計を開始する", type="primary", use_container_width=True):
                st.session_state.run_calc = True
                st.rerun()

        # 集計実行フラグがONのときに処理を行う
        if st.session_state.run_calc:
            try:
                # 1. 店コード表の読み込み (B列:コード, C列:店舗名)
                df_store_raw = pd.read_excel(store_code_file, header=None)
                store_map = {}
                for idx, row in df_store_raw.iterrows():
                    if pd.notna(row[1]) and pd.notna(row[2]):
                        raw_code = str(row[1]).strip()
                        if raw_code.replace(".0", "").isdigit():
                            code = str(int(float(raw_code))).zfill(3)
                        else:
                            code = raw_code.zfill(3)
                        name = str(row[2]).strip()
                        store_map[name] = code

                # 2. 商品カタログの読み込み
                df_cat = pd.read_excel(catalog_file, header=None)
                row11 = df_cat.iloc[11].values  # 12行目 (店舗名など)
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

                # 実データ部分（14行目以降）
                df_data = df_cat.iloc[13:].copy()
                unit_prices = pd.to_numeric(df_data[unit_price_col_idx], errors="coerce").fillna(0)

                # 4. 各店舗の「売上数」列の位置と店舗名を特定
                store_columns = []
                for col_idx in range(len(row11)):
                    top_val = str(row11[col_idx]).strip() if pd.notna(row11[col_idx]) else ""
                    if top_val != "" and top_val not in ["品番", "枝番", "取引先"]:
                        for offset in range(3):
                            check_idx = col_idx + offset
                            if check_idx < len(row12):
                                sub_val = str(row12[check_idx]).strip() if pd.notna(row12[check_idx]) else ""
                                if sub_val == "売上数":
                                    store_columns.append((top_val, check_idx))
                                    break

                # 5. 店舗別集計 ＆ 行（商品）ごとの全店売上金額の計算
                row_total_sales = pd.Series(0.0, index=df_data.index)
                store_results = []

                for store_name, col_idx in store_columns:
                    qty_series = pd.to_numeric(df_data[col_idx], errors="coerce").fillna(0)
                    sales_amount_series = qty_series * unit_prices
                    row_total_sales += sales_amount_series  # 行ごとの売上金額を加算

                    code = store_map.get(store_name, "999")
                    store_results.append({
                        "店コード": code,
                        "店舗名": store_name,
                        "売上高（売単価×数量）": sales_amount_series.sum(),
                    })

                # 店舗別データフレーム作成・並び替え
                df_result_store = pd.DataFrame(store_results)
                df_result_store = df_result_store.sort_values(by="店コード").reset_index(drop=True)

                # 6. 取引先別集計 (取引先コード: Col 8, 取引先名: Col 9)
                df_data["vendor_code"] = df_data[8].apply(
                    lambda x: str(int(float(x))).zfill(7)
                    if pd.notna(x) and str(x).replace(".0", "").isdigit()
                    else str(x).strip().zfill(7)
                    if pd.notna(x)
                    else "0000000"
                )
                df_data["vendor_name"] = df_data[9].astype(str).str.strip()
                df_data["row_sales"] = row_total_sales

                df_result_vendor = (
                    df_data.groupby(["vendor_code", "vendor_name"])["row_sales"]
                    .sum()
                    .reset_index()
                )
                df_result_vendor.columns = ["取引先コード", "取引先名", "売上高（売単価×数量）"]
                df_result_vendor = df_result_vendor.sort_values(by="取引先コード").reset_index(drop=True)

                # 7. 品番別集計 (品番コード: Col 2, 品番名: Col 3)
                df_data["item_code"] = df_data[2].apply(
                    lambda x: str(int(float(x)))
                    if pd.notna(x) and str(x).replace(".0", "").isdigit()
                    else str(x).strip()
                    if pd.notna(x)
                    else ""
                )
                df_data["item_name"] = df_data[3].astype(str).str.strip()

                df_result_item = (
                    df_data.groupby(["item_code", "item_name"])["row_sales"]
                    .sum()
                    .reset_index()
                )
                df_result_item.columns = ["品番コード", "品番名", "売上高（売単価×数量）"]
                df_result_item = df_result_item.sort_values(by="品番コード").reset_index(drop=True)

                # 8. 結果の表示（タブで切替表示）
                st.success("✅ 集計が完了しました！")

                tab1, tab2, tab3 = st.tabs(["🏬 店舗別集計", "🏢 取引先別集計", "📦 品番別集計"])

                with tab1:
                    st.dataframe(
                        df_result_store.style.format({"売上高（売単価×数量）": "{:,.0f}"}),
                        use_container_width=True,
                    )
                with tab2:
                    st.dataframe(
                        df_result_vendor.style.format({"売上高（売単価×数量）": "{:,.0f}"}),
                        use_container_width=True,
                    )
                with tab3:
                    st.dataframe(
                        df_result_item.style.format({"売上高（売単価×数量）": "{:,.0f}"}),
                        use_container_width=True,
                    )

                st.divider()

                # 9. 🎨 装飾付き Multi-sheet Excelデータの生成 (openpyxl)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result_store.to_excel(writer, index=False, sheet_name="店舗別売上")
                    df_result_vendor.to_excel(writer, index=False, sheet_name="取引先別売上")
                    df_result_item.to_excel(writer, index=False, sheet_name="品番別売上")

                    # デザイン共通スタイルの定義
                    FONT_NAME = "メイリオ"
                    header_font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    body_font = Font(name=FONT_NAME, size=10)
                    thin_border = Border(
                        left=Side(style="thin", color="D9D9D9"),
                        right=Side(style="thin", color="D9D9D9"),
                        top=Side(style="thin", color="D9D9D9"),
                        bottom=Side(style="thin", color="D9D9D9"),
                    )

                    # 3つのシートすべてにデザインを適用する汎用処理
                    sheet_configs = [
                        ("店舗別売上", df_result_store),
                        ("取引先別売上", df_result_vendor),
                        ("品番別売上", df_result_item),
                    ]

                    for sheet_name, df_sheet in sheet_configs:
                        ws = writer.sheets[sheet_name]

                        # ヘッダー行装飾
                        for col_num in range(1, len(df_sheet.columns) + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")

                        # データ行装飾
                        for row_num in range(2, len(df_sheet) + 2):
                            # 1列目 (コード類: 中央揃え, 文字列指定)
                            c1 = ws.cell(row=row_num, column=1)
                            c1.font = body_font
                            c1.alignment = Alignment(horizontal="center")
                            c1.number_format = "@"  # 文字列形式（0埋め保持用）
                            c1.border = thin_border

                            # 2列目 (名称: 左揃え)
                            c2 = ws.cell(row=row_num, column=2)
                            c2.font = body_font
                            c2.alignment = Alignment(horizontal="left")
                            c2.border = thin_border

                            # 3列目 (売上高: 右揃え・3桁カンマ)
                            c3 = ws.cell(row=row_num, column=3)
                            c3.font = body_font
                            c3.number_format = "#,##0"
                            c3.alignment = Alignment(horizontal="right")
                            c3.border = thin_border

                        # 列幅自動調整
                        for col in ws.columns:
                            max_len = max(len(str(cell.value or "")) for cell in col)
                            col_letter = openpyxl.utils.get_column_letter(col[0].column)
                            ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

                # ダウンロード ＆ リセットボタンエリア
                col_dl, col_rst = st.columns([2, 1])
                with col_dl:
                    st.download_button(
                        label="📥 集計結果（3シート入りExcel）をダウンロード",
                        data=output.getvalue(),
                        file_name=f"店舗・取引先・品番別売上集計_{catalog_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                with col_rst:
                    st.button(
                        "🔄 最初の画面に戻る（全リセット）",
                        on_click=reset_app,
                        type="secondary",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
