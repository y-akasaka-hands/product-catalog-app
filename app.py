import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="商品カタログ 店舗・取引先別売上集計", layout="wide")
st.title("📊 商品カタログ 店舗・取引先別売上集計ツール")

st.markdown("""
**使い方:**
1. 下記のエリアに「**⑩店コード表.xlsx**」と「**商品カタログ_〜.xlsx**」を2つ同時にドラッグ＆ドロップしてください。
2. 「**🚀 集計を開始する**」ボタンを押すと、店舗別および各取引先ごとの「店コード・品番別売上高」が集計されます。
""")

# セッション状態の初期化
if "run_calc" not in st.session_state:
    st.session_state.run_calc = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# 完全リセット処理関数
def reset_app():
    st.session_state.run_calc = False
    st.session_state.uploader_key += 1
    st.rerun()


# Excelシート名で使えない記号の除去関数
def clean_sheet_name(name):
    # \ / ? * : [ ] を除去し、最大31文字に制限
    cleaned = re.sub(r"[\\/*?:\[\]]", "", str(name))
    return cleaned[:31] if cleaned else "Sheet"


# ファイルアップローダー
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

    if store_code_file and catalog_file:
        if not st.session_state.run_calc:
            st.info("💡 ファイルの準備ができました。下のボタンを押して集計を開始してください。")
            if st.button("🚀 集計を開始する", type="primary", use_container_width=True):
                st.session_state.run_calc = True
                st.rerun()

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

                # 取引先コード (Col 8), 取引先名 (Col 9), 品番 (Col 2) の抽出
                df_data["vendor_code"] = df_data[8].apply(
                    lambda x: str(int(float(x))).zfill(7)
                    if pd.notna(x) and str(x).replace(".0", "").isdigit()
                    else str(x).strip().zfill(7)
                    if pd.notna(x)
                    else "0000000"
                )
                df_data["vendor_name"] = df_data[9].astype(str).str.strip()
                df_data["item_code"] = df_data[2].apply(
                    lambda x: str(int(float(x)))
                    if pd.notna(x) and str(x).replace(".0", "").isdigit()
                    else str(x).strip()
                    if pd.notna(x)
                    else ""
                )

                # 5. 明細データの集計（取引先 × 店コード × 品番）
                records = []
                store_totals = {}

                for store_name, col_idx in store_columns:
                    qty_series = pd.to_numeric(df_data[col_idx], errors="coerce").fillna(0)
                    sales_series = qty_series * unit_prices
                    
                    store_code = store_map.get(store_name, "999")
                    store_totals[store_name] = store_totals.get(store_name, 0.0) + sales_series.sum()

                    temp_df = pd.DataFrame({
                        "vendor_code": df_data["vendor_code"],
                        "vendor_name": df_data["vendor_name"],
                        "store_code": store_code,
                        "store_name": store_name,
                        "item_code": df_data["item_code"],
                        "sales": sales_series
                    })
                    # 売上高 > 0 の行のみ保持
                    temp_df = temp_df[temp_df["sales"] > 0]
                    records.append(temp_df)

                # 全明細の結合
                if records:
                    df_all_details = pd.concat(records, ignore_index=True)
                    df_detail_grouped = df_all_details.groupby(
                        ["vendor_code", "vendor_name", "store_code", "item_code"], as_index=False
                    )["sales"].sum()
                else:
                    df_detail_grouped = pd.DataFrame(columns=["vendor_code", "vendor_name", "store_code", "item_code", "sales"])

                # 店舗別サマリーデータの作成
                df_result_store = pd.DataFrame([
                    {"店コード": store_map.get(name, "999"), "店舗名": name, "売上高（売単価×数量）": amt}
                    for name, amt in store_totals.items()
                ]).sort_values(by="店コード").reset_index(drop=True)

                # 6. 結果の表示
                st.success("✅ 集計が完了しました！")

                tab1, tab2 = st.tabs(["🏬 店舗別売上サマリー", "🏢 取引先別・店コード・品番別明細"])

                with tab1:
                    st.dataframe(
                        df_result_store.style.format({"売上高（売単価×数量）": "{:,.0f}"}),
                        use_container_width=True,
                    )
                with tab2:
                    df_preview = df_detail_grouped.copy()
                    df_preview.columns = ["取引先コード", "取引先名", "店コード", "品番", "売上高（売単価×数量）"]
                    st.dataframe(
                        df_preview.style.format({"売上高（売単価×数量）": "{:,.0f}"}),
                        use_container_width=True,
                    )

                st.divider()

                # 7. 🎨 装飾付き Multi-sheet Excelデータの生成 (openpyxl)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # ① 店舗別売上シート
                    df_result_store.to_excel(writer, index=False, sheet_name="店舗別売上")

                    # ② 取引先ごとの個別シート
                    unique_vendors = df_detail_grouped[["vendor_code", "vendor_name"]].drop_duplicates()
                    
                    # 取引先名でソートして順次シート作成
                    for _, v_row in unique_vendors.iterrows():
                        v_code = v_row["vendor_code"]
                        v_name = v_row["vendor_name"]

                        df_v = df_detail_grouped[
                            (df_detail_grouped["vendor_code"] == v_code) & 
                            (df_detail_grouped["vendor_name"] == v_name)
                        ].copy()

                        df_v_export = df_v[["vendor_code", "vendor_name", "store_code", "item_code", "sales"]].copy()
                        df_v_export.columns = ["取引先コード", "取引先名", "店コード", "品番", "売上高（売単価×数量）"]
                        
                        # 店コード -> 品番 の昇順並び替え
                        df_v_export = df_v_export.sort_values(by=["店コード", "品番"]).reset_index(drop=True)

                        sheet_title = clean_sheet_name(v_name)
                        # シート名の重複回避
                        existing_sheets = writer.sheets.keys()
                        base_title = sheet_title
                        counter = 1
                        while sheet_title in existing_sheets:
                            sheet_title = clean_sheet_name(f"{base_title}_{counter}")
                            counter += 1

                        df_v_export.to_excel(writer, index=False, sheet_name=sheet_title)

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

                    # 全シートへのデザイン一括適用
                    for sheet_name in writer.sheets.keys():
                        ws = writer.sheets[sheet_name]

                        # ヘッダー行装飾
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")

                        # データ行装飾
                        for row_num in range(2, ws.max_row + 1):
                            for col_num in range(1, ws.max_column + 1):
                                cell = ws.cell(row=row_num, column=col_num)
                                cell.font = body_font
                                cell.border = thin_border

                                if sheet_name == "店舗別売上":
                                    if col_num == 1:
                                        cell.alignment = Alignment(horizontal="center")
                                        cell.number_format = "@"
                                    elif col_num == 2:
                                        cell.alignment = Alignment(horizontal="left")
                                    elif col_num == 3:
                                        cell.alignment = Alignment(horizontal="right")
                                        cell.number_format = "#,##0"
                                else:
                                    # 取引先別シート (A:取引先コード, B:取引先名, C:店コード, D:品番, E:売上高)
                                    if col_num in [1, 3, 4]:  # コード類（取引先コード, 店コード, 品番）
                                        cell.alignment = Alignment(horizontal="center")
                                        cell.number_format = "@"
                                    elif col_num == 2:  # 取引先名
                                        cell.alignment = Alignment(horizontal="left")
                                    elif col_num == 5:  # 売上高
                                        cell.alignment = Alignment(horizontal="right")
                                        cell.number_format = "#,##0"

                        # 列幅自動調整
                        for col in ws.columns:
                            max_len = max(len(str(cell.value or "")) for cell in col)
                            col_letter = openpyxl.utils.get_column_letter(col[0].column)
                            ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

                # ダウンロード ＆ リセットボタンエリア
                col_dl, col_rst = st.columns([2, 1])
                with col_dl:
                    st.download_button(
                        label="📥 集集結果（取引先別シート入りExcel）をダウンロード",
                        data=output.getvalue(),
                        file_name=f"店舗・取引先別売上集計_{catalog_file.name}",
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
