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
2. 必要に応じて「**協賛料率（％）**」を入力してください（空欄でも集計可能です）。
3. 「**🚀 集計を開始する**」ボタンを押すと集計結果が生成されます。
""")

# セッション状態の初期化
if "run_calc" not in st.session_state:
    st.session_state.run_calc = False
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


# リセット処理関数
def reset_app():
    st.session_state.run_calc = False
    st.session_state.uploader_key += 1
    st.rerun()


# シート名修正関数
def clean_sheet_name(name):
    cleaned = re.sub(r"[\\/*?:\[\]]", "", str(name))
    return cleaned[:31] if cleaned else "Sheet"


# 入力エリアの構築
col_file, col_rate = st.columns([3, 1])

with col_file:
    uploaded_files = st.file_uploader(
        "「商品カタログ」と「店コード表」をアップロード（複数可）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state.uploader_key}",
    )

with col_rate:
    sponsor_rate_input = st.number_input(
        "協賛料率（％）",
        min_value=0.0,
        max_value=100.0,
        value=None,
        step=0.5,
        format="%.1f",
        help="例: 10 と入力すると 10% として計算されます（空欄可）",
    )

if uploaded_files:
    catalog_file = None
    store_code_file = None

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
            st.info("💡 ファイルの準備ができました。「集計を開始する」ボタンを押してください。")
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

                # 3. 「売単価」列の位置判定
                unit_price_col_idx = -1
                for idx, val in enumerate(row12):
                    if pd.notna(val) and str(val).strip() == "売単価":
                        unit_price_col_idx = idx
                        break

                if unit_price_col_idx == -1:
                    st.error("❌ 商品カタログ内に「売単価」の列が見つかりませんでした。")
                    st.stop()

                df_data = df_cat.iloc[13:].copy()
                unit_prices = pd.to_numeric(df_data[unit_price_col_idx], errors="coerce").fillna(0)

                # 4. 各店舗の「売上数」列の位置特定
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

                # 5. 明細データおよび店舗別売上の計算
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
                    temp_df = temp_df[temp_df["sales"] > 0]
                    records.append(temp_df)

                if records:
                    df_all_details = pd.concat(records, ignore_index=True)
                    df_detail_grouped = df_all_details.groupby(
                        ["vendor_code", "vendor_name", "store_code", "item_code"], as_index=False
                    )["sales"].sum()
                else:
                    df_detail_grouped = pd.DataFrame(columns=["vendor_code", "vendor_name", "store_code", "item_code", "sales"])

                # 店舗別集計サマリーの構築（000 全店 + 各店舗）
                store_list = [
                    {"店コード": store_map.get(name, "999"), "店舗名": name, "売上高（売単価×数量）": amt}
                    for name, amt in store_totals.items()
                ]
                df_stores_only = pd.DataFrame(store_list).sort_values(by="店コード").reset_index(drop=True)

                total_sales_all = df_stores_only["売上高（売単価×数量）"].sum()

                df_total_row = pd.DataFrame([{
                    "店コード": "000",
                    "店舗名": "全店",
                    "売上高（売単価×数量）": total_sales_all
                }])
                
                df_result_store = pd.concat([df_total_row, df_stores_only], ignore_index=True)

                df_result_store["構成比"] = df_result_store["売上高（売単価×数量）"] / total_sales_all if total_sales_all > 0 else 0.0

                rate_val = (sponsor_rate_input / 100.0) if sponsor_rate_input is not None else None

                if rate_val is not None:
                    total_sponsor = total_sales_all * rate_val
                    df_result_store["協賛額"] = df_result_store["構成比"] * total_sponsor
                    df_result_store["協賛料率"] = [rate_val if i == 0 else None for i in range(len(df_result_store))]
                else:
                    df_result_store["協賛額"] = None
                    df_result_store["協賛料率"] = None

                # 6. 結果の表示
                st.success("✅ 集計が完了しました！")

                tab1, tab2 = st.tabs(["🏬 店舗別集計", "🏢 取引先別・店コード・品番別明細"])

                with tab1:
                    fmt_dict = {
                        "売上高（売単価×数量）": "{:,.0f}",
                        "構成比": "{:.1%}"
                    }
                    if rate_val is not None:
                        fmt_dict["協賛額"] = "{:,.0f}"
                        fmt_dict["協賛料率"] = lambda x: f"{x:.0%}" if pd.notna(x) else ""

                    st.dataframe(
                        df_result_store.style.format(fmt_dict, na_rep=""),
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

                # 7. 🎨 装飾付き Multi-sheet Excel データの生成 (openpyxl)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # ① 店舗別集計シート
                    df_result_store.to_excel(writer, index=False, sheet_name="店舗別集計")

                    # ② 取引先別シート
                    unique_vendors = df_detail_grouped[["vendor_code", "vendor_name"]].drop_duplicates()
                    for _, v_row in unique_vendors.iterrows():
                        v_code = v_row["vendor_code"]
                        v_name = v_row["vendor_name"]

                        df_v = df_detail_grouped[
                            (df_detail_grouped["vendor_code"] == v_code) & 
                            (df_detail_grouped["vendor_name"] == v_name)
                        ].copy()

                        df_v_export = df_v[["vendor_code", "vendor_name", "store_code", "item_code", "sales"]].copy()
                        df_v_export.columns = ["取引先コード", "取引先名", "店コード", "品番", "売上高（売単価×数量）"]
                        df_v_export = df_v_export.sort_values(by=["店コード", "品番"]).reset_index(drop=True)

                        sheet_title = clean_sheet_name(v_name)
                        existing_sheets = writer.sheets.keys()
                        base_title = sheet_title
                        counter = 1
                        while sheet_title in existing_sheets:
                            sheet_title = clean_sheet_name(f"{base_title}_{counter}")
                            counter += 1

                        df_v_export.to_excel(writer, index=False, sheet_name=sheet_title)

                    # ③ 「商品カタログ」原本シートのコピー追加
                    df_cat.to_excel(writer, index=False, header=False, sheet_name="商品カタログ")

                    # デザイン共通スタイルの定義
                    FONT_NAME = "メイリオ"
                    header_font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    body_font = Font(name=FONT_NAME, size=10)
                    total_font = Font(name=FONT_NAME, size=10, bold=True)
                    
                    thin_border = Border(
                        left=Side(style="thin", color="D9D9D9"),
                        right=Side(style="thin", color="D9D9D9"),
                        top=Side(style="thin", color="D9D9D9"),
                        bottom=Side(style="thin", color="D9D9D9"),
                    )

                    total_bottom_border = Border(
                        left=Side(style="thin", color="D9D9D9"),
                        right=Side(style="thin", color="D9D9D9"),
                        top=Side(style="thin", color="D9D9D9"),
                        bottom=Side(style="double", color="000000"),
                    )

                    # 各集計シートへのデザイン適用（「商品カタログ」シートは除く）
                    for sheet_name in writer.sheets.keys():
                        if sheet_name == "商品カタログ":
                            continue

                        ws = writer.sheets[sheet_name]

                        # ヘッダー行装飾
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")

                        if sheet_name == "店舗別集計":
                            for row_num in range(2, ws.max_row + 1):
                                is_total_row = (row_num == 2)
                                current_font = total_font if is_total_row else body_font
                                current_border = total_bottom_border if is_total_row else thin_border

                                c1 = ws.cell(row=row_num, column=1)
                                c1.font = current_font
                                c1.alignment = Alignment(horizontal="center")
                                c1.number_format = "@"
                                c1.border = current_border

                                c2 = ws.cell(row=row_num, column=2)
                                c2.font = current_font
                                c2.alignment = Alignment(horizontal="left")
                                c2.border = current_border

                                c3 = ws.cell(row=row_num, column=3)
                                c3.font = current_font
                                c3.alignment = Alignment(horizontal="right")
                                c3.number_format = "#,##0"
                                c3.border = current_border

                                c4 = ws.cell(row=row_num, column=4)
                                c4.font = current_font
                                c4.alignment = Alignment(horizontal="right")
                                c4.number_format = "0.0%"
                                if is_total_row:
                                    c4.value = None
                                c4.border = current_border

                                c5 = ws.cell(row=row_num, column=5)
                                c5.font = current_font
                                c5.alignment = Alignment(horizontal="right")
                                if rate_val is not None:
                                    c5.number_format = "#,##0"
                                else:
                                    c5.value = None
                                c5.border = current_border

                                c6 = ws.cell(row=row_num, column=6)
                                c6.font = current_font
                                c6.alignment = Alignment(horizontal="right")
                                if is_total_row and rate_val is not None:
                                    c6.number_format = "0%"
                                else:
                                    c6.value = None
                                c6.border = current_border

                        else:
                            for row_num in range(2, ws.max_row + 1):
                                for col_num in range(1, ws.max_column + 1):
                                    cell = ws.cell(row=row_num, column=col_num)
                                    cell.font = body_font
                                    cell.border = thin_border

                                    if col_num in [1, 3, 4]:
                                        cell.alignment = Alignment(horizontal="center")
                                        cell.number_format = "@"
                                    elif col_num == 2:
                                        cell.alignment = Alignment(horizontal="left")
                                    elif col_num == 5:
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
                        label="📥 集計結果（Excel）をダウンロード",
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
