import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="商品カタログ 店舗・取引先別売上集計", layout="wide")

# Custom CSS: ダウンロードボタン＆リセットボタン共通の立体的・大文字デザイン
st.markdown("""
<style>
div.stDownloadButton > button, div.stButton > button {
    background: linear-gradient(180deg, #28a745 0%, #1e7e34 100%) !important;
    color: #ffffff !important;
    font-size: 20px !important;
    font-weight: 800 !important;
    padding: 16px 28px !important;
    border-radius: 8px !important;
    border: none !important;
    border-bottom: 5px solid #145222 !important;
    box-shadow: 0px 5px 10px rgba(0, 0, 0, 0.3) !important;
    transition: all 0.15s ease-in-out !important;
    width: 100% !important;
}

div.stDownloadButton > button:hover, div.stButton > button:hover {
    background: linear-gradient(180deg, #34ce57 0%, #218838 100%) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0px 7px 14px rgba(0, 0, 0, 0.35) !important;
    color: #ffffff !important;
}

div.stDownloadButton > button:active, div.stButton > button:active {
    transform: translateY(3px) !important;
    border-bottom: 2px solid #145222 !important;
    box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.2) !important;
}
</style>
""", unsafe_allow_html=True)

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
    if f"sponsor_rate_{st.session_state.uploader_key-1}" in st.session_state:
        del st.session_state[f"sponsor_rate_{st.session_state.uploader_key-1}"]


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
        key=f"sponsor_rate_{st.session_state.uploader_key}",
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
                # ----------------------------------------------------
                # 1. 店コード表の自動解析・マッピング作成
                # ----------------------------------------------------
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

                # ----------------------------------------------------
                # 2. 商品カタログの自動解析（ヘッダー行の動的特定）
                # ----------------------------------------------------
                df_cat = pd.read_excel(catalog_file, header=None)
                
                header_row_idx = -1
                for i in range(min(30, len(df_cat))):
                    row_vals = [str(v).strip() for v in df_cat.iloc[i].values if pd.notna(v)]
                    if "売単価" in row_vals or "原単価" in row_vals or "商品名称" in row_vals:
                        header_row_idx = i
                        break

                if header_row_idx == -1:
                    st.error("❌ エラー: 商品カタログ内に見出し行（「売単価」等の列）が見つかりませんでした。ファイル形式を確認してください。")
                    st.stop()

                # 店舗名行（通常はヘッダーの1行上）と詳細見出し行
                store_name_row_idx = max(0, header_row_idx - 1)
                row_store_names = df_cat.iloc[store_name_row_idx].values
                row_headers = df_cat.iloc[header_row_idx].values

                # ----------------------------------------------------
                # 3. 必要なカラム列（品番・取引先・売単価）の動的特定
                # ----------------------------------------------------
                item_code_col = -1
                vendor_code_col = -1
                vendor_name_col = -1
                unit_price_col = -1

                for idx, h_val in enumerate(row_headers):
                    if pd.isna(h_val):
                        continue
                    h_str = str(h_val).strip()

                    # 売単価列の特定
                    if h_str == "売単価" and unit_price_col == -1:
                        unit_price_col = idx

                    # 取引先名称列
                    if h_str in ["取引先名称", "取引先名"] and vendor_name_col == -1:
                        vendor_name_col = idx

                    # 取引先コード列（「取引先」カテゴリ下の「コード」）
                    top_h = str(row_store_names[idx]).strip() if pd.notna(row_store_names[idx]) else ""
                    if "取引先" in top_h and h_str in ["ｺｰﾄﾞ", "コード"] and vendor_code_col == -1:
                        vendor_code_col = idx

                    # 品番コード列（「品番」カテゴリ下の「コード」）
                    if "品番" in top_h and h_str in ["ｺｰﾄﾞ", "コード"] and item_code_col == -1:
                        item_code_col = idx

                # フォールバック（自動特定が漏れた場合のバックアップ）
                if vendor_code_col == -1 and vendor_name_col != -1 and vendor_name_col > 0:
                    vendor_code_col = vendor_name_col - 1
                if item_code_col == -1:
                    for idx, h_val in enumerate(row_headers):
                        if str(h_val).strip() in ["品番", "品番ｺｰﾄﾞ", "品番コード"]:
                            item_code_col = idx
                            break

                # カラムが見つからない場合のエラー出力
                missing_cols = []
                if unit_price_col == -1: missing_cols.append("売単価")
                if vendor_code_col == -1: missing_cols.append("取引先コード")
                if vendor_name_col == -1: missing_cols.append("取引先名称")
                if item_code_col == -1: missing_cols.append("品番コード")

                if missing_cols:
                    st.error(f"❌ エラー: 商品カタログから以下の必要な列が特定できませんでした: {', '.join(missing_cols)}")
                    st.stop()

                # データ実効行（ヘッダーの次の行以降）
                df_data = df_cat.iloc[header_row_idx + 1:].copy()
                unit_prices = pd.to_numeric(df_data[unit_price_col], errors="coerce").fillna(0)

                # ----------------------------------------------------
                # 4. 店舗列の動的特定（「売上数」列を持つ店舗のみ抽出）
                # ----------------------------------------------------
                store_columns = []
                for col_idx in range(len(row_store_names)):
                    top_val = str(row_store_names[col_idx]).strip() if pd.notna(row_store_names[col_idx]) else ""
                    if top_val != "" and top_val not in ["品番", "枝番", "取引先", "全店"]:
                        for offset in range(3):
                            check_idx = col_idx + offset
                            if check_idx < len(row_headers):
                                sub_val = str(row_headers[check_idx]).strip() if pd.notna(row_headers[check_idx]) else ""
                                if sub_val == "売上数":
                                    store_columns.append((top_val, check_idx))
                                    break

                if not store_columns:
                    st.error("❌ エラー: 商品カタログ内に集計対象となる店舗（「売上数」列）が見つかりませんでした。")
                    st.stop()

                # ----------------------------------------------------
                # 5. 明細データおよび店舗別売上の計算
                # ----------------------------------------------------
                df_data["vendor_code"] = df_data[vendor_code_col].apply(
                    lambda x: str(int(float(x))).zfill(7)
                    if pd.notna(x) and str(x).replace(".0", "").isdigit()
                    else str(x).strip().zfill(7)
                    if pd.notna(x)
                    else "0000000"
                )
                df_data["vendor_name"] = df_data[vendor_name_col].astype(str).str.strip()
                df_data["item_code"] = df_data[item_code_col].apply(
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
                        ["vendor_code", "vendor_name", "store_code", "store_name", "item_code"], as_index=False
                    )["sales"].sum()
                else:
                    df_detail_grouped = pd.DataFrame(columns=["vendor_code", "vendor_name", "store_code", "store_name", "item_code", "sales"])

                # 店舗別集計サマリーの構築（000 全店 + 各店舗）
                store_list = [
                    {"店コード": store_map.get(name, "999"), "店舗名": name, "売上高\n(売単価×数量)": amt}
                    for name, amt in store_totals.items()
                ]
                df_stores_only = pd.DataFrame(store_list).sort_values(by="店コード").reset_index(drop=True)

                total_sales_all = df_stores_only["売上高\n(売単価×数量)"].sum()

                df_total_row = pd.DataFrame([{
                    "店コード": "000",
                    "店舗名": "全店",
                    "売上高\n(売単価×数量)": total_sales_all
                }])
                
                df_result_store = pd.concat([df_total_row, df_stores_only], ignore_index=True)

                df_result_store["構成比"] = df_result_store["売上高\n(売単価×数量)"] / total_sales_all if total_sales_all > 0 else 0.0

                rate_val = (sponsor_rate_input / 100.0) if sponsor_rate_input is not None else None

                # 協賛額の端数四捨五入 ＆ 最多売上店舗での差額調整処理
                if rate_val is not None:
                    target_total_sponsor = round(total_sales_all * rate_val)
                    store_sponsors = [round(amt * rate_val) for amt in df_stores_only["売上高\n(売単価×数量)"]]

                    current_sum = sum(store_sponsors)
                    diff = target_total_sponsor - current_sum

                    if diff != 0 and len(store_sponsors) > 0:
                        max_idx = df_stores_only["売上高\n(売単価×数量)"].idxmax()
                        store_sponsors[max_idx] += diff

                    df_result_store["協賛額"] = [target_total_sponsor] + store_sponsors
                    df_result_store["協賛料率"] = [rate_val if i == 0 else None for i in range(len(df_result_store))]
                else:
                    df_result_store["協賛額"] = None
                    df_result_store["協賛料率"] = None

                # 6. 結果の表示
                st.success("✅ 集計が完了しました！")

                tab1, tab2 = st.tabs(["🏬 店舗別集計", "🏢 取引先別・店コード・店舗名・品番別明細"])

                with tab1:
                    fmt_dict = {
                        "売上高\n(売単価×数量)": "{:,.0f}",
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
                    df_preview.columns = ["取引先コード", "取引先名", "店コード", "店舗名", "品番", "売上高（売単価×数量）"]
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

                    # ② 各取引先別シートの構築
                    unique_vendors = df_detail_grouped[["vendor_code", "vendor_name"]].drop_duplicates()
                    for _, v_row in unique_vendors.iterrows():
                        v_code = v_row["vendor_code"]
                        v_name = v_row["vendor_name"]

                        df_v = df_detail_grouped[
                            (df_detail_grouped["vendor_code"] == v_code) & 
                            (df_detail_grouped["vendor_name"] == v_name)
                        ].copy()

                        df_v_export = df_v[["vendor_code", "vendor_name", "store_code", "store_name", "item_code", "sales"]].copy()
                        df_v_export = df_v_export.sort_values(by=["store_code", "item_code"]).reset_index(drop=True)

                        v_total_sales = df_v_export["sales"].sum()

                        # 000 全店合計行（2行目）の作成
                        df_v_total_row = pd.DataFrame([{
                            "vendor_code": None,
                            "vendor_name": None,
                            "store_code": "000",
                            "store_name": "全店",
                            "item_code": None,
                            "sales": v_total_sales
                        }])

                        df_v_full = pd.concat([df_v_total_row, df_v_export], ignore_index=True)

                        # 構成比
                        df_v_full["構成比"] = df_v_full["sales"] / v_total_sales if v_total_sales > 0 else 0.0

                        # 協賛額 ＆ 端数調整
                        if rate_val is not None:
                            v_target_sponsor = round(v_total_sales * rate_val)
                            v_store_sponsors = [round(amt * rate_val) for amt in df_v_export["sales"]]

                            v_diff = v_target_sponsor - sum(v_store_sponsors)
                            if v_diff != 0 and len(v_store_sponsors) > 0:
                                max_v_idx = df_v_export["sales"].idxmax()
                                v_store_sponsors[max_v_idx] += v_diff

                            df_v_full["協賛額"] = [v_target_sponsor] + v_store_sponsors
                            df_v_full["協賛料率"] = [rate_val if i == 0 else None for i in range(len(df_v_full))]
                        else:
                            df_v_full["協賛額"] = None
                            df_v_full["協賛料率"] = None

                        df_v_full.columns = [
                            "取引先コード", "取引先名", "店コード", "店舗名", "品番",
                            "売上高\n(売単価×数量)", "構成比", "協賛額", "協賛料率"
                        ]

                        sheet_title = clean_sheet_name(v_name)
                        existing_sheets = writer.sheets.keys()
                        base_title = sheet_title
                        counter = 1
                        while sheet_title in existing_sheets:
                            sheet_title = clean_sheet_name(f"{base_title}_{counter}")
                            counter += 1

                        df_v_full.to_excel(writer, index=False, sheet_name=sheet_title)

                    # ③ 「商品カタログ」原本シートのコピー追加
                    df_cat.to_excel(writer, index=False, header=False, sheet_name="商品カタログ")

                    # デザイン共通スタイルの定義
                    FONT_NAME = "メイリオ"
                    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    body_font = Font(name=FONT_NAME, size=10)
                    total_font = Font(name=FONT_NAME, size=10, bold=True)
                    
                    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

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

                    # 各集計シートへのデザイン適用（フォント10ptメイリオ）
                    for sheet_name in writer.sheets.keys():
                        if sheet_name == "商品カタログ":
                            continue

                        ws = writer.sheets[sheet_name]

                        # ヘッダー行装飾 (1行目)
                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                        # データ行装飾
                        for row_num in range(2, ws.max_row + 1):
                            is_total_row = (row_num == 2)  # 2行目が000 全店
                            current_font = total_font if is_total_row else body_font
                            current_border = total_bottom_border if is_total_row else thin_border

                            if sheet_name == "店舗別集計":
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
                                c3.alignment = Alignment(horizontal="center")
                                c3.number_format = "@"
                                c3.border = current_border

                                c4 = ws.cell(row=row_num, column=4)
                                c4.font = current_font
                                c4.alignment = Alignment(horizontal="left")
                                c4.border = current_border

                                c5 = ws.cell(row=row_num, column=5)
                                c5.font = current_font
                                c5.alignment = Alignment(horizontal="center")
                                c5.number_format = "@"
                                c5.border = current_border

                                c6 = ws.cell(row=row_num, column=6)
                                c6.font = current_font
                                c6.alignment = Alignment(horizontal="right")
                                c6.number_format = "#,##0"
                                c6.border = current_border

                                c7 = ws.cell(row=row_num, column=7)
                                c7.font = current_font
                                c7.alignment = Alignment(horizontal="right")
                                c7.number_format = "0.0%"
                                if is_total_row:
                                    c7.value = None
                                c7.border = current_border

                                c8 = ws.cell(row=row_num, column=8)
                                c8.font = current_font
                                c8.alignment = Alignment(horizontal="right")
                                if rate_val is not None:
                                    c8.number_format = "#,##0"
                                else:
                                    c8.value = None
                                c8.border = current_border

                                c9 = ws.cell(row=row_num, column=9)
                                c9.font = current_font
                                c9.alignment = Alignment(horizontal="right")
                                if is_total_row and rate_val is not None:
                                    c9.number_format = "0%"
                                else:
                                    c9.value = None
                                c9.border = current_border

                        # 協賛額の全店合計セルの黄色ハイライト指定
                        if sheet_name == "店舗別集計":
                            ws.cell(row=2, column=5).fill = yellow_fill
                        else:
                            ws.cell(row=2, column=8).fill = yellow_fill

                        # 列幅調整
                        for col in ws.columns:
                            col_letter = openpyxl.utils.get_column_letter(col[0].column)
                            if (sheet_name == "店舗別集計" and col_letter == "D") or (sheet_name != "店舗別集計" and col_letter == "G"):
                                ws.column_dimensions[col_letter].width = 10
                            else:
                                max_len = 0
                                for cell in col:
                                    val_str = str(cell.value or "")
                                    lines = val_str.split("\n")
                                    for l in lines:
                                        length = sum(2 if ord(c) > 256 else 1 for c in l)
                                        if length > max_len:
                                            max_len = length
                                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                # ダウンロード ＆ リセットボタンエリア
                col_dl, col_rst = st.columns([1, 1])
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
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"❌ 予期せぬエラーが発生しました: {e}")
