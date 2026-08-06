import io
import math
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="売上集計ツール", layout="wide")

# Custom CSS: ボタン装飾
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

st.title("📊 売上集計ツール（店舗別・取引先別）")

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


# 会員番号マスク関数（下4桁を残し前半を*で置換）
def mask_member_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if len(s) > 4:
        return "*" * (len(s) - 4) + s[-4:]
    return s


# モード選択
calc_mode = st.radio(
    "集計パターンを選択してください：",
    [
        "① 商品カタログ ＋ 店コード表（通常集計）",
        "② 共同販促パターン ＋ 商品カタログ（共同販促集計）"
    ],
    horizontal=True,
    key=f"calc_mode_{st.session_state.uploader_key}"
)

st.markdown("---")

# 入力エリアの構築
col_file, col_opts = st.columns([3, 1])

with col_file:
    if "①" in calc_mode:
        file_help = "「商品カタログ」と「⑩店コード表」の2つのファイルをドラッグ＆ドロップしてください。"
    else:
        file_help = "「共同販促パターン」と「商品カタログ」の2つのファイルをドラッグ＆ドロップしてください。"

    uploaded_files = st.file_uploader(
        file_help,
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state.uploader_key}",
    )

with col_opts:
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

    # 共同販促モードのみ按分基準の選択肢を表示
    if "②" in calc_mode:
        apportion_base = st.selectbox(
            "補填額（協賛額）の按分基準：",
            ["売上高", "売上原価", "付与ポイント"],
            key=f"apportion_base_{st.session_state.uploader_key}"
        )
    else:
        apportion_base = "売上高"

if uploaded_files:
    catalog_file = None
    store_code_file = None
    promo_file = None

    for f in uploaded_files:
        if "店コード" in f.name:
            store_code_file = f
        elif "共同販促" in f.name or "販促" in f.name:
            promo_file = f
        elif "商品カタログ" in f.name or "カタログ" in f.name:
            catalog_file = f

    ready_to_run = False
    if "①" in calc_mode:
        if not store_code_file: st.warning("⚠️ 「店コード表」が選択されていません。")
        if not catalog_file: st.warning("⚠️ 「商品カタログ」が選択されていません。")
        if store_code_file and catalog_file: ready_to_run = True
    else:
        if not promo_file: st.warning("⚠️ 「共同販促パターン」ファイルが選択されていません。")
        if not catalog_file: st.warning("⚠️ 「商品カタログ」ファイルが選択されていません。")
        if promo_file and catalog_file: ready_to_run = True

    if ready_to_run:
        if not st.session_state.run_calc:
            st.info("💡 ファイルの準備ができました。「集計を開始する」ボタンを押してください。")
            if st.button("🚀 集計を開始する", type="primary", use_container_width=True):
                st.session_state.run_calc = True
                st.rerun()

        if st.session_state.run_calc:
            try:
                # ----------------------------------------------------
                # 商品カタログの読み込みとヘッダー特定
                # ----------------------------------------------------
                df_cat = pd.read_excel(catalog_file, header=None)

                header_row_idx = -1
                for i in range(min(30, len(df_cat))):
                    row_vals = [str(v).strip() for v in df_cat.iloc[i].values if pd.notna(v)]
                    if "売単価" in row_vals or "原単価" in row_vals or "商品名称" in row_vals or "商品ｺｰﾄﾞ" in row_vals:
                        header_row_idx = i
                        break

                if header_row_idx == -1:
                    st.error("❌ エラー: 商品カタログ内に見出し行が見つかりませんでした。")
                    st.stop()

                store_name_row_idx = max(0, header_row_idx - 1)
                row_store_names = df_cat.iloc[store_name_row_idx].values
                row_headers = df_cat.iloc[header_row_idx].values

                # カラム特定
                item_code_col, vendor_code_col, vendor_name_col, unit_price_col, cost_price_col, jan_col = -1, -1, -1, -1, -1, -1

                for idx, h_val in enumerate(row_headers):
                    if pd.isna(h_val): continue
                    h_str = str(h_val).strip()

                    if h_str == "売単価" and unit_price_col == -1: unit_price_col = idx
                    if h_str == "原単価" and cost_price_col == -1: cost_price_col = idx
                    if h_str in ["取引先名称", "取引先名"] and vendor_name_col == -1: vendor_name_col = idx
                    if h_str in ["商品ｺｰﾄﾞ", "商品コード", "JAN", "jan"] and jan_col == -1: jan_col = idx

                    top_h = str(row_store_names[idx]).strip() if pd.notna(row_store_names[idx]) else ""
                    if "取引先" in top_h and h_str in ["ｺｰﾄﾞ", "コード"] and vendor_code_col == -1: vendor_code_col = idx
                    if "品番" in top_h and h_str in ["ｺｰﾄﾞ", "コード"] and item_code_col == -1: item_code_col = idx

                if vendor_code_col == -1 and vendor_name_col != -1 and vendor_name_col > 0:
                    vendor_code_col = vendor_name_col - 1
                if item_code_col == -1:
                    for idx, h_val in enumerate(row_headers):
                        if str(h_val).strip() in ["品番", "品番ｺｰﾄﾞ", "品番コード"]:
                            item_code_col = idx; break

                df_cat_data = df_cat.iloc[header_row_idx + 1:].copy()
                rate_val = (sponsor_rate_input / 100.0) if sponsor_rate_input is not None else None

                # ====================================================
                # モード①：通常集計（カタログ ＋ 店コード表）
                # ====================================================
                if "①" in calc_mode:
                    df_store_raw = pd.read_excel(store_code_file, header=None)
                    store_map = {}
                    for idx, row in df_store_raw.iterrows():
                        if pd.notna(row[1]) and pd.notna(row[2]):
                            raw_code = str(row[1]).strip()
                            code = str(int(float(raw_code))).zfill(3) if raw_code.replace(".0", "").isdigit() else raw_code.zfill(3)
                            store_map[str(row[2]).strip()] = code

                    unit_prices = pd.to_numeric(df_cat_data[unit_price_col], errors="coerce").fillna(0)

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

                    df_cat_data["vendor_code"] = df_cat_data[vendor_code_col].apply(
                        lambda x: str(int(float(x))).zfill(7) if pd.notna(x) and str(x).replace(".0", "").isdigit()
                        else str(x).strip().zfill(7) if pd.notna(x) else "0000000"
                    )
                    df_cat_data["vendor_name"] = df_cat_data[vendor_name_col].astype(str).str.strip()
                    df_cat_data["item_code"] = df_cat_data[item_code_col].apply(
                        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace(".0", "").isdigit()
                        else str(x).strip() if pd.notna(x) else ""
                    )

                    records = []
                    store_totals = {}

                    for store_name, col_idx in store_columns:
                        qty_series = pd.to_numeric(df_cat_data[col_idx], errors="coerce").fillna(0)
                        sales_series = qty_series * unit_prices
                        store_code = store_map.get(store_name, "999")
                        store_totals[store_name] = store_totals.get(store_name, 0.0) + sales_series.sum()

                        temp_df = pd.DataFrame({
                            "vendor_code": df_cat_data["vendor_code"],
                            "vendor_name": df_cat_data["vendor_name"],
                            "store_code": store_code,
                            "store_name": store_name,
                            "item_code": df_cat_data["item_code"],
                            "sales": sales_series
                        })
                        records.append(temp_df[temp_df["sales"] > 0])

                    df_all_details = pd.concat(records, ignore_index=True) if records else pd.DataFrame()
                    df_detail_grouped = df_all_details.groupby(
                        ["vendor_code", "vendor_name", "store_code", "store_name", "item_code"], as_index=False
                    )["sales"].sum() if not df_all_details.empty else pd.DataFrame(columns=["vendor_code", "vendor_name", "store_code", "store_name", "item_code", "sales"])

                    store_list = [
                        {"店コード": store_map.get(name, "999"), "店舗名": name, "売上高\n(売単価×数量)": amt}
                        for name, amt in store_totals.items()
                    ]
                    df_promo_export = None  # 原本コピー不要

                # ====================================================
                # モード②：共同販促集計（共同販促 ＋ カタログ）
                # ====================================================
                else:
                    # 商品カタログマッピング (JAN -> 取引先コード, 取引先名, 品番, 原単価)
                    jan_map = {}
                    for _, row in df_cat_data.iterrows():
                        j_val = str(row[jan_col]).strip() if pd.notna(row[jan_col]) else ""
                        if j_val:
                            if j_val.endswith(".0"): j_val = j_val[:-2]
                            v_code = str(int(float(row[vendor_code_col]))).zfill(7) if pd.notna(row[vendor_code_col]) and str(row[vendor_code_col]).replace('.0','').isdigit() else str(row[vendor_code_col]).strip().zfill(7) if pd.notna(row[vendor_code_col]) else "0000000"
                            v_name = str(row[vendor_name_col]).strip() if pd.notna(row[vendor_name_col]) else ""
                            i_code = str(int(float(row[item_code_col]))) if pd.notna(row[item_code_col]) and str(row[item_code_col]).replace('.0','').isdigit() else str(row[item_code_col]).strip() if pd.notna(row[item_code_col]) else ""
                            c_price = float(row[cost_price_col]) if cost_price_col != -1 and pd.notna(row[cost_price_col]) else 0.0
                            jan_map[j_val] = {
                                "vendor_code": v_code,
                                "vendor_name": v_name,
                                "item_code": i_code,
                                "cost_price": c_price
                            }

                    df_prom_raw = pd.read_excel(promo_file)
                    df_prom = df_prom_raw.copy()

                    # 列特定
                    master_price_col = next((c for c in df_prom.columns if "マスター単価" in str(c) or "単価" in str(c)), None)
                    qty_col = next((c for c in df_prom.columns if "数量" in str(c)), None)
                    store_code_col_prom = next((c for c in df_prom.columns if "店舗コード" in str(c) or "店コード" in str(c)), None)
                    store_name_col_prom = next((c for c in df_prom.columns if "店舗名" in str(c) or "店名" in str(c)), None)
                    jan_col_prom = next((c for c in df_prom.columns if "jan" in str(c).lower() or "商品コード" in str(c)), None)
                    item_code_col_prom = next((c for c in df_prom.columns if "品番" in str(c)), None)
                    point_col_prom = next((c for c in df_prom.columns if "付与ポイント" in str(c) or "ポイント" in str(c)), None)
                    member_col_prom = next((c for c in df_prom.columns if "会員" in str(c)), None)

                    if not (master_price_col and qty_col and store_name_col_prom and jan_col_prom):
                        st.error("❌ エラー: 共同販促パターン内に必要な列（マスター単価・数量・店舗名・JAN）が見つかりませんでした。")
                        st.stop()

                    # 店舗コード3桁整形
                    df_prom["store_code_clean"] = df_prom[store_code_col_prom].apply(
                        lambda x: str(int(float(x))).zfill(3) if pd.notna(x) and str(x).replace(".0","").isdigit() else str(x).strip().zfill(3) if pd.notna(x) else "999"
                    )

                    # 除外店舗（052 名古屋店）のフィルタリング
                    df_prom_calc = df_prom[df_prom["store_code_clean"] != "052"].copy()

                    # JAN文字列変換
                    df_prom_calc["jan_str"] = df_prom_calc[jan_col_prom].apply(
                        lambda x: str(int(float(x))) if pd.notna(x) and str(x).replace(".0","").isdigit() else str(x).strip() if pd.notna(x) else ""
                    )

                    # 売上高
                    df_prom_calc["sales"] = pd.to_numeric(df_prom_calc[master_price_col], errors="coerce").fillna(0) * pd.to_numeric(df_prom_calc[qty_col], errors="coerce").fillna(0)

                    # 売上原価 (カタログ原単価 × 数量)
                    df_prom_calc["cost_price"] = df_prom_calc["jan_str"].apply(lambda j: jan_map.get(j, {}).get("cost_price", 0.0))
                    df_prom_calc["cost_total"] = df_prom_calc["cost_price"] * pd.to_numeric(df_prom_calc[qty_col], errors="coerce").fillna(0)

                    # 付与ポイント（行ごとに端数切り捨て）
                    if point_col_prom and point_col_prom in df_prom_calc.columns:
                        df_prom_calc["points_clean"] = df_prom_calc[point_col_prom].apply(
                            lambda x: math.floor(float(x)) if pd.notna(x) and not math.isnan(float(x)) else 0
                        )
                    else:
                        df_prom_calc["points_clean"] = 0

                    df_prom_calc["store_name"] = df_prom_calc[store_name_col_prom].astype(str).str.strip()

                    # 取引先情報マッピング
                    df_prom_calc["vendor_code"] = df_prom_calc["jan_str"].apply(lambda j: jan_map.get(j, {}).get("vendor_code", "0000000"))
                    df_prom_calc["vendor_name"] = df_prom_calc["jan_str"].apply(lambda j: jan_map.get(j, {}).get("vendor_name", "不明"))
                    df_prom_calc["item_code"] = df_prom_calc.apply(
                        lambda r: jan_map.get(r["jan_str"], {}).get("item_code", str(r[item_code_col_prom]) if item_code_col_prom and pd.notna(r[item_code_col_prom]) else ""), axis=1
                    )

                    # 明細グループ化
                    df_detail_grouped = df_prom_calc.groupby(
                        ["vendor_code", "vendor_name", "store_code_clean", "store_name", "item_code"], as_index=False
                    ).agg({
                        "sales": "sum",
                        "cost_total": "sum",
                        "points_clean": "sum"
                    }).rename(columns={"store_code_clean": "store_code"})

                    # 按分基準の列を「apportion_val」として準備
                    if apportion_base == "売上原価":
                        df_prom_calc["apportion_val"] = df_prom_calc["cost_total"]
                        df_detail_grouped["apportion_val"] = df_detail_grouped["cost_total"]
                    elif apportion_base == "付与ポイント":
                        df_prom_calc["apportion_val"] = df_prom_calc["points_clean"]
                        df_detail_grouped["apportion_val"] = df_detail_grouped["points_clean"]
                    else:
                        df_prom_calc["apportion_val"] = df_prom_calc["sales"]
                        df_detail_grouped["apportion_val"] = df_detail_grouped["sales"]

                    # 店舗別サマリー構築
                    store_summary = df_prom_calc.groupby(["store_code_clean", "store_name"], as_index=False).agg({
                        "sales": "sum",
                        "apportion_val": "sum"
                    }).rename(columns={"store_code_clean": "store_code"})

                    store_list = [
                        {
                            "店コード": r["store_code"],
                            "店舗名": r["store_name"],
                            "売上高\n(売単価×数量)": r["sales"],
                            "apportion_val": r["apportion_val"]
                        }
                        for _, r in store_summary.iterrows()
                    ]

                    # 共同販促パターン原本シート用の整形コピー（会員番号マスク・3桁店コード適用）
                    df_promo_export = df_prom_raw.copy()
                    if member_col_prom and member_col_prom in df_promo_export.columns:
                        df_promo_export[member_col_prom] = df_promo_export[member_col_prom].apply(mask_member_id)
                    if store_code_col_prom and store_code_col_prom in df_promo_export.columns:
                        df_promo_export[store_code_col_prom] = df_prom["store_code_clean"]

                # ====================================================
                # 共通集計データフレーム作成＆端数調整（按分基準対応）
                # ====================================================
                df_stores_only = pd.DataFrame(store_list).sort_values(by="店コード").reset_index(drop=True)
                total_sales_all = df_stores_only["売上高\n(売単価×数量)"].sum()

                if "apportion_val" not in df_stores_only.columns:
                    df_stores_only["apportion_val"] = df_stores_only["売上高\n(売単価×数量)"]

                total_apportion_all = df_stores_only["apportion_val"].sum()

                df_total_row = pd.DataFrame([{
                    "店コード": "000",
                    "店舗名": "全店",
                    "売上高\n(売単価×数量)": total_sales_all,
                    "apportion_val": total_apportion_all
                }])
                df_result_store = pd.concat([df_total_row, df_stores_only], ignore_index=True)

                # 構成比（選ばれた按分基準に基づく）
                df_result_store["構成比"] = df_result_store["apportion_val"] / total_apportion_all if total_apportion_all > 0 else 0.0

                if rate_val is not None:
                    target_total_sponsor = round(total_apportion_all * rate_val)
                    store_sponsors = [round(amt * rate_val) for amt in df_stores_only["apportion_val"]]
                    diff = target_total_sponsor - sum(store_sponsors)
                    if diff != 0 and len(store_sponsors) > 0:
                        max_idx = df_stores_only["apportion_val"].idxmax()
                        store_sponsors[max_idx] += diff

                    df_result_store["協賛額"] = [target_total_sponsor] + store_sponsors
                    df_result_store["協賛料率"] = [rate_val if i == 0 else None for i in range(len(df_result_store))]
                else:
                    df_result_store["協賛額"] = None
                    df_result_store["協賛料率"] = None

                # 不要になった計算列の削除
                df_result_store = df_result_store.drop(columns=["apportion_val"])

                # 画面表示
                st.success(f"✅ 集計が完了しました！（按分基準: {apportion_base}）")
                tab1, tab2 = st.tabs(["🏬 店舗別集計", "🏢 取引先別・店コード・店舗名・品番別明細"])

                with tab1:
                    fmt_dict = {"売上高\n(売単価×数量)": "{:,.0f}", "構成比": "{:.1%}"}
                    if rate_val is not None:
                        fmt_dict["協賛額"] = "{:,.0f}"
                        fmt_dict["協賛料率"] = lambda x: f"{x:.0%}" if pd.notna(x) else ""
                    st.dataframe(df_result_store.style.format(fmt_dict, na_rep=""), use_container_width=True)

                with tab2:
                    df_preview = df_detail_grouped[["vendor_code", "vendor_name", "store_code", "store_name", "item_code", "sales"]].copy()
                    df_preview.columns = ["取引先コード", "取引先名", "店コード", "店舗名", "品番", "売上高（売単価×数量）"]
                    st.dataframe(df_preview.style.format({"売上高（売単価×数量）": "{:,.0f}"}), use_container_width=True)

                st.divider()

                # ====================================================
                # Excel生成 (openpyxl)
                # ====================================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # ① 店舗別集計シート
                    df_result_store.to_excel(writer, index=False, sheet_name="店舗別集計")

                    # ② 取引先別集計シート
                    unique_vendors = df_detail_grouped[["vendor_code", "vendor_name"]].drop_duplicates()
                    for _, v_row in unique_vendors.iterrows():
                        v_code, v_name = v_row["vendor_code"], v_row["vendor_name"]

                        df_v = df_detail_grouped[
                            (df_detail_grouped["vendor_code"] == v_code) & (df_detail_grouped["vendor_name"] == v_name)
                        ].copy()

                        df_v_export = df_v[["vendor_code", "vendor_name", "store_code", "store_name", "item_code", "sales", "apportion_val"]].sort_values(by=["store_code", "item_code"]).reset_index(drop=True)
                        v_total_sales = df_v_export["sales"].sum()
                        v_total_apportion = df_v_export["apportion_val"].sum()

                        df_v_total_row = pd.DataFrame([{
                            "vendor_code": None, "vendor_name": None, "store_code": "000",
                            "store_name": "全店", "item_code": None, "sales": v_total_sales, "apportion_val": v_total_apportion
                        }])

                        df_v_full = pd.concat([df_v_total_row, df_v_export], ignore_index=True)
                        df_v_full["構成比"] = df_v_full["apportion_val"] / v_total_apportion if v_total_apportion > 0 else 0.0

                        if rate_val is not None:
                            v_target_sponsor = round(v_total_apportion * rate_val)
                            v_store_sponsors = [round(amt * rate_val) for amt in df_v_export["apportion_val"]]
                            v_diff = v_target_sponsor - sum(v_store_sponsors)
                            if v_diff != 0 and len(v_store_sponsors) > 0:
                                max_v_idx = df_v_export["apportion_val"].idxmax()
                                v_store_sponsors[max_v_idx] += v_diff

                            df_v_full["協賛額"] = [v_target_sponsor] + v_store_sponsors
                            df_v_full["協賛料率"] = [rate_val if i == 0 else None for i in range(len(df_v_full))]
                        else:
                            df_v_full["協賛額"] = None
                            df_v_full["協賛料率"] = None

                        df_v_full = df_v_full.drop(columns=["apportion_val"])
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

                    # ③ 「共同販促パターン」原本シートコピー（モード②の場合）
                    if df_promo_export is not None:
                        df_promo_export.to_excel(writer, index=False, sheet_name="共同販促パターン")

                    # ④ 「商品カタログ」原本シートのコピー
                    df_cat.to_excel(writer, index=False, header=False, sheet_name="商品カタログ")

                    # スタイル装飾設定
                    FONT_NAME = "メイリオ"
                    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    body_font = Font(name=FONT_NAME, size=10)
                    total_font = Font(name=FONT_NAME, size=10, bold=True)
                    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
                    total_bottom_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="double", color="000000"))

                    for sheet_name in writer.sheets.keys():
                        if sheet_name in ["商品カタログ", "共同販促パターン"]:
                            continue
                        ws = writer.sheets[sheet_name]

                        for col_num in range(1, ws.max_column + 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                        for row_num in range(2, ws.max_row + 1):
                            is_total_row = (row_num == 2)
                            current_font = total_font if is_total_row else body_font
                            current_border = total_bottom_border if is_total_row else thin_border

                            if sheet_name == "店舗別集計":
                                ws.cell(row=row_num, column=1).alignment = Alignment(horizontal="center")
                                ws.cell(row=row_num, column=2).alignment = Alignment(horizontal="left")
                                ws.cell(row=row_num, column=3).alignment = Alignment(horizontal="right")
                                ws.cell(row=row_num, column=3).number_format = "#,##0"
                                ws.cell(row=row_num, column=4).alignment = Alignment(horizontal="right")
                                ws.cell(row=row_num, column=4).number_format = "0.0%"
                                if is_total_row: ws.cell(row=row_num, column=4).value = None

                                c5 = ws.cell(row=row_num, column=5)
                                c5.alignment = Alignment(horizontal="right")
                                if rate_val is not None: c5.number_format = "#,##0"
                                else: c5.value = None

                                c6 = ws.cell(row=row_num, column=6)
                                c6.alignment = Alignment(horizontal="right")
                                if is_total_row and rate_val is not None: c6.number_format = "0%"
                                else: c6.value = None

                                for c in range(1, 7):
                                    cell = ws.cell(row=row_num, column=c)
                                    cell.font = current_font
                                    cell.border = current_border
                            else:
                                for c in range(1, 10):
                                    cell = ws.cell(row=row_num, column=c)
                                    cell.font = current_font
                                    cell.border = current_border
                                    if c in [1, 3, 5]: cell.alignment = Alignment(horizontal="center")
                                    elif c in [2, 4]: cell.alignment = Alignment(horizontal="left")
                                    elif c in [6, 8]:
                                        cell.alignment = Alignment(horizontal="right")
                                        if c == 6: cell.number_format = "#,##0"
                                        elif c == 8 and rate_val is not None: cell.number_format = "#,##0"
                                        elif c == 8: cell.value = None
                                    elif c == 7:
                                        cell.alignment = Alignment(horizontal="right")
                                        cell.number_format = "0.0%"
                                        if is_total_row: cell.value = None
                                    elif c == 9:
                                        cell.alignment = Alignment(horizontal="right")
                                        if is_total_row and rate_val is not None: cell.number_format = "0%"
                                        else: cell.value = None

                        if sheet_name == "店舗別集計":
                            ws.cell(row=2, column=5).fill = yellow_fill
                        else:
                            ws.cell(row=2, column=8).fill = yellow_fill

                        for col in ws.columns:
                            col_letter = openpyxl.utils.get_column_letter(col[0].column)
                            if (sheet_name == "店舗別集計" and col_letter == "D") or (sheet_name != "店舗別集計" and col_letter == "G"):
                                ws.column_dimensions[col_letter].width = 10
                            else:
                                max_len = 0
                                for cell in col:
                                    val_str = str(cell.value or "")
                                    for l in val_str.split("\n"):
                                        length = sum(2 if ord(c) > 256 else 1 for c in l)
                                        if length > max_len: max_len = length
                                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                out_filename = f"共同販促売上集計_{promo_file.name}" if "②" in calc_mode else f"店舗・取引先別売上集計_{catalog_file.name}"

                col_dl, col_rst = st.columns([1, 1])
                with col_dl:
                    st.download_button(
                        label="📥 集計結果（Excel）をダウンロード",
                        data=output.getvalue(),
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                with col_rst:
                    st.button("🔄 最初の画面に戻る（全リセット）", on_click=reset_app, use_container_width=True)

            except Exception as e:
                st.error(f"❌ 予期せぬエラーが発生しました: {e}")
