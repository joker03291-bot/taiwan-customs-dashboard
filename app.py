"""
台湾関税統計 Dashboard ジェネレーター (Streamlit Web UI)

起動: streamlit run app.py

主要機能:
  - 関税署からダウンロードした原始 .xls をアップロード
  - 自動で品目・期間・国家を検出してプレビュー
  - 輸入/輸出を選択
  - ダッシュボード生成 → ブラウザでプレビュー → ダウンロード
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from dashboard_generator import (
    generate_dashboard,
    load_country_mapping,
    load_commodity_settings,
    normalize_time,
    detect_commodity_info,
)

# ========== ページ設定 ==========

st.set_page_config(
    page_title="関税統計ダッシュボード生成ツール",
    page_icon="📊",
    layout="wide",
)

# ========== ヘッダー ==========

st.title("📊 台湾関税統計 ダッシュボード生成ツール")
st.caption(
    "財政部関税統計の原始データから、日本語版分析ダッシュボードを自動生成します。"
)

# ========== サイドバー: 使用ガイド ==========

with st.sidebar:
    st.header("📖 使い方")
    st.markdown(
        """
        1. **原始データ取得**
           [財政部関税統計](https://portal.sw.nat.gov.tw/APGA/GA03)から
           品目別の月別輸出入統計を `.xls` 形式でダウンロード

        2. **アップロード**
           本ツールに `.xls` ファイルをドラッグ&ドロップ

        3. **設定確認**
           品目情報・期間・国家リストが自動検出されます

        4. **生成**
           「ダッシュボード生成」ボタンをクリック

        5. **ダウンロード**
           生成された `.xlsx` をダウンロード
        """
    )

    st.divider()
    st.subheader("🛠 マスタ管理")
    st.markdown(
        """
        - **国名対照** — 上部メニュー「国名対照」ページ
        - **品目設定** — 上部メニュー「品目設定」ページ
        """
    )

    st.divider()
    st.caption("v2.0 / 2026-04")

# ========== メイン: ファイルアップロード ==========

uploaded = st.file_uploader(
    "原始ファイルをアップロード",
    type=["xls", "xlsx"],
    help="財政部関税統計から取得した .xls ファイル",
)

if uploaded is None:
    st.info("👆 上のエリアから原始ファイルをアップロードしてください。")
    with st.expander("💡 サンプルデータ形式について"):
        st.markdown(
            """
            原始データは関税統計サイトのダウンロード形式そのまま (long-format) を想定しています:

            | Imports/Exports | Time | Commodity Code | Description | Country | Value(USD$1,000) | Weight(TNE) |
            |---|---|---|---|---|---|---|
            | Imports | 2025/1 | 27101964009 | Naphtha, mineral | Russian Federation | 119,197 | 199,892 |
            | Imports | 2025/1 | 27101964009 | Naphtha, mineral | Singapore | 13,453 | 22,178 |
            | ... | ... | ... | ... | ... | ... | ... |

            速報値の月は `2026/3(preliminary)` のように `(preliminary)` が付与されている前提です。
            """
        )
    st.stop()

# ========== ファイル解析 ==========

# 一時ファイルに保存
ext = Path(uploaded.name).suffix.lower()
with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
    tmp.write(uploaded.read())
    tmp_path = tmp.name

# 原始データ読み込み
try:
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df_raw = pd.read_excel(tmp_path, engine=engine, header=0)
    df_raw.columns = ["Imports/Exports", "Time", "Commodity Code", "Description",
                      "Country", "Value(USD$1,000)", "Weight(TNE)"]
except Exception as e:
    st.error(f"❌ ファイル読み込みエラー: {e}")
    st.stop()

# 時間正規化 (3-tuple)
parsed = df_raw["Time"].map(normalize_time)
df_raw["Time_clean"] = [p[0] for p in parsed]
df_raw["Is_provisional"] = [p[1] for p in parsed]
df_raw["Is_monthly"] = [p[2] for p in parsed]
df_raw["Time"] = df_raw["Time_clean"]

# 粒度判定 (monthly / yearly / mixed)
monthly_count = df_raw["Is_monthly"].sum()
total_count = len(df_raw)
if monthly_count == total_count:
    detected_granularity = "monthly"
    granularity_label = "月別"
elif monthly_count == 0:
    detected_granularity = "yearly"
    granularity_label = "年度別"
else:
    st.error(
        f"❌ **時間値の形式が混在しています。**\n\n"
        f"月別形式: {monthly_count} 件、非月別形式: {total_count - monthly_count} 件。\n\n"
        "本ツールは月別または年度別のいずれか単一形式のみに対応しています。"
        "財政部関税統計サイトでダウンロードする際、いずれか一方を選択してください。"
    )
    st.stop()

# ========== プレビュー ==========

st.success(f"✅ 読み込み完了: {len(df_raw)} 件")

col1, col2, col3, col4 = st.columns(4)

with col1:
    n_records = len(df_raw)
    st.metric("レコード数", f"{n_records:,}")

with col2:
    # 粒度に応じた並べ替え
    import re as _re_local
    def _sk(s):
        s = str(s)
        m = _re_local.fullmatch(r"(\d{4})/(\d{1,2})", s)
        if m: return (int(m.group(1)), int(m.group(2)), 0)
        m = _re_local.fullmatch(r"(\d{4})/(\d{1,2})~\d{4}/\d{1,2}", s)
        if m: return (int(m.group(1)), int(m.group(2)), 1)
        m = _re_local.fullmatch(r"(\d{4})", s)
        if m: return (int(m.group(1)), 0, 2)
        return (9999, 99, 9)
    months = sorted(df_raw["Time"].unique(), key=_sk)
    if detected_granularity == "yearly":
        st.metric("期間数", f"{len(months)} 年")
    else:
        st.metric("月数", f"{len(months)} ヶ月")

with col3:
    n_countries = df_raw["Country"].nunique()
    st.metric("国家数", f"{n_countries}")

with col4:
    types = df_raw["Imports/Exports"].unique()
    st.metric("区分", " / ".join(types))

# 粒度の通知
if detected_granularity == "yearly":
    st.info(
        f"📅 **年度別データを検出しました** — 各表の時間軸は「年」単位で集計されます "
        f"(月別ピボット・前月比は表示されません)。"
    )

# ========== 設定セクション ==========

st.divider()
st.subheader("⚙️ 生成設定")

col_a, col_b = st.columns(2)

with col_a:
    available_types = list(df_raw["Imports/Exports"].unique())
    has_imports = "Imports" in available_types
    has_exports = "Exports" in available_types

    st.markdown("**輸出入区分** *(複数選択可)*")
    cb_col1, cb_col2 = st.columns(2)
    with cb_col1:
        sel_imports = st.checkbox(
            "Imports (輸入)",
            value=has_imports,
            disabled=not has_imports,
            help="原始データに含まれていない場合は選択不可",
        )
    with cb_col2:
        sel_exports = st.checkbox(
            "Exports (輸出)",
            value=has_exports and not has_imports,  # Imports があれば Imports のみデフォルト
            disabled=not has_exports,
        )

    # 既定挙動: 一つもない場合は両方選択 (両方 available なら両方既定で選択)
    selected_trades = []
    if sel_imports:
        selected_trades.append("Imports")
    if sel_exports:
        selected_trades.append("Exports")

    if not selected_trades:
        st.warning("⚠️ 少なくとも1つの区分を選択してください。")

with col_b:
    if selected_trades:
        df_filtered = df_raw[df_raw["Imports/Exports"].isin(selected_trades)]
        period_str = (
            f"{df_filtered['Time'].min()} 〜 {df_filtered['Time'].max()}"
        )
    else:
        df_filtered = df_raw  # 警告表示のためのフォールバック
        period_str = ""
    st.text_input("対象期間", value=period_str, disabled=True)

# 品目情報
commodity_info = detect_commodity_info(df_filtered)
commodity_code = commodity_info["primary_code"]
commodity_settings = load_commodity_settings(commodity_code)

with st.expander("📋 品目情報", expanded=True):
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.text(f"HS code: {commodity_code}")
        st.text(f"英文品名: {commodity_info['primary_desc']}")
    with col_p2:
        jp_name = commodity_settings.get("jp_name") or commodity_info["primary_desc"]
        is_registered = commodity_settings.get("jp_name") is not None
        st.text(f"日文品名: {jp_name}")
        if is_registered:
            st.text(
                f"単価レンジ: {commodity_settings['price_min']}〜"
                f"{commodity_settings['price_max']} USD/TNE"
            )
        else:
            st.warning("⚠️ この品目は未登録です。「品目設定」ページで登録すると、"
                       "日文名と単価市況レンジが反映されます。")

    if commodity_info["is_multi"]:
        st.warning(
            f"⚠️ 原始データに複数の品目コード ({len(commodity_info['codes'])} 件) が含まれています。"
            "本ツールはすべて合算して処理します。"
        )

# 国家対照チェック
mapping = load_country_mapping()
raw_countries_in_data = set(df_filtered["Country"].dropna().unique())
unmapped = raw_countries_in_data - set(mapping.keys())

if unmapped:
    with st.expander("🆕 新規国家の検出", expanded=True):
        st.warning(
            f"対照表に未登録の国家が **{len(unmapped)} 件** あります: "
            f"{', '.join(sorted(unmapped))}"
        )
        st.markdown(
            "生成は可能ですが、これらの国家は英文名のまま表示されます。"
            "日文名を設定するには「**国名対照**」ページで編集してください。"
        )

# ========== 生成ボタン ==========

st.divider()

button_disabled = len(selected_trades) == 0
if st.button("🚀 ダッシュボード生成", type="primary",
             use_container_width=True, disabled=button_disabled):
    with st.spinner("生成中..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_tmp:
                output_path = out_tmp.name

            result = generate_dashboard(
                input_path=tmp_path,
                output_path=output_path,
                trade_type=selected_trades,
            )

            # ファイル名を区分に応じて生成
            trade_part = "_".join(
                "輸入" if t == "Imports" else "輸出"
                for t in result["trade_types"]
            )
            st.session_state["last_result"] = result
            st.session_state["last_output_path"] = output_path
            st.session_state["last_filename"] = (
                f"{result['commodity_jp']}_{trade_part}_Dashboard.xlsx"
            )
            st.success("✅ 生成完了!")
        except Exception as e:
            st.error(f"❌ 生成エラー: {e}")
            st.exception(e)

# ========== 結果表示 ==========

if "last_result" in st.session_state:
    result = st.session_state["last_result"]
    output_path = st.session_state["last_output_path"]
    filename = st.session_state["last_filename"]

    st.divider()
    st.subheader("📥 生成結果")

    col_r1, col_r2 = st.columns([2, 1])

    with col_r1:
        st.markdown(f"**サマリー**: {result['summary']}")
        st.caption(
            f"期間: {result['period']}　|　"
            f"国家: {result['n_countries']}　|　"
            f"レコード: {result['n_records']}"
        )
        if result["new_countries"]:
            st.warning(
                f"新規国家 (英文名のまま表示): {', '.join(result['new_countries'])}"
            )
        if result["provisional_months"]:
            st.info(
                f"速報値月: {', '.join(result['provisional_months'])}"
            )

    with col_r2:
        with open(output_path, "rb") as f:
            st.download_button(
                "💾 ダウンロード",
                data=f.read(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

    # ダッシュボードのプレビュー (主要な表のみ)
    st.divider()
    st.subheader("🔍 プレビュー")
    st.caption("生成されたダッシュボードの主要表を表示します。完全版はダウンロードしてご確認ください。")

    try:
        from openpyxl import load_workbook
        wb_preview = load_workbook(output_path, data_only=True)

        # 各区分の表 3 (ランキング) を順に表示
        for ts in result["per_trade_summaries"]:
            t = ts["trade_type"]
            sheet_name = "輸入分析ダッシュボード" if t == "Imports" else "輸出分析ダッシュボード"
            trade_label = "輸入" if t == "Imports" else "輸出"
            ws = wb_preview[sheet_name]

            # 表 3 ヘッダ位置を探す
            t3_header_row = None
            for r in range(1, ws.max_row + 1):
                if ws.cell(row=r, column=2).value == "順位":
                    t3_header_row = r
                    break

            if t3_header_row:
                ranking_data = []
                for i in range(ts["n_countries"]):
                    r = t3_header_row + 1 + i
                    ranking_data.append({
                        "順位": ws.cell(row=r, column=2).value,
                 
