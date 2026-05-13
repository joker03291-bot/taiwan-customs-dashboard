"""
品目設定 管理ページ

HS code ごとに、日文品名・単価妥当レンジ・備考を編集できます。
"""

import json
from pathlib import Path

import pandas as pd
import streamlit as st

DATA_DIR = Path(__file__).parent.parent / "data"
COMMODITY_FILE = DATA_DIR / "commodity_settings.json"

st.set_page_config(page_title="品目設定", page_icon="🧪", layout="wide")
st.title("🧪 品目設定")
st.caption("HS code ごとに日文品名と単価妥当レンジを管理します。")

# 読み込み
with open(COMMODITY_FILE, encoding="utf-8") as f:
    data = json.load(f)

# 「_」で始まるサンプルキーは除外
real_commodities = {
    k: v for k, v in data["commodities"].items()
    if not k.startswith("_")
}

# DataFrame 化
rows = []
for code, info in real_commodities.items():
    rows.append({
        "HS code": code,
        "日文品名": info.get("jp_name", ""),
        "単価下限 (USD/TNE)": info.get("price_min", 0),
        "単価上限 (USD/TNE)": info.get("price_max", 999999),
        "備考": info.get("note", ""),
    })

df = pd.DataFrame(rows) if rows else pd.DataFrame(
    columns=["HS code", "日文品名", "単価下限 (USD/TNE)", "単価上限 (USD/TNE)", "備考"]
)

st.markdown(f"**登録品目数**: {len(df)} 件　|　**最終更新**: {data.get('_last_updated', '-')}")

# 編集可能なテーブル
st.markdown("### ✏️ 品目編集")
st.caption("HS code は財政部の商品分類コード (CCC code)。下の行に新規追加できます。")

edited = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "HS code": st.column_config.TextColumn(
            "HS code", help="財政部の商品分類コード", width="medium"
        ),
        "日文品名": st.column_config.TextColumn("日文品名", width="medium"),
        "単価下限 (USD/TNE)": st.column_config.NumberColumn(
            "単価下限 (USD/TNE)",
            help="この値より低い平均単価は『要確認』マークが表示されます",
            min_value=0, step=10,
        ),
        "単価上限 (USD/TNE)": st.column_config.NumberColumn(
            "単価上限 (USD/TNE)",
            help="この値より高い平均単価は『要確認』マークが表示されます",
            min_value=0, step=10,
        ),
        "備考": st.column_config.TextColumn("備考", width="large"),
    },
)

col1, col2, col3 = st.columns([1, 1, 4])

with col1:
    if st.button("💾 保存", type="primary"):
        if edited["HS code"].duplicated().any():
            st.error("HS code に重複があります")
        elif edited["HS code"].isna().any() or edited["日文品名"].isna().any():
            st.error("HS code と日文品名は必須です")
        elif (edited["単価下限 (USD/TNE)"] >= edited["単価上限 (USD/TNE)"]).any():
            st.error("単価下限は上限より小さい値にしてください")
        else:
            new_commodities = {}
            for _, row in edited.iterrows():
                new_commodities[str(row["HS code"])] = {
                    "jp_name": row["日文品名"],
                    "price_min": int(row["単価下限 (USD/TNE)"]),
                    "price_max": int(row["単価上限 (USD/TNE)"]),
                    "note": row["備考"] or "",
                }

            new_data = {
                "_comment": data.get("_comment", ""),
                "_last_updated": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "commodities": new_commodities,
                "default": data["default"],
            }
            with open(COMMODITY_FILE, "w", encoding="utf-8") as f:
                json.dump(new_data, f, ensure_ascii=False, indent=2)
            st.success("✅ 保存しました")
            st.rerun()

with col2:
    st.download_button(
        "📥 JSONダウンロード",
        data=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="commodity_settings.json",
        mime="application/json",
    )

st.divider()

st.markdown("### 📌 使い方のヒント")
st.markdown(
    """
    - **HS code (CCC)**: 財政部の商品分類コード。原始データの「Commodity Code」欄の値と一致させてください
    - **日文品名**: ダッシュボードのタイトルに表示される名称
    - **単価妥当レンジ**: 平均単価がこのレンジを外れると、表3の「備考」欄に「※単価要確認」と
      自動的に表示されます。市況急変・データ異常の早期発見用
    - **未登録品目**: ここに登録されていない HS code の場合、英文品名がそのまま使われ、
      単価チェックは実施されません (上限・下限が `0` と `999999` のため)
    """
)

st.divider()

st.markdown("### 💡 単価レンジ設定の目安")
st.markdown(
    """
    化学品ごとの単価レンジ設定例:

    - **ナフサ系 (基礎化学品)**: 400〜800 USD/TNE
    - **触媒・添加剤 (TEDA, TOYOCAT 等)**: 2,500〜5,500 USD/TNE
    - **特殊化学品 / 添加剤**: 5,000〜15,000 USD/TNE 程度

    実際の市況データを参照しながら適宜調整してください。
    一度設定すれば次回以降も適用されます。
    """
)
