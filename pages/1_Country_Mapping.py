"""
国名対照表 管理ページ

英文国名 → 日文国名 の対照を編集・追加できます。
"""

import json
from pathlib import Path

import pandas as pd
import streamlit as st

DATA_DIR = Path(__file__).parent.parent / "data"
COUNTRY_FILE = DATA_DIR / "country_mapping.json"

st.set_page_config(page_title="国名対照", page_icon="🌏", layout="wide")
st.title("🌏 国名対照表")
st.caption("英文国名 → 日文国名の対照を管理します。順序は表示順位に影響します（小さいほど上）。")

# 読み込み
with open(COUNTRY_FILE, encoding="utf-8") as f:
    data = json.load(f)

entries = [e for e in data["countries"] if "en" in e and "jp" in e]

# DataFrame 化
df = pd.DataFrame(entries)
df = df[["order", "en", "jp"]].copy()
df.columns = ["順序", "国名(英)", "国名(日)"]
df = df.sort_values("順序").reset_index(drop=True)

st.markdown(f"**現在の登録数**: {len(df)} 件　|　**最終更新**: {data.get('_last_updated', '-')}")

# 編集可能なテーブル
st.markdown("### ✏️ 対照表編集")
st.caption("セルをダブルクリックで編集。下の行に追加することもできます。")

edited = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "順序": st.column_config.NumberColumn(
            "順序", help="表示順 (小さいほど上)", min_value=1, step=1, width="small"
        ),
        "国名(英)": st.column_config.TextColumn("国名(英)", width="medium"),
        "国名(日)": st.column_config.TextColumn("国名(日)", width="medium"),
    },
)

col1, col2, col3 = st.columns([1, 1, 4])

with col1:
    if st.button("💾 保存", type="primary"):
        # 入力検証
        if edited["国名(英)"].duplicated().any():
            st.error("英文国名に重複があります")
        elif edited["国名(英)"].isna().any() or edited["国名(日)"].isna().any():
            st.error("空白のセルがあります")
        else:
            new_data = {
                "_comment": data.get("_comment", ""),
                "_last_updated": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "countries": [
                    {"en": row["国名(英)"], "jp": row["国名(日)"], "order": int(row["順序"])}
                    for _, row in edited.iterrows()
                ],
            }
            with open(COUNTRY_FILE, "w", encoding="utf-8") as f:
                json.dump(new_data, f, ensure_ascii=False, indent=2)
            st.success("✅ 保存しました")
            st.rerun()

with col2:
    # JSON ダウンロード (バックアップ用)
    st.download_button(
        "📥 JSONダウンロード",
        data=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="country_mapping.json",
        mime="application/json",
    )

st.divider()
st.markdown("### 📌 使い方のヒント")
st.markdown(
    """
    - **新規追加**: テーブル最終行の下のスペースに直接入力できます
    - **削除**: 行を選択して `Delete` キーを押すか、行頭のチェックボックスから削除
    - **順序の意味**: 「順序」が小さい国家がダッシュボードの表 1・表 2 で上に表示されます
    - **★新規マーク**: 原始データに出現したが対照表にない国家は、ダッシュボード生成時に自動で
      末尾に追加されます（橙色マーカー付き）。後でこのページで日文名を設定してください
    - **予備対照**: 順序 100 以上は「予備」(よく出る国家を先に登録しておくもの)。
      実際の原始データに出現するまでダッシュボードには表示されません
    """
)
