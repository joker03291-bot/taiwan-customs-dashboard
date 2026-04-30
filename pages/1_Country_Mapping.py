"""
国名対照表 管理ページ

英文国名 → 日文国名 の対照を編集・追加できます。
"""

import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

# 親ディレクトリの _theme をインポート可能にする
sys.path.insert(0, str(Path(__file__).parent.parent))
from _theme import inject_theme, hero, footer

DATA_DIR = Path(__file__).parent.parent / "data"
COUNTRY_FILE = DATA_DIR / "country_mapping.json"

st.set_page_config(
    page_title="国名対照",
    page_icon="🌏",
    layout="wide",
    initial_sidebar_state="collapsed",
)

inject_theme()

hero(
    eyebrow="MASTER DATA",
    title="国名対照表",
    desc=(
        "英文国名と日文国名の対照を管理します。"
        "ここで設定した内容は、すべてのダッシュボード生成に反映されます。"
    ),
)

# ========== 読み込み ==========

with open(COUNTRY_FILE, encoding="utf-8") as f:
    data = json.load(f)

entries = [e for e in data["countries"] if "en" in e and "jp" in e]

df = pd.DataFrame(entries)
df = df[["order", "en", "jp"]].copy()
df.columns = ["順序", "国名(英)", "国名(日)"]
df = df.sort_values("順序").reset_index(drop=True)

# ========== サマリー ==========

st.markdown(
    f"""
    <div class="section-card">
      <div class="section-label">CURRENT STATUS</div>
      <div style="display: flex; gap: 2.5rem; flex-wrap: wrap; margin-top: 0.4rem;">
        <div>
          <div style="color:#64748b; font-size:0.78rem;">登録数</div>
          <div style="color:#1a202c; font-size:1.05rem; font-weight:600;">{len(df)} 件</div>
        </div>
        <div>
          <div style="color:#64748b; font-size:0.78rem;">最終更新</div>
          <div style="color:#1a202c; font-size:1.05rem; font-weight:600;">
            {data.get('_last_updated', '-')}
          </div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ========== 編集テーブル ==========

st.markdown("### 対照表編集")
st.caption("セルをダブルクリックで編集。最終行の下のスペースに新規追加できます。")

edited = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "順序": st.column_config.NumberColumn(
            "順序", help="表示順（小さいほど上）", min_value=1, step=1, width="small"
        ),
        "国名(英)": st.column_config.TextColumn("国名(英)", width="medium"),
        "国名(日)": st.column_config.TextColumn("国名(日)", width="medium"),
    },
)

# ========== アクションボタン ==========

col1, col2, col3 = st.columns([1.2, 1.2, 4])

with col1:
    if st.button("変更を保存", type="primary", use_container_width=True):
        if edited["国名(英)"].duplicated().any():
            st.error(
                "英文国名に重複があります。\n\n"
                "重複している行を確認し、修正または削除してから再度保存してください。"
            )
        elif edited["国名(英)"].isna().any() or edited["国名(日)"].isna().any():
            st.error(
                "空白のセルがあります。\n\n"
                "「国名(英)」と「国名(日)」はすべて入力必須です。"
            )
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
            st.success("保存しました。")
            st.rerun()

with col2:
    st.download_button(
        "JSONバックアップ",
        data=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="country_mapping.json",
        mime="application/json",
        use_container_width=True,
    )

# ========== ヒント ==========

st.divider()

with st.expander("使い方のヒント", expanded=False):
    st.markdown(
        """
        **新規追加**
        　テーブル最終行の下のスペースに直接入力できます。

        **削除**
        　行頭のチェックボックスで選択 → `Delete` キーで削除。

        **順序の意味**
        　「順序」が小さい国家が、ダッシュボードの表 1・表 2 で上に表示されます。

        **新規マークについて**
        　原始データに出現したが対照表にない国家は、ダッシュボード生成時に自動で
        　末尾に追加されます（橙色マーカー付き）。後でこのページで日文名を設定してください。

        **予備対照について**
        　順序 100 以上は「予備」（よく出る国家を先に登録しておくもの）。
        　実際の原始データに出現するまでダッシュボードには表示されません。
        """
    )

# ========== フッター ==========

footer()
