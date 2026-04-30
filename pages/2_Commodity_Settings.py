"""
品目設定 管理ページ

HS code ごとに、日文品名・単価妥当レンジ・備考を編集できます。
"""

import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent.parent))
from _theme import inject_theme, hero, footer

DATA_DIR = Path(__file__).parent.parent / "data"
COMMODITY_FILE = DATA_DIR / "commodity_settings.json"

st.set_page_config(
    page_title="品目設定",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

inject_theme()

hero(
    eyebrow="MASTER DATA",
    title="品目設定",
    desc=(
        "HS code ごとに、日文品名と単価妥当レンジを管理します。"
        "ここで設定した単価レンジは、表3「備考」欄の異常値検出に使用されます。"
    ),
)

# ========== 読み込み ==========

with open(COMMODITY_FILE, encoding="utf-8") as f:
    data = json.load(f)

# 「_」で始まるサンプルキーは除外
real_commodities = {
    k: v for k, v in data["commodities"].items()
    if not k.startswith("_")
}

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

# ========== サマリー ==========

st.markdown(
    f"""
    <div class="section-card">
      <div class="section-label">CURRENT STATUS</div>
      <div style="display: flex; gap: 2.5rem; flex-wrap: wrap; margin-top: 0.4rem;">
        <div>
          <div style="color:#64748b; font-size:0.78rem;">登録品目数</div>
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

st.markdown("### 品目編集")
st.caption(
    "HS code は財政部の商品分類コード（CCC code）です。"
    "最終行の下のスペースに新規追加できます。"
)

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

# ========== アクションボタン ==========

col1, col2, col3 = st.columns([1.2, 1.2, 4])

with col1:
    if st.button("変更を保存", type="primary", use_container_width=True):
        if edited["HS code"].duplicated().any():
            st.error(
                "HS code に重複があります。\n\n"
                "重複している行を確認し、修正または削除してから再度保存してください。"
            )
        elif edited["HS code"].isna().any() or edited["日文品名"].isna().any():
            st.error(
                "必須項目が未入力です。\n\n"
                "「HS code」と「日文品名」はすべての行で入力が必要です。"
            )
        elif (edited["単価下限 (USD/TNE)"] >= edited["単価上限 (USD/TNE)"]).any():
            st.error(
                "単価レンジが正しくありません。\n\n"
                "「単価下限」は「単価上限」より小さい値にしてください。"
            )
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
            st.success("保存しました。")
            st.rerun()

with col2:
    st.download_button(
        "JSONバックアップ",
        data=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="commodity_settings.json",
        mime="application/json",
        use_container_width=True,
    )

# ========== ヒント ==========

st.divider()

with st.expander("項目の説明", expanded=False):
    st.markdown(
        """
        **HS code（CCC）**
        　財政部の商品分類コード。原始データの「Commodity Code」欄の値と一致させてください。

        **日文品名**
        　ダッシュボードのタイトルに表示される名称です。

        **単価妥当レンジ**
        　平均単価がこのレンジを外れると、表3 の「備考」欄に
        　「※単価要確認」と自動的に表示されます。市況急変・データ異常の早期発見にご利用ください。

        **未登録品目の扱い**
        　ここに登録されていない HS code の場合、英文品名がそのまま使われ、
        　単価チェックは実施されません（上限・下限が `0` と `999999` のため）。
        """
    )

with st.expander("単価レンジ設定の目安", expanded=False):
    st.markdown(
        """
        化学品ごとの単価レンジ設定例：

        - **ナフサ系（基礎化学品）** … 400 〜 800 USD/TNE
        - **触媒・添加剤（TEDA、TOYOCAT 等）** … 2,500 〜 5,500 USD/TNE
        - **特殊化学品 / 添加剤** … 5,000 〜 15,000 USD/TNE 程度

        実際の市況データを参照しながら適宜調整してください。
        一度設定すれば次回以降も適用されます。
        """
    )

# ========== フッター ==========

footer()
