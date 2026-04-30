"""
共通テーマ・スタイル定義

すべてのページで `from _theme import inject_theme` してから
`inject_theme()` を最初に呼び出してください。
"""

import streamlit as st


# ========== カラーパレット ==========
# 配色: 落ち着いた商務ブルー基調。彩度は意図的に抑えめ。
COLOR_PRIMARY = "#1e3a5f"        # 深ブルー (タイトル・主要アクション)
COLOR_PRIMARY_LIGHT = "#3b5a85"  # ホバー・アクセント
COLOR_BG = "#f5f7fa"             # ページ背景 (薄グレーブルー)
COLOR_SURFACE = "#ffffff"        # カード・セクション背景
COLOR_BORDER = "#e2e8f0"         # 罫線
COLOR_TEXT = "#1a202c"           # 本文
COLOR_TEXT_MUTED = "#64748b"     # 補助テキスト
COLOR_SUCCESS = "#0f7c5e"        # 成功 (落ち着いたグリーン)
COLOR_WARNING = "#b45309"        # 警告 (落ち着いたアンバー)
COLOR_ERROR = "#b91c1c"          # エラー


_CSS = f"""
<style>
/* ========== 全体 ========== */
html, body, [class*="css"] {{
    font-family: -apple-system, "Hiragino Sans", "Hiragino Kaku Gothic ProN",
                 "Yu Gothic UI", "Meiryo", "MS PGothic", "Helvetica Neue", sans-serif;
    color: {COLOR_TEXT};
}}

.stApp {{
    background-color: {COLOR_BG};
}}

/* メインコンテンツの最大幅 */
.main .block-container {{
    max-width: 1100px;
    padding-top: 2.5rem;
    padding-bottom: 4rem;
}}

/* ========== タイトル類 ========== */
h1 {{
    color: {COLOR_PRIMARY};
    font-weight: 600;
    letter-spacing: -0.01em;
}}

h2, h3 {{
    color: {COLOR_PRIMARY};
    font-weight: 600;
}}

/* ========== ボタン ========== */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {{
    background-color: {COLOR_PRIMARY};
    color: white;
    border: none;
    font-weight: 500;
    transition: background-color 0.15s ease;
}}

.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {{
    background-color: {COLOR_PRIMARY_LIGHT};
    color: white;
}}

.stButton > button[kind="primary"]:disabled {{
    background-color: #cbd5e1;
    color: #94a3b8;
}}

.stButton > button:not([kind="primary"]),
.stDownloadButton > button:not([kind="primary"]) {{
    border: 1px solid {COLOR_BORDER};
    background: white;
    color: {COLOR_TEXT};
}}

/* ========== ファイルアップローダー ========== */
section[data-testid="stFileUploaderDropzone"] {{
    border: 2px dashed {COLOR_PRIMARY_LIGHT};
    background-color: white;
    border-radius: 8px;
    padding: 1.5rem 1rem;
    transition: all 0.2s ease;
}}

section[data-testid="stFileUploaderDropzone"]:hover {{
    border-color: {COLOR_PRIMARY};
    background-color: #fafbfd;
}}

/* ========== Hero セクション ========== */
.hero {{
    background: linear-gradient(135deg, #ffffff 0%, #eef2f8 100%);
    border: 1px solid {COLOR_BORDER};
    border-radius: 12px;
    padding: 2rem 2rem 1.75rem;
    margin-bottom: 1.75rem;
}}

.hero-eyebrow {{
    color: {COLOR_PRIMARY_LIGHT};
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}}

.hero-title {{
    color: {COLOR_PRIMARY};
    font-size: 1.65rem;
    font-weight: 700;
    margin: 0 0 0.5rem 0;
    letter-spacing: -0.01em;
}}

.hero-desc {{
    color: {COLOR_TEXT_MUTED};
    font-size: 0.95rem;
    margin: 0 0 1rem 0;
    line-height: 1.6;
}}

.hero-cta {{
    display: inline-block;
    color: {COLOR_PRIMARY};
    font-size: 0.85rem;
    font-weight: 500;
    border-top: 1px solid {COLOR_BORDER};
    padding-top: 0.75rem;
    margin-top: 0.25rem;
}}

/* ========== ステップ・タイムライン ========== */
.timeline {{
    position: relative;
    margin: 0.5rem 0 1rem 0;
    padding-left: 0;
}}

.step {{
    position: relative;
    padding: 0 0 1.25rem 3rem;
    border-left: 2px solid {COLOR_BORDER};
    margin-left: 1.1rem;
}}

.step:last-child {{
    border-left-color: transparent;
    padding-bottom: 0.25rem;
}}

.step-number {{
    position: absolute;
    left: -1.2rem;
    top: -0.1rem;
    width: 2.2rem;
    height: 2.2rem;
    border-radius: 50%;
    background: {COLOR_PRIMARY};
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 600;
    font-size: 0.95rem;
    box-shadow: 0 0 0 4px {COLOR_BG};
}}

.step-title {{
    color: {COLOR_PRIMARY};
    font-size: 1.02rem;
    font-weight: 600;
    margin: 0.1rem 0 0.4rem 0;
}}

.step-body {{
    color: {COLOR_TEXT};
    font-size: 0.9rem;
    line-height: 1.65;
    margin: 0;
}}

.step-body a {{
    color: {COLOR_PRIMARY};
    text-decoration: underline;
    text-underline-offset: 2px;
}}

.step-note {{
    background: #f8fafc;
    border-left: 3px solid {COLOR_PRIMARY_LIGHT};
    padding: 0.6rem 0.9rem;
    margin-top: 0.6rem;
    border-radius: 4px;
    font-size: 0.85rem;
    color: {COLOR_TEXT_MUTED};
    line-height: 1.6;
}}

.step-note strong {{
    color: {COLOR_TEXT};
}}

.step-note table {{
    border-collapse: collapse;
    margin-top: 0.4rem;
    font-size: 0.83rem;
    width: 100%;
}}

.step-note th, .step-note td {{
    padding: 0.35rem 0.6rem;
    border: 1px solid {COLOR_BORDER};
    text-align: left;
}}

.step-note th {{
    background: white;
    color: {COLOR_PRIMARY};
    font-weight: 600;
}}

/* ========== セクションカード ========== */
.section-card {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}}

.section-label {{
    color: {COLOR_PRIMARY_LIGHT};
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    margin-bottom: 0.4rem;
}}

/* ========== メトリック ========== */
[data-testid="stMetric"] {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
    padding: 0.85rem 1.1rem;
}}

[data-testid="stMetricLabel"] {{
    color: {COLOR_TEXT_MUTED};
    font-size: 0.8rem;
}}

[data-testid="stMetricValue"] {{
    color: {COLOR_PRIMARY};
    font-weight: 600;
}}

/* ========== alert系 ========== */
div[data-testid="stAlert"] {{
    border-radius: 8px;
    border-left-width: 4px;
}}

/* ========== expander ========== */
.streamlit-expanderHeader,
[data-testid="stExpander"] summary {{
    font-weight: 500;
    color: {COLOR_TEXT};
}}

[data-testid="stExpander"] {{
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
    background: {COLOR_SURFACE};
}}

/* ========== 区切り線 ========== */
hr {{
    border: none;
    border-top: 1px solid {COLOR_BORDER};
    margin: 1.5rem 0;
}}

/* ========== サイドバー ========== */
section[data-testid="stSidebar"] {{
    background-color: #fafbfd;
    border-right: 1px solid {COLOR_BORDER};
}}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    color: {COLOR_PRIMARY};
}}

/* ========== フッター ========== */
.app-footer {{
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid {COLOR_BORDER};
    color: {COLOR_TEXT_MUTED};
    font-size: 0.8rem;
    display: flex;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 0.5rem;
}}

/* ========== 結果バナー ========== */
.result-banner {{
    background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
    border: 1px solid #bbf7d0;
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    margin: 1rem 0;
}}

.result-banner-title {{
    color: {COLOR_SUCCESS};
    font-size: 1.05rem;
    font-weight: 600;
    margin: 0 0 0.35rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}}

.result-banner-desc {{
    color: {COLOR_TEXT};
    font-size: 0.88rem;
    margin: 0;
    line-height: 1.6;
}}

/* ========== Streamlit デフォルト要素を隠す ========== */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}

/* ========== info card (Step下のサポート情報など) ========== */
.info-card {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin: 0.5rem 0;
}}

.info-card-title {{
    color: {COLOR_PRIMARY};
    font-size: 0.92rem;
    font-weight: 600;
    margin: 0 0 0.35rem 0;
}}

.info-card-body {{
    color: {COLOR_TEXT_MUTED};
    font-size: 0.85rem;
    margin: 0;
    line-height: 1.6;
}}
</style>
"""


def inject_theme():
    """各ページの先頭で呼び出してテーマを適用する。"""
    st.markdown(_CSS, unsafe_allow_html=True)


def hero(eyebrow: str, title: str, desc: str, cta: str = ""):
    """Hero セクションを表示。"""
    cta_html = f'<div class="hero-cta">{cta}</div>' if cta else ""
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-eyebrow">{eyebrow}</div>
            <h1 class="hero-title">{title}</h1>
            <p class="hero-desc">{desc}</p>
            {cta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def footer():
    """フッターを表示。"""
    st.markdown(
        """
        <div class="app-footer">
            <div>関税統計ダッシュボード自動生成ツール</div>
            <div>お問い合わせ：Mars　|　v2.0</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
