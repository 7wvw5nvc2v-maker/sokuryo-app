import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

from st_aggrid import (
    AgGrid,
    GridOptionsBuilder,
    GridUpdateMode,
    DataReturnMode,
    JsCode,
)

# =========================================================
# ページ設定
# =========================================================
st.set_page_config(
    page_title="水準測量・器高式計算アプリ",
    layout="wide",
)

# =========================================================
# デザイン
# =========================================================
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f1f5f9;
    }
    h1 {
        color: #1e3a5f;
    }
    div[data-testid="stNumberInput"] {
        background-color: #ffffff;
        padding: 10px;
        border-radius: 8px;
    }
    .stButton > button {
        border-radius: 6px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# タイトル・使い方
# =========================================================
st.title("水準測量・器高式計算アプリ")
st.write("水準測量で使用する器高式の計算を自動化するアプリです。")

with st.expander("使い方", expanded=True):
    st.write(
        """
        ① 基準GHを入力します。

        ② 水色のセルに測定値を入力します。

        ③ 黄色のセルは器高式によって自動計算されます。

        ④ 行が足りない場合は「＋ 行を追加」を押します。

        ⑤ 不要な行は選択して「− 選択行を削除」を押します。

        ⑥ 必要に応じてExcelファイルとして保存できます。
        """
    )

st.markdown("🟦 **入力**　　🟨 **自動計算**　　⬜ **累計距離**")

# =========================================================
# 基準GH
# =========================================================
base_gh = st.number_input(
    "基準地盤高（GH）",
    value=500.000,
    step=0.001,
    format="%.3f",
)

# =========================================================
# データ作成・入力データ整理
# =========================================================
def create_empty_data(rows=10):
    return pd.DataFrame(
        {
            "測点": ["" for _ in range(rows)],
            "距離": [None for _ in range(rows)],
            "累計距離": [None for _ in range(rows)],
            "BS": [None for _ in range(rows)],
            "TP": [None for _ in range(rows)],
            "IP": [None for _ in range(rows)],
        }
    )
