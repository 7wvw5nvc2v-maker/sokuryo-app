import streamlit as st
import pandas as pd

from io import BytesIO
from openpyxl import Workbook

from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder
from st_aggrid import JsCode


st.set_page_config(
    page_title="測量自動計算",
    layout="wide"
)


# =========================
# 全体のデザイン
# =========================

st.markdown("""
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
    background-color: #2f6690;
    color: white;
    border-radius: 6px;
    border: none;
}

.stButton > button:hover {
    background-color: #245273;
}
</style>
""", unsafe_allow_html=True)


# =========================
# タイトル
# =========================

st.title("測量自動計算")


# =========================
# 基準GH
# =========================

base_gh = st.number_input(
    "基準GH",
    value=500.000,
    step=0.001,
    format="%.3f"
)


# =========================
# 初期データ
# =========================

if "data" not in st.session_state:

    st.session_state.data = pd.DataFrame({
        "測点": [""] * 10,
        "距離": [None] * 10,
        "累計距離": [None] * 10,
        "BS": [None] * 10,
        "IH": [None] * 10,
        "TP": [None] * 10,
        "IP": [None] * 10,
        "GH": [None] * 10,
    })


# =========================
# 累計距離を計算
# =========================

data = st.session_state.data.copy()

total = 0.0
cumulative = []

for distance in data["距離"]:

    if pd.isna(distance):

        cumulative.append(None)

    else:

        try:

            total += float(distance)
            cumulative.append(total)

        except (ValueError, TypeError):

            cumulative.append(None)


# =========================
# IH・GHを計算
# =========================

ih_values = []
gh_values = []

current_gh = float(base_gh)
current_ih = None


for i, (bs, tp, ip) in enumerate(
    zip(
        data["BS"],
        data["TP"],
        data["IP"]
    )
):

    # BSがあればIHを更新

    if pd.notna(bs):

        current_ih = current_gh + float(bs)
        ih_values.append(current_ih)

    else:

        ih_values.append(None)


    # TPがあればGHを更新

    if current_ih is not None and pd.notna(tp):

        current_gh = current_ih - float(tp)
        gh_values.append(current_gh)


    # IPがあればGHを更新

    elif current_ih is not None and pd.notna(ip):

        current_gh = current_ih - float(ip)
        gh_values.append(current_gh)


    # 1行目は基準GH

    elif i == 0:

        gh_values.append(current_gh)


    else:

        gh_values.append(None)


# =========================
# 計算結果をデータに反映
# =========================

data["累計距離"] = cumulative
data["IH"] = ih_values
data["GH"] = gh_values


# =========================
# セルの色
# =========================

blue_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#a7ddf5'
    };
}
""")


yellow_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#ffff00'
    };
}
""")


# =========================
# 表の設定
# =========================

gb = GridOptionsBuilder.from_dataframe(data)


# 基本設定
gb.configure_default_column(
    editable=True,
    resizable=True,
    sortable=False
)


# 測点
gb.configure_column(
    "測点",
    editable=True,
    width=120
)


# 距離
gb.configure_column(
    "距離",
    editable=True,
    width=110
)


# 累計距離
gb.configure_column(
    "累計距離",
    editable=False,
    width=120
)


# BS → 水色
gb.configure_column(
    "BS",
    editable=True,
    width=110,
    cellStyle=blue_style
)


# IH → 黄色
gb.configure_column(
    "IH",
    editable=False,
    width=110,
    cellStyle=yellow_style
)


# TP → 水色
gb.configure_column(
    "TP",
    editable=True,
    width=110,
    cellStyle=blue_style
)


# IP → 水色
gb.configure_column(
    "IP",
    editable=True,
    width=110,
    cellStyle=blue_style
)


# GH → 黄色
gb.configure_column(
    "GH",
    editable=False,
    width=110,
    cellStyle=yellow_style
)


# 行追加・削除
gb.configure_grid_options(
    stopEditingWhenCellsLoseFocus=True
)


grid_options = gb.build()


# =========================
# 表を表示
# =========================

grid_return = AgGrid(
    data,
    gridOptions=grid_options,
    height=500,
    allow_unsafe_jscode=True,
    fit_columns_on_grid_load=True
)


# =========================
# 編集されたデータを取得
# =========================

edited = pd.DataFrame(grid_return["data"])


# 数値列を数値に変換

for column in [
    "距離",
    "累計距離",
    "BS",
    "IH",
    "TP",
    "IP",
    "GH"
]:

    edited[column] = pd.to_numeric(
        edited[column],
        errors="coerce"
    )


# =========================
# データを保存
# =========================

st.session_state.data = edited


# =========================
# Excel保存
# =========================

if st.button("Excelに保存"):

    output = BytesIO()

    wb = Workbook()

    ws = wb.active

    ws.title = "測量野帳"


    # 見出し

    for col_num, column_name in enumerate(
        edited.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )


    # データ

    for row_num, row in enumerate(
        edited.itertuples(index=False),
        2
    ):

        for col_num, value in enumerate(
            row,
            1
        ):

            if pd.notna(value):

                ws.cell(
                    row=row_num,
                    column=col_num,
                    value=value
                )


    # 列幅

    for column in ws.columns:

        ws.column_dimensions[
            column[0].column_letter
        ].width = 12


    wb.save(output)

    output.seek(0)


    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="測量野帳.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
