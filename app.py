import streamlit as st
import pandas as pd

from io import BytesIO
from openpyxl import Workbook

from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder
from st_aggrid import JsCode


# =========================
# ページ設定
# =========================

st.set_page_config(
    page_title="測量自動計算",
    layout="wide"
)


# =========================
# 全体デザイン
# =========================

st.markdown("""
<style>

.stApp {
    background-color: #f1f5f9;
}

h1 {
    color: #1e3a5f;
}

h2, h3 {
    color: #1e3a5f;
}


/* 説明ボックス */

.info-box {
    background-color: #ffffff;
    border-radius: 10px;
    padding: 18px 20px;
    margin-bottom: 15px;
    border: 1px solid #d6dee8;
}


/* 入力の説明 */

.input-box {
    background-color: #d9eef7;
    border-radius: 8px;
    padding: 10px 15px;
    margin-bottom: 8px;
}


/* 自動計算の説明 */

.calc-box {
    background-color: #fff4cc;
    border-radius: 8px;
    padding: 10px 15px;
    margin-bottom: 8px;
}


/* 基準GH */

.gh-box {
    background-color: #ffffff;
    border-radius: 10px;
    padding: 12px 15px;
    margin-top: 10px;
    margin-bottom: 15px;
    border: 1px solid #d6dee8;
}


/* ボタン */

.stButton > button {
    background-color: #2f6690;
    color: white;
    border-radius: 6px;
    border: none;
    padding: 8px 20px;
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
# 使い方
# =========================

st.markdown("""
<div class="info-box">

<h3>📖 使い方</h3>

<p>
このアプリでは、水準測量の計算を自動で行います。
</p>

<p>
<b>①</b> 基準GHを入力します。<br>
<b>②</b> 水色のセルに測定した数値を入力します。<br>
<b>③</b> 黄色のセルは自動で計算されます。<br>
<b>④</b> 計算結果をExcelファイルとして保存できます。
</p>

</div>
""", unsafe_allow_html=True)


# =========================
# 色の凡例
# =========================

st.markdown("""
<div class="info-box">

<h3>🎨 入力・計算項目</h3>

<div class="input-box">
🟦 <b>水色：入力する項目</b><br>
BS・TP・IPなど、測量して得られた数値を入力します。
</div>

<div class="calc-box">
🟨 <b>黄色：自動計算される項目</b><br>
IH・GHなど、入力した数値をもとに自動で計算されます。
</div>

</div>
""", unsafe_allow_html=True)


# =========================
# 基準GH
# =========================

st.markdown("""
<div class="gh-box">

<b>基準GHについて</b><br>
測量開始地点の既知の地盤高を入力してください。

</div>
""", unsafe_allow_html=True)


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


data = st.session_state.data.copy()


# =========================
# 累計距離を計算
# =========================

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
# 計算結果を反映
# =========================

data["累計距離"] = cumulative
data["IH"] = ih_values
data["GH"] = gh_values


# =========================
# セルの色
# =========================

# 入力セル → 水色

input_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#d9eef7'
    };
}
""")


# 自動計算セル → 薄い黄色

calculation_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#fff4cc'
    };
}
""")


# 累計距離 → 薄いグレー

distance_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#eeeeee'
    };
}
""")


# =========================
# 表の設定
# =========================

gb = GridOptionsBuilder.from_dataframe(data)


gb.configure_default_column(
    editable=True,
    resizable=True,
    sortable=False
)


# =========================
# 測点
# =========================

gb.configure_column(
    "測点",
    headerName="測点【入力】",
    editable=True,
    width=140
)


# =========================
# 距離
# =========================

gb.configure_column(
    "距離",
    headerName="距離【入力】",
    editable=True,
    width=140
)


# =========================
# 累計距離
# =========================

gb.configure_column(
    "累計距離",
    headerName="累計距離【自動計算】",
    editable=False,
    width=190,
    cellStyle=distance_style
)


# =========================
# BS
# =========================

gb.configure_column(
    "BS",
    headerName="BS【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# =========================
# IH
# =========================

gb.configure_column(
    "IH",
    headerName="IH【自動計算】",
    editable=False,
    width=190,
    cellStyle=calculation_style
)


# =========================
# TP
# =========================

gb.configure_column(
    "TP",
    headerName="TP【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# =========================
# IP
# =========================

gb.configure_column(
    "IP",
    headerName="IP【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# =========================
# GH
# =========================

gb.configure_column(
    "GH",
    headerName="GH【自動計算】",
    editable=False,
    width=190,
    cellStyle=calculation_style
)


# =========================
# 表の設定
# =========================

gb.configure_grid_options(
    stopEditingWhenCellsLoseFocus=True
)


grid_options = gb.build()


# =========================
# 測量野帳
# =========================

st.markdown("### 📋 測量野帳")

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

edited = pd.DataFrame(
    grid_return["data"]
)


# =========================
# 数値列を数値に変換
# =========================

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

st.markdown("### 💾 データ保存")

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
        ].width = 15


    # Excel作成

    wb.save(output)

    output.seek(0)


    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="測量野帳.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
