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
    page_title="水準測量・器高式計算アプリ",
    layout="wide"
)


# =========================
# デザイン
# =========================
st.markdown("""
<style>
.stApp {
    background-color: #f1f5f9;
}

h1 {
    color: #1e3a5f;
}

h2 {
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

.info-box {
    background-color: #ffffff;
    padding: 18px;
    border-radius: 10px;
    border-left: 5px solid #2f6690;
    margin-bottom: 15px;
}

.legend-box {
    background-color: #ffffff;
    padding: 15px;
    border-radius: 10px;
    margin-bottom: 15px;
}

.input-color {
    background-color: #d9eef7;
    padding: 5px 12px;
    border-radius: 5px;
}

.calc-color {
    background-color: #fff4cc;
    padding: 5px 12px;
    border-radius: 5px;
}

.other-color {
    background-color: #eeeeee;
    padding: 5px 12px;
    border-radius: 5px;
}
</style>
""", unsafe_allow_html=True)


# =========================
# タイトル
# =========================
st.title("水準測量・器高式計算アプリ")

st.write(
    "水準測量で使用する器高式の計算を自動化するアプリです。"
)


# =========================
# 使い方
# =========================
st.markdown("""
<div class="info-box">

<h3>📖 使い方</h3>

<p>① 基準GHを入力します。</p>
<p>② 水色のセルに測定値を入力します。</p>
<p>③ 黄色のセルは器高式によって自動計算されます。</p>
<p>④ 必要に応じて「Excelに保存」を押してください。</p>

</div>
""", unsafe_allow_html=True)


# =========================
# 色の説明
# =========================
st.markdown("""
<div class="legend-box">

<h3>🎨 セルの色について</h3>

<p>
<span class="input-color">🟦 水色</span>
＝ 使用者が入力する項目
</p>

<p>
<span class="calc-color">🟨 黄色</span>
＝ 自動で計算される項目
</p>

<p>
<span class="other-color">⬜ 灰色</span>
＝ 累計距離などの自動計算項目
</p>

</div>
""", unsafe_allow_html=True)


# =========================
# 基準GH
# =========================
st.subheader("基準GH")

st.caption(
    "測量を開始する地点の既知の地盤高（GH）を入力してください。"
)

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
        "GH": [None] * 10
    })


# =========================
# 計算前データ
# =========================
data = st.session_state.data.copy()


# =========================
# 累計距離の計算
# =========================
cumulative = []

total_distance = 0.0

for distance in data["距離"]:

    if pd.notna(distance):

        total_distance += float(distance)
        cumulative.append(total_distance)

    else:

        cumulative.append(None)


# =========================
# IH・GHの計算
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

    # -------------------------
    # IH
    # -------------------------
    if pd.notna(bs):

        current_ih = current_gh + float(bs)

        ih_values.append(current_ih)

    else:

        ih_values.append(None)


    # -------------------------
    # GH
    # -------------------------
    if current_ih is not None and pd.notna(tp):

        current_gh = current_ih - float(tp)

        gh_values.append(current_gh)

    elif current_ih is not None and pd.notna(ip):

        current_gh = current_ih - float(ip)

        gh_values.append(current_gh)

    elif i == 0:

        gh_values.append(current_gh)

    else:

        gh_values.append(None)


# =========================
# 計算結果
# =========================
result = data.copy()

result["累計距離"] = cumulative
result["IH"] = ih_values
result["GH"] = gh_values


# =========================
# 測量野帳
# =========================
st.subheader("📋 測量野帳")


# =========================
# セルの色
# =========================

# 入力セル（水色）
input_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#d9eef7'
    };
}
""")


# 自動計算セル（黄色）
calculation_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#fff4cc'
    };
}
""")


# 累計距離（灰色）
distance_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#eeeeee'
    };
}
""")


# =========================
# AgGrid設定
# =========================
gb = GridOptionsBuilder.from_dataframe(result)


# 測点
gb.configure_column(
    "測点",
    headerName="測点【入力】",
    editable=True,
    width=140
)


# 距離
gb.configure_column(
    "距離",
    headerName="距離【入力】",
    editable=True,
    width=140
)


# 累計距離
gb.configure_column(
    "累計距離",
    headerName="累計距離【自動計算】",
    editable=False,
    width=180,
    cellStyle=distance_style
)


# BS
gb.configure_column(
    "BS",
    headerName="BS【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# IH
gb.configure_column(
    "IH",
    headerName="IH【自動計算】",
    editable=False,
    width=180,
    cellStyle=calculation_style
)


# TP
gb.configure_column(
    "TP",
    headerName="TP【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# IP
gb.configure_column(
    "IP",
    headerName="IP【入力】",
    editable=True,
    width=140,
    cellStyle=input_style
)


# GH
gb.configure_column(
    "GH",
    headerName="GH【自動計算】",
    editable=False,
    width=180,
    cellStyle=calculation_style
)


# =========================
# Grid表示
# =========================
grid_options = gb.build()

grid_return = AgGrid(
    result,
    gridOptions=grid_options,
    height=500,
    allow_unsafe_jscode=True,
    fit_columns_on_grid_load=True
)


# =========================
# 編集後データ
# =========================
edited = pd.DataFrame(
    grid_return["data"]
)


# 数値列を数値化
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


# セッションに保存
st.session_state.data = edited


# =========================
# Excel保存
# =========================
st.subheader("💾 データ保存")

if st.button("Excelに保存"):

    output = BytesIO()

    wb = Workbook()

    ws = wb.active

    ws.title = "測量野帳"


    # -------------------------
    # ヘッダー
    # -------------------------
    for col_num, column_name in enumerate(
        result.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )


    # -------------------------
    # データ
    # -------------------------
    for row_num, row in enumerate(
        result.itertuples(index=False),
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


    # -------------------------
    # 列幅
    # -------------------------
    for column in ws.columns:

        ws.column_dimensions[
            column[0].column_letter
        ].width = 12


    # -------------------------
    # Excel作成
    # -------------------------
    wb.save(output)

    output.seek(0)


    # -------------------------
    # ダウンロード
    # -------------------------
    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="水準測量_器高式計算.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
