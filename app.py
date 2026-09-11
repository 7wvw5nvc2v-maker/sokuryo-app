import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode


# =========================================================
# ページ設定
# =========================================================

st.set_page_config(
    page_title="水準測量・器高式計算アプリ",
    layout="wide"
)


# =========================================================
# デザイン
# =========================================================

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
    border-radius: 6px;
    border: none;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# タイトル
# =========================================================

st.title("水準測量・器高式計算アプリ")

st.write(
    "水準測量で使用する器高式の計算を自動化するアプリです。"
)


# =========================================================
# 使い方
# =========================================================

with st.expander("使い方", expanded=True):

    st.write("""
① 基準GHを入力します。

② 水色のセルに測定値を入力します。

③ 黄色のセルは器高式によって自動計算されます。

④ 行が足りない場合は「＋ 行を追加」を押します。

⑤ 不要な行はチェックして「− 選択行を削除」を押します。

⑥ 必要に応じて「Excelに保存」からExcelファイルを保存できます。
""")


# =========================================================
# 色の説明
# =========================================================

st.markdown("""
**セルの色**

🟦 水色：入力する項目  
🟨 黄色：自動計算される項目  
⬜ 灰色：自動計算（累計距離）
""")


# =========================================================
# 基準GH
# =========================================================

base_gh = st.number_input(
    "基準GH",
    value=500.000,
    step=0.001,
    format="%.3f"
)


# =========================================================
# セッション状態の初期化
# =========================================================

editable_columns = [
    "測点",
    "距離",
    "BS",
    "TP",
    "IP"
]

numeric_columns = [
    "距離",
    "BS",
    "TP",
    "IP"
]


if "data" not in st.session_state:

    st.session_state.data = pd.DataFrame({
        "_row_id": range(10),
        "測点": [""] * 10,
        "距離": [None] * 10,
        "BS": [None] * 10,
        "TP": [None] * 10,
        "IP": [None] * 10
    })


if "grid_version" not in st.session_state:
    st.session_state.grid_version = 0


# =========================================================
# データを整える関数
# =========================================================

def normalize_input_data(df):

    result = df.copy()

    # 必要な列がなければ追加
    for col in editable_columns:
        if col not in result.columns:
            result[col] = None

    if "_row_id" not in result.columns:
        result["_row_id"] = range(len(result))

    # 測点
    result["測点"] = (
        result["測点"]
        .fillna("")
        .astype(str)
    )

    # 数値項目
    for col in numeric_columns:
        result[col] = pd.to_numeric(
            result[col],
            errors="coerce"
        )

    return result[
        ["_row_id"] + editable_columns
    ].reset_index(drop=True)


# =========================================================
# データ比較用
# =========================================================

def data_signature(df):

    normalized = normalize_input_data(df)

    return normalized.to_json(
        orient="records",
        force_ascii=False
    )


# =========================================================
# 行追加
# =========================================================

if st.button("＋ 行を追加"):

    data = normalize_input_data(
        st.session_state.data
    )

    if len(data) > 0:
        new_id = int(data["_row_id"].max()) + 1
    else:
        new_id = 0

    new_row = pd.DataFrame({
        "_row_id": [new_id],
        "測点": [""],
        "距離": [None],
        "BS": [None],
        "TP": [None],
        "IP": [None]
    })

    st.session_state.data = pd.concat(
        [data, new_row],
        ignore_index=True
    )

    # AgGridを新しく読み直す
    st.session_state.grid_version += 1

    st.rerun()


# =========================================================
# 現在の入力データ
# =========================================================

input_data = normalize_input_data(
    st.session_state.data
)


# =========================================================
# 表示用データの作成
# =========================================================

cumulative = []

total_distance = 0.0

for distance in input_data["距離"]:

    if pd.notna(distance):

        total_distance += float(distance)
        cumulative.append(total_distance)

    else:

        cumulative.append(None)


# =========================================================
# IH・GHの計算
# =========================================================

ih_values = []
gh_values = []

current_gh = float(base_gh)
current_ih = None


for i, (bs, tp, ip) in enumerate(
    zip(
        input_data["BS"],
        input_data["TP"],
        input_data["IP"]
    )
):

    # ---------------------------------------------
    # IH
    # ---------------------------------------------

    if pd.notna(bs):

        current_ih = current_gh + float(bs)

        ih_values.append(current_ih)

    else:

        ih_values.append(None)


    # ---------------------------------------------
    # GH
    # ---------------------------------------------

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


# =========================================================
# 表示用テーブル
# =========================================================

result = input_data.copy()

result["累計距離"] = cumulative
result["IH"] = ih_values
result["GH"] = gh_values


# 表示順をExcelのようにする
result = result[
    [
        "_row_id",
        "測点",
        "距離",
        "累計距離",
        "BS",
        "IH",
        "TP",
        "IP",
        "GH"
    ]
]


# =========================================================
# AgGrid設定
# =========================================================

gb = GridOptionsBuilder.from_dataframe(result)


# ---------------------------------------------------------
# 行選択
# ---------------------------------------------------------

gb.configure_selection(
    selection_mode="multiple",
    use_checkbox=True
)


# ---------------------------------------------------------
# 行IDを非表示
# ---------------------------------------------------------

gb.configure_column(
    "_row_id",
    hide=True
)


# ---------------------------------------------------------
# 入力列
# ---------------------------------------------------------

gb.configure_column(
    "測点",
    header_name="測点",
    editable=True
)

gb.configure_column(
    "距離",
    header_name="距離【入力】",
    editable=True,
    type=["numericColumn"]
)

gb.configure_column(
    "BS",
    header_name="BS【入力】",
    editable=True,
    type=["numericColumn"]
)

gb.configure_column(
    "TP",
    header_name="TP【入力】",
    editable=True,
    type=["numericColumn"]
)

gb.configure_column(
    "IP",
    header_name="IP【入力】",
    editable=True,
    type=["numericColumn"]
)


# ---------------------------------------------------------
# 自動計算列
# ---------------------------------------------------------

gb.configure_column(
    "累計距離",
    header_name="累計距離【自動計算】",
    editable=False,
    type=["numericColumn"]
)

gb.configure_column(
    "IH",
    header_name="IH【自動計算】",
    editable=False,
    type=["numericColumn"]
)

gb.configure_column(
    "GH",
    header_name="GH【自動計算】",
    editable=False,
    type=["numericColumn"]
)


# =========================================================
# セルの色
# =========================================================

input_cell_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#d9eef7'
    };
}
""")


auto_cell_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#fff4cc'
    };
}
""")


distance_cell_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#eeeeee'
    };
}
""")


# 入力セル
for column in [
    "距離",
    "BS",
    "TP",
    "IP"
]:

    gb.configure_column(
        column,
        cellStyle=input_cell_style
    )


# 自動計算セル
for column in [
    "IH",
    "GH"
]:

    gb.configure_column(
        column,
        cellStyle=auto_cell_style
    )


# 累計距離
gb.configure_column(
    "累計距離",
    cellStyle=distance_cell_style
)


grid_options = gb.build()


# =========================================================
# AgGrid表示
# =========================================================

grid_return = AgGrid(
    result,
    gridOptions=grid_options,

    # ★重要
    # 行追加・削除時だけGridを新しくする
    key=f"survey_grid_{st.session_state.grid_version}",

    height=500,

    fit_columns_on_grid_load=True,

    allow_unsafe_jscode=True
)


# =========================================================
# AgGridから返ってきたデータ
# =========================================================

returned_data = grid_return.get("data")


if returned_data is not None:

    returned_df = pd.DataFrame(returned_data)

    returned_inputs = normalize_input_data(
        returned_df
    )


    # =====================================================
    # 選択された行
    # =====================================================

    selected_rows = grid_return.get(
        "selected_rows",
        []
    )


    selected_ids = []

    for row in selected_rows:

        if "_row_id" in row:

            try:
                selected_ids.append(
                    int(row["_row_id"])
                )
            except:
                pass


    # =====================================================
    # 行削除
    # =====================================================

    if selected_ids:

        if st.button("− 選択行を削除"):

            # AgGridから返ってきた最新の入力内容を使う
            new_data = returned_inputs[
                ~returned_inputs["_row_id"].isin(
                    selected_ids
                )
            ].reset_index(drop=True)

            st.session_state.data = new_data

            # Gridを新しくする
            st.session_state.grid_version += 1

            st.rerun()


    # =====================================================
    # ★重要
    # 入力内容が変わったら即座に保存
    # =====================================================

    old_signature = data_signature(
        st.session_state.data
    )

    new_signature = data_signature(
        returned_inputs
    )


    if old_signature != new_signature:

        # ★計算結果ではなく
        # 入力データだけを保存する
        st.session_state.data = returned_inputs

        # 保存後にもう一度画面を描画
        st.rerun()


# =========================================================
# Excel保存
# =========================================================

st.markdown("---")

if st.button("Excelに保存"):

    # 最新の入力データから再計算
    input_data = normalize_input_data(
        st.session_state.data
    )

    cumulative = []

    total_distance = 0.0

    for distance in input_data["距離"]:

        if pd.notna(distance):

            total_distance += float(distance)
            cumulative.append(total_distance)

        else:

            cumulative.append(None)


    ih_values = []
    gh_values = []

    current_gh = float(base_gh)
    current_ih = None


    for i, (bs, tp, ip) in enumerate(
        zip(
            input_data["BS"],
            input_data["TP"],
            input_data["IP"]
        )
    ):

        if pd.notna(bs):

            current_ih = current_gh + float(bs)
            ih_values.append(current_ih)

        else:

            ih_values.append(None)


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


    excel_result = input_data.copy()

    excel_result["累計距離"] = cumulative
    excel_result["IH"] = ih_values
    excel_result["GH"] = gh_values


    # _row_idはExcelには出さない
    excel_result = excel_result[
        [
            "測点",
            "距離",
            "累計距離",
            "BS",
            "IH",
            "TP",
            "IP",
            "GH"
        ]
    ]


    # =====================================================
    # Excel作成
    # =====================================================

    output = BytesIO()

    wb = Workbook()

    ws = wb.active

    ws.title = "測量野帳"


    # ヘッダー
    for col_num, column_name in enumerate(
        excel_result.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )


    # データ
    for row_num, row in enumerate(
        excel_result.itertuples(index=False),
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


    # =====================================================
    # ダウンロード
    # =====================================================

    st.download_button(
        label="Excelファイルをダウンロード",

        data=output,

        file_name="水準測量_器高式計算.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
