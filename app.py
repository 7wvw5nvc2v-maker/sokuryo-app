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
⬜ 灰色：累計距離（自動計算）
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
# 列
# =========================================================

input_columns = [
    "測点",
    "距離",
    "BS",
    "TP",
    "IP"
]

number_columns = [
    "距離",
    "BS",
    "TP",
    "IP"
]


# =========================================================
# 入力データを整える
# =========================================================

def normalize_data(df):

    df = df.copy()

    # 必要な列を作成
    for col in input_columns:

        if col not in df.columns:
            df[col] = None


    # 行ID
    if "_row_id" not in df.columns:

        df["_row_id"] = range(len(df))


    # 行IDを整数にする
    df["_row_id"] = pd.to_numeric(
        df["_row_id"],
        errors="coerce"
    )


    # IDが空の場合
    if df["_row_id"].isna().any():

        missing = df["_row_id"].isna()

        existing = df["_row_id"].dropna()

        if len(existing) > 0:
            next_id = int(existing.max()) + 1
        else:
            next_id = 0

        for index in df.index[missing]:

            df.loc[index, "_row_id"] = next_id
            next_id += 1


    df["_row_id"] = df["_row_id"].astype(int)


    # 測点
    df["測点"] = (
        df["測点"]
        .fillna("")
        .astype(str)
    )


    # 数値項目
    for col in number_columns:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        )


    return df[
        ["_row_id"] + input_columns
    ].reset_index(drop=True)


# =========================================================
# 初期データ
# =========================================================

if "data" not in st.session_state:

    st.session_state.data = pd.DataFrame({

        "_row_id": range(10),

        "測点": [""] * 10,

        "距離": [None] * 10,

        "BS": [None] * 10,

        "TP": [None] * 10,

        "IP": [None] * 10

    })


# =========================================================
# データを整理
# =========================================================

st.session_state.data = normalize_data(
    st.session_state.data
)


# =========================================================
# 行追加
# =========================================================

if st.button("＋ 行を追加"):

    data = normalize_data(
        st.session_state.data
    )

    if len(data) > 0:

        new_id = int(
            data["_row_id"].max()
        ) + 1

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
        [
            data,
            new_row
        ],
        ignore_index=True
    )


    st.rerun()


# =========================================================
# 計算関数
# =========================================================

def calculate_data(input_data, base_gh):

    data = normalize_data(input_data)


    # -----------------------------------------------------
    # 累計距離
    # -----------------------------------------------------

    cumulative = []

    total_distance = 0.0


    for distance in data["距離"]:

        if pd.notna(distance):

            total_distance += float(distance)

            cumulative.append(
                total_distance
            )

        else:

            cumulative.append(None)


    # -----------------------------------------------------
    # IH・GH
    # -----------------------------------------------------

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

        # IH
        if pd.notna(bs):

            current_ih = (
                current_gh + float(bs)
            )

            ih_values.append(
                current_ih
            )

        else:

            ih_values.append(None)


        # GH
        if (
            current_ih is not None
            and pd.notna(tp)
        ):

            current_gh = (
                current_ih - float(tp)
            )

            gh_values.append(
                current_gh
            )

        elif (
            current_ih is not None
            and pd.notna(ip)
        ):

            current_gh = (
                current_ih - float(ip)
            )

            gh_values.append(
                current_gh
            )

        elif i == 0:

            gh_values.append(
                current_gh
            )

        else:

            gh_values.append(None)


    # -----------------------------------------------------
    # 結果
    # -----------------------------------------------------

    result = data.copy()

    result["累計距離"] = cumulative

    result["IH"] = ih_values

    result["GH"] = gh_values


    return result[
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
# 現在の計算結果
# =========================================================

result = calculate_data(
    st.session_state.data,
    base_gh
)


# =========================================================
# AgGrid
# =========================================================

gb = GridOptionsBuilder.from_dataframe(
    result
)


# ---------------------------------------------------------
# 選択
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
# 色
# =========================================================

input_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#d9eef7'
    };
}
""")


auto_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#fff4cc'
    };
}
""")


gray_style = JsCode("""
function(params) {
    return {
        'backgroundColor': '#eeeeee'
    };
}
""")


for column in [
    "距離",
    "BS",
    "TP",
    "IP"
]:

    gb.configure_column(
        column,
        cellStyle=input_style
    )


for column in [
    "IH",
    "GH"
]:

    gb.configure_column(
        column,
        cellStyle=auto_style
    )


gb.configure_column(
    "累計距離",
    cellStyle=gray_style
)


# =========================================================
# 編集イベント
# =========================================================

gb.configure_grid_options(
    onCellValueChanged=JsCode("""
    function(params) {

        params.api.refreshCells({
            force: true
        });

    }
    """)
)


grid_options = gb.build()


# =========================================================
# AgGrid表示
# =========================================================

grid_return = AgGrid(

    result,

    gridOptions=grid_options,

    key="survey_grid",

    height=500,

    fit_columns_on_grid_load=True,

    allow_unsafe_jscode=True,

    update_on=["cellValueChanged"]
)


# =========================================================
# AgGridからデータ取得
# =========================================================

returned_data = grid_return.get("data")


if returned_data is not None:

    returned_df = pd.DataFrame(
        returned_data
    )


    returned_inputs = normalize_data(
        returned_df
    )


    # -----------------------------------------------------
    # 入力値が変更された場合だけ保存
    # -----------------------------------------------------

    current_data = normalize_data(
        st.session_state.data
    )


    current_compare = current_data[
        [
            "_row_id",
            "測点",
            "距離",
            "BS",
            "TP",
            "IP"
        ]
    ].copy()


    returned_compare = returned_inputs[
        [
            "_row_id",
            "測点",
            "距離",
            "BS",
            "TP",
            "IP"
        ]
    ].copy()


    current_compare = (
        current_compare
        .fillna("")
        .reset_index(drop=True)
    )


    returned_compare = (
        returned_compare
        .fillna("")
        .reset_index(drop=True)
    )


    if not current_compare.equals(
        returned_compare
    ):

        st.session_state.data = (
            returned_inputs
        )

        st.rerun()


# =========================================================
# 選択行取得
# =========================================================

selected_rows = grid_return.get(
    "selected_rows"
)


selected_ids = []


# ---------------------------------------------------------
# リストの場合
# ---------------------------------------------------------

if isinstance(
    selected_rows,
    list
):

    for row in selected_rows:

        if (
            isinstance(row, dict)
            and "_row_id" in row
        ):

            try:

                selected_ids.append(
                    int(row["_row_id"])
                )

            except (
                ValueError,
                TypeError
            ):

                pass


# ---------------------------------------------------------
# DataFrameの場合
# ---------------------------------------------------------

elif isinstance(
    selected_rows,
    pd.DataFrame
):

    if "_row_id" in selected_rows.columns:

        for row_id in selected_rows[
            "_row_id"
        ].tolist():

            try:

                selected_ids.append(
                    int(row_id)
                )

            except (
                ValueError,
                TypeError
            ):

                pass


# =========================================================
# 選択行削除
# =========================================================

if selected_ids:

    if st.button(
        "− 選択行を削除"
    ):

        current_data = normalize_data(
            st.session_state.data
        )


        new_data = current_data[
            ~current_data[
                "_row_id"
            ].isin(selected_ids)
        ].reset_index(
            drop=True
        )


        st.session_state.data = new_data

        st.rerun()


# =========================================================
# Excel保存
# =========================================================

st.markdown("---")


if st.button("Excelに保存"):

    # 最新データを計算
    excel_result = calculate_data(
        st.session_state.data,
        base_gh
    )


    # _row_idはExcelに出さない
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


    # -----------------------------------------------------
    # Excel作成
    # -----------------------------------------------------

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
        excel_result.itertuples(
            index=False
        ),
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


    # -----------------------------------------------------
    # 列幅
    # -----------------------------------------------------

    for column in ws.columns:

        ws.column_dimensions[
            column[0].column_letter
        ].width = 12


    # -----------------------------------------------------
    # Excel保存
    # -----------------------------------------------------

    wb.save(output)

    output.seek(0)


    # -----------------------------------------------------
    # ダウンロード
    # -----------------------------------------------------

    st.download_button(

        label="Excelファイルをダウンロード",

        data=output,

        file_name="水準測量_器高式計算.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
