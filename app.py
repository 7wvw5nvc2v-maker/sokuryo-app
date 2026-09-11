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

⑤ 不要な行は選択して「− 選択行を削除」を押します。

⑥ 必要に応じてExcelファイルとして保存できます。
""")


st.markdown(
    "🟦 **入力**　　🟨 **自動計算**　　⬜ **累計距離**"
)


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
# データ初期化
# =========================================================

columns = [
    "測点",
    "距離",
    "累計距離",
    "BS",
    "IH",
    "TP",
    "IP",
    "GH",
]


if "survey_data" not in st.session_state:

    st.session_state.survey_data = pd.DataFrame({
        "測点": ["" for _ in range(10)],
        "距離": [None for _ in range(10)],
        "累計距離": [None for _ in range(10)],
        "BS": [None for _ in range(10)],
        "IH": [None for _ in range(10)],
        "TP": [None for _ in range(10)],
        "IP": [None for _ in range(10)],
        "GH": [None for _ in range(10)],
    })


# =========================================================
# 入力データだけ取り出す
# =========================================================

def get_input_data(df):

    result = pd.DataFrame()

    result["測点"] = df["測点"].fillna("").astype(str)

    for col in ["距離", "BS", "TP", "IP"]:

        result[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        )

    return result


# =========================================================
# 計算
# =========================================================

def calculate_data(input_df, base_gh):

    result = input_df.copy()

    # -----------------------------
    # 累計距離
    # -----------------------------

    cumulative = []

    total = 0.0

    for value in result["距離"]:

        if pd.notna(value):

            total += float(value)

            cumulative.append(total)

        else:

            cumulative.append(None)

    result["累計距離"] = cumulative


    # -----------------------------
    # IH・GH
    # -----------------------------

    ih_values = []
    gh_values = []

    current_gh = float(base_gh)
    current_ih = None

    for i in range(len(result)):

        bs = result.loc[i, "BS"]
        tp = result.loc[i, "TP"]
        ip = result.loc[i, "IP"]


        # BSが入力されたらIHを計算
        if pd.notna(bs):

            current_ih = current_gh + float(bs)

            ih_values.append(current_ih)

        else:

            ih_values.append(None)


        # TP
        if current_ih is not None and pd.notna(tp):

            current_gh = current_ih - float(tp)

            gh_values.append(current_gh)


        # IP
        elif current_ih is not None and pd.notna(ip):

            current_gh = current_ih - float(ip)

            gh_values.append(current_gh)


        # 最初のGH
        elif i == 0:

            gh_values.append(current_gh)


        else:

            gh_values.append(None)


    result["IH"] = ih_values
    result["GH"] = gh_values


    # 列順
    result = result[
        [
            "測点",
            "距離",
            "累計距離",
            "BS",
            "IH",
            "TP",
            "IP",
            "GH",
        ]
    ]

    return result


# =========================================================
# 現在の入力データ
# =========================================================

input_data = get_input_data(
    st.session_state.survey_data
)


# =========================================================
# AgGrid
# =========================================================

gb = GridOptionsBuilder.from_dataframe(
    input_data
)


# -----------------------------
# 列設定
# -----------------------------

gb.configure_column(
    "測点",
    headerName="測点",
    editable=True,
    width=120,
)


gb.configure_column(
    "距離",
    headerName="距離【入力】",
    editable=True,
    type=["numericColumn"],
    width=130,
)


gb.configure_column(
    "累計距離",
    headerName="累計距離【自動計算】",
    editable=False,
    width=150,
)


gb.configure_column(
    "BS",
    headerName="BS【入力】",
    editable=True,
    type=["numericColumn"],
    width=130,
)


gb.configure_column(
    "IH",
    headerName="IH【自動計算】",
    editable=False,
    width=140,
)


gb.configure_column(
    "TP",
    headerName="TP【入力】",
    editable=True,
    type=["numericColumn"],
    width=130,
)


gb.configure_column(
    "IP",
    headerName="IP【入力】",
    editable=True,
    type=["numericColumn"],
    width=130,
)


gb.configure_column(
    "GH",
    headerName="GH【自動計算】",
    editable=False,
    width=140,
)


# =========================================================
# セルの色
# =========================================================

input_style = JsCode("""
function(params) {
    return {
        'background-color': '#d9eef7'
    };
}
""")


auto_style = JsCode("""
function(params) {
    return {
        'background-color': '#fff4cc'
    };
}
""")


gray_style = JsCode("""
function(params) {
    return {
        'background-color': '#eeeeee'
    };
}
""")


gb.configure_column(
    "距離",
    cellStyle=input_style
)

gb.configure_column(
    "BS",
    cellStyle=input_style
)

gb.configure_column(
    "TP",
    cellStyle=input_style
)

gb.configure_column(
    "IP",
    cellStyle=input_style
)

gb.configure_column(
    "累計距離",
    cellStyle=gray_style
)

gb.configure_column(
    "IH",
    cellStyle=auto_style
)

gb.configure_column(
    "GH",
    cellStyle=auto_style
)


# -----------------------------
# 行選択
# -----------------------------

gb.configure_selection(
    selection_mode="multiple",
    use_checkbox=True
)


# -----------------------------
# 編集したら値を返す
# -----------------------------

gb.configure_grid_options(
    stopEditingWhenCellsLoseFocus=True
)


grid_options = gb.build()


# =========================================================
# AgGrid表示
# =========================================================

grid_return = AgGrid(
    input_data,
    gridOptions=grid_options,
    height=500,
    width="100%",
    data_return_mode=DataReturnMode.AS_INPUT,
    update_mode=GridUpdateMode.VALUE_CHANGED,
    allow_unsafe_jscode=True,
    fit_columns_on_grid_load=False,
    reload_data=False,
    key="survey_grid",
)


# =========================================================
# AgGridから入力値を取得
# =========================================================

returned_data = grid_return.get("data")


if returned_data is not None:

    if isinstance(returned_data, pd.DataFrame):

        edited_input = get_input_data(
            returned_data
        )

    else:

        edited_input = pd.DataFrame(
            returned_data
        )

        edited_input = get_input_data(
            edited_input
        )


    # 現在の入力値と比較
    old_input = get_input_data(
        st.session_state.survey_data
    )


    if not edited_input.equals(old_input):

        st.session_state.survey_data = edited_input

        st.rerun()


# =========================================================
# 計算結果
# =========================================================

result = calculate_data(
    get_input_data(
        st.session_state.survey_data
    ),
    base_gh
)


# =========================================================
# 行追加・削除
# =========================================================

col1, col2 = st.columns(2)


with col1:

    if st.button("＋ 行を追加", use_container_width=True):

        new_row = pd.DataFrame({
            "測点": [""],
            "距離": [None],
            "BS": [None],
            "TP": [None],
            "IP": [None],
        })

        current = get_input_data(
            st.session_state.survey_data
        )

        st.session_state.survey_data = pd.concat(
            [current, new_row],
            ignore_index=True
        )

        st.rerun()


with col2:

    if st.button(
        "− 選択行を削除",
        use_container_width=True
    ):

        selected_rows = grid_return.get(
            "selected_rows",
            []
        )

        if isinstance(
            selected_rows,
            pd.DataFrame
        ):

            selected_rows = selected_rows.to_dict(
                "records"
            )

        if not isinstance(
            selected_rows,
            list
        ):

            selected_rows = []


        if len(selected_rows) > 0:

            current = get_input_data(
                st.session_state.survey_data
            )


            indexes = []

            for row in selected_rows:

                if isinstance(row, dict):

                    if "測点" in row:

                        # 行の位置を特定
                        for i in range(
                            len(current)
                        ):

                            if (
                                current.loc[i, "測点"]
                                == str(row["測点"])
                            ):

                                indexes.append(i)

                                break


            if indexes:

                current = current.drop(
                    indexes
                ).reset_index(drop=True)

                st.session_state.survey_data = current

                st.rerun()


# =========================================================
# 計算結果表示
# =========================================================

st.subheader("計算結果")


display_result = result.copy()


for col in [
    "距離",
    "累計距離",
    "BS",
    "IH",
    "TP",
    "IP",
    "GH",
]:

    display_result[col] = display_result[col].apply(
        lambda x: (
            "" if pd.isna(x)
            else f"{float(x):.3f}"
        )
    )


st.dataframe(
    display_result,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# Excel保存
# =========================================================

st.subheader("Excel保存")


if st.button(
    "Excelに保存",
    use_container_width=True
):

    output = BytesIO()

    wb = Workbook()

    ws = wb.active

    ws.title = "測量野帳"


    # ヘッダー
    for col_num, column_name in enumerate(
        result.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )


    # データ
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


    # 列幅
    for column in ws.columns:

        ws.column_dimensions[
            column[0].column_letter
        ].width = 14


    wb.save(output)

    output.seek(0)


    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="水準測量_器高式計算.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
