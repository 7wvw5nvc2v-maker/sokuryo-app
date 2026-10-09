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
    "基準GH",
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
            "BS": [None for _ in range(rows)],
            "TP": [None for _ in range(rows)],
            "IP": [None for _ in range(rows)],
        }
    )


if "survey_data" not in st.session_state:
    st.session_state.survey_data = create_empty_data(10)


def get_input_data(df):
    df = df.copy()

    for col in ["測点", "距離", "BS", "TP", "IP"]:
        if col not in df.columns:
            df[col] = "" if col == "測点" else None

    result = pd.DataFrame()
    result["測点"] = df["測点"].fillna("").astype(str)

    for col in ["距離", "BS", "TP", "IP"]:
        result[col] = pd.to_numeric(df[col], errors="coerce")

    return result.reset_index(drop=True)


# =========================================================
# 器高式計算
# =========================================================
def calculate_data(input_df, base_gh):
    result = input_df.copy().reset_index(drop=True)

    cumulative = []
    total_distance = 0.0

    for i in range(len(result)):
        distance = result.iloc[i]["距離"]
        if pd.notna(distance):
            total_distance += float(distance)
            cumulative.append(total_distance)
        else:
            cumulative.append(None)

    result["累計距離"] = cumulative

    ih_values = []
    gh_values = []
    current_gh = float(base_gh)
    current_ih = None

    for i in range(len(result)):
        bs = result.iloc[i]["BS"]
        tp = result.iloc[i]["TP"]
        ip = result.iloc[i]["IP"]

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

    result["IH"] = ih_values
    result["GH"] = gh_values

    result = result[
        ["測点", "距離", "累計距離", "BS", "IH", "TP", "IP", "GH"]
    ]
    return result.reset_index(drop=True)


# =========================================================
# AgGrid設定
# =========================================================
input_data = get_input_data(st.session_state.survey_data)
gb = GridOptionsBuilder.from_dataframe(input_data)

gb.configure_column(
    "測点", headerName="測点", editable=True, width=120
)
gb.configure_column(
    "距離", headerName="距離【入力】", editable=True,
    type=["numericColumn"], width=130
)
gb.configure_column(
    "累計距離", headerName="累計距離【自動計算】",
    editable=False, width=150
)
gb.configure_column(
    "BS", headerName="BS【入力】", editable=True,
    type=["numericColumn"], width=130
)
gb.configure_column(
    "IH", headerName="IH【自動計算】", editable=False, width=140
)
gb.configure_column(
    "TP", headerName="TP【入力】", editable=True,
    type=["numericColumn"], width=130
)
gb.configure_column(
    "IP", headerName="IP【入力】", editable=True,
    type=["numericColumn"], width=130
)
gb.configure_column(
    "GH", headerName="GH【自動計算】", editable=False, width=140
)

# セルの基本色
input_style = JsCode("""
function(params) {
    return {'background-color': '#d9eef7'};
}
""")
auto_style = JsCode("""
function(params) {
    return {'background-color': '#fff4cc'};
}
""")
gray_style = JsCode("""
function(params) {
    return {'background-color': '#eeeeee'};
}
""")

for col in ["測点", "距離", "BS", "TP", "IP"]:
    gb.configure_column(
        col,
        cellClass="input-cell",
        cellStyle=input_style,
    )

for col in ["IH", "GH"]:
    gb.configure_column(
        col,
        cellClass="auto-cell",
        cellStyle=auto_style,
    )

gb.configure_column(
    "累計距離",
    cellClass="gray-cell",
    cellStyle=gray_style,
)

# チェックボックスによる行選択（行削除用）
gb.configure_selection(
    selection_mode="multiple",
    use_checkbox=True,
)

# 行全体のホバー強調を無効化
gb.configure_grid_options(
    stopEditingWhenCellsLoseFocus=True,
    suppressRowHoverHighlight=True,
    rowSelection="multiple",
)

grid_options = gb.build()

# =========================================================
# AgGrid表示
# =========================================================
grid_return = AgGrid(
    input_data,
    gridOptions=grid_options,
    custom_css={
        # 行ホバー／行選択で、行全体に色を付けない
        ".ag-row-hover": {
            "background-color": "transparent !important",
        },
        ".ag-row-hover .ag-cell": {
            "background-color": "inherit !important",
        },
        ".ag-row-selected": {
            "background-color": "transparent !important",
        },
        ".ag-row-selected .ag-cell": {
            "background-color": "inherit !important",
        },

        # 各セルの通常色を維持する
        ".ag-cell.input-cell": {
            "background-color": "#d9eef7 !important",
        },
        ".ag-cell.auto-cell": {
            "background-color": "#fff4cc !important",
        },
        ".ag-cell.gray-cell": {
            "background-color": "#eeeeee !important",
        },

        # マウスが乗っているセルだけ薄灰色にする
        ".ag-cell.input-cell:hover": {
            "background-color": "#e5e7eb !important",
        },
        ".ag-cell.auto-cell:hover": {
            "background-color": "#e5e7eb !important",
        },
        ".ag-cell.gray-cell:hover": {
            "background-color": "#e5e7eb !important",
        },

        # 選択中のセルは枠線のみ表示し、行全体は強調しない
        ".ag-cell-focus": {
            "border": "2px solid #2563eb !important",
            "outline": "none !important",
        },
    },
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
# AgGridからデータ取得
# =========================================================
returned_data = grid_return.get("data")

if returned_data is not None:
    try:
        if isinstance(returned_data, pd.DataFrame):
            edited_input = get_input_data(returned_data)
        else:
            edited_input = get_input_data(pd.DataFrame(returned_data))

        old_input = get_input_data(st.session_state.survey_data)

        if not edited_input.equals(old_input):
            st.session_state.survey_data = edited_input
            st.rerun()
    except Exception:
        pass

# =========================================================
# 計算結果
# =========================================================
result = calculate_data(
    get_input_data(st.session_state.survey_data),
    base_gh,
)

# =========================================================
# 行追加・削除
# =========================================================
col1, col2 = st.columns(2)

with col1:
    if st.button("＋ 行を追加", use_container_width=True):
        current = get_input_data(st.session_state.survey_data)
        new_row = pd.DataFrame(
            {"測点": [""], "距離": [None], "BS": [None], "TP": [None], "IP": [None]}
        )
        st.session_state.survey_data = pd.concat(
            [current, new_row], ignore_index=True
        )
        st.rerun()

with col2:
    if st.button("− 選択行を削除", use_container_width=True):
        selected_rows = grid_return.get("selected_rows", [])

        if isinstance(selected_rows, pd.DataFrame):
            selected_rows = selected_rows.to_dict("records")
        if not isinstance(selected_rows, list):
            selected_rows = []

        if selected_rows:
            current = get_input_data(st.session_state.survey_data)
            selected_indexes = []

            for row in selected_rows:
                if not isinstance(row, dict):
                    continue
                info = row.get("_selectedRowNodeInfo", {})
                node = info.get("node", {}) if isinstance(info, dict) else {}
                if "rowIndex" in node:
                    try:
                        selected_indexes.append(int(node["rowIndex"]))
                    except (TypeError, ValueError):
                        pass

            if not selected_indexes:
                # 測点名が重複している場合に誤削除しないよう、
                # rowIndexが取得できないときは削除を実行しない
                st.warning("選択した行の番号を取得できませんでした。もう一度選択してください。")
            else:
                current = current.drop(index=sorted(set(selected_indexes)))
                st.session_state.survey_data = current.reset_index(drop=True)
                st.rerun()

# =========================================================
# 計算結果表示
# =========================================================
st.subheader("計算結果")
display_result = result.copy()

for col in ["距離", "累計距離", "BS", "IH", "TP", "IP", "GH"]:
    display_result[col] = display_result[col].apply(
        lambda x: "" if pd.isna(x) else f"{float(x):.3f}"
    )

st.dataframe(
    display_result,
    use_container_width=True,
    hide_index=True,
)

# =========================================================
# Excel保存
# =========================================================
st.subheader("Excel保存")

if st.button("Excelに保存", use_container_width=True):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "測量野帳"

    for col_num, column_name in enumerate(result.columns, 1):
        ws.cell(row=1, column=col_num, value=column_name)

    for row_num, row in enumerate(result.itertuples(index=False), 2):
        for col_num, value in enumerate(row, 1):
            if pd.notna(value):
                ws.cell(row=row_num, column=col_num, value=value)

    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = 14

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="水準測量_器高式計算.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
