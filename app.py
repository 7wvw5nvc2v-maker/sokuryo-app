import streamlit as st
import pandas as pd

st.set_page_config(page_title="測量自動計算", layout="wide")

st.title("測量自動計算")
base_gh = st.number_input(
    "基準GH",
    value=500.000,
    step=0.001,
    format="%.3f"
)

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

edited = st.data_editor(
    st.session_state.data,
    width="stretch",
    num_rows="dynamic",
    key="survey_table",
    disabled=["累計距離", "IH", "GH"]
)

# 累計距離を計算
total = 0.0
cumulative = []

for distance in edited["距離"]:
    if pd.isna(distance):
        cumulative.append(None)
    else:
        try:
            total += float(distance)
            cumulative.append(total)
        except (ValueError, TypeError):
            cumulative.append(None)

# IH・GHを計算
ih_values = []
gh_values = []

current_gh = float(base_gh)
current_ih = None

for i, (bs, tp, ip) in enumerate(
    zip(
        edited["BS"],
        edited["TP"],
        edited["IP"]
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


# 計算結果を保存
result = edited.copy()
result["累計距離"] = cumulative
result["IH"] = ih_values
result["GH"] = gh_values


st.session_state.data = result


# Excel保存ボタン
from io import BytesIO
from openpyxl import Workbook

if st.button("Excelに保存"):
    output = BytesIO()

    wb = Workbook()
    ws = wb.active
    ws.title = "測量野帳"

    # 見出し
    for col_num, column_name in enumerate(result.columns, 1):
        ws.cell(row=1, column=col_num, value=column_name)

    # データ
    for row_num, row in enumerate(result.itertuples(index=False), 2):
        for col_num, value in enumerate(row, 1):
            if pd.notna(value):
                ws.cell(row=row_num, column=col_num, value=value)

    # 列幅
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = 12

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="測量野帳.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )