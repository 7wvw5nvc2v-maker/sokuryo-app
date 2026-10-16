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

display_result = display_result.rename(columns={
    "BS": "後視（BS）",
    "IH": "器械高（IH）",
    "TP": "もりかえ点（TP）",
    "IP": "中間点（IP）",
    "GH": "地盤高（GH）",
})

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
