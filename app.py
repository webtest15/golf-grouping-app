import streamlit as st
import pandas as pd
import math
import xlsxwriter
from io import BytesIO

st.title("ゴルフ組み合わせ自動生成アプリ")

# ファイルアップロード
uploaded_file = st.file_uploader("プレイヤー情報ファイル（Excel）をアップロード", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df = df.sort_values(by='AverageScore')

    # グループ分け
    num_players = len(df)
    num_groups = math.ceil(num_players / 4)
    groups = [[] for _ in range(num_groups)]

    for i, (_, row) in enumerate(df.iterrows()):
        group_index = i % num_groups
        groups[group_index].append((row['Name'], row['AverageScore']))

    # IN/OUT割り当て + 時間
    start_time_in = 7 * 60 + 30
    start_time_out = 7 * 60 + 30
    interval = 7

    group_info = []
    for i in range(num_groups):
        in_out = "IN" if i % 2 == 0 else "OUT"
        if in_out == "IN":
            time = f"{start_time_in // 60}:{start_time_in % 60:02d}"
            start_time_in += interval
        else:
            time = f"{start_time_out // 60}:{start_time_out % 60:02d}"
            start_time_out += interval
        group_info.append((f"{i+1}組{in_out}\n{time}", in_out))

    # 結果表示
    st.subheader("組み合わせ結果")
    for idx, (label, _) in enumerate(group_info):
        st.markdown(f"### {label.replace(chr(10), ' ')}")
        group_df = pd.DataFrame(groups[idx], columns=["名前", "平均スコア"])
        st.table(group_df)

    # Excel出力
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    # 書式設定
    in_format = workbook.add_format({
        'border': 1, 'bg_color': '#00FFFF', 'align': 'center',
        'valign': 'vcenter', 'font_name': 'メイリオ', 'font_size': 14, 'text_wrap': True
    })
    out_format = workbook.add_format({
        'border': 1, 'bg_color': '#FFFF00', 'align': 'center',
        'valign': 'vcenter', 'font_name': 'メイリオ', 'font_size': 14, 'text_wrap': True
    })
    player_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'メイリオ', 'font_size': 14
    })

    worksheet.set_column('A:F', 18)

    for row_idx, (label, in_out) in enumerate(group_info):
        fmt = in_format if in_out == "IN" else out_format
        worksheet.set_row(row_idx, 60)
        worksheet.write(row_idx, 0, label, fmt)

        for col_idx, (player, score) in enumerate(groups[row_idx]):
            worksheet.write(row_idx, col_idx + 1, player, player_format)
            worksheet.write(row_idx, col_idx + 1 + 4, score, player_format)  # F列にスコア

    workbook.close()
    output.seek(0)

    st.download_button(
        label="Excelファイルをダウンロード",
        data=output,
        file_name="grouping_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
