import streamlit as st
import pandas as pd
import io

# ページ設定
st.set_page_config(page_title="Factor Score Calculator")

# アプリタイトルと作成者
st.title("Factor Score Calculator")
st.title("因子得点算出Webアプリケーション")
st.caption("Created by Dit-Lab.(Daiki Ito)")

# ひな形ファイルのパス
template_file_path = '尺度情報.xlsx'

# ファイルアップロードセクション
st.subheader("尺度情報ファイルのアップロード")
uploaded_file_scale_info = st.file_uploader("尺度情報ファイルをアップロードしてください", type=['xlsx'], key="scale_info")
# 尺度情報のひな形ファイルのダウンロード
with open(template_file_path, "rb") as file:
    btn = st.download_button(
        label="尺度情報のフォーマットをダウンロード",
        data=file,
        file_name="尺度情報.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.subheader("因子得点を算出するファイルのアップロード")
uploaded_file_data = st.file_uploader("データファイルをアップロードしてください", type=['xlsx'], key="data")

# マージン
st.write("")
# 何件法の入力
n_point_scale = st.number_input("何件法を使用していますか？", value=4, min_value=3, step=1)


# 因子得点の計算とダウンロードボタンの表示
if uploaded_file_scale_info and uploaded_file_data:
    # ファイルの読み込み
    scale_info = pd.read_excel(uploaded_file_scale_info)
    data = pd.read_excel(uploaded_file_data)

    # ファイル内容の表示
    st.write("尺度情報プレビュー:")
    st.dataframe(scale_info)

    st.write("データプレビュー:")
    st.dataframe(data)

    # 因子得点を計算するボタン
    if st.button("因子得点を計算"):
        # 反転項目の処理
        for index, row in scale_info.iterrows():
            if row['反転'] == 1:
                data[row['設問名']] = n_point_scale - data[row['設問名']]

        # 因子得点の算出
        for factor in scale_info['因子名'].unique():
            relevant_questions = scale_info[scale_info['因子名'] == factor]['設問名']
            data[factor + ' 因子得点'] = data[relevant_questions].mean(axis=1)

        # 不要な列の削除
        data.drop(columns=scale_info['設問名'], inplace=True)

        # 成功メッセージ
        st.success("因子得点の計算が完了しました。")

        # 更新されたデータをExcelファイルとしてダウンロード
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            data.to_excel(writer, index=False, sheet_name='Updated Data')
            writer.save()
        
        # ダウンロードするファイル名の設定
        download_file_name = uploaded_file_data.name.split('.')[0] + "因子得点算出.xlsx"
        st.download_button(
            label="更新されたデータをダウンロード",
            data=output.getvalue(),
            file_name=download_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 著作権表示
st.markdown('© 2022-2023 Dit-Lab.(Daiki Ito). All Rights Reserved.')
