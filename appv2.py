import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

# タイトル
st.title("🕒 勤務時間抽出＆フォーマット修正ツール")

# ファイルアップロード
uploaded_file = st.file_uploader("📂 Excelファイルをアップロード", type=["xlsx"])

# 検索する名前を入力
search_name = st.text_input("🔎 検索する名前を入力:", "宇都宮美香")

# 時間フォーマット修正関数
def format_time(time_str):
    """ 勤務時間のフォーマットを修正 """
    pattern = re.compile(r"(\d{1,2}):?(\d{2})?\s?-\s?(\d{1,2}):?(\d{2})?")
    match = pattern.search(str(time_str))
    
    if not match:
        return None, None, None  # フォーマットが不明なら None を返す
    
    start_hour, start_minute, end_hour, end_minute = match.groups()
    start_time = f"{start_hour}:{start_minute or '00'}"
    end_time = f"{end_hour}:{end_minute or '00'}"
    
    # 勤務時間の計算
    start_dt = datetime.strptime(start_time, "%H:%M")
    end_dt = datetime.strptime(end_time, "%H:%M")
    if end_dt < start_dt:
        end_dt += timedelta(days=1)  # 翌日またぎ対応
    work_hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
    
    return start_time, end_time, work_hours

# ファイルがアップロードされたら処理を開始
if uploaded_file:
    # Excelの読み込み（全シートを取得）
    df = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

    # シートの選択
    sheet_name = st.selectbox("📄 シートを選択:", options=df.keys())
    data = df[sheet_name]

    # **データのプレビュー**
    st.subheader("📋 アップロードしたExcelデータのプレビュー")
    st.dataframe(data.head(10))  # 上位10行のみ表示

    # **名前検索＆勤務時間の抽出**
    memo_list = []
    total_work_hours = 0  # 合計勤務時間
    work_days = 0  # 勤務日数カウント
    
    for row_idx in range(data.shape[0]):  # 行ごとにスキャン
        for col_idx in range(data.shape[1]):  # 列ごとにスキャン
            cell_value = str(data.iloc[row_idx, col_idx]).strip()  # セルの値を取得
            if search_name in cell_value:  # ユーザー指定の名前を検索
                # **3～4行上の日付データを取得（桁数チェック付き）**
                date_value = None
                for offset in range(3, 5):
                    if row_idx - offset >= 0:
                        potential_date = str(data.iloc[row_idx - offset, col_idx]).strip()
                        if potential_date.isdigit() and 1 <= len(potential_date) <= 2:
                            date_value = potential_date  # 2桁以内の数値のみ日付として認識
                            break
                
                # **隣のセルから時間データ取得**
                if col_idx + 1 < data.shape[1]:  # 右隣のセルがある場合のみ取得
                    time_value = str(data.iloc[row_idx, col_idx + 1]).strip()
                    start_time, end_time, work_hours = format_time(time_value)
                    if start_time and end_time:
                        work_days += 1  # 勤務回数をカウント
                        memo_list.append([date_value, search_name, start_time, end_time, work_hours])
                        total_work_hours += work_hours  # 合計時間を加算

    # **結果を表示**
    st.subheader(f"📋 『{search_name}』の勤務時間（フォーマット修正後）")
    if memo_list:
        df_result = pd.DataFrame(memo_list, columns=["日付", "名前", "開始時間", "終了時間", "勤務時間（時間）"])
        st.dataframe(df_result)

        # **合計勤務時間を表示**
        st.subheader("⏳ 勤務データ集計")
        st.write(f"🔢 合計勤務時間: {total_work_hours} 時間")
        st.write(f"📅 勤務日数: {work_days} 日")

        # **ダウンロード可能なCSVを作成**
        csv = df_result.to_csv(index=False).encode("utf-8")
        st.download_button(label="📥 修正データをダウンロード", data=csv, file_name="fixed_work_hours.csv", mime="text/csv")
    else:
        st.warning(f"⚠️ 『{search_name}』が見つかりませんでした。")
