"""====================
■ Excelファイルの読み書き
===================="""

import pandas as pd


def deside_setting():
    start_row = 1  # カレンダーが始まる行
    end_row = 26  # カレンダーが終わる行
    row_step = 5  # 一週間のデータが何行あるか
    start_col = 1  # カレンダーが始まる列番号
    end_col = 6  # カレンダーが終わる列番号
    setting_list = [start_row, end_row, row_step, start_col, end_col]
    return setting_list


def creat_event_dic(df, date_row, week_col):
    """
    指定のセルから1日分のイベント情報を抽出
    """
    date = df.iloc[date_row, week_col]
    if pd.isna(date):
        return None
    event_data = {
        "日付": df.iloc[date_row, week_col],
        "曜日": df.iloc[date_row + 1, week_col],
        "プログラム名": df.iloc[date_row + 2, week_col],
        "参加目安": df.iloc[date_row + 3, week_col],
        "出席予定": df.iloc[date_row + 4, week_col],
    }
    return event_data


def creat_month_schdule(df, start_row, end_row, row_step, start_col, end_col):
    all_events = []
    for date_row in range(start_row, end_row, row_step):
        for week_col in range(start_col, end_col + 1):

            # 情報を抽出
            event_data = creat_event_dic(df, date_row, week_col)

            if event_data:
                # all_eventのlistに追加
                all_events.append(event_data)
    return all_events
