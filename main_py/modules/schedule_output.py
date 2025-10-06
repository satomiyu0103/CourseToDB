"""====================
■ スケジュールリストの作成
===================="""

import pandas as pd

from modules.calendar_io import find_calendar_name, input_file_to_df
from modules.calendar_parser import creat_month_schdule, deside_setting


def list_to_df(all_events):
    # リストを表に変換する
    if not all_events:
        print("結果を抽出できませんでした。設定を確認してください")
        return pd.DataFrame()

    result_df = pd.DataFrame(all_events)

    # プログラム名の改行\nを空白に変換
    result_df["プログラム名"] = result_df["プログラム名"].str.replace(
        "\n", " ", regex=False
    )

    print("\n --- 結果を表示する ---/")
    print(result_df)
    return result_df


def creat_schedule_df(file_path):
    # シート名を探す
    sheet_name = find_calendar_name(file_path)

    # Excelを読み込む
    df = input_file_to_df(file_path, sheet_name)
    start_row, end_row, row_step, start_col, end_col = deside_setting()
    if df.empty:
        return

    # スケジュールの抽出
    all_events = creat_month_schdule(
        df, start_row, end_row, row_step, start_col, end_col
    )
    result_df = list_to_df(all_events)
    return result_df


def attend_course_schedule_df(result_df):
    """
    出席予定が〇の行だけ抽出し、日付とプログラム名だけのDataFrameを返す
    """
    # 出席予定が〇の行を抽出
    attend_df = result_df[result_df["出席予定"] == "〇"]
    # 日付とプログラム名だけ抽出
    attend_df = attend_df[["日付", "プログラム名"]]
    attend_df["日付"] = attend_df["日付"].apply(
        lambda x: x.replace(hour=10, minute=0, second=0) if pd.notnull(x) else x
    )
    print("\n --- 結果を表示する ---/")
    print(attend_df)
    return attend_df


def output_to_new_sheet(file_path, result_df, attend_df):
    # できたスケジュールを元のExcelに作成した新シートに出力する
    new_sheet_name = "今月の講座一覧表"
    attend_sheet_name = "参加講座一覧表"
    with pd.ExcelWriter(
        file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
    ) as writer:
        result_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
        print(f"{new_sheet_name}に保存しました")

    with pd.ExcelWriter(
        file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
    ) as writer:
        attend_df.to_excel(writer, sheet_name=attend_sheet_name, index=False)
        print(f"{attend_sheet_name}に保存しました")
