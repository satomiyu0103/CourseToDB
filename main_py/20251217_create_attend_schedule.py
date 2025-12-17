"""
毎月の職業準備性講座スケジュールを
表形式にリスト化し、参加講座の一覧を作る
"""

from pathlib import Path
from dotenv import load_dotenv
import os

from modules.calendar_io import (
    creat_file_path_list,
    find_calendar_name,
    input_file_to_df,
)

# from modules.calendar_parser import creat_event_dic, creat_month_schdule
from modules.schedule_output import (
    creat_schedule_df,
    attend_course_schedule_df,
    output_to_new_sheet,
)


def main():
    try:
        PROJECT_ROOT = Path(__file__).resolve().parent.parent
    except NameError:
        PROJECT_ROOT = Path.cwd()
    env_path = PROJECT_ROOT / "config" / ".env"
    load_dotenv(env_path)
    folda_path = os.getenv("FOLDA_PATH")

    # デバッグ出力
    print(f"PROJECT_ROOT: {PROJECT_ROOT}")
    print(f"env_path: {env_path}")
    print(f"folda_path: {folda_path}")

    data_list = creat_file_path_list(folda_path)
    print(f"data_list: {data_list}")  # ← ここで確認
    if not data_list:  # リストが空の場合
        print("エラー: ファイルが見つかりません")
        return

    file_path = data_list[-1]  # 最新のファイルを選択
    result_df = creat_schedule_df(file_path)
    attend_df = attend_course_schedule_df(result_df)
    output_to_new_sheet(file_path, result_df, attend_df)


if __name__ == "__main__":
    main()
