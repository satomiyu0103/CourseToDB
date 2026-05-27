"""
毎月の職業準備性講座スケジュールを
表形式にリスト化し、参加講座の一覧を作る
"""

import argparse
import os
from pathlib import Path

from dotenv import load_dotenv

from modules.calendar_io import resolve_file_path
from modules.schedule_output import (
    ATTEND_SHEET_NAME,
    attend_course_schedule_df,
    creat_schedule_df,
    default_csv_path,
    export_attend_csv_from_excel,
    output_to_new_sheet,
)


def load_project_env() -> Path:
    try:
        project_root = Path(__file__).resolve().parent.parent
    except NameError:
        project_root = Path.cwd()
    env_path = project_root / "config" / ".env"
    load_dotenv(env_path)
    return project_root


def build_excel(file_path: str) -> bool:
    result_df = creat_schedule_df(file_path)
    if result_df is None or result_df.empty:
        print("エラー: スケジュールを抽出できませんでした")
        return False

    attend_df = attend_course_schedule_df(result_df)
    if attend_df.empty:
        print("エラー: 参加講座がありません")
        return False

    output_to_new_sheet(file_path, result_df, attend_df)
    return True


def export_csv(file_path: str, csv_path: str | None = None) -> None:
    output_path = export_attend_csv_from_excel(file_path, csv_path)
    if output_path.exists():
        print(f"{ATTEND_SHEET_NAME} から CSV を出力しました")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="職業準備性講座スケジュールから参加講座一覧を作成する"
    )
    parser.add_argument(
        "--export-csv",
        action="store_true",
        help="Excelの確認・修正後にCSVだけ再出力する",
    )
    args = parser.parse_args()

    project_root = load_project_env()
    folda_path = os.getenv("FOLDA_PATH")
    csv_output_path = os.getenv("CSV_OUTPUT_PATH")

    print(f"PROJECT_ROOT: {project_root}")
    print(f"FOLDA_PATH: {folda_path}")
    print(f"FILE_PATH: {os.getenv('FILE_PATH')}")

    file_path = resolve_file_path(folda_path)
    if not file_path:
        print("エラー: ファイルが見つかりません")
        return

    print(f"使用ファイル: {file_path}")

    csv_path = csv_output_path or default_csv_path(file_path)

    if args.export_csv:
        export_csv(file_path, csv_path)
        return

    if build_excel(file_path):
        export_csv(file_path, csv_path)


if __name__ == "__main__":
    main()
