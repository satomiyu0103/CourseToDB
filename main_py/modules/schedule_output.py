"""====================
■ スケジュールリストの作成
===================="""

from pathlib import Path

import pandas as pd

from modules.calendar_io import find_calendar_name, input_file_to_df, read_attend_sheet
from modules.calendar_parser import creat_month_schdule, deside_setting

ATTEND_SHEET_NAME = "参加講座一覧表"
JST_OFFSET = "+09:00"
ATTEND_COLUMNS = ["日付", "講座名", "開始日時", "終了日時"]
CSV_COLUMNS = ["title", "start", "end"]


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


def _build_course_datetimes(date_series):
    dates = pd.to_datetime(date_series, errors="coerce")
    normalized = dates.dt.normalize()
    start = normalized + pd.Timedelta(hours=10)
    end = normalized + pd.Timedelta(hours=12)
    return normalized, start, end


def format_jst_datetime(value) -> str:
    ts = pd.Timestamp(value)
    if pd.isna(ts):
        return ""
    return ts.strftime("%Y/%m/%d %H:%M") + JST_OFFSET


def attend_course_schedule_df(result_df):
    """
    出席予定が〇の行だけ抽出し、確認用の4列DataFrameを返す
    """
    if result_df is None or result_df.empty:
        print("結果を抽出できませんでした。設定を確認してください")
        return pd.DataFrame(columns=ATTEND_COLUMNS)

    attend_df = result_df[result_df["出席予定"] == "〇"][["日付", "プログラム名"]].copy()
    dates, start, end = _build_course_datetimes(attend_df["日付"])
    attend_df["日付"] = dates
    attend_df["開始日時"] = start
    attend_df["終了日時"] = end
    attend_df = attend_df.rename(columns={"プログラム名": "講座名"})
    attend_df = attend_df[ATTEND_COLUMNS]

    print("\n --- 結果を表示する ---/")
    print(attend_df)
    return attend_df


def attend_df_to_csv_df(attend_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "title": attend_df["講座名"],
            "start": attend_df["開始日時"].apply(format_jst_datetime),
            "end": attend_df["終了日時"].apply(format_jst_datetime),
        }
    )


def default_csv_path(file_path: str) -> Path:
    excel_path = Path(file_path)
    stem = excel_path.stem.replace("配布用_", "登録用_", 1)
    return excel_path.with_name(f"{stem}_notion.csv")


def legacy_csv_paths(file_path: str) -> list[Path]:
    """以前の命名規則で出力されたCSVの候補"""
    excel_path = Path(file_path)
    return [excel_path.with_name(f"{excel_path.stem}_notion.csv")]


def cleanup_legacy_csv_files(file_path: str, current_csv_path: Path) -> None:
    for legacy_path in legacy_csv_paths(file_path):
        if legacy_path == current_csv_path or not legacy_path.exists():
            continue
        legacy_path.unlink()
        print(f"旧CSVを削除しました: {legacy_path}")


def export_attend_csv(attend_df: pd.DataFrame, csv_path: str | Path) -> Path:
    csv_path = Path(csv_path)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if csv_path.exists():
        print(f"既存CSVを上書きします: {csv_path}")

    csv_df = attend_df_to_csv_df(attend_df)
    csv_df.to_csv(csv_path, index=False, encoding="utf-8-sig", columns=CSV_COLUMNS)
    print(f"CSVを保存しました: {csv_path}")
    return csv_path


def export_attend_csv_from_excel(file_path: str, csv_path: str | Path | None = None) -> Path:
    attend_df = read_attend_sheet(file_path, ATTEND_SHEET_NAME)
    output_path = Path(csv_path) if csv_path else default_csv_path(file_path)

    if attend_df.empty:
        print(f"エラー: {ATTEND_SHEET_NAME} にデータがありません")
        return output_path

    cleanup_legacy_csv_files(file_path, output_path)
    return export_attend_csv(attend_df, output_path)


def output_to_new_sheet(file_path, result_df, attend_df):
    # できたスケジュールを元のExcelに作成した新シートに出力する
    new_sheet_name = "今月の講座一覧表"
    print(f"Excelシートを上書き保存します: {file_path}")
    with pd.ExcelWriter(
        file_path,
        mode="a",
        engine="openpyxl",
        if_sheet_exists="replace",
        datetime_format="yyyy/mm/dd hh:mm",
        date_format="yyyy/mm/dd",
    ) as writer:
        result_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
        attend_df.to_excel(writer, sheet_name=ATTEND_SHEET_NAME, index=False)
        print(f"{new_sheet_name}に保存しました")
        print(f"{ATTEND_SHEET_NAME}に保存しました")
