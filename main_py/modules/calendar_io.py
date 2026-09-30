"""====================
■ Excelファイルの読み書き
===================="""

import glob
import os
import re
from pathlib import Path

import pandas as pd

ATTEND_COLUMNS = ["日付", "講座名", "開始日時", "終了日時"]
_SCHEDULE_YYYYMM_RE = re.compile(r"(\d{4})年(\d{1,2})月")


def yyyymm_from_path(path: str | Path) -> int | None:
    match = _SCHEDULE_YYYYMM_RE.search(Path(path).name)
    if not match:
        return None
    year, month = int(match.group(1)), int(match.group(2))
    return year * 100 + month


def _schedule_sort_key(path: str) -> tuple[int, float, str]:
    yyyymm = yyyymm_from_path(path)
    try:
        mtime = Path(path).stat().st_mtime
    except OSError:
        mtime = 0.0
    return (yyyymm if yyyymm is not None else -1, mtime, path)


def sort_schedule_paths(paths: list[str]) -> list[str]:
    return sorted(paths, key=_schedule_sort_key, reverse=True)


def pick_latest_schedule_path(paths: list[str]) -> str | None:
    ranked = sort_schedule_paths(paths)
    return ranked[0] if ranked else None


def choose_schedule_path_interactive(paths: list[str]) -> str | None:
    ranked = sort_schedule_paths(paths)
    if not ranked:
        return None

    print("処理する Excel を選んでください:")
    for index, file_path in enumerate(ranked, start=1):
        label = Path(file_path).name
        yyyymm = yyyymm_from_path(file_path)
        if yyyymm is not None:
            year, month = divmod(yyyymm, 100)
            label = f"{label}  （{year}年{month}月）"
        print(f"  {index}. {label}")

    while True:
        try:
            choice = input("番号を入力（空 Enter で先頭）: ").strip()
        except EOFError:
            print("入力がキャンセルされました")
            return None

        if not choice:
            return ranked[0]

        if not choice.isdigit():
            print("1 から数字で選んでください")
            continue

        selected = int(choice)
        if 1 <= selected <= len(ranked):
            return ranked[selected - 1]

        print(f"1 〜 {len(ranked)} の番号を入力してください")


def creat_file_path_list(folda_path):
    """
    フォルダから職業準備性講座スケジュールの
    ファイルパスを抽出してリスト化
    Args:
        folda_path (_type_): _description_

    Returns:
        _type_: _description_
    """
    file_paths = glob.glob(folda_path)
    data_list = []
    for file_path in file_paths:
        # print(f"\n{file_path}")
        data_list.append(file_path)
    return data_list


def find_calendar_name(file_path):
    """
    講座予定のカレンダーが入っているシート名を探す
    Args:
        file_path (_type_): _description_
    """
    try:
        xl = pd.ExcelFile(file_path)
        candidates = ["今月の予定", "今月", "月"] + [f"{i}月" for i in range(1, 13)]
        # 正規表現で候補に近いものも拾う
        for sheet in xl.sheet_names:
            for cand in candidates:
                if re.search(cand, sheet):
                    return sheet
        # どれも該当しなければ最初のシート
        return xl.sheet_names[0]
    except FileNotFoundError:
        print(f"error: ファイル{file_path}が見つかりません")


def _normalize_excel_path(raw_path: str) -> Path:
    return Path(raw_path.strip().strip('"')).expanduser()


def resolve_file_path(
    folda_path: str | None = None,
    explicit_path: str | None = None,
    *,
    interactive: bool = True,
) -> str | None:
    if explicit_path:
        path = _normalize_excel_path(explicit_path)
        if path.is_file():
            return str(path.resolve())
        print(f"error: 指定されたファイルが見つかりません: {explicit_path}")
        return None

    file_path = os.getenv("FILE_PATH")
    if file_path:
        path = _normalize_excel_path(file_path)
        if path.is_file():
            return str(path.resolve())
        print(f"error: FILE_PATH のファイルが見つかりません: {file_path}")

    if not folda_path:
        return None

    data_list = creat_file_path_list(folda_path)
    if not data_list:
        return None

    if interactive:
        import sys

        if sys.stdin.isatty():
            chosen = choose_schedule_path_interactive(data_list)
            if chosen:
                return chosen

    latest = pick_latest_schedule_path(data_list)
    if latest:
        yyyymm = yyyymm_from_path(latest)
        if yyyymm is not None:
            year, month = divmod(yyyymm, 100)
            print(
                f"FOLDA_PATH から自動選択: {Path(latest).name} （{year}年{month}月・暦の最新）"
            )
        else:
            print(f"FOLDA_PATH から自動選択: {Path(latest).name}")
    return latest


def read_attend_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"error: ファイル{file_path}が見つかりません")
        return pd.DataFrame(columns=ATTEND_COLUMNS)
    except ValueError as exc:
        print(f"error: シート {sheet_name} が見つかりません ({exc})")
        return pd.DataFrame(columns=ATTEND_COLUMNS)

    missing = [col for col in ATTEND_COLUMNS if col not in df.columns]
    if missing:
        print(f"error: {sheet_name} に必要な列がありません: {', '.join(missing)}")
        return pd.DataFrame(columns=ATTEND_COLUMNS)

    attend_df = df[ATTEND_COLUMNS].copy()
    attend_df["日付"] = pd.to_datetime(attend_df["日付"], errors="coerce").dt.normalize()
    attend_df["開始日時"] = pd.to_datetime(attend_df["開始日時"], errors="coerce")
    attend_df["終了日時"] = pd.to_datetime(attend_df["終了日時"], errors="coerce")
    attend_df["講座名"] = attend_df["講座名"].astype(str).str.replace("\n", " ", regex=False)
    return attend_df


def input_file_to_df(file_path, sheet_name):
    try:
        input_file_path = file_path
        sheet_name = sheet_name
        df = pd.read_excel(input_file_path, sheet_name=sheet_name, header=None)
        return df
    except FileNotFoundError:
        print(f"error: ファイル{file_path}が見つかりません")
