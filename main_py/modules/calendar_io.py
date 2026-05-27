"""====================
■ Excelファイルの読み書き
===================="""

import os
from pathlib import Path

import pandas as pd
import glob
import re

ATTEND_COLUMNS = ["日付", "講座名", "開始日時", "終了日時"]

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


def resolve_file_path(folda_path: str | None = None) -> str | None:
    file_path = os.getenv("FILE_PATH")
    if file_path:
        path = Path(file_path)
        if path.exists():
            return str(path)
        print(f"error: FILE_PATH のファイルが見つかりません: {file_path}")

    if not folda_path:
        return None

    data_list = creat_file_path_list(folda_path)
    if not data_list:
        return None
    return data_list[-1]


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
