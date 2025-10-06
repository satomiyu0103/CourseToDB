"""====================
■ Excelファイルの読み書き
===================="""

import pandas as pd
import glob
import re

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


def input_file_to_df(file_path, sheet_name):
    try:
        input_file_path = file_path
        sheet_name = sheet_name
        df = pd.read_excel(input_file_path, sheet_name=sheet_name, header=None)
        return df
    except FileNotFoundError:
        print(f"error: ファイル{file_path}が見つかりません")
