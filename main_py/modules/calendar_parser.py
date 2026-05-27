"""====================
■ Excelファイルの読み書き
===================="""

import pandas as pd

WEEKDAYS = {"月", "火", "水", "木", "金", "土"}


def deside_setting():
    start_row = 1  # カレンダーが始まる行
    end_row = 26  # カレンダーが終わる行
    row_step = 5  # 一週間のデータが何行あるか
    start_col = 1  # カレンダーが始まる列番号
    end_col = 6  # カレンダーが終わる列番号
    setting_list = [start_row, end_row, row_step, start_col, end_col]
    return setting_list


def _cell_text(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        text = value.strip()
        return text or None
    return value


def _is_weekday_label(value):
    text = _cell_text(value)
    return text in WEEKDAYS if text else False


def _to_date(value):
    if pd.isna(value):
        return None
    try:
        return pd.Timestamp(value).normalize()
    except (ValueError, TypeError):
        return None


def _count_weekday_labels(df, row, start_col, end_col):
    return sum(
        1
        for col in range(start_col, end_col + 1)
        if _is_weekday_label(df.iloc[row, col])
    )


def _find_weekday_rows(df, start_row, end_row, start_col, end_col):
    weekday_rows = []
    for row in range(start_row, end_row):
        weekday_count = _count_weekday_labels(df, row, start_col, end_col)
        has_date_above = row > 0 and any(
            _to_date(df.iloc[row - 1, col]) is not None
            for col in range(start_col, end_col + 1)
        )
        if weekday_count >= 3 or (weekday_count >= 1 and has_date_above):
            weekday_rows.append(row)
    return weekday_rows


def _extract_date_cells(df, row, start_col, end_col):
    return [_to_date(df.iloc[row, col]) for col in range(start_col, end_col + 1)]


def _resolve_week_dates(df, weekday_row, start_col, end_col, previous_dates):
    col_count = end_col - start_col + 1
    if weekday_row > 0:
        above_dates = _extract_date_cells(df, weekday_row - 1, start_col, end_col)
        if any(date is not None for date in above_dates):
            anchor_idx = next(
                idx for idx, date in enumerate(above_dates) if date is not None
            )
            anchor_date = above_dates[anchor_idx]
            resolved = []
            for idx in range(col_count):
                if above_dates[idx] is not None:
                    resolved.append(above_dates[idx])
                else:
                    resolved.append(
                        anchor_date + pd.Timedelta(days=idx - anchor_idx)
                    )
            return resolved

    if previous_dates:
        return [
            date + pd.Timedelta(days=7) if date is not None else None
            for date in previous_dates
        ]

    return [None] * col_count


def creat_event_dic(df, weekday_row, week_col, date_value):
    program_row = weekday_row + 1
    target_row = weekday_row + 2
    attend_row = weekday_row + 3

    if attend_row >= len(df):
        return None

    program_name = _cell_text(df.iloc[program_row, week_col])
    if program_name is None:
        return None

    return {
        "日付": date_value,
        "曜日": _cell_text(df.iloc[weekday_row, week_col]),
        "プログラム名": df.iloc[program_row, week_col],
        "参加目安": df.iloc[target_row, week_col],
        "出席予定": df.iloc[attend_row, week_col],
    }


def creat_month_schdule(df, start_row, end_row, row_step, start_col, end_col):
    all_events = []
    previous_dates = None

    for weekday_row in _find_weekday_rows(df, start_row, end_row, start_col, end_col):
        week_dates = _resolve_week_dates(
            df, weekday_row, start_col, end_col, previous_dates
        )
        previous_dates = week_dates

        for offset, week_col in enumerate(range(start_col, end_col + 1)):
            date_value = week_dates[offset]
            if date_value is None:
                continue

            event_data = creat_event_dic(df, weekday_row, week_col, date_value)
            if event_data:
                all_events.append(event_data)

    return all_events
