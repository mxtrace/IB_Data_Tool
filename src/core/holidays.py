"""
holidays.py — 2026年中国法定节假日 + 工作日计算
"""
from __future__ import annotations

from datetime import date, timedelta

# 2026年中国法定节假日（休息日）
# 来源：国务院办公厅发布的放假安排
HOLIDAYS_2026 = {
    # 元旦
    date(2026, 1, 1), date(2026, 1, 2), date(2026, 1, 3),
    # 春节
    date(2026, 2, 17), date(2026, 2, 18), date(2026, 2, 19),
    date(2026, 2, 20), date(2026, 2, 21), date(2026, 2, 22),
    date(2026, 2, 23),
    # 清明
    date(2026, 4, 5), date(2026, 4, 6), date(2026, 4, 7),
    # 劳动节
    date(2026, 5, 1), date(2026, 5, 2), date(2026, 5, 3),
    date(2026, 5, 4), date(2026, 5, 5),
    # 端午
    date(2026, 5, 31), date(2026, 6, 1), date(2026, 6, 2),
    # 中秋+国庆
    date(2026, 10, 1), date(2026, 10, 2), date(2026, 10, 3),
    date(2026, 10, 4), date(2026, 10, 5), date(2026, 10, 6),
    date(2026, 10, 7), date(2026, 10, 8),
}

# 2026年调休上班日（周末但需上班）
WORKDAYS_2026 = {
    date(2026, 2, 14), date(2026, 2, 15),  # 春节调休
    date(2026, 4, 4),   # 清明调休
    date(2026, 4, 26),  # 劳动节调休
    date(2026, 5, 30),  # 端午调休
    date(2026, 9, 27),  # 国庆调休
    date(2026, 10, 10), # 国庆调休
}


def is_workday(d: date) -> bool:
    """判断是否为工作日"""
    if d in HOLIDAYS_2026:
        return False
    if d in WORKDAYS_2026:
        return True
    # 周六日
    return d.weekday() < 5


def next_workday(d: date) -> date:
    """返回 d 之后的下一个工作日（不含 d 本身）"""
    candidate = d + timedelta(days=1)
    while not is_workday(candidate):
        candidate += timedelta(days=1)
    return candidate
