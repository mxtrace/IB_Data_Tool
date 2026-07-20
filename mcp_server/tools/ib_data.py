"""
ib_data.py — 进仓数据读取工具
从共享盘读取最新 Excel，按 bc_login 过滤，返回结构化数据。
"""
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Optional

from tools.config import get_config

# 进仓数据列映射（0-indexed）
COL_MAP = {
    "al0": 0,             # A
    "hbl_number": 1,      # B
    "shipper_id": 2,      # C
    "pol": 5,             # F
    "pod": 6,             # G
    "bc_login": 8,        # I
    "flipped_fc": 10,     # K
    "received_cartons": 13,  # N
    "received_volume": 14,   # O
    "received_weight": 15,   # P
}


def read_ib_data(selected_logins: list[str], max_count: int = 30) -> dict:
    """
    读取共享盘最新进仓 Excel，过滤并返回数据。

    Returns:
        {
            "al0_list": [str],
            "ib_data": {al0: {field: value}},
            "total_before_filter": int,
            "total_after_filter": int,
            "truncated": bool,
            "source_file": str,
        }
    """
    config = get_config()
    if "error" in config:
        return config

    shared_path = Path(config.get("shared_drive_path", ""))
    if not shared_path.exists():
        return {"error": f"共享盘路径不可达: {shared_path}"}

    # 定位最新文件
    xlsx_files = list(shared_path.glob("*.xlsx"))
    if not xlsx_files:
        return {"error": f"共享盘下无 .xlsx 文件: {shared_path}"}

    # 按文件名中的日期降序（假设文件名含日期）
    # 如果无法从文件名解析日期，退回使用修改时间
    latest_file = _find_latest_file(xlsx_files)
    if not latest_file:
        return {"error": "无法定位最新进仓数据文件"}

    # 读取 Excel
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(latest_file), read_only=True, data_only=True)
        ws = wb.active
    except Exception as e:
        return {"error": f"读取 Excel 失败: {e}"}

    # 解析行
    rows = list(ws.iter_rows(min_row=2, values_only=True))  # 跳过表头
    total_before = len(rows)

    # 按 bc_login 过滤
    selected_set = set(login.lower().strip() for login in selected_logins)
    filtered_rows = []
    for row in rows:
        bc_login_val = str(row[COL_MAP["bc_login"]] or "").strip().lower()
        if bc_login_val in selected_set:
            filtered_rows.append(row)

    # 去重（按 AL0）
    seen_al0 = set()
    unique_rows = []
    for row in filtered_rows:
        al0 = str(row[COL_MAP["al0"]] or "").strip()
        if al0 and al0 not in seen_al0:
            seen_al0.add(al0)
            unique_rows.append(row)

    total_after = len(unique_rows)
    truncated = total_after > max_count
    if truncated:
        unique_rows = unique_rows[:max_count]

    # 构建输出
    al0_list = []
    ib_data = {}
    for row in unique_rows:
        al0 = str(row[COL_MAP["al0"]] or "").strip()
        al0_list.append(al0)
        ib_data[al0] = {
            "al0": al0,
            "hbl_number": str(row[COL_MAP["hbl_number"]] or "").strip(),
            "shipper_id": str(row[COL_MAP["shipper_id"]] or "").strip(),
            "pol": str(row[COL_MAP["pol"]] or "").strip(),
            "pod": str(row[COL_MAP["pod"]] or "").strip(),
            "bc_login": str(row[COL_MAP["bc_login"]] or "").strip(),
            "flipped_fc": str(row[COL_MAP["flipped_fc"]] or "").strip(),
            "received_cartons": _to_float(row[COL_MAP["received_cartons"]]),
            "received_volume": _to_float(row[COL_MAP["received_volume"]]),
            "received_weight": _to_float(row[COL_MAP["received_weight"]]),
        }

    wb.close()

    return {
        "al0_list": al0_list,
        "ib_data": ib_data,
        "total_before_filter": total_before,
        "total_after_filter": total_after,
        "truncated": truncated,
        "source_file": str(latest_file),
    }


def _find_latest_file(files: list[Path]) -> Optional[Path]:
    """按文件名日期或修改时间找最新文件。"""
    # 尝试从文件名提取日期 (常见格式: YYYYMMDD 或 YYYY-MM-DD)
    date_pattern = re.compile(r"(\d{8}|\d{4}-\d{2}-\d{2})")
    dated_files = []
    for f in files:
        match = date_pattern.search(f.stem)
        if match:
            date_str = match.group(1).replace("-", "")
            dated_files.append((date_str, f))

    if dated_files:
        dated_files.sort(key=lambda x: x[0], reverse=True)
        return dated_files[0][1]

    # 退回使用修改时间
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _to_float(val) -> float:
    """安全转换为 float。"""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
