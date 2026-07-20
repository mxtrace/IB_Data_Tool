"""
excel_helper.py — openpyxl 安全包装
统一使用 context manager，防止文件句柄泄漏。
"""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import openpyxl


@contextmanager
def open_workbook(path: Path, read_only=False, data_only=False):
    """
    安全打开 Excel 工作簿。
    自动在退出时关闭，即使发生异常。
    """
    wb = openpyxl.load_workbook(str(path), read_only=read_only, data_only=data_only)
    try:
        yield wb
    finally:
        wb.close()


@contextmanager
def open_workbook_write(path: Path):
    """
    打开 Excel 用于写入。退出时自动保存并关闭。
    异常时仅关闭不保存。
    """
    wb = openpyxl.load_workbook(str(path))
    try:
        yield wb
        wb.save(str(path))
    finally:
        wb.close()
