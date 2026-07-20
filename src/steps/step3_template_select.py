"""
step3_template_select.py — 模板选择（AMS / ENS）
"""
from __future__ import annotations


def select_template(pod: str) -> str:
    """
    根据 POD 前两位判断模板类型。
    pod[:2] == "US" → "AMS"
    其他 → "ENS"
    """
    region = (pod or "")[:2].upper()
    return "AMS" if region == "US" else "ENS"
