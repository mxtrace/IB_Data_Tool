"""
logger.py — 结构化审计日志
每个操作一条记录，不记录PII（邮箱/地址/电话）。
"""
from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path

_logger: logging.Logger | None = None


def init_logger(base_dir: Path):
    """初始化日志系统"""
    global _logger
    log_dir = base_dir / "logs"
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"

    _logger = logging.getLogger("ib_data_tool")
    _logger.setLevel(logging.DEBUG)
    _logger.handlers.clear()

    # 文件 handler
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s|%(levelname)s|%(message)s"))
    _logger.addHandler(fh)

    # 控制台 handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(asctime)s|%(levelname)s|%(message)s"))
    _logger.addHandler(ch)


def audit(al0: str, step: str, status: str, detail: str = ""):
    """审计日志：不含PII，只记录操作结果"""
    entry = {
        "al0": al0,
        "step": step,
        "status": status,
        "detail": detail[:200],
    }
    if _logger:
        _logger.info(json.dumps(entry, ensure_ascii=False))


def log_info(msg: str):
    if _logger:
        _logger.info(msg)


def log_error(msg: str):
    if _logger:
        _logger.error(msg)


def log_warning(msg: str):
    if _logger:
        _logger.warning(msg)
