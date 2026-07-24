"""
logger.py — 结构化日志系统
文件日志记录 DEBUG 级别（详细诊断），控制台 INFO 级别。
"""
from __future__ import annotations

import json
import logging
import os
import traceback
from datetime import datetime
from pathlib import Path

_logger: logging.Logger | None = None
_log_file: Path | None = None


def init_logger(base_dir: Path):
    """初始化日志系统"""
    global _logger, _log_file
    log_dir = base_dir / "logs"
    log_dir.mkdir(exist_ok=True)

    _log_file = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"

    _logger = logging.getLogger("ib_data_tool")
    _logger.setLevel(logging.DEBUG)
    _logger.handlers.clear()

    # 文件 handler — DEBUG 级别（详细）
    fh = logging.FileHandler(_log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s|%(levelname)s|%(message)s"))
    _logger.addHandler(fh)

    # 控制台 handler — INFO 级别
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(asctime)s|%(levelname)s|%(message)s"))
    _logger.addHandler(ch)

    # 启动时记录环境信息
    _logger.info("=" * 50)
    _logger.debug(f"日志文件：{_log_file}")
    _logger.debug(f"Python：{os.sys.executable}")
    _logger.debug(f"用户：{os.environ.get('USERNAME', 'unknown')}")
    _logger.debug(f"工作目录：{base_dir}")


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


def log_debug(msg: str):
    """详细诊断信息（仅写入文件，不显示在控制台）"""
    if _logger:
        _logger.debug(msg)


def log_error(msg: str):
    if _logger:
        _logger.error(msg)


def log_warning(msg: str):
    if _logger:
        _logger.warning(msg)


def log_exception(msg: str):
    """记录异常（含完整 traceback）"""
    if _logger:
        _logger.error(f"{msg}\n{traceback.format_exc()}")


def get_log_file() -> Path | None:
    return _log_file
