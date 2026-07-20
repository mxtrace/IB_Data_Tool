"""
wal.py — Write-Ahead Log（崩溃恢复）
每票处理完立即追加一行，启动时检查恢复未回写的状态。
防止崩溃导致重复发邮件。
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

WAL_FILENAME = "wal.jsonl"


def append_wal(base_dir: Path, al0: str, status: str, reason: str = ""):
    """每票处理完立即追加一行"""
    wal = base_dir / WAL_FILENAME
    entry = {
        "al0": al0,
        "status": status,
        "reason": reason,
        "ts": datetime.now().isoformat(),
    }
    with open(wal, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def load_wal(base_dir: Path) -> list[dict]:
    """读取 WAL 全部记录"""
    wal = base_dir / WAL_FILENAME
    if not wal.exists():
        return []
    entries = []
    for line in wal.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            try:
                entries.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return entries


def recover_pending_list_from_wal(base_dir: Path):
    """
    启动时调用：将 WAL 中 success/odm 但未回写到 Pending List 的记录补写。
    执行完毕后清空 WAL。
    """
    entries = load_wal(base_dir)
    if not entries:
        return

    from steps.step1_sync import write_single_ticket
    for entry in entries:
        if entry["status"] in ("success", "odm"):
            write_single_ticket(base_dir, entry["al0"], entry["status"])

    # 恢复完毕，清空 WAL
    clear_wal(base_dir)


def clear_wal(base_dir: Path):
    """清空 WAL 文件"""
    wal = base_dir / WAL_FILENAME
    if wal.exists():
        wal.write_text("", encoding="utf-8")
