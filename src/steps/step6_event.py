"""
step6_event.py — 打卡（生成 Event CSV + 打开 Pending Tasks 页面）
"""
from __future__ import annotations

import csv
import webbrowser
from datetime import datetime
from pathlib import Path

from core.batch_controller import TicketResult


def generate_event_csv(records: list[TicketResult], base_dir: Path) -> Path:
    """
    生成 Event CSV 文件。
    字段：booking_id, timestamp, bc_login, event_code
    """
    output_dir = base_dir / "Output"
    output_dir.mkdir(parents=True, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    csv_path = output_dir / f"event_list_{today}.csv"

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["booking_id", "timestamp", "bc_login", "event_code"])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in records:
            event_code = "IB_DATA_ODM" if r.status == "odm" else "IB_DATA_EMAIL_TO_SELLER"
            writer.writerow([
                r.al0,
                now_str,
                r.bc_login,
                event_code,
            ])

    return csv_path


def open_pending_tasks():
    """打开 OC Pending Tasks 页面"""
    url = "https://trans-logistics-cn.amazon.com/aglt/appViews/app#/pending-tasks"
    webbrowser.open(url)


def cleanup_output(base_dir: Path) -> int:
    """
    将 Output/ 下的 .xlsx/.csv 文件移入回收站。
    返回清理的文件数。失败时移到 Output/archive/{date}/。
    """
    import tkinter as tk
    from tkinter import messagebox

    output_dir = base_dir / "Output"
    if not output_dir.exists():
        return 0

    files = list(output_dir.glob("*.xlsx")) + list(output_dir.glob("*.csv"))
    if not files:
        return 0

    # 弹窗确认
    root = tk.Tk()
    root.withdraw()
    confirm = messagebox.askyesno(
        "清理 Output",
        f"请确认打卡已完成。

将 {len(files)} 个文件移入回收站：
"
        + "
".join(f"  • {f.name}" for f in files[:10])
        + ("
  ..." if len(files) > 10 else ""),
    )
    root.destroy()

    if not confirm:
        return 0

    cleaned = 0
    for f in files:
        if _send_to_recycle_bin(str(f)):
            cleaned += 1
        else:
            # Fallback: 移到 archive 子目录
            archive = output_dir / "archive" / datetime.now().strftime("%Y%m%d")
            archive.mkdir(parents=True, exist_ok=True)
            try:
                f.rename(archive / f.name)
                cleaned += 1
            except Exception:
                pass

    return cleaned


def _send_to_recycle_bin(file_path: str) -> bool:
    """使用 Windows Shell API 将文件移入回收站"""
    try:
        from win32com.shell import shell, shellcon
        result = shell.SHFileOperation((
            0,
            shellcon.FO_DELETE,
            file_path,
            None,
            shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION | shellcon.FOF_SILENT,
            None,
            None,
        ))
        return result[0] == 0
    except Exception:
        return False
