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
            writer.writerow([
                r.al0,
                now_str,
                r.bc_login,
                "IB_DATA_EMAIL_TO_SELLER",
            ])

    return csv_path


def open_pending_tasks():
    """打开 OC Pending Tasks 页面"""
    url = "https://trans-logistics-cn.amazon.com/aglt/appViews/app#/pending-tasks"
    webbrowser.open(url)
