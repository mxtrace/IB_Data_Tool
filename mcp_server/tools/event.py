"""
event.py — 打卡 CSV 生成工具
"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
OUTPUT_DIR = DATA_DIR / "output"


def gen_event_csv(sent_records: list[dict]) -> dict:
    """
    生成打卡 CSV 文件。

    Args:
        sent_records: [{al0, actual_time, user_name}]

    Returns:
        {success, file_path, record_count}
    """
    if not sent_records:
        return {"success": False, "error": "无已发送记录，无需生成 CSV"}

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Event_{timestamp}.csv"
    file_path = OUTPUT_DIR / filename

    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["bookingID", "actualTime", "userName", "eventCode"])
        for record in sent_records:
            writer.writerow([
                record.get("al0", ""),
                record.get("actual_time", ""),
                record.get("user_name", ""),
                "IB_DATA_EMAIL_TO_SELLER",
            ])

    return {
        "success": True,
        "file_path": str(file_path),
        "record_count": len(sent_records),
    }
