"""
step6_event.py - Event CSV generation and Pending Tasks page
"""
from __future__ import annotations

import csv
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

from core.batch_controller import TicketResult


def _is_headless() -> bool:
    return "--headless" in sys.argv


def generate_event_csv(records: list[TicketResult], base_dir: Path) -> Path:
    """Generate Event CSV for batch upload."""
    output_dir = base_dir / "Output"
    output_dir.mkdir(parents=True, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    csv_path = output_dir / f"event_list_{today}.csv"

    # Append mode: write header only if file doesn't exist
    file_exists = csv_path.exists()
    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["bookingID", "actualTime", "userName", "eventCode"])
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in records:
            event_code = "IB_DATA_EMAIL_TO_SELLER"
            writer.writerow([
                r.al0,
                now_str,
                r.bc_login,
                event_code,
            ])

    return csv_path


def open_pending_tasks():
    """Open OC Pending Tasks page (skip in headless mode)."""
    if _is_headless():
        print("[INFO] Headless: skip opening Pending Tasks page")
        return
    url = "https://trans-logistics-cn.amazon.com/aglt/appViews/app#/pending-tasks"
    webbrowser.open(url)


def cleanup_output(base_dir: Path) -> int:
    """Move .xlsx/.csv files in Output/ to recycle bin."""
    output_dir = base_dir / "Output"
    if not output_dir.exists():
        return 0

    files = list(output_dir.glob("*.xlsx")) + list(output_dir.glob("*.csv"))
    if not files:
        return 0

    # Headless: auto-clean without confirmation
    if not _is_headless():
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        confirm = messagebox.askyesno(
            "Clean Output",
            "Please confirm event upload is done.\n\n"
            f"Move {len(files)} file(s) to recycle bin:\n"
            + "\n".join(f"  * {f.name}" for f in files[:10])
            + ("\n  ..." if len(files) > 10 else ""),
        )
        root.destroy()
        if not confirm:
            return 0

    cleaned = 0
    for f in files:
        if _send_to_recycle_bin(str(f)):
            cleaned += 1
        else:
            archive = output_dir / "archive" / datetime.now().strftime("%Y%m%d")
            archive.mkdir(parents=True, exist_ok=True)
            try:
                f.rename(archive / f.name)
                cleaned += 1
            except Exception:
                pass

    return cleaned


def _send_to_recycle_bin(file_path: str) -> bool:
    """Send file to Windows recycle bin via Shell API."""
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
