"""
progress_window.py — 处理进度窗口
显示当前批次进度 N/M + AL0 编号，主线程手动刷新。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class ProgressWindow:
    """非模态进度窗口，含 Input Zone 显示"""

    def __init__(self, total: int):
        self._root = tk.Toplevel() if _has_tk_root() else tk.Tk()
        self._root.title("IB Data Tool — 处理进度")
        self._root.geometry("500x360")
        self._root.resizable(True, True)
        self._root.attributes("-topmost", True)

        self._root.update_idletasks()
        x = (self._root.winfo_screenwidth() - 500) // 2
        y = (self._root.winfo_screenheight() - 360) // 2 - 50
        self._root.geometry(f"+{x}+{y}")

        # 进度区
        frame_top = ttk.Frame(self._root, padding=(15, 10, 15, 5))
        frame_top.pack(fill="x")

        self._label = ttk.Label(frame_top, text="准备中...", font=("", 11, "bold"))
        self._label.pack(anchor="w")

        self._progress = ttk.Progressbar(frame_top, maximum=total, length=460)
        self._progress.pack(fill="x", pady=(8, 4))

        self._detail = ttk.Label(frame_top, text="", foreground="gray")
        self._detail.pack(anchor="w")

        # Input Zone 区
        frame_iz = ttk.LabelFrame(self._root, text="Input Zone", padding=5)
        frame_iz.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        self._text = tk.Text(frame_iz, height=10, wrap="word", font=("Consolas", 9),
                             state="disabled", bg="#f8f8f8")
        iz_scroll = ttk.Scrollbar(frame_iz, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=iz_scroll.set)
        self._text.pack(side="left", fill="both", expand=True)
        iz_scroll.pack(side="right", fill="y")

        self._total = total
        self._root.update()

    def update(self, idx: int, al0: str, status: str = ""):
        """更新进度（主线程调用）"""
        self._progress["value"] = idx + 1
        self._label.config(text=f"[{idx + 1}/{self._total}] {al0}")
        if status:
            self._detail.config(text=status)
        self._root.update()

    def set_input_zone(self, text: str):
        """显示 Input Zone 内容"""
        self._text.config(state="normal")
        self._text.delete("1.0", "end")
        # 只显示关键信息（Parties 段），避免太长
        self._text.insert("1.0", text[:2000])
        self._text.config(state="disabled")
        self._root.update()

    def finish(self, msg: str = "批次完成"):
        self._label.config(text=msg)
        self._progress["value"] = self._total
        self._root.update()

    def close(self):
        try:
            self._root.destroy()
        except Exception:
            pass


def _has_tk_root() -> bool:
    """检查是否已有 Tk root"""
    try:
        return tk._default_root is not None
    except Exception:
        return False
