"""
startup_gui.py — 启动配置界面（简化版）
仅选择搜索邮箱 + 批次大小。
"""
from __future__ import annotations

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path


def show_startup_dialog(config_path: Path) -> dict | None:
    """显示启动配置窗口。返回配置 dict 或 None（取消）。"""
    defaults = {}
    if config_path.exists():
        try:
            defaults = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    current_user = os.environ.get("USERNAME", os.getlogin()).lower()
    result = {"cancelled": True}

    # ── 窗口 ──
    root = tk.Tk()
    root.title("IB Data Tool")
    root.geometry("440x350")
    root.resizable(False, True)

    root.update_idletasks()
    x = (root.winfo_screenwidth() - 440) // 2
    y = (root.winfo_screenheight() - 350) // 2
    root.geometry(f"+{x}+{y}")

    # ── Actor 显示 ──
    frame_actor = ttk.Frame(root, padding=(12, 10, 12, 0))
    frame_actor.pack(fill="x")
    ttk.Label(frame_actor, text=f"执行人：{current_user}",
              font=("", 10, "bold")).pack(anchor="w")

    # ── 搜索邮箱选择 ──
    frame_stores = ttk.LabelFrame(root, text="搜索邮箱（用于提取收件人）", padding=8)
    frame_stores.pack(fill="x", padx=12, pady=5)

    from core.outlook_helper import list_outlook_stores
    available_stores = list_outlook_stores()
    prev_stores = set(defaults.get("search_stores", []))
    store_vars = {}
    if available_stores:
        for store_name in available_stores:
            var = tk.BooleanVar(value=(store_name in prev_stores) if prev_stores else True)
            store_vars[store_name] = var
            ttk.Checkbutton(frame_stores, text=store_name, variable=var).pack(anchor="w", pady=1)
    else:
        ttk.Label(frame_stores, text="（未检测到邮箱账号）", foreground="gray").pack(anchor="w")

    # ── 批次大小 ──
    frame_batch = ttk.Frame(root, padding=(12, 5))
    frame_batch.pack(fill="x")
    ttk.Label(frame_batch, text="每批处理数量：").pack(side="left")
    batch_var = tk.IntVar(value=defaults.get("batch_size", 10))
    ttk.Spinbox(frame_batch, from_=1, to=50, textvariable=batch_var, width=5).pack(side="left")

    # ── 按钮 ──
    btn_frame = ttk.Frame(root, padding=(12, 10))
    btn_frame.pack(fill="x")

    def on_start():
        config_data = {
            "actor": current_user,
            "batch_size": batch_var.get(),
            "search_stores": [s for s, v in store_vars.items() if v.get()],
        }
        config_path.write_text(json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8")

        result["cancelled"] = False
        result["config"] = config_data
        root.destroy()

    def on_cancel():
        root.destroy()

    ttk.Button(btn_frame, text="▶ 开始运行", command=on_start, width=15).pack(side="left", padx=(0, 10))
    ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side="left")

    root.bind("<Return>", lambda e: on_start())
    root.bind("<Escape>", lambda e: on_cancel())
    root.mainloop()

    if result["cancelled"]:
        return None
    return result["config"]
