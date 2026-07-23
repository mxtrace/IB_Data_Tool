"""
startup_gui.py — 启动配置界面
三种运行范围：self / team / custom
"""
from __future__ import annotations

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
from pathlib import Path


def _is_weekend() -> bool:
    return date.today().weekday() >= 5


def show_startup_dialog(config_path: Path) -> dict | None:
    """
    显示启动配置窗口。
    返回完整配置 dict 或 None（取消）。
    """
    # 加载 roles 配置
    roles_path = config_path.parent / "roles.json"
    roles = _load_roles(roles_path)

    # 加载上次运行配置
    defaults = {}
    if config_path.exists():
        try:
            defaults = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    result = {"cancelled": True}

    # ── 窗口 ──
    root = tk.Tk()
    root.title("IB Data Tool")
    root.geometry("440x600")
    root.resizable(False, True)

    root.update_idletasks()
    x = (root.winfo_screenwidth() - 440) // 2
    y = (root.winfo_screenheight() - 600) // 2
    root.geometry(f"+{x}+{y}")

    # ── Actor 显示 ──
    frame_actor = ttk.Frame(root, padding=(12, 10, 12, 0))
    frame_actor.pack(fill="x")
    ttk.Label(frame_actor, text=f"执行人：{roles['actor']}",
              font=("", 10, "bold")).pack(anchor="w")

    # ── Scope 选择 ──
    frame_scope = ttk.LabelFrame(root, text="运行范围", padding=10)
    frame_scope.pack(fill="x", padx=12, pady=(8, 5))

    # 默认 scope：周末→team，工作日→self
    default_scope = "team" if _is_weekend() else "self"
    scope_var = tk.StringVar(value=defaults.get("scope", default_scope))

    scope_options = [
        ("self", f"仅自己（{', '.join(roles['self_logins'])}）", "weekday_default"),
        ("team", f"全组（{len(roles['team_logins'])} 人）", "weekend_overtime"),
        ("custom", "自定义选择", "manual_override"),
    ]

    for val, text, _ in scope_options:
        ttk.Radiobutton(frame_scope, text=text, variable=scope_var, value=val,
                        command=lambda: _update_custom_state()).pack(anchor="w", pady=2)

    # ── Custom 多选列表 ──
    frame_custom = ttk.LabelFrame(root, text="自定义选择（仅 custom 模式生效）", padding=8)
    frame_custom.pack(fill="both", expand=True, padx=12, pady=5)

    canvas = tk.Canvas(frame_custom, highlightthickness=0, height=100)
    scrollbar = ttk.Scrollbar(frame_custom, orient="vertical", command=canvas.yview)
    inner_frame = ttk.Frame(canvas)
    inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    # 生成 checkbox
    prev_custom = set(defaults.get("selected_logins", []))
    check_vars = {}
    for login in roles["team_logins"]:
        var = tk.BooleanVar(value=(login in prev_custom))
        check_vars[login] = var
        ttk.Checkbutton(inner_frame, text=login, variable=var).pack(anchor="w", pady=1)

    def _update_custom_state():
        state = "normal" if scope_var.get() == "custom" else "disabled"
        for child in inner_frame.winfo_children():
            child.configure(state=state)

    _update_custom_state()

    # ── Login 管理按钮 ──
    frame_login_mgmt = ttk.Frame(root, padding=(12, 0))
    frame_login_mgmt.pack(fill="x")

    def _add_login():
        from tkinter import simpledialog
        new_login = simpledialog.askstring("添加 Login", "输入新 Login：", parent=root)
        if not new_login or not new_login.strip():
            return
        new_login = new_login.strip().lower()
        if new_login in check_vars:
            messagebox.showinfo("提示", f"{new_login} 已存在")
            return
        roles["team_logins"].append(new_login)
        var = tk.BooleanVar(value=True)
        check_vars[new_login] = var
        ttk.Checkbutton(inner_frame, text=new_login, variable=var).pack(anchor="w", pady=1)
        _save_roles(roles_path, roles)
        _update_custom_state()

    def _del_login():
        to_del = [l for l, v in check_vars.items() if v.get()]
        if not to_del:
            messagebox.showinfo("提示", "请先勾选要删除的 Login")
            return
        if not messagebox.askyesno("确认", f"删除 {len(to_del)} 个 Login？\n{', '.join(to_del)}"):
            return
        for login in to_del:
            if login in roles["team_logins"]:
                roles["team_logins"].remove(login)
            del check_vars[login]
        for child in inner_frame.winfo_children():
            child.destroy()
        for login in roles["team_logins"]:
            if login not in check_vars:
                check_vars[login] = tk.BooleanVar(value=False)
            ttk.Checkbutton(inner_frame, text=login, variable=check_vars[login]).pack(anchor="w", pady=1)
        _save_roles(roles_path, roles)
        _update_custom_state()

    ttk.Button(frame_login_mgmt, text="+ 添加", command=_add_login, width=8).pack(side="left", padx=(0, 5))
    ttk.Button(frame_login_mgmt, text="- 删除选中", command=_del_login, width=10).pack(side="left")

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
        scope = scope_var.get()
        reason_map = {"self": "weekday_default", "team": "weekend_overtime", "custom": "manual_override"}

        if scope == "self":
            selected = roles["self_logins"]
        elif scope == "team":
            selected = roles["team_logins"]
        else:  # custom
            selected = [l for l, v in check_vars.items() if v.get()]
            if not selected:
                messagebox.showwarning("提示", "请至少选择一个 Login")
                return

        config_data = {
            "actor": roles["actor"],
            "scope": scope,
            "reason": reason_map[scope],
            "selected_logins": selected,
            "batch_size": batch_var.get(),
            "search_stores": [s for s, v in store_vars.items() if v.get()],
            "shared_drive_path": roles.get("shared_drive_path",
                r"\\ant.amazon.com\dept-as\sha11\ILS\LCL_INBOUND_DATA_ETL\IBDATACONFIRM\DATA"),
        }
        config_path.write_text(json.dumps(config_data, ensure_ascii=False, indent=2), encoding="utf-8")

        result["cancelled"] = False
        result["config"] = config_data
        root.destroy()

    def on_cancel():
        root.destroy()

    ttk.Button(btn_frame, text="▶ 开始运行", command=on_start, width=15).pack(side="left", padx=(0, 10))
    ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side="left")

    # 提示
    hint = "💡 周末/加班自动选择 team 模式" if _is_weekend() else "💡 工作日默认 self 模式"
    ttk.Label(btn_frame, text=hint, foreground="gray").pack(side="right")

    root.bind("<Return>", lambda e: on_start())
    root.bind("<Escape>", lambda e: on_cancel())
    root.mainloop()

    if result["cancelled"]:
        return None
    return result["config"]


def _load_roles(roles_path: Path) -> dict:
    """
    加载 roles.json 并自动识别当前 Windows 用户作为 actor。
    """
    current_user = os.environ.get("USERNAME", os.getlogin()).lower()
    defaults = {
        "actor": current_user,
        "self_logins": [current_user],
        "team_logins": [],
        "shared_drive_path": r"\\ant.amazon.com\dept-as\sha11\ILS\LCL_INBOUND_DATA_ETL\IBDATACONFIRM\DATA",
    }
    if roles_path.exists():
        try:
            data = json.loads(roles_path.read_text(encoding="utf-8"))
            # team_logins 和 shared_drive_path 从文件读取
            defaults["team_logins"] = data.get("team_logins", [])
            defaults["shared_drive_path"] = data.get("shared_drive_path", defaults["shared_drive_path"])
            # actor 和 self_logins 始终用当前系统用户
        except Exception:
            pass
    return defaults


def _save_roles(roles_path: Path, roles: dict):
    """保存 roles.json（仅写 team_logins 和 shared_drive_path）"""
    data = {
        "team_logins": roles["team_logins"],
        "search_stores": [s for s, v in store_vars.items() if v.get()],
            "shared_drive_path": roles.get("shared_drive_path", ""),
    }
    roles_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
