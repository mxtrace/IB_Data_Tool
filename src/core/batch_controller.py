"""
batch_controller.py — 批次控制
管理批次进度、记录结果、回写 Pending List、汇总弹窗。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class TicketResult:
    al0: str
    status: str   # "success" | "skipped" | "odm"
    reason: str   # 失败原因（skipped时有值）
    bc_login: str = ""  # 该票对应的BC Login


@dataclass
class BatchData:
    """一个批次的数据"""
    al0_list: list[str]
    rows: dict  # {al0: {field: value}}  Pending List 行数据
    row_indices: dict  # {al0: excel_row_number}  用于回写
    total_pending: int = 0  # Pending List 中总待处理数

    def get_row(self, al0: str) -> dict:
        return self.rows.get(al0, {})

    def write_odm_flag(self, al0: str):
        """标记 S 列为 ODM（实际写入在批次结束时统一执行）"""
        self._odm_flags.add(al0)

    _odm_flags: set = field(default_factory=set, init=False, repr=False)


class BatchController:
    """批次状态管理器"""

    def __init__(self, batch_size: int):
        self.batch_size = batch_size
        self.results: list[TicketResult] = []
        self._current_batch: BatchData | None = None
        self._progress_win = None
        # 累计统计（跨批次）
        self.total_processed: int = 0
        self.total_pending: int = 0

    def start_batch(self, batch: BatchData):
        self._current_batch = batch
        self.results = []
        # 创建进度窗口
        from core.progress_window import ProgressWindow
        try:
            self._progress_win = ProgressWindow(len(batch.al0_list))
        except Exception:
            self._progress_win = None

    def set_total_pending(self, total: int):
        """设置 Pending List 中总待处理数（首次调用时）"""
        if self.total_pending == 0:
            self.total_pending = total

    def update_progress(self, idx: int, al0: str):
        """更新进度显示"""
        if self._progress_win:
            self._progress_win.update(idx, al0)

    def set_input_zone(self, text: str):
        """显示 Input Zone 内容"""
        if self._progress_win:
            self._progress_win.set_input_zone(text)

    def finish_progress(self):
        """关闭进度窗口"""
        if self._progress_win:
            self._progress_win.close()
            self._progress_win = None

    def record_result(self, result: TicketResult):
        self.results.append(result)

    def get_success_records(self) -> list[TicketResult]:
        return [r for r in self.results if r.status == "success"]

    def write_back_pending_list(self, base_dir: Path):
        """
        批次完成后统一回写 Pending List：
        - success → T列 = "Yes"
        - odm → S列 = "ODM", T列 = "Yes"
        - skipped → 不写（T列保持空）
        """
        from steps.step1_sync import write_back_results
        write_back_results(self._current_batch, self.results, base_dir)

    def show_summary_dialog(self) -> str:
        """
        显示汇总弹窗，包含失败明细。
        返回 "continue" 或 "finish"。
        """
        import tkinter as tk
        from tkinter import ttk

        success = [r for r in self.results if r.status == "success"]
        odm = [r for r in self.results if r.status == "odm"]
        skipped = [r for r in self.results if r.status == "skipped"]

        result = {"action": "finish"}

        root = tk.Tk()
        root.title("批次完成")
        root.geometry("420x320")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        root.update_idletasks()
        x = (root.winfo_screenwidth() - 420) // 2
        y = (root.winfo_screenheight() - 320) // 2
        root.geometry(f"+{x}+{y}")

        # 累计统计
        self.total_processed += len(self.results)

        # 统计区
        frame_stats = ttk.Frame(root, padding=15)
        frame_stats.pack(fill="x")
        ttk.Label(frame_stats, text=f"已完成 {self.total_processed} 票 / 共 {self.total_pending} 票",
                  font=("", 12, "bold")).pack(anchor="w")
        ttk.Separator(frame_stats, orient="horizontal").pack(fill="x", pady=5)
        ttk.Label(frame_stats, text=f"本批：✅ 成功 {len(success)} | ⏭️ ODM {len(odm)} | ❌ 失败 {len(skipped)}").pack(anchor="w", pady=2)

        # 失败明细
        if skipped:
            frame_detail = ttk.LabelFrame(root, text="失败明细", padding=8)
            frame_detail.pack(fill="both", expand=True, padx=15, pady=5)
            text_widget = tk.Text(frame_detail, height=6, wrap="word", font=("", 9))
            text_widget.pack(fill="both", expand=True)
            for r in skipped:
                text_widget.insert("end", f"{r.al0} — {r.reason}\n")
            text_widget.config(state="disabled")

        # 按钮
        btn_frame = ttk.Frame(root, padding=(15, 10))
        btn_frame.pack(fill="x")

        def on_continue():
            result["action"] = "continue"
            root.destroy()

        def on_finish():
            result["action"] = "finish"
            root.destroy()

        ttk.Button(btn_frame, text="▶ 继续下一批", command=on_continue, width=15).pack(side="left", padx=(0, 10))
        ttk.Button(btn_frame, text="✔ 全部完成", command=on_finish, width=15).pack(side="left")

        root.protocol("WM_DELETE_CLOSE", on_finish)
        root.mainloop()
        return result["action"]
