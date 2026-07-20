"""
IB Data Batch Sending Tool — 主入口
流程：Step0(配置) → 自检 → WAL恢复 → Step1(数据同步) → Per-AL0循环(Step2~5) → Step6(打卡)
"""
from __future__ import annotations

import gc
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from core.config import load_config, AppConfig, ConfigError
from core.startup_gui import show_startup_dialog
from core.logger import init_logger, log_info, log_error, audit
from core.preflight import run_preflight
from core.wal import append_wal, recover_pending_list_from_wal, clear_wal
from core.batch_controller import BatchController, TicketResult
from core.browser_manager import BrowserManager
from steps.step1_sync import sync_to_pending_list, load_batch, append_to_history
from steps.step2_oc_scrape import scrape_booking_summary
from steps.step3_template_select import select_template
from steps.step4_fill_template import fill_template
from steps.step5_email import generate_email
from steps.step6_event import generate_event_csv, open_pending_tasks

# exe 模式：exe 所在目录；开发模式：src 的上级目录
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent.parent


def main():
    # ═══════════════════════════════════════════════════════════════
    # Step 0: 启动配置（GUI）
    # ═══════════════════════════════════════════════════════════════
    config_path = BASE_DIR / "config.json"
    gui_result = show_startup_dialog(config_path)
    if gui_result is None:
        return  # 用户取消

    try:
        config = load_config(config_path)
    except ConfigError as e:
        show_blocker(str(e))
        return

    # 初始化日志
    init_logger(BASE_DIR)
    log_info("=" * 50)
    log_info(f"IB Data Tool 启动，Logins={config.selected_logins}, batch_size={config.batch_size}")

    # ═══════════════════════════════════════════════════════════════
    # 启动前自检
    # ═══════════════════════════════════════════════════════════════
    issues = run_preflight(BASE_DIR, config)
    if issues:
        msg = "启动自检未通过：\n" + "\n".join(f"  • {i}" for i in issues)
        log_error(msg)
        show_blocker(msg)
        return

    # ═══════════════════════════════════════════════════════════════
    # WAL 崩溃恢复
    # ═══════════════════════════════════════════════════════════════
    recover_pending_list_from_wal(BASE_DIR)

    # ═══════════════════════════════════════════════════════════════
    # Step 1: 读取进仓数据 → 同步到 Pending List
    # ═══════════════════════════════════════════════════════════════
    try:
        sync_result = sync_to_pending_list(config, BASE_DIR)
        log_info(f"Step1 完成：来源={sync_result['source']}，追加={sync_result['appended']}行")
    except Exception as e:
        log_error(f"Step1 失败：{e}")
        show_blocker(f"Step 1 数据同步失败：{e}")
        return

    # ═══════════════════════════════════════════════════════════════
    # 启动浏览器
    # ═══════════════════════════════════════════════════════════════
    browser = BrowserManager()
    try:
        browser.start()
        browser.warmup()  # 自动导航到 OC 首页，确保登录态
        log_info("浏览器已启动")
    except Exception as e:
        log_error(f"浏览器启动失败：{e}")
        show_blocker(f"浏览器启动失败：{e}")
        return

    # ═══════════════════════════════════════════════════════════════
    # 批次循环
    # ═══════════════════════════════════════════════════════════════
    batch_ctrl = BatchController(config.batch_size)
    all_event_records = []

    try:
        while True:
            batch = load_batch(config, BASE_DIR)
            if not batch.al0_list:
                log_info("无待处理 AL0，流程结束")
                break

            log_info(f"开始新批次：{len(batch.al0_list)} 票")
            batch_ctrl.set_total_pending(batch.total_pending)
            batch_ctrl.start_batch(batch)

            # Per-AL0 循环
            for idx, al0 in enumerate(batch.al0_list):
                batch_ctrl.update_progress(idx, al0)
                result = process_single_al0(al0, batch, config, BASE_DIR, browser, batch_ctrl)
                batch_ctrl.record_result(result)

                # 🟡 跳过时弹窗提示 BC 确认
                if result.status == "skipped":
                    _show_skip_warning(al0, result.reason)

                # P0: 立即写 WAL（防崩溃丢失）
                append_wal(BASE_DIR, al0, result.status, result.reason)
                # 写入 history 防重复
                if result.status in ("success", "odm"):
                    append_to_history(BASE_DIR, al0, result.status)

                # 释放 COM 引用
                gc.collect()

            # 批次完成：关闭进度窗口 + 回写 Pending List
            batch_ctrl.finish_progress()
            batch_ctrl.write_back_pending_list(BASE_DIR)
            clear_wal(BASE_DIR)  # 回写成功后清除 WAL
            all_event_records.extend(batch_ctrl.get_success_records())

            log_info(f"批次完成：成功={len(batch_ctrl.get_success_records())},"
                     f" 跳过={len([r for r in batch_ctrl.results if r.status == 'skipped'])},"
                     f" ODM={len([r for r in batch_ctrl.results if r.status == 'odm'])}")

            # 汇总弹窗 — 等待BC手动点击
            action = batch_ctrl.show_summary_dialog()
            if action == "finish":
                break

            # 批次间重置：下一批首单强制完整导航（防 SPA 页面过期）
            browser._oc_warmed = False
            browser._devtools_open = False

    finally:
        browser.close()
        log_info("浏览器已关闭")

    # ═══════════════════════════════════════════════════════════════
    # Step 6: 打卡
    # ═══════════════════════════════════════════════════════════════
    if all_event_records:
        csv_path = generate_event_csv(all_event_records, BASE_DIR)
        log_info(f"Step6: Event CSV 已生成 → {csv_path}")
        open_pending_tasks()

    log_info("IB Data Tool 运行完毕")


def process_single_al0(al0, batch, config, base_dir, browser, batch_ctrl) -> TicketResult:
    """单票 AL0 的完整处理流程（Step 2~5）"""
    ib_row = batch.get_row(al0)
    bc_login = ib_row.get("bc_login", "") if ib_row else ""
    if not ib_row:
        return TicketResult(al0, "skipped", "Pending List 中无此 AL0 数据")

    # ── Step 2: OC 抓取 + 类型判断 ──
    scrape_result = scrape_booking_summary(al0, browser)
    if scrape_result.error:
        audit(al0, "step2", "skipped", scrape_result.error)
        return TicketResult(al0, "skipped", scrape_result.error, bc_login=bc_login)

    input_zone = scrape_result.input_zone
    audit(al0, "step2", "success", f"ODM={input_zone.odm_booking},CDA={input_zone.cda_booking}")

    # 显示 Input Zone 到进度窗口
    iz_summary = (
        f"Shipper: {input_zone.shipper.company}\n"
        f"Consignee: {input_zone.consignee.company}\n"
        f"ODM: {input_zone.odm_booking} | CDA: {input_zone.cda_booking}\n"
        f"Email: {input_zone.shipper.email}"
    )
    batch_ctrl.set_input_zone(iz_summary)

    # ODM → 跳过
    if input_zone.odm_booking:
        audit(al0, "step2", "odm", "")
        return TicketResult(al0, "odm", "", bc_login=bc_login)

    # CDA → 下载ASI
    if input_zone.cda_booking:
        asi_path = scrape_result.download_asi(browser, base_dir)
        if not asi_path:
            audit(al0, "step2", "skipped", "ASI下载失败")
            return TicketResult(al0, "skipped", "ASI 文件下载失败", bc_login=bc_login)
        template_type = "ASI"
        template_file = asi_path
    else:
        # ── Step 3: 模板选择 ──
        template_type = select_template(ib_row.get("pod", ""))
        template_file = None

    # ── Step 4: 填充模板 ──
    fill_result = fill_template(
        al0=al0,
        template_type=template_type,
        ib_row=ib_row,
        input_zone=input_zone,
        base_dir=base_dir,
        asi_file=template_file,
    )
    if fill_result.error:
        audit(al0, "step4", "skipped", fill_result.error)
        return TicketResult(al0, "skipped", fill_result.error, bc_login=bc_login)

    # ── Step 5: 生成邮件 ──
    email_result = generate_email(
        al0=al0,
        ib_row=ib_row,
        input_zone=input_zone,
        attachment_path=fill_result.output_file,
        config=config,
        base_dir=base_dir,
    )
    if email_result.error:
        audit(al0, "step5", "skipped", email_result.error)
        return TicketResult(al0, "skipped", email_result.error, bc_login=bc_login)

    audit(al0, "step5", "success", "邮件已弹出")
    return TicketResult(al0, "success", "", bc_login=bc_login)


def show_blocker(msg: str):
    """🔴 阻断弹窗"""
    log_error(f"[BLOCKER] {msg}")
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("IB Data Tool — 阻断", msg)
    root.destroy()


def _show_skip_warning(al0: str, reason: str):
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning("跳过提示", f"[{al0}] 已跳过\n\n原因：{reason}")
    root.destroy()


if __name__ == "__main__":
    main()





