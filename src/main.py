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
from steps.step1_sync import load_batch, append_to_history
from core.oc_api_client import build_session, check_session_valid
from steps.step2_oc_scrape import scrape_booking_summary
from steps.step3_template_select import select_template
from steps.step4_fill_template import fill_template
from steps.step5_email import generate_email
from steps.step6_event import generate_event_csv, open_pending_tasks, cleanup_output

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
    log_info(f"IB Data Tool 启动，batch_size={config.batch_size}")

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
    # Step 1: 直接从源 PendingList 读取（load_batch 内完成）
    # ═══════════════════════════════════════════════════════════════
    log_info("Step1: 将直接从源 PendingList 读取数据")

    # ═══════════════════════════════════════════════════════════════
    # 建立 OC API Session（替代浏览器抓取，仅 CDA 订单才启动浏览器）
    # ═══════════════════════════════════════════════════════════════
    try:
        oc_session = build_session()
        if not check_session_valid(oc_session):
            show_blocker("OC Cookie 已过期，请先在 Firefox 登录 trans-logistics-cn.amazon.com 后重试")
            return
        log_info("OC API Session 已建立")
    except Exception as e:
        log_error(f"OC Session 建立失败：{e}")
        show_blocker(f"OC Session 建立失败：{e}")
        return

    # 浏览器仅 CDA 订单 ASI 下载时按需启动
    browser = None

    # ═══════════════════════════════════════════════════════════════
    # 清理 Output 旧文件（新文件生成前自动执行）
    # ═══════════════════════════════════════════════════════════════
    cleaned = cleanup_output(BASE_DIR)
    if cleaned:
        log_info(f"已清理 Output 中 {cleaned} 个旧文件到回收站")

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

            # ═══════════════════════════════════════════════════════
            # Phase 1: 全量 OC 查询（收集 input_zone）
            # ═══════════════════════════════════════════════════════
            scrape_results = {}  # al0 -> scrape_result
            phase1_skipped = {}  # al0 -> reason

            for idx, al0 in enumerate(batch.al0_list):
                batch_ctrl.update_progress(idx, al0)
                ib_row = batch.get_row(al0)
                if not ib_row:
                    phase1_skipped[al0] = "Pending List 中无此 AL0 数据"
                    continue

                scrape_result = scrape_booking_summary(al0, oc_session)
                if scrape_result.error:
                    audit(al0, "step2", "skipped", scrape_result.error)
                    phase1_skipped[al0] = scrape_result.error
                    continue

                scrape_results[al0] = scrape_result
                input_zone = scrape_result.input_zone
                audit(al0, "step2", "success", f"ODM={input_zone.odm_booking},CDA={input_zone.cda_booking}")

            log_info(f"Phase1 完成：查询成功={len(scrape_results)}，跳过={len(phase1_skipped)}")

            # ═══════════════════════════════════════════════════════
            # Phase 2: 批量模板填充
            # ═══════════════════════════════════════════════════════
            fill_results = {}  # al0 -> fill_result
            phase2_odm = []
            phase2_cda = []
            phase2_skipped = {}

            for al0, scrape_result in scrape_results.items():
                ib_row = batch.get_row(al0)
                input_zone = scrape_result.input_zone

                if input_zone.odm_booking:
                    audit(al0, "step2", "odm", "")
                    phase2_odm.append(al0)
                    continue

                if input_zone.cda_booking:
                    audit(al0, "step2", "cda", "")
                    phase2_cda.append(al0)
                    continue

                # 普通订单：选模板
                template_type = select_template(ib_row.get("pod", ""))
                template_file = None

                fill_result = fill_template(
                    al0=al0,
                    template_type=template_type,
                    ib_row=ib_row,
                    input_zone=input_zone,
                    base_dir=BASE_DIR,
                    asi_file=template_file,
                )
                if fill_result.error:
                    audit(al0, "step4", "skipped", fill_result.error)
                    phase2_skipped[al0] = fill_result.error
                    continue

                fill_results[al0] = fill_result

            log_info(f"Phase2 完成：填充成功={len(fill_results)}，ODM={len(phase2_odm)}，跳过={len(phase2_skipped)}")

            # ═══════════════════════════════════════════════════════
            # Phase 3: 连续弹出邮件
            # ═══════════════════════════════════════════════════════
            for al0, fill_result in fill_results.items():
                ib_row = batch.get_row(al0)
                input_zone = scrape_results[al0].input_zone

                email_result = generate_email(
                    al0=al0,
                    ib_row=ib_row,
                    input_zone=input_zone,
                    attachment_path=fill_result.output_file,
                    config=config,
                    base_dir=BASE_DIR,
                )
                bc_login = ib_row.get("bc_login", "")
                if email_result.error:
                    audit(al0, "step5", "skipped", email_result.error)
                    result = TicketResult(al0, "skipped", email_result.error, bc_login=bc_login)
                else:
                    audit(al0, "step5", "success", "邮件已弹出")
                    result = TicketResult(al0, "success", "", bc_login=bc_login)

                batch_ctrl.record_result(result)
                append_wal(BASE_DIR, al0, result.status, result.reason)
                if result.status == "success":
                    append_to_history(BASE_DIR, al0, result.status)

            # 记录 Phase1/2 跳过和 ODM
            for al0, reason in phase1_skipped.items():
                bc_login = (batch.get_row(al0) or {}).get("bc_login", "")
                result = TicketResult(al0, "skipped", reason, bc_login=bc_login)
                batch_ctrl.record_result(result)
                append_wal(BASE_DIR, al0, "skipped", reason)
                _show_skip_warning(al0, reason)

            for al0 in phase2_odm:
                bc_login = (batch.get_row(al0) or {}).get("bc_login", "")
                result = TicketResult(al0, "odm", "", bc_login=bc_login)
                batch_ctrl.record_result(result)
                append_wal(BASE_DIR, al0, "odm", "")
                append_to_history(BASE_DIR, al0, "odm")

            for al0 in phase2_cda:
                bc_login = (batch.get_row(al0) or {}).get("bc_login", "")
                result = TicketResult(al0, "cda", "", bc_login=bc_login)
                batch_ctrl.record_result(result)
                append_to_history(BASE_DIR, al0, "cda")

            for al0, reason in phase2_skipped.items():
                bc_login = (batch.get_row(al0) or {}).get("bc_login", "")
                result = TicketResult(al0, "skipped", reason, bc_login=bc_login)
                batch_ctrl.record_result(result)
                append_wal(BASE_DIR, al0, "skipped", reason)
                _show_skip_warning(al0, reason)

            gc.collect()

            # 批次完成
            batch_ctrl.finish_progress()
            batch_ctrl.write_back_pending_list(BASE_DIR)
            clear_wal(BASE_DIR)
            all_event_records.extend(batch_ctrl.get_event_records())

            log_info(f"批次完成：成功={len(batch_ctrl.get_success_records())}，"
                     f" 跳过={len([r for r in batch_ctrl.results if r.status == 'skipped'])}，"
                     f" ODM={len([r for r in batch_ctrl.results if r.status == 'odm'])}")

            action = batch_ctrl.show_summary_dialog()
            if action == "finish":
                break

    finally:
        if browser:
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
    import traceback
    try:
        main()
    except Exception as e:
        try:
            from core.logger import log_error
            log_error(f"未捕获异常：{e}\n{traceback.format_exc()}")
        except Exception:
            pass
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("IB Data Tool 崩溃", f"{e}\n\n详情见 logs/ 目录")
        root.destroy()





