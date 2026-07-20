"""
IB Data Tool 完整流程测试
模拟实际运行：GUI启动 → Step1~6
"""
import sys
sys.path.insert(0, "src")

# 验证所有模块可导入
print("=" * 50)
print("IB Data Tool — 导入检查")
print("=" * 50)

modules = [
    ("core.config", "AppConfig, load_config"),
    ("core.startup_gui", "show_startup_dialog"),
    ("core.logger", "init_logger, log_info"),
    ("core.preflight", "run_preflight"),
    ("core.wal", "append_wal, recover_pending_list_from_wal"),
    ("core.batch_controller", "BatchController, TicketResult"),
    ("core.browser_manager", "BrowserManager"),
    ("core.progress_window", "ProgressWindow"),
    ("steps.step1_sync", "sync_to_pending_list, load_batch"),
    ("steps.step2_oc_scrape", "scrape_booking_summary"),
    ("steps.step3_template_select", "select_template"),
    ("steps.step4_fill_template", "fill_template"),
    ("steps.step5_email", "generate_email"),
    ("steps.step6_event", "generate_event_csv, open_pending_tasks"),
]

all_ok = True
for mod, names in modules:
    try:
        __import__(mod)
        print(f"  ✓ {mod}")
    except Exception as e:
        print(f"  ✗ {mod} — {e}")
        all_ok = False

if not all_ok:
    print("\n❌ 导入失败，无法运行")
    sys.exit(1)

print("\n✅ 所有模块导入成功")
print("\n启动 GUI...")

# 实际启动
from main import main
main()
