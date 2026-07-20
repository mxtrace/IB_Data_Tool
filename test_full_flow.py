"""
test_full_flow.py - Full pipeline test (Step 1-5)
Bypasses preflight, runs the complete batch loop with real Firefox + Outlook.
"""
from __future__ import annotations
import gc, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pathlib import Path
from core.config import load_config, ConfigError
from core.logger import init_logger, log_info, log_error, audit
from core.wal import append_wal, recover_pending_list_from_wal, clear_wal
from core.batch_controller import BatchController, TicketResult
from core.browser_manager import BrowserManager
from steps.step1_sync import sync_to_pending_list, load_batch, append_to_history
from steps.step2_oc_scrape import scrape_booking_summary
from steps.step3_template_select import select_template
from steps.step4_fill_template import fill_template
from steps.step5_email import generate_email

BASE_DIR = Path(__file__).parent


def process_single_al0(al0, batch, config, base_dir, browser):
    ib_row = batch.get_row(al0)
    if not ib_row:
        return TicketResult(al0, 'skipped', 'Pending List no data')

    # Step 2
    scrape_result = scrape_booking_summary(al0, browser)
    if scrape_result.error:
        audit(al0, 'step2', 'skipped', scrape_result.error)
        return TicketResult(al0, 'skipped', scrape_result.error)

    input_zone = scrape_result.input_zone
    audit(al0, 'step2', 'success', f'ODM={input_zone.odm_booking},CDA={input_zone.cda_booking}')

    if input_zone.odm_booking:
        return TicketResult(al0, 'odm', '')

    # CDA -> skip ASI for now
    if input_zone.cda_booking:
        return TicketResult(al0, 'skipped', 'CDA order - ASI download not yet implemented')

    # Step 3
    template_type = select_template(ib_row.get('pod', ''))

    # Step 4
    fill_result = fill_template(
        al0=al0,
        template_type=template_type,
        ib_row=ib_row,
        input_zone=input_zone,
        base_dir=base_dir,
        asi_file=None,
    )
    if fill_result.error:
        audit(al0, 'step4', 'skipped', fill_result.error)
        return TicketResult(al0, 'skipped', fill_result.error)

    # Step 5
    email_result = generate_email(
        al0=al0,
        ib_row=ib_row,
        input_zone=input_zone,
        attachment_path=fill_result.output_file,
        config=config,
        base_dir=base_dir,
    )
    if email_result.error:
        audit(al0, 'step5', 'skipped', email_result.error)
        return TicketResult(al0, 'skipped', email_result.error)

    audit(al0, 'step5', 'success', '')
    return TicketResult(al0, 'success', '')


def main():
    print(f'\n{"="*60}')
    print(f'  IB Data Tool - Full Flow Test')
    print(f'{"="*60}\n')

    config = load_config(BASE_DIR / 'config.json')
    init_logger(BASE_DIR)
    log_info(f'Logins={config.selected_logins}, batch={config.batch_size}')

    # WAL recovery
    recover_pending_list_from_wal(BASE_DIR)

    # Step 1: Sync
    print('[Step 1] Syncing...')
    try:
        result = sync_to_pending_list(config, BASE_DIR)
        print(f'  Synced {result["appended"]} rows')
    except Exception as e:
        print(f'  FAIL: {e}')
        return

    # Load batch
    batch = load_batch(config, BASE_DIR)
    if not batch.al0_list:
        print('No pending tickets. Done.')
        return
    print(f'  Batch: {len(batch.al0_list)} tickets')
    for al0 in batch.al0_list:
        row = batch.rows[al0]
        print(f'    {al0} | {row["pod"]} | FC={row["flipped_fc"]}')

    # Start browser
    print(f'\n[Browser] Starting Firefox...')
    browser = BrowserManager()
    browser.start()
    print(f'  HWND={browser._hwnd}')

    # Warmup: load OC homepage first to avoid cold-start timeout
    print(f'\n[Warmup] Loading OC homepage...')
    browser.warmup()
    print(f'  OK OC ready')

    print(f'\n[Ready] Press Enter to begin processing {len(batch.al0_list)} tickets...')
    input()

    # Process
    batch_ctrl = BatchController(config.batch_size)
    batch_ctrl.start_batch(batch)

    try:
        for idx, al0 in enumerate(batch.al0_list):
            batch_ctrl.update_progress(idx, al0)
            result = process_single_al0(al0, batch, config, BASE_DIR, browser)
            batch_ctrl.record_result(result)
            append_wal(BASE_DIR, al0, result.status, result.reason)
            if result.status in ("success", "odm"):
                append_to_history(BASE_DIR, al0, result.status)
            gc.collect()

            # Brief pause between tickets
            import time
            time.sleep(1)

        # Write back
        batch_ctrl.write_back_pending_list(BASE_DIR)
        clear_wal(BASE_DIR)

        # Summary
        batch_ctrl.show_summary_dialog()

    finally:
        browser.close()
        print('\nBrowser closed. Done.')


if __name__ == '__main__':
    main()


