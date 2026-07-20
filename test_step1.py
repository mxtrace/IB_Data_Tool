"""
test_step1.py - Step 1 Integration Test (local data)
Tests: sync shared drive -> filter by BC -> create Pending List -> load batch
"""
from __future__ import annotations
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pathlib import Path
from core.config import load_config
from core.logger import init_logger, log_info
from steps.step1_sync import sync_to_pending_list, load_batch

BASE_DIR = Path(__file__).parent


def main():
    print(f'\n{"="*60}')
    print(f'  Step 1 Integration Test (Local Data)')
    print(f'{"="*60}\n')

    # Load config
    config = load_config(BASE_DIR / 'config.json')
    print(f'[config] logins={config.selected_logins}, batch_size={config.batch_size}')
    print(f'[config] shared_drive={config.shared_drive_path}')

    # Init logger
    init_logger(BASE_DIR)

    # === Step 1.1-1.3: Sync ===
    print(f'\n[Step 1.1-1.3] Sync from shared drive...')
    try:
        result = sync_to_pending_list(config, BASE_DIR)
        print(f'  OK Source: {result["source"]}')
        print(f'  OK Appended: {result["appended"]} rows')
    except Exception as e:
        print(f'  FAIL: {e}')
        import traceback
        traceback.print_exc()
        return

    # === Step 1.4: Load Batch ===
    print(f'\n[Step 1.4] Load batch (size={config.batch_size})...')
    try:
        batch = load_batch(config, BASE_DIR)
        print(f'  AL0 list ({len(batch.al0_list)} tickets):')
        for al0 in batch.al0_list:
            row = batch.rows[al0]
            print(f'    {al0} | FC={row["flipped_fc"]} | POD={row["pod"]} | CTN={row["received_cartons"]}')
    except Exception as e:
        print(f'  FAIL: {e}')
        import traceback
        traceback.print_exc()
        return

    # === Verify Pending List file ===
    pl_dir = BASE_DIR / config.pending_list_dir
    pl_files = list(pl_dir.glob('*.xlsx'))
    if pl_files:
        print(f'\n[verify] Pending List: {pl_files[0].name} ({pl_files[0].stat().st_size} bytes)')
    else:
        print(f'\n[verify] WARNING: No Pending List file found in {pl_dir}')

    print(f'\n{"="*60}')
    print(f'  Step 1 Test Complete!')
    print(f'{"="*60}')


if __name__ == '__main__':
    main()
