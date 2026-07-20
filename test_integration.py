"""
test_integration.py - Step2 Integration Test
Usage:
    set IB_DEBUG=1
    python test_integration.py AL0-XXXXXXXX
"""
from __future__ import annotations

import os
import sys
import time

os.environ["IB_DEBUG"] = "1"

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from core.browser_manager import BrowserManager
from core.input_zone_parser import InputZoneData
from steps.step2_oc_scrape import scrape_booking_summary, ScrapeResult
from pathlib import Path

BASE_DIR = Path(__file__).parent


def print_party(name, party):
    print(f"  [{name}]")
    if party.company:
        print(f"    Company: {party.company}")
    if party.email:
        print(f"    Email:   {party.email}")
    if party.address_raw:
        print(f"    Address: {party.address_raw}")
        print(f"    -> Street={party.street}, City={party.city}, State={party.state}, Zip={party.zip}, Country={party.country}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python test_integration.py <AL0>")
        print("Example: python test_integration.py AL0-VQGFKGNKSXKNS")
        sys.exit(1)

    al0 = sys.argv[1].strip()
    print(f"\n{'='*60}")
    print(f"  IB Data Tool - Step2 Integration Test")
    print(f"  AL0: {al0}")
    print(f"{'='*60}\n")

    print("[1/4] Starting Chrome (isolated profile)...")
    browser = BrowserManager()
    try:
        browser.start()
        print(f"  OK Chrome started, HWND={browser._hwnd}")
    except Exception as e:
        print(f"  FAIL Chrome start failed: {e}")
        sys.exit(1)

    try:
        print("\n[NOTE] If first time with isolated profile, please login to OC in Chrome.")
        print("       Press Enter here when ready...")
        input()

        print(f"\n[2/4] Scraping Booking Summary...")
        result = scrape_booking_summary(al0, browser)

        if result.error:
            print(f"\n  FAIL: {result.error}")
            print(f"  -> Check debug/{al0}_raw.txt for raw content")
            return

        print(f"  OK Scrape successful!")

        iz = result.input_zone
        print(f"\n[3/4] Parse result:")
        print(f"  ODM Booking: {iz.odm_booking}")
        print(f"  CDA Booking: {iz.cda_booking}")
        print()
        print_party("Shipper", iz.shipper)
        print_party("Consignee", iz.consignee)
        print_party("Notify Party", iz.notify)
        print_party("DIE", iz.die)
        print_party("Primary Contact", iz.primary_contact)

        if iz.cda_booking:
            print(f"\n[4/4] CDA order -> attempting ASI download...")
            asi_path = result.download_asi(browser, BASE_DIR)
            if asi_path:
                print(f"  OK ASI downloaded: {asi_path}")
            else:
                print(f"  FAIL ASI download failed (check logs)")
        else:
            print(f"\n[4/4] Not CDA, skipping ASI download")

        print(f"\n{'='*60}")
        print(f"  Test complete!")
        print(f"  Debug file: debug/{al0}_raw.txt")
        print(f"{'='*60}")

    finally:
        print("\nClosing browser...")
        browser.close()
        print("Done.")


if __name__ == "__main__":
    main()
