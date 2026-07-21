"""
step2_oc_scrape.py — OC 数据获取 + 订单类型判断 + ASI 下载
v2.0: 使用 REST API 直接获取数据（替代浏览器抓取，~1s/单）
"""
from __future__ import annotations

import os
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path

import requests

from core.browser_manager import BrowserManager
from core.input_zone_parser import InputZoneData, PartyInfo
from core.logger import audit, log_warning
from core.oc_api_client import BookingData, fetch_booking_data


# Firefox 默认下载目录
DOWNLOAD_DIR = Path(os.environ.get("USERPROFILE", "")) / "Downloads"


@dataclass
class ScrapeResult:
    al0: str = ""
    input_zone: InputZoneData | None = None
    error: str = ""

    def download_asi(self, browser: BrowserManager, base_dir: Path) -> Path | None:
        """
        CDA 订单：通过浏览器下载 ASI 文件（仍需 browser）。
        """
        if not self.al0:
            return None

        try:
            browser.navigate_shipment_document(self.al0)
            time.sleep(3)

            text = browser.copy_page_text_stable(min_length=100, max_retries=3)
            if "Document Type" not in text and "Shipping document" not in text:
                log_warning(f"[{self.al0}] Shipment Document 页面加载超时")
                return None

            if "ASI" not in text and "Arrival and Shipping Information" not in text:
                log_warning(f"[{self.al0}] 未找到 ASI 文档")
                return None

            before_files = _snapshot_downloads()
            _keyboard_download_asi(browser)

            downloaded = _wait_for_new_download(before_files, timeout=30)
            if not downloaded:
                log_warning(f"[{self.al0}] ASI 下载超时")
                return None

            output_dir = base_dir / "Output" / self.al0
            output_dir.mkdir(parents=True, exist_ok=True)
            ext = downloaded.suffix or ".xlsx"
            dest = output_dir / f"ASI_{self.al0}{ext}"
            shutil.move(str(downloaded), str(dest))

            audit(self.al0, "step2", "asi_downloaded", dest.name)
            return dest

        except Exception as e:
            log_warning(f"[{self.al0}] ASI 下载异常：{e}")
            return None


def scrape_booking_summary(al0: str, session: requests.Session) -> ScrapeResult:
    """
    对单个 AL0 调用 OC API 获取 Booking Summary 数据。
    v2.0: 纯 API 调用，耗时 ~0.5-1s（替代原 ~10s 浏览器抓取）。
    """
    result = ScrapeResult(al0=al0)

    # 调用 API
    booking = fetch_booking_data(al0, session)

    if booking.error:
        result.error = booking.error
        return result

    # 转换为 InputZoneData（保持下游接口兼容）
    input_zone = _booking_to_input_zone(booking)

    # 校验：至少有 Shipper company
    if not input_zone.shipper.company:
        result.error = "API 返回数据异常：Shipper Company 为空"
        return result

    result.input_zone = input_zone
    return result


def _booking_to_input_zone(booking: BookingData) -> InputZoneData:
    """将 API BookingData 转换为 InputZoneData（兼容现有下游逻辑）"""
    iz = InputZoneData()
    iz.odm_booking = booking.odm_booking
    iz.cda_booking = booking.cda_booking

    # Shipper
    iz.shipper = PartyInfo(
        company=booking.shipper.company or booking.shipper_company,
        email=booking.shipper.email,
        address_raw=_build_address_raw(booking.shipper),
        street=booking.shipper.address_line1,
        city=booking.shipper.city,
        state=booking.shipper.state,
        zip=booking.shipper.zip,
        country=booking.shipper.country,
    )

    # Consignee
    iz.consignee = PartyInfo(
        company=booking.consignee.company,
        email=booking.consignee.email,
        address_raw=_build_address_raw(booking.consignee),
        street=booking.consignee.address_line1,
        city=booking.consignee.city,
        state=booking.consignee.state,
        zip=booking.consignee.zip,
        country=booking.consignee.country,
    )

    # Notify Party
    iz.notify = PartyInfo(
        company=booking.notify.company,
        email=booking.notify.email,
    )

    # DIE (Destination Importer Entity)
    iz.die = PartyInfo(
        company=booking.die.company,
        email=booking.die.email,
        address_raw=_build_address_raw(booking.die),
        street=booking.die.address_line1,
        city=booking.die.city,
        state=booking.die.state,
        zip=booking.die.zip,
        country=booking.die.country,
    )

    # Primary Contact
    iz.primary_contact = PartyInfo(
        company=booking.primary_contact.company,
        email=booking.primary_contact.email,
    )

    return iz


def _build_address_raw(party) -> str:
    """拼接地址为单行文本"""
    parts = [
        party.address_line1,
        party.address_line2 if hasattr(party, 'address_line2') else "",
        party.city,
        party.state,
        party.zip,
        party.country,
    ]
    return ", ".join(p for p in parts if p)


# ══════════════════════════════════════════════════════════════════════
# ASI 下载辅助函数（仍用浏览器）
# ══════════════════════════════════════════════════════════════════════

def _keyboard_download_asi(browser: BrowserManager):
    """通过 F12 Console JS 下载 ASI 文件"""
    js_code = (
        "var asiSpan=null;"
        "document.querySelectorAll('span').forEach(function(s){"
        "  if(s.textContent.trim()==='ASI'&&s.children.length===0) asiSpan=s;"
        "});"
        "if(asiSpan){"
        "  var p=asiSpan;"
        "  for(var i=0;i<10;i++){"
        "    p=p.parentElement;"
        "    if(!p)break;"
        "    var cbs=p.querySelectorAll('input[type=checkbox]');"
        "    if(cbs.length===1){"
        "      var label=p.querySelector('label');"
        "      if(label){label.click()}"
        "      else{cbs[0].click()}"
        "      break;"
        "    }"
        "  }"
        "}"
        "setTimeout(function(){"
        "  var btns=document.querySelectorAll('button');"
        "  for(var i=0;i<btns.length;i++){"
        "    if(btns[i].textContent.trim()==='Download'&&!btns[i].disabled){"
        "      btns[i].click();"
        "      document.title='DL_CLICKED';"
        "      break;"
        "    }"
        "  }"
        "},1000);"
    )
    browser._run_js_in_console(js_code)
    time.sleep(3)


def _snapshot_downloads() -> set[str]:
    if not DOWNLOAD_DIR.exists():
        return set()
    return set(os.listdir(DOWNLOAD_DIR))


def _wait_for_new_download(before_snapshot: set[str], timeout: int = 30) -> Path | None:
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(1)
        if not DOWNLOAD_DIR.exists():
            continue
        current = set(os.listdir(DOWNLOAD_DIR))
        new_files = current - before_snapshot
        for fname in new_files:
            if fname.endswith(".crdownload") or fname.endswith(".tmp"):
                continue
            full_path = DOWNLOAD_DIR / fname
            try:
                size1 = full_path.stat().st_size
                time.sleep(0.5)
                size2 = full_path.stat().st_size
                if size1 == size2 and size1 > 0:
                    return full_path
            except OSError:
                continue
    return None
