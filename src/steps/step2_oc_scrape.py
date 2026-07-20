"""
step2_oc_scrape.py — OC 数据抓取 + 订单类型判断 + ASI 下载
流程：导航 Booking Summary → 全选复制 → 解析 → (CDA则下载ASI)
"""
from __future__ import annotations

import glob
import os
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path

from core.browser_manager import BrowserManager
from core.input_zone_parser import InputZoneData, parse_input_zone
from core.logger import audit, log_warning


# Firefox 默认下载目录
DOWNLOAD_DIR = Path(os.environ.get("USERPROFILE", "")) / "Downloads"

# Debug 模式：设为 True 时将 raw_text 写入文件供离线分析
DEBUG_DUMP = os.environ.get("IB_DEBUG", "") == "1"
DEBUG_DIR = Path(os.environ.get("USERPROFILE", "")) / "Desktop" / "IB_Data_Tool" / "debug"

# 页面加载检测关键词（Booking Summary 必含）
CONTENT_MARKERS = ["Booking Summary", "Booking ID", "Mode"]


@dataclass
class ScrapeResult:
    al0: str = ""
    input_zone: InputZoneData | None = None
    error: str = ""

    def download_asi(self, browser: BrowserManager, base_dir: Path) -> Path | None:
        """
        CDA 订单：导航到 Shipment Document 页面下载 ASI 文件。
        返回下载到的文件路径，失败返回 None。

        流程（纯键盘）：
        1. 导航到 Shipping document 页面
        2. 等待页面加载（含 "Document Type" 关键词）
        3. Ctrl+F 搜索 "ASI" 确认存在
        4. 使用 Tab 键定位 checkbox + Download 按钮
        5. 等待下载完成
        """
        if not self.al0:
            return None

        try:
            # Step 1: 导航
            browser.navigate_shipment_document(self.al0)
            time.sleep(3)

            # Step 2: 等待页面加载
            text = browser.copy_page_text_stable(min_length=100, max_retries=3)
            if "Document Type" not in text and "Shipping document" not in text:
                log_warning(f"[{self.al0}] Shipment Document 页面加载超时")
                return None

            # Step 3: 检查是否有 ASI 文档
            if "ASI" not in text and "Arrival and Shipping Information" not in text:
                log_warning(f"[{self.al0}] 未找到 ASI 文档")
                return None

            # Step 4: 记录下载前的文件
            before_files = _snapshot_downloads()

            # Step 5: 用键盘流触发下载
            # Ctrl+F "ASI" → 定位到行 → Tab到checkbox → Space勾选 → Tab到Download → Enter
            _keyboard_download_asi(browser)

            # Step 6: 等待下载完成（最多30秒）
            downloaded = _wait_for_new_download(before_files, timeout=30)
            if not downloaded:
                log_warning(f"[{self.al0}] ASI 下载超时")
                return None

            # Step 7: 移动到 Output 目录，重命名为 ASI_{al0}.{ext}
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


def scrape_booking_summary(al0: str, browser: BrowserManager) -> ScrapeResult:
    """
    对单个 AL0 执行 OC Booking Summary 抓取。

    策略：
    1. 导航到 Booking Summary URL
    2. 等待页面稳定（SPA 路由跳转 + 数据加载 需要时间）
    3. 全选复制 + 解析
    4. 校验关键字段存在性
    5. 失败则重试一次（页面可能未加载完）
    """
    result = ScrapeResult(al0=al0)

    # ── 导航 ──
    try:
        browser.navigate_booking_summary(al0)
    except Exception as e:
        result.error = f"OC 页面导航失败：{e}"
        return result

    # ── 抓取（带重试）──
    raw_text = ""
    for attempt in range(2):
        raw_text = browser.copy_page_text_stable(
            min_length=300,
            max_retries=3,
            interval=1.0,
        )

        # 校验：页面是否包含 Booking Summary 内容
        if raw_text and all(m in raw_text for m in CONTENT_MARKERS):
            break

        # 重试：用 JS location.reload() 强制刷新
        if attempt == 0:
            log_warning(f"[{al0}] 第一次抓取内容不完整，等待重试...")
            browser._run_js_in_console("location.reload();")
            time.sleep(4)

    # ── Debug dump ──
    if DEBUG_DUMP:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        dump_file = DEBUG_DIR / f"{al0}_raw.txt"
        dump_file.write_text(raw_text or "(empty)", encoding="utf-8")

    # ── 校验最终内容 ──
    if not raw_text or len(raw_text) < 200:
        result.error = "Booking Summary 内容为空或过短"
        return result

    if not all(m in raw_text for m in CONTENT_MARKERS):
        # 可能 cookie 过期 → HTTP 500
        if "500" in raw_text or "Internal Server Error" in raw_text:
            result.error = "OC 返回 500 错误，请重新登录 trans-logistics-cn.amazon.com"
        else:
            result.error = "页面内容不含 Booking Summary 字段，可能未加载完成"
        return result

    # ── 解析 ──
    try:
        input_zone = parse_input_zone(raw_text)
    except Exception as e:
        result.error = f"Booking Summary 解析失败：{e}"
        return result

    # 二次校验：Parties 段至少有 Shipper company
    if not input_zone.shipper.company:
        result.error = "解析结果异常：Shipper Company 为空"
        return result

    result.input_zone = input_zone
    return result


# ══════════════════════════════════════════════════════════════════════
# ASI 下载辅助函数
# ══════════════════════════════════════════════════════════════════════

def _keyboard_download_asi(browser: BrowserManager):
    """
    在 Shipment Document 页面下载 ASI 文件。
    通过单次 F12 Console JS 执行：
    1. 找到 Document Type = ASI 的行，点击 label 勾选 checkbox
    2. setTimeout 等 React 状态更新
    3. 点击 Download 按钮
    前提：Firefox 已允许 trans-logistics-cn.amazon.com 弹窗（一次性设置）。
    """
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
    """获取当前下载目录中的文件列表快照"""
    if not DOWNLOAD_DIR.exists():
        return set()
    return set(os.listdir(DOWNLOAD_DIR))


def _wait_for_new_download(before_snapshot: set[str], timeout: int = 30) -> Path | None:
    """
    等待下载目录出现新文件（非 .crdownload 临时文件）。
    返回新文件路径，超时返回 None。
    """
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(1)
        if not DOWNLOAD_DIR.exists():
            continue

        current = set(os.listdir(DOWNLOAD_DIR))
        new_files = current - before_snapshot

        for fname in new_files:
            # 跳过浏览器下载中的临时文件
            if fname.endswith(".crdownload") or fname.endswith(".tmp"):
                continue
            full_path = DOWNLOAD_DIR / fname
            # 确保文件大小稳定（不在写入中）
            try:
                size1 = full_path.stat().st_size
                time.sleep(0.5)
                size2 = full_path.stat().st_size
                if size1 == size2 and size1 > 0:
                    return full_path
            except OSError:
                continue

    return None


