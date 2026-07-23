"""
step5_email.py — 生成邮件（主题/正文/收件人/抄送/附件 → Display弹窗）
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import openpyxl

from core.config import AppConfig
from core.holidays import next_workday
from core.input_zone_parser import InputZoneData
from core.outlook_helper import search_al0_email, display_new_mail


@dataclass
class EmailResult:
    success: bool = False
    error: str = ""


def generate_email(
    al0: str,
    ib_row: dict,
    input_zone: InputZoneData,
    attachment_path: str,
    config: AppConfig,
    base_dir: Path,
) -> EmailResult:
    """Step 5 完整流程"""

    # 5.1 邮件主题
    pod = ib_row.get("pod", "")
    region = pod[:2].upper() if pod else ""
    subject = (
        f"{region} [进仓数据确认] {al0} "
        f"{ib_row.get("shipper_company_name", "")} "
        f"{ib_row.get("pol", "")} "
        f"{ib_row.get("hbl_number", "")}"
    )

    # 5.2 Blurb 正文
    pol = ib_row.get("pol", "")
    blurb_path = base_dir / config.blurb_root / pol / "template.txt"
    if not blurb_path.exists():
        return EmailResult(error=f"找不到 Blurb 模板：IB txt/{pol}/template.txt")

    html_body = blurb_path.read_text(encoding="utf-8")
    html_body = html_body.replace("{ShipmentID}", al0)
    html_body = html_body.replace("{received_CTN}", str(int(ib_row.get("received_cartons", 0))))
    html_body = html_body.replace("{received_KGS}", str(ib_row.get("received_weight", 0)))
    html_body = html_body.replace("{received_CBM}", str(ib_row.get("received_volume", 0)))
    html_body = html_body.replace("{sicut}", _calc_sicut())

    # 5.3 收件人：优先从 Outlook 搜索 AL0 邮件提取
    search_stores = config.search_stores if config.search_stores else None
    mail_info = search_al0_email(al0, search_stores)

    if mail_info and mail_info["emails"]:
        to_addr = ";".join(mail_info["emails"])
    else:
        # Fallback: OC API Input Zone 邮箱
        emails = []
        if input_zone.shipper.email:
            emails.append(input_zone.shipper.email)
        if input_zone.primary_contact.email:
            emails.append(input_zone.primary_contact.email)
        emails = [e for e in emails if not e.lower().endswith("@amazon.com")]
        emails = list(dict.fromkeys(emails))

        if not emails:
            return EmailResult(error="无法获取收件人邮箱")
        to_addr = ";".join(emails)

    # 5.4 抄送
    cc_final = _build_cc(al0, ib_row, config, base_dir)

    # 5.5 新建邮件弹窗
    result = display_new_mail(
        subject=subject,
        to=to_addr,
        html_body=html_body,
        cc=cc_final,
        attachment_path=attachment_path,
    )

    if not result.get("success"):
        return EmailResult(error=result.get("error", "邮件弹窗生成失败"))

    return EmailResult(success=True)


# ══════════════════════════════════════════════════════════════════════

def _calc_sicut() -> str:
    """截止时间计算：<12→13:00, 12~16→17:00, >16→次工作日10:00"""
    now = datetime.now()
    if now.hour < 12:
        sicut = now.replace(hour=13, minute=0, second=0, microsecond=0)
    elif now.hour < 16:
        sicut = now.replace(hour=17, minute=0, second=0, microsecond=0)
    else:
        nwd = next_workday(now.date())
        sicut = datetime(nwd.year, nwd.month, nwd.day, 10, 0, 0)
    return sicut.strftime("%Y-%m-%d %H:%M:%S")


def _build_cc(al0: str, ib_row: dict, config: AppConfig, base_dir: Path) -> str:
    """
    组装 CC 字段。
    格式：{actor_login}-lclbc@amazon.com {notes}
    """
    import os
    actor_login = os.environ.get("USERNAME", "").lower()
    cc_email = f"{actor_login}-lclbc@amazon.com" if actor_login else ""

    notes = []

    # 特殊需求
    seller_req = _lookup_seller_request(
        ib_row.get("shipper_company_name", ""),
        base_dir / config.seller_request_file,
    )
    if seller_req:
        notes.append(seller_req)

    if notes:
        return f"{cc_email}; {'/'.join(notes)}"
    return cc_email


def _lookup_seller_request(shipper_id: str, file_path: Path) -> str:
    """查询 Seller request list，返回 B 列 Request 文本"""
    if not shipper_id or not file_path.exists():
        return ""
    wb = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0] or "").strip().upper() == shipper_id.upper():
            wb.close()
            return str(row[1] or "").strip()
    wb.close()
    return ""
