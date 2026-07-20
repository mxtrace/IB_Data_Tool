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
from core.outlook_helper import (
    search_inbound_notification,
    display_reply_mail,
    display_new_mail,
)


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
        f"{ib_row.get('shipper_company_name', '')} "
        f"{ib_row.get('pol', '')} "
        f"{ib_row.get('hbl_number', '')}"
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

    # 5.4 收件人
    mail_info = search_inbound_notification(al0)
    use_reply = mail_info is not None and mail_info["recipients"]

    if not use_reply:
        # Fallback 1: Input Zone emails
        emails = []
        if input_zone.shipper.email:
            emails.append(input_zone.shipper.email)
        if input_zone.primary_contact.email:
            emails.append(input_zone.primary_contact.email)
        # 去 amazon 邮箱 + 去重
        emails = [e for e in emails if not e.lower().endswith("@amazon.com")]
        emails = list(dict.fromkeys(emails))

        if not emails:
            return EmailResult(error="无法获取收件人邮箱")
        to_addr = ";".join(emails)

    # 5.5 抄送
    cc_final = _build_cc(al0, ib_row, config, base_dir)

    # 5.7 弹窗
    if use_reply:
        # 修改主题：进仓通知 → 进仓数据确认
        reply_subject = mail_info["subject"].replace("进仓通知", "进仓数据确认")
        if "进仓数据确认" not in reply_subject:
            reply_subject = subject  # fallback 用标准主题

        result = display_reply_mail(
            original_mail_info=mail_info,
            subject=reply_subject,
            html_body=html_body,
            cc=cc_final,
            attachment_path=attachment_path,
        )
    else:
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
    格式（PRD §8.6~8.8）：{actor_login}-lclbc@amazon.com {notes}
    notes = special_request/分批进仓（用 / 连接）
    CC 使用当前操作者 login（config.selected_logins[0]），非数据源 BC 列。
    """
    actor_login = config.selected_logins[0] if config.selected_logins else ""
    cc_email = f"{actor_login}-lclbc@amazon.com" if actor_login else ""

    # 备注部分
    notes = []

    # 特殊需求（§8.7）
    seller_req = _lookup_seller_request(
        ib_row.get("shipper_company_name", ""),
        base_dir / config.seller_request_file,
    )
    if seller_req:
        notes.append(seller_req)

    # 分批进仓（§8.8）
    if _check_partial_inbound(al0, base_dir / config.loading_file):
        notes.append("分批进仓")

    if notes:
        return f"{cc_email} {'/'.join(notes)}"
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


def _check_partial_inbound(al0: str, loading_path: Path) -> bool:
    """检查 Loading 表 R 列(备注)是否含"分批" """
    if not loading_path.exists():
        return False
    wb = openpyxl.load_workbook(str(loading_path), read_only=True, data_only=True)
    ws = wb["List"] if "List" in wb.sheetnames else wb.active
    # 表头 Row5，数据 Row6+
    for row in ws.iter_rows(min_row=6, values_only=True):
        if str(row[0] or "").strip() == al0:
            remark = str(row[17] or "").strip()  # R列 = index 17
            wb.close()
            return "分批" in remark
    wb.close()
    return False
