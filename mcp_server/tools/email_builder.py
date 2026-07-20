"""
email_builder.py — 邮件构建工具
组装邮件主题、正文、收件人，通过 Outlook COM 弹窗。
"""
from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

from tools.config import get_config

DATA_DIR = Path(__file__).parent.parent / "data"


def gen_email(
    al0: str,
    ib_row: dict,
    template_file: str,
    recipient_email: str = "",
    seller_request: str = "",
) -> dict:
    """
    组装邮件并通过 Outlook COM 弹窗。

    Args:
        al0: 订单 ID
        ib_row: 进仓数据行
        template_file: 填充完成的附件路径
        recipient_email: 收件人邮箱（可为空，BC手动填）
        seller_request: 特殊需求文本

    Returns:
        {success, subject, to, cc, attachment, warning}
    """
    config = get_config()
    if "error" in config:
        return {"success": False, "error": config["error"]}

    # ── 邮件主题 ──
    pod = ib_row.get("pod", "")
    region = pod[:2].upper() if pod else ""
    shipper_id = ib_row.get("shipper_id", "")
    pol = ib_row.get("pol", "")
    hbl = ib_row.get("hbl_number", "")
    subject = f"{region} [进仓数据确认] {al0} {shipper_id} {pol} {hbl}"

    # ── 邮件正文（HTML Blurb）──
    blurb_root = Path(config.get("blurb_root", "data/IB txt"))
    if not blurb_root.is_absolute():
        blurb_root = DATA_DIR.parent / blurb_root
    blurb_path = blurb_root / pol / "template.txt"

    html_body = ""
    if blurb_path.exists():
        html_body = blurb_path.read_text(encoding="utf-8")
        # 占位符替换
        html_body = html_body.replace("{ShipmentID}", al0)
        html_body = html_body.replace("{received_CTN}", str(ib_row.get("received_cartons", "")))
        html_body = html_body.replace("{received_KGS}", str(ib_row.get("received_weight", "")))
        html_body = html_body.replace("{received_CBM}", str(ib_row.get("received_volume", "")))
        html_body = html_body.replace("{sicut}", _calc_sicut())
    else:
        html_body = f"<p>[WARNING] 未找到 Blurb 模板: {blurb_path}</p>"

    # ── 特殊需求追加 ──
    if seller_request:
        html_body += f"\n<br/><p><b>Seller Request:</b> {seller_request}</p>"

    # ── 收件人 ──
    warning = ""
    if not recipient_email:
        warning = f"[WARNING] AL0={al0} 无收件人邮箱，请人工补充"

    # ── 抄送 ──
    bc_login = ib_row.get("bc_login", "")
    cc_email = f"{bc_login}-lclbc@amazon.com" if bc_login else ""

    # ── Outlook COM 弹窗 ──
    attachment = template_file
    outlook_result = _display_outlook_mail(
        subject=subject,
        html_body=html_body,
        to=recipient_email,
        cc=cc_email,
        attachment=attachment,
    )

    return {
        "success": outlook_result.get("success", False),
        "al0": al0,
        "subject": subject,
        "to": recipient_email,
        "cc": cc_email,
        "attachment": attachment,
        "warning": warning,
        "outlook_error": outlook_result.get("error", ""),
    }


def _calc_sicut() -> str:
    """计算截止时间 sicut。"""
    now = datetime.now()
    hour = now.hour

    if hour <= 12:
        sicut = now.replace(hour=13, minute=0, second=0)
    elif hour <= 15:
        sicut = now.replace(hour=17, minute=0, second=0)
    else:
        tomorrow = now + timedelta(days=1)
        sicut = tomorrow.replace(hour=10, minute=0, second=0)

    return sicut.strftime("%Y-%m-%d %H:%M:%S")


def _display_outlook_mail(subject: str, html_body: str, to: str, cc: str, attachment: str) -> dict:
    """通过 Outlook COM 创建并 Display 邮件。"""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem

        mail.Subject = subject
        mail.HTMLBody = html_body
        if to:
            mail.To = to
        if cc:
            mail.CC = cc
        if attachment and Path(attachment).exists():
            mail.Attachments.Add(str(Path(attachment).resolve()))

        mail.Display()  # 弹窗，不发送
        return {"success": True}

    except Exception as e:
        return {"success": False, "error": str(e)}
