"""
outlook_helper.py — Outlook COM 操作
搜索"进仓通知"邮件 + Reply + 新建邮件 + Display。
"""
from __future__ import annotations

from pathlib import Path


def check_outlook_running() -> bool:
    """检查 Outlook 是否正在运行"""
    try:
        import win32com.client
        win32com.client.Dispatch("Outlook.Application")
        return True
    except Exception:
        return False


def search_inbound_notification(al0: str) -> dict | None:
    """
    搜索 Outlook 全部文件夹：主题含 al0 AND 含 "进仓通知"。
    返回最新匹配的邮件对象包装，或 None。
    """
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")

    # 遍历所有 Store 的所有文件夹
    found_items = []
    for store in mapi.Stores:
        try:
            _search_folder_recursive(store.GetRootFolder(), al0, found_items)
        except Exception:
            continue

    if not found_items:
        return None

    # 按发送时间降序，取最新
    found_items.sort(key=lambda x: x.SentOn, reverse=True)
    mail = found_items[0]

    # 提取收件人（去除 @amazon.com）
    recipients = []
    for i in range(1, mail.Recipients.Count + 1):
        addr = mail.Recipients.Item(i).Address or ""
        if addr and not _is_amazon_email(addr):
            recipients.append(addr)
    # 去重
    recipients = list(dict.fromkeys(recipients))

    return {
        "mail_item": mail,
        "recipients": recipients,
        "subject": mail.Subject or "",
    }


def display_reply_mail(
    original_mail_info: dict,
    subject: str,
    html_body: str,
    cc: str,
    attachment_path: str,
) -> dict:
    """在原邮件上 Reply，修改主题/正文/CC/附件，Display弹出"""
    try:
        mail = original_mail_info["mail_item"].Reply()
        mail.Subject = subject
        mail.To = ";".join(original_mail_info["recipients"])
        mail.HTMLBody = html_body
        mail.CC = cc
        if attachment_path and Path(attachment_path).exists():
            mail.Attachments.Add(str(Path(attachment_path).resolve()))
        mail.Display()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def display_new_mail(
    subject: str,
    to: str,
    html_body: str,
    cc: str,
    attachment_path: str,
) -> dict:
    """新建邮件，Display弹出"""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = to
        mail.HTMLBody = html_body
        mail.CC = cc
        if attachment_path and Path(attachment_path).exists():
            mail.Attachments.Add(str(Path(attachment_path).resolve()))
        mail.Display()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── 内部辅助 ──

def _search_folder_recursive(folder, al0: str, results: list):
    """递归搜索文件夹"""
    try:
        items = folder.Items
        # Restrict 条件：主题含 al0 且含 "进仓通知"
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{al0}%' "
            f"AND \"urn:schemas:httpmail:subject\" LIKE '%进仓通知%'"
        )
        filtered = items.Restrict(filter_str)
        for item in filtered:
            results.append(item)
    except Exception:
        pass

    # 递归子文件夹
    try:
        for i in range(1, folder.Folders.Count + 1):
            _search_folder_recursive(folder.Folders.Item(i), al0, results)
    except Exception:
        pass


def _is_amazon_email(addr: str) -> bool:
    lower = addr.lower()
    return lower.endswith("@amazon.com") or lower.endswith(".amazon.com")
