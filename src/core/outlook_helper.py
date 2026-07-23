"""
outlook_helper.py — Outlook COM 操作
搜索邮件提取收件人 + 新建邮件 Display。
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


def list_outlook_stores() -> list[str]:
    """枚举 Outlook 中所有已配置的邮箱账号（Stores + Accounts + Folders 三级探测）"""
    import win32com.client
    names = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")
    except Exception:
        return names

    # 方式1: Stores
    try:
        for i in range(1, mapi.Stores.Count + 1):
            store = mapi.Stores.Item(i)
            names.append(store.DisplayName)
    except Exception:
        pass

    # 方式2: Accounts 补充
    try:
        for i in range(1, mapi.Accounts.Count + 1):
            acc = mapi.Accounts.Item(i)
            display = acc.SmtpAddress or acc.DisplayName
            if display and display not in names:
                names.append(display)
    except Exception:
        pass

    # 方式3: Root Folders 补充（订阅邮箱可能只在这里出现）
    try:
        for i in range(1, mapi.Folders.Count + 1):
            folder = mapi.Folders.Item(i)
            if folder.Name and folder.Name not in names:
                names.append(folder.Name)
    except Exception:
        pass

    return names


def search_al0_email(al0: str, store_names: list[str] | None = None) -> dict | None:
    """
    在指定 Store 中搜索主题含 AL0 的最新一封邮件。
    返回提取的非 amazon.com 邮箱列表，或 None。

    返回格式：
    {
        "emails": ["addr1", "addr2", ...],
        "subject": "原邮件主题",
    }
    """
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")

    # 确定搜索范围
    target_stores = []
    if store_names:
        for i in range(1, mapi.Stores.Count + 1):
            store = mapi.Stores.Item(i)
            if store.DisplayName in store_names:
                target_stores.append(store)
    else:
        # 未指定则搜索全部
        for i in range(1, mapi.Stores.Count + 1):
            target_stores.append(mapi.Stores.Item(i))

    # 搜索
    found_items = []
    for store in target_stores:
        try:
            _search_folder_recursive(store.GetRootFolder(), al0, found_items)
        except Exception:
            continue

    if not found_items:
        return None

    # 按发送时间降序，取最新
    found_items.sort(key=lambda x: x.SentOn, reverse=True)
    mail = found_items[0]

    # 提取邮箱：Recipients(To+CC) 中非 amazon
    emails = []
    for i in range(1, mail.Recipients.Count + 1):
        recip = mail.Recipients.Item(i)
        addr = _resolve_smtp_address(recip)
        if addr and not _is_amazon_email(addr):
            emails.append(addr)

    # 发件人非 amazon 也追加
    sender_addr = _resolve_sender_address(mail)
    if sender_addr and not _is_amazon_email(sender_addr):
        emails.append(sender_addr)

    # 去重（保持顺序）
    emails = list(dict.fromkeys(emails))

    if not emails:
        return None

    return {
        "emails": emails,
        "subject": mail.Subject or "",
    }


def display_new_mail(
    subject: str,
    to: str,
    html_body: str,
    cc: str,
    attachment_path: str,
) -> dict:
    """新建邮件，Display弹出，自动追加 work 签名"""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = to
        mail.CC = cc
        if attachment_path and Path(attachment_path).exists():
            mail.Attachments.Add(str(Path(attachment_path).resolve()))

        # 追加 work 签名
        signature_html = _load_work_signature()
        if signature_html:
            mail.HTMLBody = html_body + "<br><br>" + signature_html
        else:
            mail.HTMLBody = html_body

        mail.Display()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}



def _load_work_signature() -> str:
    """
    读取 Outlook 签名目录中名为 work 的签名 HTML。
    匹配规则：文件名以 'work' 开头（不区分大小写），后缀 .htm
    图片相对路径替换为绝对路径。
    """
    import os
    sig_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""

    # 查找 work 签名文件（可能是 "work (email).htm" 或 "work.htm"）
    htm_file = None
    for f in sig_dir.glob("*.htm"):
        if f.stem.lower().startswith("work"):
            htm_file = f
            break

    if not htm_file:
        return ""

    try:
        # 尝试多种编码读取
        for enc in ("utf-8", "gb2312", "gbk", "latin-1"):
            try:
                raw_html = htm_file.read_text(encoding=enc)
                break
            except (UnicodeDecodeError, ValueError):
                continue
        else:
            return ""

        # 提取 <body>...</body> 内容
        import re
        body_match = re.search(r"<body[^>]*>(.*)</body>", raw_html, re.DOTALL | re.IGNORECASE)
        if not body_match:
            return raw_html

        body_content = body_match.group(1)

        # 修正图片/资源的相对路径为绝对路径
        files_dir = htm_file.stem + "_files"
        abs_files_dir = str(sig_dir / files_dir).replace("\\", "/")
        body_content = body_content.replace(files_dir + "/", "file:///" + abs_files_dir + "/")

        return body_content
    except Exception:
        return ""

# ── 内部辅助 ──

def _search_folder_recursive(folder, al0: str, results: list):
    """递归搜索文件夹：主题含 AL0 单号"""
    try:
        items = folder.Items
        filter_str = (
            '@SQL="urn:schemas:httpmail:subject" LIKE \'%' + al0 + '%\''
        )
        filtered = items.Restrict(filter_str)
        for item in filtered:
            results.append(item)
    except Exception:
        pass

    try:
        for i in range(1, folder.Folders.Count + 1):
            _search_folder_recursive(folder.Folders.Item(i), al0, results)
    except Exception:
        pass


def _resolve_smtp_address(recipient) -> str:
    """从 Recipient 对象解析 SMTP 邮箱地址"""
    try:
        # 优先尝试 PropertyAccessor 取 SMTP
        PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        addr = recipient.PropertyAccessor.GetProperty(PR_SMTP)
        if addr:
            return addr.strip()
    except Exception:
        pass
    try:
        # Fallback: 直接取 Address（可能是 Exchange DN）
        addr = recipient.Address or ""
        if "@" in addr:
            return addr.strip()
        # Exchange DN → 尝试 GetExchangeUser
        eu = recipient.AddressEntry.GetExchangeUser()
        if eu:
            return (eu.PrimarySmtpAddress or "").strip()
    except Exception:
        pass
    return ""


def _resolve_sender_address(mail) -> str:
    """解析邮件发件人 SMTP 地址"""
    try:
        # 优先 SenderEmailAddress（SMTP 类型直接可用）
        if mail.SenderEmailType == "SMTP":
            return (mail.SenderEmailAddress or "").strip()
        # Exchange 类型 → GetExchangeUser
        sender = mail.Sender
        if sender:
            eu = sender.GetExchangeUser()
            if eu:
                return (eu.PrimarySmtpAddress or "").strip()
    except Exception:
        pass
    try:
        addr = mail.SenderEmailAddress or ""
        if "@" in addr:
            return addr.strip()
    except Exception:
        pass
    return ""


def _is_amazon_email(addr: str) -> bool:
    lower = addr.lower()
    return lower.endswith("@amazon.com") or ".amazon.com" in lower
