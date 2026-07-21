"""
oc_api_client.py — OC REST API 客户端
直接调用 OC 后端接口获取 Booking 数据，替代浏览器抓取。

已验证端点：
  GET  /aglt/rest/bookingV2/getBookingById/{al0}
  POST /aglt/rest/getAddressInfosWithContactOnly
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import urllib3
import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

OC_BASE = "https://trans-logistics-cn.amazon.com"
BOOKING_DETAIL_API = f"{OC_BASE}/aglt/rest/bookingV2/getBookingById/{{al0}}"
PARTIES_API = f"{OC_BASE}/aglt/rest/getAddressInfosWithContactOnly"

# Firefox profile 路径
FIREFOX_PROFILE_DIR = (
    Path(os.environ.get("APPDATA", ""))
    / "Mozilla" / "Firefox" / "Profiles"
)

# Party contact 字段映射
PARTY_CONTACT_FIELDS = {
    "shipperContactId": "shipper",
    "consigneeContactId": "consignee",
    "notifyPartyContactId": "notify",
    "importerParty": "die",
    "overallContactId": "primary_contact",
}


@dataclass
class PartyData:
    company: str = ""
    email: str = ""
    address_line1: str = ""
    address_line2: str = ""
    city: str = ""
    state: str = ""
    zip: str = ""
    country: str = ""


@dataclass
class BookingData:
    """从 API 获取的 Booking 完整数据"""
    al0: str = ""
    odm_booking: bool = False
    cda_booking: bool = False
    shipper_company: str = ""
    pol: str = ""
    dst_country: str = ""
    shipper: PartyData = field(default_factory=PartyData)
    consignee: PartyData = field(default_factory=PartyData)
    notify: PartyData = field(default_factory=PartyData)
    die: PartyData = field(default_factory=PartyData)
    primary_contact: PartyData = field(default_factory=PartyData)
    error: str = ""


def _find_firefox_profile() -> Path:
    """查找 Firefox ESR profile 目录"""
    if not FIREFOX_PROFILE_DIR.exists():
        raise FileNotFoundError(f"Firefox Profiles 目录不存在：{FIREFOX_PROFILE_DIR}")
    # 优先 default-esr，其次 default-release，最后 default
    for suffix in ("default-esr", "default-release", "default"):
        for d in FIREFOX_PROFILE_DIR.iterdir():
            if d.is_dir() and d.name.endswith(suffix):
                return d
    # Fallback: 第一个目录
    dirs = [d for d in FIREFOX_PROFILE_DIR.iterdir() if d.is_dir()]
    if dirs:
        return dirs[0]
    raise FileNotFoundError("未找到任何 Firefox Profile 目录")


def _load_cookies() -> dict[str, str]:
    """从 Firefox cookie 数据库提取 OC 会话 cookie"""
    profile = _find_firefox_profile()
    cookies_db = profile / "cookies.sqlite"
    if not cookies_db.exists():
        raise FileNotFoundError(f"Cookie 文件不存在：{cookies_db}")

    # 复制避免锁冲突
    tmp = Path(tempfile.mkdtemp()) / "cookies.sqlite"
    shutil.copy2(cookies_db, tmp)

    try:
        conn = sqlite3.connect(str(tmp))
        cur = conn.cursor()
        cur.execute(
            "SELECT name, value FROM moz_cookies "
            "WHERE host LIKE '%trans-logistics%'"
        )
        cookies = {name: value for name, value in cur.fetchall()}
        conn.close()
    finally:
        try:
            tmp.unlink()
            tmp.parent.rmdir()
        except Exception:
            pass

    if not cookies:
        raise RuntimeError(
            "未找到 OC Cookie，请先在 Firefox 登录 trans-logistics-cn.amazon.com"
        )
    return cookies


def build_session() -> requests.Session:
    """构建带 OC 认证的 requests Session"""
    session = requests.Session()
    session.cookies.update(_load_cookies())
    session.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br",
        "Referer": f"{OC_BASE}/aglt/appViews/app",
        "Content-Type": "application/json",
    })
    session.verify = False
    return session


def check_session_valid(session: requests.Session) -> bool:
    """验证 session 是否有效（快速测试一个已知端点）"""
    try:
        resp = session.get(
            f"{OC_BASE}/aglt/rest/bookingV2/getBookingById/AL0-TEST",
            timeout=10,
        )
        if resp.status_code == 200:
            body = resp.json()
            # ERROR + BookingNotFoundException = 认证通过，AL0不存在（正常）
            return body.get("status") in ("SUCCESS", "ERROR")
        # 500 + HTML 页面 = cookie 过期
        if resp.status_code == 500 and "<!DOCTYPE" in resp.text[:50]:
            return False
        # 302 redirect = 需要重新登录
        return resp.status_code not in (302, 401, 403)
    except Exception:
        return False


def fetch_booking_data(al0: str, session: requests.Session) -> BookingData:
    """
    获取单个 AL0 的完整 Booking 数据（Step A + Step B）。
    耗时约 0.5-1s（两次 HTTP）。
    """
    result = BookingData(al0=al0)

    # ── Step A: 获取 Booking 详情 ──
    url = BOOKING_DETAIL_API.format(al0=al0)
    try:
        resp = session.get(url, timeout=15)
    except requests.RequestException as e:
        result.error = f"API 请求失败：{e}"
        return result

    if resp.status_code == 500:
        result.error = "OC 返回 500，Cookie 可能过期，请重新登录 Firefox"
        return result
    if resp.status_code != 200:
        result.error = f"getBookingById 返回 HTTP {resp.status_code}"
        return result

    body = resp.json()
    if body.get("status") != "SUCCESS":
        result.error = f"getBookingById status={body.get('status')}"
        return result

    data = body.get("data", {})

    # 提取基础字段
    result.odm_booking = bool(data.get("isODM2Enabled", False))
    result.cda_booking = bool(data.get("isCDAEnabled", False))
    result.shipper_company = data.get("shipperCompany", "")
    result.pol = data.get("from", "")
    result.dst_country = data.get("dstCountry", "")

    # ── Step B: 获取 Parties 详情 ──
    contact_ids = []
    id_to_role = {}
    for field_name, role in PARTY_CONTACT_FIELDS.items():
        cid = (data.get(field_name) or "").strip()
        if cid and cid not in id_to_role:
            contact_ids.append(cid)
            id_to_role[cid] = role

    if not contact_ids:
        # 没有 contact ID 但有 shipperCompany → 部分数据可用
        result.shipper.company = result.shipper_company
        return result

    # 批量请求 Parties（某个 ID 可能触发 500，则逐个兜底）
    entities = _fetch_parties_batch(session, contact_ids)
    if entities is None:
        entities = _fetch_parties_individual(session, contact_ids)

    # 填充各 Party 数据
    for cid, info in entities.items():
        role = id_to_role.get(cid)
        if not role or not isinstance(info, dict):
            continue
        party = _parse_party(info)
        setattr(result, role, party)

    # 确保 shipper.company 有值
    if not result.shipper.company:
        result.shipper.company = result.shipper_company

    return result


def _fetch_parties_batch(
    session: requests.Session, contact_ids: list[str]
) -> Optional[dict]:
    """批量获取 Parties，失败返回 None"""
    payload = {"addressIds": contact_ids, "contactOnlyAddressIds": []}
    try:
        resp = session.post(PARTIES_API, json=payload, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            if isinstance(data, dict):
                return data
        return None
    except Exception:
        return None


def _fetch_parties_individual(
    session: requests.Session, contact_ids: list[str]
) -> dict:
    """逐个获取 Parties（容错模式）"""
    result = {}
    for cid in contact_ids:
        payload = {"addressIds": [cid], "contactOnlyAddressIds": []}
        try:
            resp = session.post(PARTIES_API, json=payload, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, dict) and data:
                    result.update(data)
        except Exception:
            continue
    return result


def _parse_party(info: dict) -> PartyData:
    """将 API 响应转为 PartyData"""
    return PartyData(
        company=info.get("companyName") or "",
        email=info.get("email") or "",
        address_line1=info.get("addressLine1") or "",
        address_line2=info.get("addressLine2") or "",
        city=info.get("city") or "",
        state=info.get("stateOrRegion") or "",
        zip=info.get("postalCode") or "",
        country=info.get("countryCode") or "",
    )
