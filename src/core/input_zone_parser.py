"""
input_zone_parser.py — OC Booking Summary 文本解析器
从 Ctrl+A Ctrl+C 复制的文本中提取结构化数据。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from core.address_parser import parse_address


@dataclass
class PartyInfo:
    company: str = ""
    email: str = ""
    address_raw: str = ""
    # 拆分后地址字段
    street: str = ""
    city: str = ""
    state: str = ""
    zip: str = ""
    country: str = ""


@dataclass
class InputZoneData:
    odm_booking: bool = False
    cda_booking: bool = False
    shipper: PartyInfo = field(default_factory=PartyInfo)
    consignee: PartyInfo = field(default_factory=PartyInfo)
    notify: PartyInfo = field(default_factory=PartyInfo)
    die: PartyInfo = field(default_factory=PartyInfo)  # Destination Importer Entity
    primary_contact: PartyInfo = field(default_factory=PartyInfo)


def parse_input_zone(raw_text: str) -> InputZoneData:
    """
    解析 OC Booking Summary 全选复制的文本。
    返回结构化的 InputZoneData。
    """
    data = InputZoneData()

    # ── 1. ODM / CDA 字段 ──
    data.odm_booking = _extract_yes_no(raw_text, "ODM Booking")
    data.cda_booking = _extract_yes_no(raw_text, "CDA Booking")

    # ── 2. Parties 分段解析 ──
    parties_text = _extract_parties_section(raw_text)
    if parties_text:
        sections = _split_party_sections(parties_text)

        # Shipper: company + email + address(拆分)
        shipper_raw = sections.get("shipper", "")
        data.shipper.company = _extract_field(shipper_raw, "Company")
        data.shipper.email = _extract_field(shipper_raw, "Email")
        data.shipper.address_raw = _extract_field(shipper_raw, "Address")
        if data.shipper.address_raw:
            addr = parse_address(data.shipper.address_raw)
            data.shipper.street = addr["street"]
            data.shipper.city = addr["city"]
            data.shipper.state = addr["state"]
            data.shipper.zip = addr["zip"]
            data.shipper.country = addr["country"]

        # Consignee: company only
        consignee_raw = sections.get("consignee", "")
        data.consignee.company = _extract_field(consignee_raw, "Company")

        # Notify Party: company only
        notify_raw = sections.get("notify", "")
        data.notify.company = _extract_field(notify_raw, "Company")

        # DIE: company + email + address(拆分)
        die_raw = sections.get("die", "")
        data.die.company = _extract_field(die_raw, "Company")
        data.die.email = _extract_field(die_raw, "Email")
        data.die.address_raw = _extract_field(die_raw, "Address")
        if data.die.address_raw:
            addr = parse_address(data.die.address_raw)
            data.die.street = addr["street"]
            data.die.city = addr["city"]
            data.die.state = addr["state"]
            data.die.zip = addr["zip"]
            data.die.country = addr["country"]

        # Primary Contact: email only
        primary_raw = sections.get("primary_contact", "")
        data.primary_contact.email = _extract_field(primary_raw, "Email")

    return data


# ══════════════════════════════════════════════════════════════════════
# 内部辅助函数
# ══════════════════════════════════════════════════════════════════════

def _extract_yes_no(text: str, label: str) -> bool:
    """从文本中提取 label 对应的 Yes/No 值"""
    pattern = re.compile(rf"{re.escape(label)}\s*\n\s*(Yes|No)", re.IGNORECASE)
    match = pattern.search(text)
    if match:
        return match.group(1).strip().lower() == "yes"
    return False


def _extract_parties_section(text: str) -> str:
    """提取 'Parties' 标记之后的文本段落"""
    # Parties 段落从 "Parties" 独立行开始，到 "Cargo Information" 结束
    match = re.search(r"\nParties\s*\n(.*?)(?:\nCargo Information|\Z)", text, re.DOTALL)
    return match.group(1) if match else ""


def _split_party_sections(parties_text: str) -> dict[str, str]:
    """
    将 Parties 文本按 Party 名称分段。
    Party 名称行特征：缩进 + 已知前缀。
    """
    # 定义 Party 前缀和对应 key
    PARTY_PREFIXES = [
        ("Destination Importer Entity", "die"),
        ("Shipper", "shipper"),
        ("Consignee", "consignee"),
        ("Notify Party", "notify"),
        ("Primary Contact for this Booking", "primary_contact"),
        ("ISF Buyer", "_isf_buyer"),
        ("ISF Seller", "_isf_seller"),
        ("ISF Manufacture", "_isf_manufacture"),
    ]

    sections = {}
    lines = parties_text.split("\n")
    current_key = None
    current_lines = []

    for line in lines:
        stripped = line.strip()
        matched = False
        for prefix, key in PARTY_PREFIXES:
            if stripped.startswith(prefix):
                # 保存前一个 section
                if current_key:
                    sections[current_key] = "\n".join(current_lines)
                current_key = key
                current_lines = []
                matched = True
                break
        if not matched and current_key:
            current_lines.append(line)

    # 保存最后一个 section
    if current_key:
        sections[current_key] = "\n".join(current_lines)

    return sections


def _extract_field(section_text: str, field_name: str) -> str:
    """
    从 Party section 文本中提取指定字段的值。
    格式：
        {field_name}
        {value}
    值为 "--" 视为空。
    """
    pattern = re.compile(
        rf"^\s*{re.escape(field_name)}\s*$\n^\s*(.+?)\s*$",
        re.MULTILINE
    )
    match = pattern.search(section_text)
    if match:
        value = match.group(1).strip()
        return "" if value == "--" else value
    return ""
