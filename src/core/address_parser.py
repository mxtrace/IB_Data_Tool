"""
address_parser.py — 地址拆分
策略：从后往前：Country → Zip → State → City → 剩余=Street

非 US/UK 地址：邮编前倒数第1个有效词=省份，倒数第2个有效词=城市
有效词：不含数字、不是街道关键词
"""
from __future__ import annotations

import re

# 国家关键词（尾部匹配）
COUNTRIES = [
    "China", "United States", "United Kingdom", "Germany", "Japan",
    "Canada", "Australia", "France", "Italy", "Spain", "Netherlands",
    "India", "South Korea", "Singapore", "Hong Kong", "Taiwan",
    "Thailand", "Vietnam", "Indonesia", "Malaysia", "Philippines",
    "Brazil", "Mexico", "Poland", "Czech Republic", "Sweden",
]

# US 州缩写
US_STATES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
    "DC",
}

# 街道关键词（小写），命中则不作为 city/state
STREET_KEYWORDS = {
    "road", "rd", "street", "st", "avenue", "ave", "building", "bldg",
    "plaza", "tower", "floor", "fl", "block", "blk", "unit", "suite", "ste",
    "lane", "ln", "drive", "dr", "boulevard", "blvd", "way", "court", "ct",
    "park", "center", "centre", "industrial", "zone", "district", "area",
    "no", "room", "rm", "house", "mansion", "garden", "phase", "section",
    "compound", "estate", "square", "highway", "hwy", "close", "crescent",
    "terrace", "place", "pl", "village", "town", "city",
}


def _is_valid_geo_word(word: str) -> bool:
    """判断一个词是否可作为城市/省份（非街道关键词、不含数字）"""
    if not word:
        return False
    clean = word.rstrip(",")
    if not clean:
        return False
    if any(c.isdigit() for c in clean):
        return False
    if clean.lower() in STREET_KEYWORDS:
        return False
    return True


def parse_address(raw: str) -> dict:
    """
    解析单行地址，拆为 {street, city, state, zip, country}。
    策略：从后往前提取。
    """
    result = {"street": "", "city": "", "state": "", "zip": "", "country": ""}

    if not raw or not raw.strip():
        return result

    working = raw.strip()
    working = working.replace("\uff0c", ",")
    working = re.sub(r"\s+", " ", working)

    # 1. 提取 Country（尾部匹配）
    for country in COUNTRIES:
        if working.upper().endswith(country.upper()):
            result["country"] = country
            working = working[: -len(country)].strip().rstrip(",").strip()
            break

    # 2. 提取 Zip
    # UK: AA9A 9AA / A9A 9AA / AA9 9AA / A9 9AA etc.
    # Others: 4-6 digits
    zip_match = re.search(
        r"([A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}|\d{4,6})\s*$",
        working, re.IGNORECASE
    )
    if zip_match:
        result["zip"] = zip_match.group(1).strip()
        working = working[: zip_match.start()].strip().rstrip(",").strip()

    # 3 & 4. 提取 State 和 City
    is_us = result["country"] == "United States"
    is_uk = result["country"] == "United Kingdom"

    if is_us:
        # US: 末尾2字母州码
        state_match = re.search(r"\b([A-Z]{2})\s*$", working)
        if state_match and state_match.group(1) in US_STATES:
            result["state"] = state_match.group(1)
            working = working[: state_match.start()].strip().rstrip(",").strip()
        # City: 逗号前的词
        city_match = re.search(r",\s*([^,]+?)\s*$", working)
        if city_match:
            result["city"] = city_match.group(1).strip()
            working = working[: city_match.start()].strip()
        result["street"] = working.strip().rstrip(",").strip()

    elif is_uk:
        # UK: 逗号分隔，最后两段为 city/county
        parts = [p.strip() for p in working.split(",") if p.strip()]
        if len(parts) >= 3:
            result["state"] = parts[-1]
            result["city"] = parts[-2]
            result["street"] = ", ".join(parts[:-2])
        elif len(parts) == 2:
            result["city"] = parts[-1]
            result["street"] = parts[0]
        else:
            result["street"] = working

    else:
        # 非 US/UK：邮编前倒数第1个有效词=省份，倒数第2个有效词=城市
        words = working.split()
        state_idx = None
        city_idx = None

        # 从末尾往前找省份
        for i in range(len(words) - 1, -1, -1):
            if _is_valid_geo_word(words[i]):
                result["state"] = words[i].rstrip(",")
                state_idx = i
                break
            else:
                break

        # 继续往前找城市
        if state_idx is not None and state_idx > 0:
            for i in range(state_idx - 1, -1, -1):
                if _is_valid_geo_word(words[i]):
                    result["city"] = words[i].rstrip(",")
                    city_idx = i
                    break
                else:
                    break

        # 剩余 = street
        if city_idx is not None:
            result["street"] = " ".join(words[:city_idx]).strip().rstrip(",").strip()
        elif state_idx is not None:
            result["street"] = " ".join(words[:state_idx]).strip().rstrip(",").strip()
        else:
            result["street"] = working.strip().rstrip(",").strip()

    return result
