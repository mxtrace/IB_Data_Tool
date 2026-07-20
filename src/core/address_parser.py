"""
address_parser.py — 地址拆分
策略：从后往前（Country → Zip → State → City → 剩余=Street）
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

# 中国省份（拼音，含常见变体）
PROVINCES = [
    "GUANGDONGSHENG", "GUANGDONG", "ZHEJIANGSHENG", "ZHEJIANG",
    "JIANGSUSHENG", "JIANGSU", "SHANGHAISHI", "SHANGHAI",
    "BEIJINGSHI", "BEIJING", "SHANDONGSHENG", "SHANDONG",
    "FUJIANGSHENG", "FUJIAN", "HEBEISHENG", "HEBEI",
    "HUNANSHENG", "HUNAN", "HUBEISHENG", "HUBEI",
    "SICHUANSHENG", "SICHUAN", "ANHUI", "JIANGXI",
    "LIAONING", "SHANXI", "HENAN", "CHONGQING",
    "TIANJIN", "GUANGXI", "YUNNAN", "GUIZHOU",
]

# 中国城市（拼音，含常见变体）
CITIES = [
    "SHENZHENSHI", "SHENZHEN", "GUANGZHOUSHI", "GUANGZHOU",
    "SHANGHAISHI", "SHANGHAI", "BEIJINGSHI", "BEIJING",
    "DONGGUAN", "DONGGUANSHI", "FOSHAN", "FOSHANSHI",
    "HANGZHOUSHI", "HANGZHOU", "NANJINGSHI", "NANJING",
    "SUZHOUSHI", "SUZHOU", "WUXISHI", "WUXI",
    "NINGBOSHI", "NINGBO", "XIAMENSHI", "XIAMEN",
    "QINGDAOSHI", "QINGDAO", "CHENGDUSHI", "CHENGDU",
    "WUHANSHI", "WUHAN", "ZHONGSHANSHI", "ZHONGSHAN",
    "TIANJINSHI", "TIANJIN", "CHONGQINGSHI", "CHONGQING",
    "YIWUSHI", "YIWU", "JINHUASHI", "JINHUA",
    "WENZHOU", "TAIZHOU", "FUZHOU", "QUANZHOU",
    "HEFEI", "CHANGSHA", "NANCHANG", "ZHENGZHOU",
    # 特别行政区/境外
    "HONG KONG", "HONGKONG", "KOWLOON", "MACAU", "TAIPEI",
]

# 特殊地区代码（非中国非美国）
TERRITORY_CODES = {
    "HK", "MO", "TW", "SG", "MY", "TH", "VN", "PH", "ID",
    "JP", "KR", "AU", "NZ", "GB", "DE", "FR", "IT", "ES",
    "NL", "SE", "PL", "CZ", "BR", "MX", "IN",
}

# US 州缩写
US_STATES = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
    "DC",
]


def _find_spaced_match(text: str, target_nospace: str) -> str:
    """
    在 text 中找到去空格后等于 target_nospace 的子串。
    例如 text="Guang Zhou" target="GUANGZHOU" -> "Guang Zhou"
    从后往前滑动窗口查找。
    """
    text_upper = text.upper()
    # 滑动窗口：从text后面开始，找连续字符去空格后==target
    for end in range(len(text), 0, -1):
        for start in range(max(0, end - len(target_nospace) - 5), end):
            segment = text_upper[start:end]
            if segment.replace(" ", "") == target_nospace:
                return text[start:end]
    return ""


def parse_address(raw: str) -> dict:
    """
    将单行地址拆分为 {street, city, state, zip, country}。
    策略：从后往前提取。
    """
    result = {"street": "", "city": "", "state": "", "zip": "", "country": ""}

    if not raw or not raw.strip():
        return result

    working = raw.strip()
    # 归一化：中文逗号→英文，多空格压缩
    working = working.replace("，", ",")
    working = re.sub(r"\s+", " ", working)

    # 1. 提取 Country（尾部匹配）
    for country in COUNTRIES:
        if working.upper().endswith(country.upper()):
            result["country"] = country
            working = working[: -len(country)].strip().rstrip(",").strip()
            break

    # 2. 提取 Zip（最后出现的 5-6 位数字）
    zip_match = re.search(r"(\d{5,6})\s*$", working)
    if zip_match:
        result["zip"] = zip_match.group(1)
        working = working[: zip_match.start()].strip().rstrip(",").strip()

    # 3. 提取 State
    upper_working = working.upper()
    # 先尝试中国省份（长词优先）
    for prov in sorted(PROVINCES, key=len, reverse=True):
        if prov in upper_working:
            # 找到位置，移除
            idx = upper_working.rfind(prov)
            result["state"] = working[idx: idx + len(prov)]
            working = (working[:idx] + working[idx + len(prov):]).strip().rstrip(",").strip()
            break
    else:
        # 尝试末尾 2 字母代码（US州 + 国际地区）
        state_match = re.search(r"\b([A-Z]{2})\s*$", working)
        if state_match:
            code = state_match.group(1)
            if code in US_STATES or code in TERRITORY_CODES:
                result["state"] = code
                working = working[: state_match.start()].strip().rstrip(",").strip()

    # 4. 提取 City（去空格比较，支持 "Guang Zhou" 匹配 "GUANGZHOU"）
    upper_working = working.upper()
    upper_nospace = upper_working.replace(" ", "")
    for city in sorted(CITIES, key=len, reverse=True):
        if city in upper_working:
            idx = upper_working.rfind(city)
            result["city"] = working[idx: idx + len(city)]
            working = (working[:idx] + working[idx + len(city):]).strip().rstrip(",").strip()
            break
        elif city in upper_nospace:
            # 去空格匹配成功，需要在原文中找到对应位置
            # 从后往前扫描原文找带空格的城市名
            city_with_spaces = _find_spaced_match(working, city)
            if city_with_spaces:
                idx = working.upper().rfind(city_with_spaces.upper())
                if idx >= 0:
                    result["city"] = working[idx: idx + len(city_with_spaces)].strip()
                    working = (working[:idx] + working[idx + len(city_with_spaces):]).strip().rstrip(",").strip()
                    break

    # 5. 剩余 = Street
    result["street"] = working.strip().rstrip(",").strip()

    return result




