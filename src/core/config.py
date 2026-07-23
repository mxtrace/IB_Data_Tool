"""
config.py — 配置管理
读取 config.json，验证必填字段。
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


class ConfigError(Exception):
    pass


@dataclass
class AppConfig:
    selected_logins: list[str] = field(default_factory=list)
    batch_size: int = 10

    # 路径常量（相对于 base_dir）
    shared_drive_path: str = r"\\ant.amazon.com\dept-as\sha11\ILS\LCL_INBOUND_DATA_ETL\IBDATACONFIRM\DATA"
    pending_list_dir: str = "Pending list"
    fc_address_file: str = "FC_Address.xlsx"
    seller_request_file: str = "Seller request list LCL.xlsx"
    loading_file: str = "Loading_V2.2.xlsm"
    blurb_root: str = "IB txt"
    template_dir: str = "模板"
    output_dir: str = "Output"
    search_stores: list[str] = field(default_factory=list)


def load_config(config_path: Path) -> AppConfig:
    """加载并验证配置文件"""
    if not config_path.exists():
        raise ConfigError(f"配置文件缺失：{config_path}")

    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise ConfigError(f"配置文件格式错误：{e}")

    config = AppConfig(
        selected_logins=raw.get("selected_logins", []),
        batch_size=raw.get("batch_size", 10),
    )

    # 允许 JSON 覆盖路径常量
    if "shared_drive_path" in raw:
        config.shared_drive_path = raw["shared_drive_path"]

    if "search_stores" in raw:
        config.search_stores = raw["search_stores"]

    if config.batch_size < 1:
        config.batch_size = 10

    return config

