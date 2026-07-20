"""
config.py — 团队配置管理工具
"""
from __future__ import annotations

import json
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
CONFIG_PATH = DATA_DIR / "config.json"


def get_config() -> dict:
    """读取当前配置。"""
    if not CONFIG_PATH.exists():
        return {"error": "config.json 不存在，请先调用 init_config 初始化"}
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def init_config(
    my_login: str,
    team_name: str,
    team_logins: list,
    shared_drive_path: str = r"\\ant.amazon.com\dept-as\sha11\ILS\LCL_INBOUND_DATA_ETL\IBDATACONFIRM\DATA",
    max_emails_per_batch: int = 30,
    fc_address_path: str = "data/FC_Address.xlsx",
    agent_space_path: str = "data/Agent_Space.xlsx",
    seller_request_path: str = "data/Seller request list LCL.xlsx",
) -> dict:
    """首次初始化团队配置。"""
    if CONFIG_PATH.exists():
        return {"error": "config.json 已存在，请用 update_config 修改"}

    config = {
        "team_name": team_name,
        "my_login": my_login,
        "team_logins": team_logins,
        "max_emails_per_batch": max_emails_per_batch,
        "shared_drive_path": shared_drive_path,
        "fc_address_path": fc_address_path,
        "agent_space_path": agent_space_path,
        "seller_request_path": seller_request_path,
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"success": True, "path": str(CONFIG_PATH), "config": config}


def update_config(action: str, payload: dict) -> dict:
    """
    修改配置。
    action: add_member | remove_member | set_field
    """
    if not CONFIG_PATH.exists():
        return {"error": "config.json 不存在，请先 init_config"}

    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

    if action == "add_member":
        config["team_logins"].append(payload)
    elif action == "remove_member":
        login = payload.get("login", "")
        config["team_logins"] = [m for m in config["team_logins"] if m.get("login") != login]
    elif action == "set_field":
        for k, v in payload.items():
            config[k] = v
    else:
        return {"error": f"未知 action: {action}"}

    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"success": True, "config": config}
