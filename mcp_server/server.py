"""
IB Data Batch Sending Tool — MCP Server 入口
JSON-RPC over stdin/stdout，供 Amazon Quick Agent 调用。

职责：
- 接收 Agent 请求，路由到对应 tool
- 不直接执行 HTTP 请求（合规：由浏览器 fetch 执行）
- 管理 Fetch Session 生命周期
"""
from __future__ import annotations

import json
import sys
import traceback
from typing import Any

from tools.config import get_config, init_config, update_config
from tools.ib_data import read_ib_data
from tools.fetch_session import start_fetch_session, submit_fetch_batch, get_fetch_session_status
from tools.template import fill_template
from tools.email_builder import gen_email
from tools.event import gen_event_csv

# ══════════════════════════════════════════════════════════════════════
# Tool Registry
# ══════════════════════════════════════════════════════════════════════

TOOLS = {
    # 配置管理
    "get_config": get_config,
    "init_config": init_config,
    "update_config": update_config,
    # 进仓数据
    "read_ib_data": read_ib_data,
    # Fetch Session（API 数据获取）
    "start_fetch_session": start_fetch_session,
    "submit_fetch_batch": submit_fetch_batch,
    "get_fetch_session_status": get_fetch_session_status,
    # 模板填充
    "fill_template": fill_template,
    # 邮件生成
    "gen_email": gen_email,
    # 打卡 CSV
    "gen_event_csv": gen_event_csv,
}

# Tool 元信息（供 MCP Client 发现）
TOOL_SCHEMAS = {
    "get_config": {
        "description": "读取当前团队配置（team_logins, max_emails_per_batch, paths）",
        "parameters": {},
    },
    "init_config": {
        "description": "首次初始化团队配置，写入 data/config.json",
        "parameters": {
            "my_login": {"type": "string", "description": "当前用户 login", "required": True},
            "team_name": {"type": "string", "description": "团队名称", "required": True},
            "team_logins": {
                "type": "array",
                "description": "团队成员列表 [{login, name, pol:[]}]",
                "required": True,
            },
            "shared_drive_path": {"type": "string", "description": "共享盘路径", "required": False},
            "max_emails_per_batch": {"type": "integer", "description": "单次最大邮件数", "required": False},
        },
    },
    "update_config": {
        "description": "修改已有配置（增删成员、修改路径等）",
        "parameters": {
            "action": {"type": "string", "description": "add_member|remove_member|set_field", "required": True},
            "payload": {"type": "object", "description": "操作内容", "required": True},
        },
    },
    "read_ib_data": {
        "description": "读取共享盘最新进仓 Excel，按 bc_login 过滤并返回 AL0 列表及字段数据",
        "parameters": {
            "selected_logins": {"type": "array", "description": "本次处理的 bc_login 列表", "required": True},
            "max_count": {"type": "integer", "description": "最大条数（截断）", "required": False},
        },
    },
    "start_fetch_session": {
        "description": "启动 Fetch Session，返回第一批需要浏览器 fetch 的 URL 列表",
        "parameters": {
            "al0_list": {"type": "array", "description": "待处理的 AL0 列表", "required": True},
            "ib_data": {"type": "object", "description": "进仓数据（al0→字段映射）", "required": True},
        },
    },
    "submit_fetch_batch": {
        "description": "提交浏览器 fetch 结果，获取下一批 URL 或最终结果",
        "parameters": {
            "session_id": {"type": "string", "description": "session 唯一 ID", "required": True},
            "results": {"type": "object", "description": "{url_id: response_json}", "required": True},
        },
    },
    "get_fetch_session_status": {
        "description": "查询 Fetch Session 当前进度",
        "parameters": {
            "session_id": {"type": "string", "description": "session ID", "required": True},
        },
    },
    "fill_template": {
        "description": "根据模板类型填充 AMS/ENS/ASI Excel 模板",
        "parameters": {
            "al0": {"type": "string", "required": True},
            "template_type": {"type": "string", "description": "AMS|ENS|ASI", "required": True},
            "ib_row": {"type": "object", "description": "该 AL0 的进仓数据行", "required": True},
            "parties_data": {"type": "object", "description": "API 返回的 Parties 地址数据", "required": False},
            "asi_file_path": {"type": "string", "description": "ASI 下载文件路径（仅CDA）", "required": False},
        },
    },
    "gen_email": {
        "description": "组装邮件参数（主题、正文、收件人、附件），通过 Outlook COM 弹窗",
        "parameters": {
            "al0": {"type": "string", "required": True},
            "ib_row": {"type": "object", "description": "进仓数据行", "required": True},
            "template_file": {"type": "string", "description": "填充完成的模板文件路径", "required": True},
            "recipient_email": {"type": "string", "description": "收件人邮箱", "required": False},
            "seller_request": {"type": "string", "description": "特殊需求文本", "required": False},
        },
    },
    "gen_event_csv": {
        "description": "生成打卡 CSV 文件（IB_DATA_EMAIL_TO_SELLER）",
        "parameters": {
            "sent_records": {
                "type": "array",
                "description": "[{al0, actual_time, user_name}]",
                "required": True,
            },
        },
    },
}


# ══════════════════════════════════════════════════════════════════════
# JSON-RPC 通信层
# ══════════════════════════════════════════════════════════════════════

def handle_request(request: dict) -> dict:
    """处理单个 JSON-RPC 请求。"""
    method = request.get("method", "")
    params = request.get("params", {})
    req_id = request.get("id")

    # MCP 协议：tools/list
    if method == "tools/list":
        tools_list = []
        for name, schema in TOOL_SCHEMAS.items():
            tools_list.append({
                "name": name,
                "description": schema["description"],
                "inputSchema": {
                    "type": "object",
                    "properties": schema.get("parameters", {}),
                },
            })
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": tools_list}}

    # MCP 协议：tools/call
    if method == "tools/call":
        tool_name = params.get("name", "")
        tool_args = params.get("arguments", {})

        if tool_name not in TOOLS:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"},
            }

        try:
            result = TOOLS[tool_name](**tool_args)
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {"content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False)}]},
            }
        except Exception as e:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32000, "message": str(e), "data": traceback.format_exc()},
            }

    # 未知方法
    return {
        "jsonrpc": "2.0",
        "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def main():
    """主循环：从 stdin 读取 JSON-RPC 请求，写响应到 stdout。"""
    # Windows pipe EOF 修复：逐行读取
    for line in iter(sys.stdin.buffer.readline, b""):
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line.decode("utf-8"))
        except json.JSONDecodeError:
            continue

        response = handle_request(request)
        out = json.dumps(response, ensure_ascii=False) + "\n"
        sys.stdout.buffer.write(out.encode("utf-8"))
        sys.stdout.buffer.flush()


if __name__ == "__main__":
    main()
