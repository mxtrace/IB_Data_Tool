"""
IB Data Batch Sending Tool — QuickSuite MCP Server
供 Amazon Quick Agent 通过 browser_run_js 调用。

核心职责：
1. 生成 fetch URL 列表（Agent 在浏览器中执行）
2. 解析 fetch 结果，推进 session
3. 本地文件操作（读 Excel、填模板、生成 CSV）
"""
from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

# 确保能找到 tools/ 和 core/
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

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
    "get_config": get_config,
    "init_config": init_config,
    "update_config": update_config,
    "read_ib_data": read_ib_data,
    "start_fetch_session": start_fetch_session,
    "submit_fetch_batch": submit_fetch_batch,
    "get_fetch_session_status": get_fetch_session_status,
    "fill_template": fill_template,
    "gen_email": gen_email,
    "gen_event_csv": gen_event_csv,
}

TOOL_SCHEMAS = [
    {
        "name": "get_config",
        "description": "读取当前团队配置",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "init_config",
        "description": "首次初始化团队配置，写入 data/config.json",
        "inputSchema": {
            "type": "object",
            "properties": {
                "my_login": {"type": "string", "description": "当前用户 login"},
                "team_name": {"type": "string", "description": "团队名称"},
                "team_logins": {"type": "array", "description": "团队成员 [{login, name, pol:[]}]"},
                "shared_drive_path": {"type": "string", "description": "共享盘路径"},
                "max_emails_per_batch": {"type": "integer", "description": "单次最大邮件数，默认30"},
            },
            "required": ["my_login", "team_name", "team_logins"],
        },
    },
    {
        "name": "update_config",
        "description": "修改配置（增删成员、修改路径）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "description": "add_member|remove_member|set_field"},
                "payload": {"type": "object", "description": "操作内容"},
            },
            "required": ["action", "payload"],
        },
    },
    {
        "name": "read_ib_data",
        "description": "读取共享盘最新进仓Excel，按bc_login过滤返回AL0列表及字段数据",
        "inputSchema": {
            "type": "object",
            "properties": {
                "selected_logins": {"type": "array", "description": "本次处理的bc_login列表"},
                "max_count": {"type": "integer", "description": "最大条数"},
            },
            "required": ["selected_logins"],
        },
    },
    {
        "name": "start_fetch_session",
        "description": "启动Fetch Session，返回第一批需要浏览器fetch的URL列表。Agent需用browser_run_js执行这些URL。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "al0_list": {"type": "array", "description": "待处理的AL0列表"},
                "ib_data": {"type": "object", "description": "进仓数据(al0→字段)"},
            },
            "required": ["al0_list"],
        },
    },
    {
        "name": "submit_fetch_batch",
        "description": "提交浏览器fetch结果，获取下一批URL或最终结果",
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string", "description": "session ID"},
                "results": {"type": "object", "description": "{url_id: response_json}"},
            },
            "required": ["session_id", "results"],
        },
    },
    {
        "name": "get_fetch_session_status",
        "description": "查询Fetch Session当前进度",
        "inputSchema": {
            "type": "object",
            "properties": {
                "session_id": {"type": "string", "description": "session ID"},
            },
            "required": ["session_id"],
        },
    },
    {
        "name": "fill_template",
        "description": "填充AMS/ENS/ASI Excel模板（数量+收发通）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "al0": {"type": "string"},
                "template_type": {"type": "string", "description": "AMS|ENS|ASI"},
                "ib_row": {"type": "object", "description": "进仓数据行"},
                "parties_data": {"type": "object", "description": "Parties地址数据（非CDA时）"},
                "asi_file_path": {"type": "string", "description": "ASI文件路径（CDA时）"},
            },
            "required": ["al0", "template_type", "ib_row"],
        },
    },
    {
        "name": "gen_email",
        "description": "组装邮件并通过Outlook COM弹窗",
        "inputSchema": {
            "type": "object",
            "properties": {
                "al0": {"type": "string"},
                "ib_row": {"type": "object", "description": "进仓数据行"},
                "template_file": {"type": "string", "description": "填充完的模板路径"},
                "recipient_email": {"type": "string", "description": "收件人邮箱"},
                "seller_request": {"type": "string", "description": "特殊需求文本"},
            },
            "required": ["al0", "ib_row", "template_file"],
        },
    },
    {
        "name": "gen_event_csv",
        "description": "生成打卡CSV文件(IB_DATA_EMAIL_TO_SELLER)",
        "inputSchema": {
            "type": "object",
            "properties": {
                "sent_records": {"type": "array", "description": "[{al0, actual_time, user_name}]"},
            },
            "required": ["sent_records"],
        },
    },
]


# ══════════════════════════════════════════════════════════════════════
# JSON-RPC over stdin/stdout（MCP 协议）
# ══════════════════════════════════════════════════════════════════════

def handle_request(request: dict) -> dict:
    method = request.get("method", "")
    params = request.get("params", {})
    req_id = request.get("id")

    # initialize
    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": "ib-data-tool", "version": "1.0.0"},
            },
        }

    # notifications（不需要响应）
    if method.startswith("notifications/"):
        return None

    # tools/list
    if method == "tools/list":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {"tools": TOOL_SCHEMAS},
        }

    # tools/call
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
                "result": {
                    "content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False, default=str)}]
                },
            }
        except Exception as e:
            error_detail = traceback.format_exc()
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": json.dumps({"error": str(e), "traceback": error_detail}, ensure_ascii=False)}],
                    "isError": True,
                },
            }

    # 未知方法
    return {
        "jsonrpc": "2.0",
        "id": req_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def main():
    """主循环：stdin 读 JSON-RPC，stdout 写响应。"""
    for line in iter(sys.stdin.buffer.readline, b""):
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line.decode("utf-8"))
        except json.JSONDecodeError:
            continue

        response = handle_request(request)
        if response is None:
            continue

        out = json.dumps(response, ensure_ascii=False) + "\n"
        sys.stdout.buffer.write(out.encode("utf-8"))
        sys.stdout.buffer.flush()


if __name__ == "__main__":
    main()
