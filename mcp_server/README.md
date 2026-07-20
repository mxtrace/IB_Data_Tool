# IB Data Batch Sending Tool — MCP Server

## QuickSuite 配置

### 1. 在 QuickSuite 中注册 MCP Server

进入 Quick Suite → Agents & Skills → Skills → Create MCP Skill

填写：
- **Name**: `ib-data-tool`
- **Transport**: `stdio`
- **Command**: `python`
- **Args**: `C:\Users\miaoyua\Desktop\IB_Data_Tool\mcp_server\quick_server.py`

### 2. 测试连接

在 Quick Agent 对话中输入：
```
调用 ib-data-tool 的 get_config 工具
```

预期返回：`{"error": "config.json 不存在，请先调用 init_config 初始化"}`

### 3. 首次初始化

```
调用 init_config，参数：
  my_login: "miaoyua"
  team_name: "CN Ops CTU RegionalSAM"
  team_logins: [{"login": "miaoyua", "name": "苗园", "pol": ["CNSHA", "CNYTN"]}]
```

### 4. 完整工作流

```
1. get_config → 确认配置
2. read_ib_data(selected_logins=["miaoyua"]) → 获取 AL0 列表
3. start_fetch_session(al0_list=[...]) → 返回 URL 列表
4. Agent 用 browser_run_js 执行 fetch
5. submit_fetch_batch(session_id, results) → 下一批或 DONE
6. fill_template(...) → 填充模板
7. gen_email(...) → Outlook 弹窗
8. gen_event_csv(...) → 打卡 CSV
```

## 文件结构

```
mcp_server/
├── quick_server.py          # QuickSuite MCP 入口（用这个）
├── server.py                # 通用 MCP 入口（备用）
├── core/
│   ├── schemas.py           # 数据模型
│   └── session_store.py     # Fetch Session 管理
├── tools/
│   ├── config.py            # 配置管理
│   ├── ib_data.py           # 进仓数据读取
│   ├── fetch_session.py     # Fetch Session 工具
│   ├── template.py          # 模板填充
│   ├── email_builder.py     # 邮件构建
│   └── event.py             # 打卡 CSV
└── data/
    ├── config.json           # 团队配置（init后生成）
    ├── FC_Address.xlsx       # FC地址表（需手动放入）
    ├── Seller request list LCL.xlsx
    ├── IB txt/               # Blurb 模板
    ├── 模板/                 # AMS/ENS Excel 模板
    ├── output/               # 填充后的文件
    └── fetch_cache/          # Session 持久化
```

## 依赖

- Python 3.10+
- openpyxl（Excel 操作）
- pywin32（Outlook COM，仅 Windows）
