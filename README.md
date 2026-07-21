# IB Data Tool

进仓数据批量发送工具 — FCL/LCL Inbound Data Confirmation Automation

## Version History

- **v2.0.0** — REST API 直接获取数据（当前版本）
  - 数据获取：直接调用 OC REST API（`getBookingById` + `getAddressInfosWithContactOnly`）
  - 认证：从 Firefox cookie 数据库读取 session token
  - 速度：~2s/单（比 v1.1 快 5-10 倍）
  - 浏览器：仅 CDA 订单 ASI 下载时按需启动
  - 邮件生成：win32com Outlook COM（不变）

- **v1.1.0** — pyautogui + Firefox 键盘流自动化
  - 浏览器控制：pyautogui 模拟键盘 + F12 Console JS 注入
  - 速度：~10-22s/单
  - 限制：需前台焦点，不能操作鼠标

## 架构（v2.0）

```
src/
├── main.py                  # 主入口 + 批次循环
├── core/
│   ├── oc_api_client.py     # OC REST API 客户端（v2.0 新增）
│   ├── browser_manager.py   # Firefox 窗口管理（仅 CDA/ASI 下载）
│   ├── input_zone_parser.py # Booking 数据结构定义
│   ├── startup_gui.py       # 启动选人 GUI
│   ├── progress_window.py   # 进度条窗口
│   ├── outlook_helper.py    # Outlook COM 邮件操作
│   └── ...
└── steps/
    ├── step1_sync.py        # 共享盘同步 → Pending List
    ├── step2_oc_scrape.py   # OC API 数据获取 + ASI 下载
    ├── step3_template_select.py  # AMS/ENS 模板选择
    ├── step4_fill_template.py    # 模板填充
    ├── step5_email.py       # 邮件生成
    └── step6_event.py       # 打卡 CSV
```

## API 端点（已验证）

| 端点 | 方法 | 用途 |
|------|------|------|
| `/aglt/rest/bookingV2/getBookingById/{al0}` | GET | Booking 详情（ODM/CDA/POL/Parties ID） |
| `/aglt/rest/getAddressInfosWithContactOnly` | POST | Parties 详情（公司/邮箱/地址） |

## 运行

```bash
cd C:\Users\miaoyua\Desktop\IB_Data_Tool
python test_full_run.py
```

## 前置条件

- Firefox 已登录 trans-logistics-cn.amazon.com（API 从 cookie 读取认证）
- Outlook 已打开
- 可访问共享盘 `\\ant.amazon.com\dept-as\...`
- （仅 CDA 订单）Firefox 允许 OC 站点弹窗
