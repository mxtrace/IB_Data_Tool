# IB Data Tool

进仓数据批量发送工具 — FCL/LCL Inbound Data Confirmation Automation

## Version

- **v1.1.0** — pyautogui + Firefox 键盘流自动化（当前版本）
  - 浏览器控制：pyautogui 模拟键盘 + F12 Console JS 注入
  - 邮件生成：win32com Outlook COM
  - GUI：tkinter 弹窗/进度窗口

## 架构

```
src/
├── main.py              # 主入口 + 批次循环
├── core/
│   ├── browser_manager.py   # Firefox 窗口管理（pyautogui）
│   ├── startup_gui.py       # 启动选人 GUI
│   ├── progress_window.py   # 进度条窗口
│   ├── input_zone_parser.py # Booking Summary 文本解析
│   ├── outlook_helper.py    # Outlook COM 邮件操作
│   └── ...
└── steps/
    ├── step1_sync.py        # 共享盘同步 → Pending List
    ├── step2_oc_scrape.py   # OC 页面抓取 + ASI下载
    ├── step3_template_select.py  # AMS/ENS 模板选择
    ├── step4_fill_template.py    # 模板填充
    ├── step5_email.py       # 邮件生成
    └── step6_event.py       # 打卡 CSV
```

## 运行

```bash
cd C:\Users\miaoyua\Desktop\IB_Data_Tool
python test_full_run.py
```

## 前置条件

- Firefox 已登录 trans-logistics-cn.amazon.com
- Firefox 允许该站点弹窗
- Outlook 已打开
- 可访问共享盘 `\\ant.amazon.com\dept-as\...`
