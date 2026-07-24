# IB Data Tool Headless — 部署说明

## 快速部署（一键）

```powershell
# 1. Clone 项目
cd ~/Desktop
git clone https://github.com/mxtrace/IB_Data_Tool.git
cd IB_Data_Tool

# 2. 运行部署脚本
powershell -ExecutionPolicy Bypass -File deploy\setup.ps1
```

## 手动部署

### 环境要求

| 项目 | 要求 |
|------|------|
| Python | 3.13（安装到 `C:\Program Files\Python313\`） |
| Git | 任意版本 |
| Outlook | 桌面版，已配置目标共享邮箱 |
| Firefox | 已登录 trans-logistics-cn.amazon.com |
| Mars工具 | PendingList 目录存在于桌面 |
| Aki | 已安装 |

### pip 依赖

```bash
"C:\Program Files\Python313\python.exe" -m pip install requests urllib3 openpyxl xlrd pywin32 pyautogui
```

### 文件检查

以下文件不在 git 中，需从已部署机器复制：
- `模板/AMS_ISF LCL.xlsx`
- `模板/ENS LCL.xlsx`
- `FC_Address.xlsx`
- `Seller request list LCL.xlsx`

### Aki Skill 安装

部署脚本会自动复制，手动方式：
```powershell
mkdir "$env:USERPROFILE\.aki\user_preference\akisa\skills\ib-data-tool"
copy deploy\SKILL.md "$env:USERPROFILE\.aki\user_preference\akisa\skills\ib-data-tool\SKILL.md"
```

## 使用方式

对 Aki 说："跑IB Data"

Aki 会引导：选邮箱 → 处理批次 → 展示结果 → 生成打卡CSV

## 注意事项

- OC Cookie 每周一过期，周一首次运行前需 Firefox 重新登录
- Python 必须是 3.13，不能用 uv 默认的 3.14
- 此分支 (headless) 独立于 master（生产），互不影响
