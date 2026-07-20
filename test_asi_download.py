"""
ASI 下载 v8：纯 JS 方案（需先允许 OC 网站弹窗）
单次 Console 执行：checkbox勾选 + 等待 + 点击 Download
"""
import sys
import time
sys.path.insert(0, "src")

from core.browser_manager import BrowserManager
from steps.step2_oc_scrape import _snapshot_downloads, _wait_for_new_download, DOWNLOAD_DIR
import win32gui

AL0 = "AL0-VQGFKGNKSXKNS"

print("=" * 50)
print(f"ASI 下载 v8（纯JS）：{AL0}")
print("=" * 50)
print("前提：Firefox 已允许 trans-logistics-cn.amazon.com 弹窗")

browser = BrowserManager()
browser.start()

# 1. 导航 + 切 tab
print("\n[1] 导航...")
url = f"https://trans-logistics-cn.amazon.com/aglt/appViews/app#/bookingV2/{AL0}"
browser.navigate(url, wait_seconds=6.0)

print("[2] 切换到 Shipment Document...")
browser._click_tab_by_text("Shipment Document")
time.sleep(3)

# 2. 单次 JS：勾选 ASI + 点击 Download（全部在一次 Console 执行中）
print("[3] 单次 JS：勾选 + 点 Download（3秒后）...")
time.sleep(3)

before = _snapshot_downloads()

combined_js = (
    # 找 ASI 行并勾选
    "var asiSpan=null;"
    "document.querySelectorAll('span').forEach(function(s){"
    "  if(s.textContent.trim()==='ASI'&&s.children.length===0) asiSpan=s;"
    "});"
    "if(asiSpan){"
    "  var p=asiSpan;"
    "  for(var i=0;i<10;i++){"
    "    p=p.parentElement;"
    "    if(!p)break;"
    "    var cbs=p.querySelectorAll('input[type=checkbox]');"
    "    if(cbs.length===1){"
    "      var label=p.querySelector('label');"
    "      if(label){label.click()}"
    "      else{cbs[0].click()}"
    "      break;"
    "    }"
    "  }"
    "}"
    # 等 React 更新后点击 Download
    "setTimeout(function(){"
    "  var btns=document.querySelectorAll('button');"
    "  for(var i=0;i<btns.length;i++){"
    "    if(btns[i].textContent.trim()==='Download'&&!btns[i].disabled){"
    "      btns[i].click();"
    "      document.title='DL_CLICKED';"
    "      break;"
    "    }"
    "  }"
    "  if(document.title!=='DL_CLICKED'){"
    "    document.title='DL_STILL_DISABLED';"
    "  }"
    "},1000);"
)
browser._run_js_in_console(combined_js)
time.sleep(3)  # 等 setTimeout + 下载触发

title = win32gui.GetWindowText(browser._hwnd)
print(f"    结果: {title}")

# 3. 等待下载
print("\n[4] 等待下载...")
downloaded = _wait_for_new_download(before, timeout=30)
if downloaded:
    print(f"    ✅ 下载成功: {downloaded.name} ({downloaded.stat().st_size} bytes)")
else:
    print("    ❌ 下载超时")
    import os
    after = set(os.listdir(DOWNLOAD_DIR))
    new = after - before
    if new:
        print(f"    新文件: {new}")
    else:
        print("    无新文件")
        if "DL_CLICKED" in title:
            print("    按钮已点击但无下载 → 可能弹窗仍被阻止")
            print("    请确认已在 Firefox 设置中允许 trans-logistics-cn.amazon.com 弹窗")
        elif "DISABLED" in title:
            print("    Download 按钮仍 disabled → checkbox 未生效")

print("\n测试完成")
