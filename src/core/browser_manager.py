"""
browser_manager.py — 浏览器窗口管理（Firefox）
复用已有 Firefox 窗口（保留 OC 登录态），纯键盘流操作。
"""
from __future__ import annotations

import os
import subprocess
import time
from pathlib import Path

import pyautogui
import win32clipboard
import win32con
import win32gui
import win32process
import win32com.client

# Firefox 默认路径
FIREFOX_PATHS = [
    r"C:\Program Files\Mozilla Firefox\firefox.exe",
    r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
]

# OC 基础 URL
OC_BASE = "https://trans-logistics-cn.amazon.com/aglt/appViews/app#"

# pyautogui 设置
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.05

# WScript.Shell 用于强制切换前台窗口
_shell = win32com.client.Dispatch("WScript.Shell")


class BrowserError(Exception):
    pass


class BrowserManager:
    """管理 Firefox 窗口，复用已有登录态"""

    def __init__(self):
        self._process: subprocess.Popen | None = None
        self._hwnd: int = 0
        self._pid: int = 0
        self._oc_warmed: bool = False
        self._devtools_open: bool = False

    # ════════════════════════════════════════════════════════════════
    # 生命周期
    # ════════════════════════════════════════════════════════════════

    def start(self):
        """复用已有 Firefox 窗口（OC cookie 已存在）"""
        self._hwnd = self._find_firefox_window()
        if not self._hwnd:
            firefox_exe = self._find_firefox()
            self._process = subprocess.Popen([firefox_exe, "about:blank"])
            self._pid = self._process.pid
            time.sleep(4)
            self._hwnd = self._find_firefox_window()
        if not self._hwnd:
            self._hwnd = win32gui.GetForegroundWindow()

    def close(self):
        """释放窗口引用（不关闭 Firefox，保留用户会话）"""
        self._process = None
        self._hwnd = 0

    # ════════════════════════════════════════════════════════════════
    # 导航
    # ════════════════════════════════════════════════════════════════

    def navigate(self, url: str, wait_seconds: float = 5.0):
        """
        导航到指定 URL：
        强制前台 → 写剪贴板 → Ctrl+L → Ctrl+V → Enter → 等待
        """
        self._focus_window()
        time.sleep(0.5)

        # 将 URL 写入剪贴板
        self._set_clipboard_text(url)
        time.sleep(0.2)

        # Ctrl+L 聚焦地址栏
        pyautogui.hotkey("ctrl", "l")
        time.sleep(0.5)

        # 全选地址栏已有内容（确保覆盖）
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)

        # 粘贴 URL
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)

        # 回车导航
        pyautogui.press("enter")
        time.sleep(wait_seconds)

        # 清空剪贴板（防止 URL 泄漏到终端）
        self._clear_clipboard()

    def warmup(self):
        """预热：先访问 OC 首页，确保登录态和页面框架加载"""
        self.navigate(OC_BASE + "/dashboard", wait_seconds=5.0)
        self._oc_warmed = False  # warmup 不算首单导航

    def navigate_booking_summary(self, al0: str):
        """
        导航到 OC Booking 页面（默认显示 Booking Summary tab）。
        首次用地址栏导航，后续用 JS hash 跳转（SPA 内路由，快 3 倍）。
        """
        if not self._oc_warmed:
            # 首次：完整导航
            url = f"{OC_BASE}/bookingV2/{al0}"
            self.navigate(url, wait_seconds=5.0)
            self._oc_warmed = True
        else:
            # SPA 内跳转：JS 修改 hash + 触发 React 路由
            self._navigate_spa(al0)


    def _navigate_spa(self, al0: str):
        """
        SPA 内部路由跳转：通过 JS 修改 hash，触发 React Router。
        比地址栏方式快 ~15 秒 → ~4 秒。
        """
        js = (
            f"window.location.hash='/bookingV2/{al0}';"
            "setTimeout(function(){window.dispatchEvent(new HashChangeEvent('hashchange'))},200);"
        )
        self._run_js_in_console(js)
        time.sleep(3)  # SPA 渲染等待

    def navigate_shipment_document(self, al0: str):
        """导航到 OC Booking 页面后点击 Shipment Document tab"""
        url = f"{OC_BASE}/bookingV2/{al0}"
        self.navigate(url, wait_seconds=5.0)
        # 点击 Shipment Document tab
        self._click_tab_by_text("Shipment Document")

    def _click_tab_by_text(self, tab_text: str):
        """通过 F12 Console 执行 JS 点击页面 tab"""
        js = (
            f"document.querySelectorAll('*').forEach(function(el){{"
            f"if(el.textContent.trim()==='{tab_text}'"
            f"&&el.offsetParent!==null"
            f"&&el.children.length===0)"
            f"{{el.click()}}"
            f"}});"
        )
        self._run_js_in_console(js)
        time.sleep(3)

    def _run_js_in_console(self, js_code: str, keep_open: bool = False):
        """
        打开 Firefox DevTools Console，执行 JS。
        keep_open=True 时不关闭 DevTools（连续执行多条 JS 时提速）。
        """
        import pyautogui
        self._focus_window()
        time.sleep(0.2)

        if not self._devtools_open:
            # F12 打开 DevTools
            pyautogui.press("f12")
            time.sleep(1.5)
            # Ctrl+Shift+K 切换到 Console 面板
            pyautogui.hotkey("ctrl", "shift", "k")
            time.sleep(0.5)
            self._devtools_open = True
        else:
            # DevTools 已开，直接聚焦 Console 输入
            pyautogui.hotkey("ctrl", "shift", "k")
            time.sleep(0.3)

        # 清空 + 粘贴 + 执行
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        self._set_clipboard_text(js_code)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.5)

        if not keep_open:
            # 关闭 DevTools
            pyautogui.press("f12")
            time.sleep(0.3)
            self._devtools_open = False
        self._clear_clipboard()

    # ════════════════════════════════════════════════════════════════
    # 页面内容抓取
    # ════════════════════════════════════════════════════════════════

    def copy_page_text(self) -> str:
        """
        Ctrl+A → Ctrl+C → 读剪贴板。
        如果 DevTools 开着，先关闭以确保焦点在页面。
        """
        self._focus_window()
        time.sleep(0.2)

        # 确保 DevTools 关闭（否则 Ctrl+A 选中的是 Console）
        if self._devtools_open:
            pyautogui.press("f12")
            time.sleep(0.3)
            self._devtools_open = False

        # 确保焦点在页面内容区
        pyautogui.press("f6")
        time.sleep(0.15)

        # 清空剪贴板
        self._clear_clipboard()

        # 全选 + 复制
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.3)

        return self._get_clipboard_text()

    def copy_page_text_stable(self, min_length: int = 200, max_retries: int = 3,
                              interval: float = 1.0) -> str:
        """
        带稳定性检测：连续两次内容一致且长度足够才返回。
        interval 默认从 2.0s 降到 1.0s（SPA 页面渲染已在导航时等过）。
        """
        prev_text = ""
        for attempt in range(max_retries):
            text = self.copy_page_text()
            if text and len(text) >= min_length:
                if text == prev_text:
                    return text
                prev_text = text
            time.sleep(interval)

        return self.copy_page_text()

    # ════════════════════════════════════════════════════════════════
    # Shipment Document 操作
    # ════════════════════════════════════════════════════════════════

    def find_and_click_tab(self, tab_text: str):
        """用 Ctrl+F 查找页面文本"""
        self._focus_window()
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.3)
        pyautogui.typewrite(tab_text, interval=0.02)
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.3)
        pyautogui.press("escape")
        time.sleep(0.3)

    def press_tab_enter(self, tab_count: int = 1):
        """Tab N 次后按 Enter"""
        for _ in range(tab_count):
            pyautogui.press("tab")
            time.sleep(0.15)
        pyautogui.press("enter")
        time.sleep(0.3)

    # ════════════════════════════════════════════════════════════════
    # 窗口管理
    # ════════════════════════════════════════════════════════════════

    def _focus_window(self):
        """强制将 Firefox 窗口置为前台"""
        if not self._hwnd:
            self._hwnd = self._find_firefox_window()
        if not self._hwnd:
            return

        try:
            # 方法1：AppActivate（最可靠的前台切换）
            title = win32gui.GetWindowText(self._hwnd)
            if title:
                _shell.AppActivate(title)
                time.sleep(0.3)
                return
        except Exception:
            pass

        try:
            # 方法2：先最小化再恢复（绕过前台锁）
            if win32gui.IsIconic(self._hwnd):
                win32gui.ShowWindow(self._hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
            win32gui.SetForegroundWindow(self._hwnd)
        except Exception:
            try:
                # 方法3：Alt键解锁 + SetForegroundWindow
                pyautogui.press("alt")
                time.sleep(0.1)
                win32gui.SetForegroundWindow(self._hwnd)
            except Exception:
                pass
        time.sleep(0.3)

    def _find_firefox(self) -> str:
        for p in FIREFOX_PATHS:
            if Path(p).exists():
                return p
        raise FileNotFoundError("未找到 Firefox 浏览器")

    def _find_firefox_window(self) -> int:
        """查找已运行的 Firefox 窗口"""
        result = [0]

        def enum_ff(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if "Mozilla Firefox" in title or "Firefox" in title:
                result[0] = hwnd
                return False
            return True

        try:
            win32gui.EnumWindows(enum_ff, None)
        except Exception:
            pass
        return result[0]

    def _find_window_by_pid(self, pid: int) -> int:
        """枚举所有顶层窗口，找属于指定 PID 的"""
        result = [0]

        def enum_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            try:
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    title = win32gui.GetWindowText(hwnd)
                    if title:
                        result[0] = hwnd
                        return False
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(enum_callback, None)
        except Exception:
            pass
        return result[0]

    def refresh_hwnd(self):
        """刷新窗口句柄"""
        new_hwnd = self._find_firefox_window()
        if new_hwnd:
            self._hwnd = new_hwnd

    # ════════════════════════════════════════════════════════════════
    # 剪贴板
    # ════════════════════════════════════════════════════════════════

    @staticmethod
    def _set_clipboard_text(text: str):
        """将文本写入剪贴板"""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
            win32clipboard.CloseClipboard()
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

    @staticmethod
    def _clear_clipboard():
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

    @staticmethod
    def _get_clipboard_text() -> str:
        try:
            win32clipboard.OpenClipboard()
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return text or ""
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return ""

