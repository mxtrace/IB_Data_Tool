"""
preflight.py — 启动前全面自检
检查依赖库、Firefox、文件完整性、共享盘、权限、Outlook。
"""
from __future__ import annotations

from pathlib import Path

from core.config import AppConfig


def run_preflight(base_dir: Path, config: AppConfig) -> list[str]:
    """启动前自检，返回所有问题列表（空列表=全部通过）"""
    issues = []

    # 1. 依赖库检查
    required_modules = [
        ("requests", "requests"),
        ("win32com.client", "pywin32"),
        ("openpyxl", "openpyxl"),
    ]
    for mod_name, pkg_name in required_modules:
        try:
            __import__(mod_name)
        except ImportError:
            issues.append(f"缺少依赖库：{pkg_name}（import {mod_name} 失败）")

    if issues:
        return issues  # 依赖缺失则后续检查无法执行

    # 2. Firefox 存在性
    firefox_paths = [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
    ]
    if not any(Path(p).exists() for p in firefox_paths):
        issues.append("未找到 Firefox 浏览器")

    # 3. 关键文件存在性 + 可读性
    file_checks = {
        "FC_Address 表": base_dir / config.fc_address_file,
        "Seller request list": base_dir / config.seller_request_file,
        "AMS 模板": base_dir / config.template_dir / "AMS_ISF LCL.xlsx",
        "ENS 模板": base_dir / config.template_dir / "ENS LCL.xlsx",
    }
    import openpyxl
    for name, path in file_checks.items():
        if not path.exists():
            issues.append(f"文件缺失：{name}（{path}）")
        else:
            try:
                wb = openpyxl.load_workbook(str(path), read_only=True)
                wb.close()
            except Exception as e:
                issues.append(f"文件损坏或被锁定：{name} — {e}")

    # 5. Output 目录写权限
    output = base_dir / config.output_dir
    output.mkdir(exist_ok=True)
    test_file = output / ".write_test"
    try:
        test_file.write_text("test")
        test_file.unlink()
    except Exception:
        issues.append(f"Output 目录无写入权限：{output}")

    # 6. Pending List 目录
    pl_dir = base_dir / config.pending_list_dir
    pl_dir.mkdir(exist_ok=True)

    # 7. Blurb 模板至少存在一个港口
    blurb_root = base_dir / config.blurb_root
    if not blurb_root.exists():
        issues.append(f"Blurb 模板目录不存在：{blurb_root}")
    else:
        ports = [d.name for d in blurb_root.iterdir() if d.is_dir()]
        if not ports:
            issues.append("Blurb 模板目录下无港口子文件夹")

    # 8. Outlook 就绪
    try:
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application")
        ns = ol.GetNamespace("MAPI")
        _ = ns.Folders.Count
    except Exception as e:
        issues.append(f"Outlook 未就绪：{e}")

    return issues


