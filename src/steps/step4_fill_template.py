"""
step4_fill_template.py — 填充模板（数量 + 收发通）
使用 excel_helper 安全管理文件句柄。
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from core.excel_helper import open_workbook, open_workbook_write
from core.input_zone_parser import InputZoneData
from core.logger import audit


@dataclass
class FillResult:
    output_file: str = ""
    error: str = ""


def fill_template(
    al0: str,
    template_type: str,
    ib_row: dict,
    input_zone: InputZoneData,
    base_dir: Path,
    asi_file: str | None = None,
) -> FillResult:
    """填充模板，返回输出文件路径"""
    output_dir = base_dir / "Output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # 4.0 前置检查：数量字段
    cartons = ib_row.get("received_cartons", 0)
    volume = ib_row.get("received_volume", 0)
    weight = ib_row.get("received_weight", 0)
    if not cartons and not volume and not weight:
        return FillResult(error="进仓数量数据缺失（件/体/重全为空）")

    # 4.1 确定输出文件
    if template_type == "ASI":
        if not asi_file or not Path(asi_file).exists():
            return FillResult(error="ASI 文件路径无效")
        output_file = Path(asi_file)
    elif template_type == "AMS":
        src = base_dir / "模板" / "AMS_ISF LCL.xlsx"
        output_file = output_dir / f"AMS_ISF LCL_{al0}.xlsx"
        shutil.copy2(src, output_file)
    else:  # ENS
        src = base_dir / "模板" / "ENS LCL.xlsx"
        output_file = output_dir / f"ENS LCL_{al0}.xlsx"
        shutil.copy2(src, output_file)

    # 4.2 填充数量
    try:
        _write_quantities(output_file, ib_row)
    except Exception as e:
        return FillResult(error=f"数量填充失败：{e}")

    # ASI 到此结束
    if template_type == "ASI":
        audit(al0, "step4", "success", "template=ASI, qty_only")
        return FillResult(output_file=str(output_file))

    # 4.3 FC 地址匹配
    flipped_fc = ib_row.get("flipped_fc", "")
    fc_addr = _match_fc_address(flipped_fc, base_dir / "FC_Address.xlsx")
    if not fc_addr:
        return FillResult(error=f"FC代码 {flipped_fc} 在 FC_Address 表中无匹配")

    # 4.4 填充收发通
    try:
        _write_parties(output_file, template_type, input_zone, fc_addr)
    except Exception as e:
        return FillResult(error=f"收发通填充失败：{e}")

    audit(al0, "step4", "success", f"template={template_type}")
    return FillResult(output_file=str(output_file))


# ══════════════════════════════════════════════════════════════════════

def _write_quantities(file_path: Path, ib_row: dict):
    """Row 23: C=件数, D=CARTON, E=重量(KG), F=体积(CBM)"""
    if file_path.suffix.lower() == ".xls":
        _write_quantities_com(file_path, ib_row)
    else:
        with open_workbook_write(file_path) as wb:
            ws = wb.active
            ws["C23"] = ib_row.get("received_cartons", 0)
            ws["D23"] = "CARTON"
            ws["E23"] = ib_row.get("received_weight", 0)
            ws["F23"] = ib_row.get("received_volume", 0)


def _write_quantities_com(file_path: Path, ib_row: dict):
    """用 Excel COM 写入 .xls 文件（保留原始格式）"""
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(file_path.resolve()))
        ws = wb.Sheets(1)

        ws.Range("C23").Value = ib_row.get("received_cartons", 0)
        ws.Range("D23").Value = "CARTON"
        ws.Range("E23").Value = ib_row.get("received_weight", 0)
        ws.Range("F23").Value = ib_row.get("received_volume", 0)

        wb.Save()
        wb.Close(False)
    finally:
        excel.Quit()


def _write_parties(file_path: Path, ttype: str, iz: InputZoneData, fc: dict):
    """填充 Shipper / Consignee / Notify / Buyer(ENS)"""
    with open_workbook_write(file_path) as wb:
        ws = wb.active

        # Shipper (Row 7-10) ← Input Zone
        ws["B7"] = iz.shipper.company
        ws["B8"] = iz.shipper.street
        ws["B9"] = iz.shipper.city
        ws["D9"] = iz.shipper.state
        ws["B10"] = iz.shipper.zip
        ws["D10"] = iz.shipper.country

        # Consignee (Row 12-15) ← company from IZ, address from FC
        company = iz.consignee.company
        if company and "C/O FBA" not in company.upper():
            company = f"{company}\nC/O FBA"
        ws["B12"] = company
        ws["B13"] = fc.get("address", "")
        ws["B14"] = fc.get("city", "")
        ws["D14"] = fc.get("state", "")
        ws["B15"] = fc.get("postal_code", "")
        ws["D15"] = fc.get("country", "")

        # Notify Party (Row 17-20) ← company from IZ, address from FC (同Consignee)
        ws["B17"] = iz.notify.company
        ws["B18"] = fc.get("address", "")
        ws["B19"] = fc.get("city", "")
        ws["D19"] = fc.get("state", "")
        ws["B20"] = fc.get("postal_code", "")
        ws["D20"] = fc.get("country", "")

        # Buyer/DIE (ENS only, G12-I15) ← Input Zone DIE
        if ttype == "ENS":
            ws["G12"] = iz.die.company
            ws["G13"] = iz.die.street
            ws["G14"] = iz.die.city
            ws["I14"] = iz.die.state
            ws["G15"] = iz.die.zip
            ws["I15"] = iz.die.country


def _match_fc_address(flipped_fc: str, fc_path: Path) -> dict | None:
    """匹配 FC_Address 表 A列，返回地址字段"""
    if not flipped_fc or not fc_path.exists():
        return None

    with open_workbook(fc_path, read_only=True, data_only=True) as wb:
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[0] or "").strip().upper() == flipped_fc.upper():
                return {
                    "company_name": str(row[2] or "").strip(),
                    "address": str(row[3] or "").strip(),
                    "city": str(row[4] or "").strip(),
                    "state": str(row[5] or "").strip(),
                    "postal_code": str(row[6] or "").strip() if row[6] else "",
                    "country": str(row[7] or "").strip(),
                }
    return None


