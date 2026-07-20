"""
schemas.py — 数据模型定义
所有 Tool 输入/输出的结构化类型，确保数据流转一致性。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


# ══════════════════════════════════════════════════════════════════════
# 配置相关
# ══════════════════════════════════════════════════════════════════════

@dataclass
class TeamMember:
    login: str
    name: str
    pol: list[str] = field(default_factory=list)


@dataclass
class Config:
    team_name: str
    my_login: str
    team_logins: list[TeamMember]
    max_emails_per_batch: int = 30
    shared_drive_path: str = r"\\ant.amazon.com\dept-as\sha11\ILS\LCL_INBOUND_DATA_ETL\IBDATACONFIRM\DATA"
    fc_address_path: str = "data/FC_Address.xlsx"
    agent_space_path: str = "data/Agent_Space.xlsx"
    seller_request_path: str = "data/Seller request list LCL.xlsx"
    blurb_root: str = "data/IB txt"
    ams_template_path: str = "data/AMS_template.xlsx"
    ens_template_path: str = "data/ENS_template.xlsx"


# ══════════════════════════════════════════════════════════════════════
# 进仓数据行
# ══════════════════════════════════════════════════════════════════════

@dataclass
class IBDataRow:
    """进仓 Excel 单行解析后的结构。"""
    al0: str
    hbl_number: str
    shipper_id: str
    pol: str
    pod: str
    bc_login: str
    flipped_fc: str
    received_cartons: float
    received_volume: float
    received_weight: float


# ══════════════════════════════════════════════════════════════════════
# Fetch Session 相关
# ══════════════════════════════════════════════════════════════════════

@dataclass
class FetchRequest:
    """单个需要浏览器 fetch 的请求。"""
    id: str
    url: str
    method: str = "GET"
    headers: dict = field(default_factory=dict)
    body: Optional[dict] = None

    def to_dict(self) -> dict:
        d = {"id": self.id, "url": self.url, "method": self.method, "headers": self.headers}
        if self.body is not None:
            d["body"] = self.body
        return d


@dataclass
class BookingCheckResult:
    """getBookingById 解析后的关键字段。"""
    al0: str
    is_cda: bool
    dst_country: str
    shipper_contact_id: str
    consignee_contact_id: str
    notify_party_contact_id: str
    importer_party_id: str  # ENS 用，可能为空
    shipper_company: str = ""
    shipper_id: str = ""


@dataclass
class PartiesAddress:
    """单个 Parties 地址条目。"""
    address_id: str
    address_name: str = ""
    contact: str = ""
    company_name: str = ""
    address_line1: str = ""
    address_line2: str = ""
    address_line3: str = ""
    city: str = ""
    state_or_region: str = ""
    postal_code: str = ""
    country_code: str = ""
    email: Optional[str] = None
    phone: str = ""


@dataclass
class DocumentInfo:
    """getFileList 返回的单个文件条目。"""
    doc_name: str
    doc_type: str  # ASI, NRA, PL, CI...
    unified_doc_id: str
    readiness_id: str
    doc_status: str
    upload_by: str = ""
    reference_id: str = ""


# ══════════════════════════════════════════════════════════════════════
# 模板填充结果
# ══════════════════════════════════════════════════════════════════════

@dataclass
class TemplateResult:
    """fill_template 的输出。"""
    al0: str
    success: bool
    template_type: str  # AMS / ENS / ASI
    output_file: str = ""  # 填充后的文件路径
    error: str = ""


# ══════════════════════════════════════════════════════════════════════
# 邮件构建
# ══════════════════════════════════════════════════════════════════════

@dataclass
class EmailParams:
    """gen_email 需要的完整邮件参数。"""
    al0: str
    subject: str
    html_body: str
    to_email: str
    cc_email: str
    attachment_path: str
    seller_request: str = ""


# ══════════════════════════════════════════════════════════════════════
# 打卡记录
# ══════════════════════════════════════════════════════════════════════

@dataclass
class EventRecord:
    """单条打卡记录。"""
    booking_id: str
    actual_time: str  # "YYYY-MM-DD HH:MM:SS"
    user_name: str
    event_code: str = "IB_DATA_EMAIL_TO_SELLER"
