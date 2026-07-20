"""
session_store.py — Fetch Session 状态管理
管理 session 生命周期：创建、推进 phase、持久化、恢复。
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Optional

from core.schemas import FetchRequest, BookingCheckResult, DocumentInfo

# API 基础 URL
API_BASE = "https://trans-logistics-cn.amazon.com"

# Phase 定义
PHASE_BOOKING_CHECK = "BOOKING_CHECK"
PHASE_PARTIES_FETCH = "PARTIES_FETCH"
PHASE_FILE_LIST = "FILE_LIST"
PHASE_ASI_DOWNLOAD = "ASI_DOWNLOAD"
PHASE_DONE = "DONE"

BATCH_SIZE = 15  # 每批最大 URL 数


class FetchSession:
    """一次完整的 IB Data Fetch 流程。"""

    def __init__(self, session_id: str = ""):
        self.session_id = session_id or uuid.uuid4().hex[:8]
        self.phase = PHASE_BOOKING_CHECK

        # 输入
        self.al0_list: list[str] = []
        self.ib_data: dict = {}  # al0 → IBDataRow dict

        # Phase 结果累积
        self.booking_results: dict[str, BookingCheckResult] = {}  # al0 → result
        self.parties_data: dict[str, dict] = {}  # al0 → {contactId: PartiesAddress}
        self.file_lists: dict[str, list[DocumentInfo]] = {}  # al0 → [docs]
        self.asi_files: dict[str, str] = {}  # al0 → local file path

        # 分流
        self.cda_al0s: list[str] = []
        self.non_cda_al0s: list[str] = []

        # 待 fetch 队列
        self.pending_urls: list[FetchRequest] = []

        # 错误记录
        self.errors: dict[str, str] = {}  # al0 → error msg
        self.retry_counts: dict[str, int] = {}  # url_id → retry count

    # ─────────────────────────────────────────────────────────────
    # Phase 1: BOOKING_CHECK
    # ─────────────────────────────────────────────────────────────

    def generate_booking_check_urls(self) -> None:
        """为每个 AL0 生成 getBookingById URL。"""
        self.pending_urls = []
        for al0 in self.al0_list:
            self.pending_urls.append(FetchRequest(
                id=f"booking_{al0}",
                url=f"{API_BASE}/aglt/v2/api/getBookingById/{al0}",
                method="GET",
                headers={"Accept": "application/json"},
            ))

    def process_booking_check(self, results: dict) -> None:
        """解析 getBookingById 响应，提取 CDA 状态和 contactIds。"""
        for url_id, response in results.items():
            if not url_id.startswith("booking_"):
                continue
            al0 = url_id.replace("booking_", "")

            if isinstance(response, dict) and "error" in response:
                self.errors[al0] = response["error"]
                continue

            # 解析 data 字段（API 返回 {status, data:{...}}）
            data = response.get("data", response)

            is_cda = bool(data.get("isCDAEnabled", False))
            result = BookingCheckResult(
                al0=al0,
                is_cda=is_cda,
                dst_country=data.get("dstCountry", ""),
                shipper_contact_id=data.get("shipperContactId", ""),
                consignee_contact_id=data.get("consigneeContactId", ""),
                notify_party_contact_id=data.get("notifyPartyContactId", ""),
                importer_party_id=data.get("importerParty", ""),
                shipper_company=data.get("shipperCompany", ""),
                shipper_id=data.get("shipperId", ""),
            )
            self.booking_results[al0] = result

            if is_cda:
                self.cda_al0s.append(al0)
            else:
                self.non_cda_al0s.append(al0)

    # ─────────────────────────────────────────────────────────────
    # Phase 2: PARTIES_FETCH（非CDA单）
    # ─────────────────────────────────────────────────────────────

    def generate_parties_urls(self) -> None:
        """为非CDA单生成 getAddressInfosWithContactOnly 请求。"""
        self.pending_urls = []
        for al0 in self.non_cda_al0s:
            br = self.booking_results.get(al0)
            if not br:
                continue
            # 收集所有有效的 contactId
            address_ids = [
                cid for cid in [
                    br.shipper_contact_id,
                    br.consignee_contact_id,
                    br.notify_party_contact_id,
                    br.importer_party_id,
                ] if cid
            ]
            if not address_ids:
                self.errors[al0] = "无 contactId，无法获取 Parties 数据"
                continue

            self.pending_urls.append(FetchRequest(
                id=f"parties_{al0}",
                url=f"{API_BASE}/aglt/rest/getAddressInfosWithContactOnly",
                method="POST",
                headers={"Content-Type": "application/json", "Accept": "application/json"},
                body={"addressIds": address_ids, "contactOnlyAddressIds": []},
            ))

    def process_parties_fetch(self, results: dict) -> None:
        """解析 Parties 响应。"""
        for url_id, response in results.items():
            if not url_id.startswith("parties_"):
                continue
            al0 = url_id.replace("parties_", "")

            if isinstance(response, dict) and "error" in response:
                self.errors[al0] = response["error"]
                continue

            # 响应是 {contactId: addressObject} 的 map
            self.parties_data[al0] = response

    # ─────────────────────────────────────────────────────────────
    # Phase 3: FILE_LIST（CDA单）
    # ─────────────────────────────────────────────────────────────

    def generate_file_list_urls(self) -> None:
        """为CDA单生成 getFileList 请求。"""
        self.pending_urls = []
        for al0 in self.cda_al0s:
            self.pending_urls.append(FetchRequest(
                id=f"filelist_{al0}",
                url=f"{API_BASE}/aglt/rest/shipmentDocument/getFileList",
                method="POST",
                headers={"Content-Type": "application/json", "Accept": "application/json"},
                body={"bookingId": al0},
            ))

    def process_file_list(self, results: dict) -> None:
        """解析 getFileList 响应，筛选 ASI 文档。"""
        for url_id, response in results.items():
            if not url_id.startswith("filelist_"):
                continue
            al0 = url_id.replace("filelist_", "")

            if isinstance(response, dict) and "error" in response:
                self.errors[al0] = response["error"]
                continue

            doc_list = response.get("documentInfoList", [])
            asi_docs = [d for d in doc_list if d.get("docType") == "ASI"]
            if not asi_docs:
                self.errors[al0] = "CDA订单但无 ASI 文件"
                self.cda_al0s.remove(al0)
                continue

            self.file_lists[al0] = [
                DocumentInfo(
                    doc_name=d.get("docName", ""),
                    doc_type=d.get("docType", ""),
                    unified_doc_id=d.get("unifiedDocId", ""),
                    readiness_id=d.get("readinessId", ""),
                    doc_status=d.get("docStatus", ""),
                    upload_by=d.get("uploadBy", ""),
                    reference_id=d.get("referenceId", ""),
                )
                for d in asi_docs
            ]

    # ─────────────────────────────────────────────────────────────
    # Phase 4: ASI_DOWNLOAD（CDA单）
    # ─────────────────────────────────────────────────────────────

    def generate_asi_download_urls(self) -> None:
        """为CDA单生成 ASI 下载请求。"""
        self.pending_urls = []
        for al0 in self.cda_al0s:
            docs = self.file_lists.get(al0, [])
            if not docs:
                continue
            # 取第一个 ASI 文档
            doc = docs[0]
            self.pending_urls.append(FetchRequest(
                id=f"asidown_{al0}",
                url=f"{API_BASE}/aglt/rest/shipmentDocument/download",
                method="POST",
                headers={"Content-Type": "application/json"},
                body={"unifiedDocId": doc.unified_doc_id},  # 参数待验证
            ))

    def process_asi_download(self, results: dict) -> None:
        """处理 ASI 下载结果（浏览器返回 base64 或 blob URL）。"""
        for url_id, response in results.items():
            if not url_id.startswith("asidown_"):
                continue
            al0 = url_id.replace("asidown_", "")

            if isinstance(response, dict) and "error" in response:
                self.errors[al0] = f"ASI 下载失败: {response['error']}"
                continue

            # TODO: 浏览器端需要特殊处理二进制下载
            # 暂存文件路径标记，实际下载逻辑待实现
            self.asi_files[al0] = f"temp/{al0}_ASI.xls"

    # ─────────────────────────────────────────────────────────────
    # Phase 推进
    # ─────────────────────────────────────────────────────────────

    def get_next_batch(self) -> list[dict]:
        """取下一批 URL（≤ BATCH_SIZE 个）。"""
        batch = self.pending_urls[:BATCH_SIZE]
        self.pending_urls = self.pending_urls[BATCH_SIZE:]
        return [req.to_dict() for req in batch]

    def advance_phase(self) -> None:
        """当前 phase 的 pending_urls 为空时，推进到下一阶段。"""
        if self.pending_urls:
            return  # 当前 phase 还有未处理的 URL

        if self.phase == PHASE_BOOKING_CHECK:
            # 分流完成，同时启动 PARTIES + FILE_LIST
            # 先处理非CDA（PARTIES），再处理CDA（FILE_LIST）
            if self.non_cda_al0s:
                self.phase = PHASE_PARTIES_FETCH
                self.generate_parties_urls()
            elif self.cda_al0s:
                self.phase = PHASE_FILE_LIST
                self.generate_file_list_urls()
            else:
                self.phase = PHASE_DONE

        elif self.phase == PHASE_PARTIES_FETCH:
            # Parties 完成，看是否还有 CDA 单需要获取文件列表
            if self.cda_al0s:
                self.phase = PHASE_FILE_LIST
                self.generate_file_list_urls()
            else:
                self.phase = PHASE_DONE

        elif self.phase == PHASE_FILE_LIST:
            # 文件列表完成，下载 ASI
            if self.cda_al0s and self.file_lists:
                self.phase = PHASE_ASI_DOWNLOAD
                self.generate_asi_download_urls()
            else:
                self.phase = PHASE_DONE

        elif self.phase == PHASE_ASI_DOWNLOAD:
            self.phase = PHASE_DONE

    def process_results(self, results: dict) -> None:
        """根据当前 phase 分发结果处理。"""
        if self.phase == PHASE_BOOKING_CHECK:
            self.process_booking_check(results)
        elif self.phase == PHASE_PARTIES_FETCH:
            self.process_parties_fetch(results)
        elif self.phase == PHASE_FILE_LIST:
            self.process_file_list(results)
        elif self.phase == PHASE_ASI_DOWNLOAD:
            self.process_asi_download(results)

    # ─────────────────────────────────────────────────────────────
    # 持久化（断点续传）
    # ─────────────────────────────────────────────────────────────

    def save(self, data_dir: Path) -> None:
        """持久化 session 到 JSON 文件。"""
        cache_dir = data_dir / "fetch_cache"
        cache_dir.mkdir(parents=True, exist_ok=True)
        path = cache_dir / f"{self.session_id}.json"

        state = {
            "session_id": self.session_id,
            "phase": self.phase,
            "al0_list": self.al0_list,
            "cda_al0s": self.cda_al0s,
            "non_cda_al0s": self.non_cda_al0s,
            "booking_results": {k: v.__dict__ for k, v in self.booking_results.items()},
            "parties_data": self.parties_data,
            "file_lists": {k: [d.__dict__ for d in v] for k, v in self.file_lists.items()},
            "asi_files": self.asi_files,
            "errors": self.errors,
        }
        path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    @classmethod
    def load(cls, session_id: str, data_dir: Path) -> Optional["FetchSession"]:
        """从文件恢复 session。"""
        path = data_dir / "fetch_cache" / f"{session_id}.json"
        if not path.exists():
            return None

        state = json.loads(path.read_text(encoding="utf-8"))
        session = cls(session_id=state["session_id"])
        session.phase = state["phase"]
        session.al0_list = state["al0_list"]
        session.cda_al0s = state["cda_al0s"]
        session.non_cda_al0s = state["non_cda_al0s"]
        session.parties_data = state.get("parties_data", {})
        session.asi_files = state.get("asi_files", {})
        session.errors = state.get("errors", {})
        # booking_results 和 file_lists 需要重建 dataclass
        # 简化处理：保持为 dict
        session.booking_results = state.get("booking_results", {})
        session.file_lists = state.get("file_lists", {})
        return session

    # ─────────────────────────────────────────────────────────────
    # 汇总
    # ─────────────────────────────────────────────────────────────

    def get_summary(self) -> dict:
        """返回最终汇总结果。"""
        return {
            "total": len(self.al0_list),
            "cda_count": len(self.cda_al0s),
            "non_cda_count": len(self.non_cda_al0s),
            "parties_fetched": len(self.parties_data),
            "asi_downloaded": len(self.asi_files),
            "errors": self.errors,
            "booking_results": {k: v.__dict__ if hasattr(v, "__dict__") else v
                                for k, v in self.booking_results.items()},
            "parties_data": self.parties_data,
            "asi_files": self.asi_files,
        }
