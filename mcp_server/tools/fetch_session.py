"""
fetch_session.py — Fetch Session 工具
管理浏览器 fetch 的请求-响应循环。
"""
from __future__ import annotations

from pathlib import Path

from core.session_store import (
    FetchSession,
    PHASE_DONE,
    BATCH_SIZE,
)

DATA_DIR = Path(__file__).parent.parent / "data"

# 内存中活跃的 sessions
_active_sessions: dict[str, FetchSession] = {}


def start_fetch_session(al0_list: list[str], ib_data: dict = None) -> dict:
    """
    启动 Fetch Session，生成第一批 URL。

    Args:
        al0_list: 待处理的 AL0 列表
        ib_data: 进仓数据（al0 → 字段 dict），可选

    Returns:
        {session_id, phase, urls: [{id, url, method, headers, body}]}
    """
    session = FetchSession()
    session.al0_list = al0_list
    session.ib_data = ib_data or {}

    # Phase 1: 为每个 AL0 生成 getBookingById URL
    session.generate_booking_check_urls()

    # 取第一批
    batch = session.get_next_batch()

    # 存入内存
    _active_sessions[session.session_id] = session

    # 持久化
    session.save(DATA_DIR)

    return {
        "session_id": session.session_id,
        "phase": session.phase,
        "urls": batch,
        "total_al0": len(al0_list),
    }


def submit_fetch_batch(session_id: str, results: dict) -> dict:
    """
    提交浏览器 fetch 结果，获取下一批 URL 或最终结果。

    Args:
        session_id: session 唯一 ID
        results: {url_id: response_json} 本批 fetch 结果

    Returns:
        phase == "DONE" 时: {session_id, phase, summary}
        phase != "DONE" 时: {session_id, phase, urls}
    """
    session = _active_sessions.get(session_id)
    if not session:
        # 尝试从磁盘恢复
        session = FetchSession.load(session_id, DATA_DIR)
        if not session:
            return {"error": f"Session {session_id} 不存在"}
        _active_sessions[session_id] = session

    # 处理本批结果
    session.process_results(results)

    # 检查当前 phase 是否还有剩余 URL
    if session.pending_urls:
        batch = session.get_next_batch()
        session.save(DATA_DIR)
        return {
            "session_id": session.session_id,
            "phase": session.phase,
            "urls": batch,
        }

    # 当前 phase 完成，推进到下一阶段
    session.advance_phase()

    if session.phase == PHASE_DONE:
        session.save(DATA_DIR)
        return {
            "session_id": session.session_id,
            "phase": PHASE_DONE,
            "summary": session.get_summary(),
        }

    # 新 phase，取第一批 URL
    batch = session.get_next_batch()
    session.save(DATA_DIR)

    return {
        "session_id": session.session_id,
        "phase": session.phase,
        "urls": batch,
    }


def get_fetch_session_status(session_id: str) -> dict:
    """查询 session 当前状态。"""
    session = _active_sessions.get(session_id)
    if not session:
        session = FetchSession.load(session_id, DATA_DIR)
        if not session:
            return {"error": f"Session {session_id} 不存在"}

    return {
        "session_id": session.session_id,
        "phase": session.phase,
        "total_al0": len(session.al0_list),
        "cda_count": len(session.cda_al0s),
        "non_cda_count": len(session.non_cda_al0s),
        "booking_checked": len(session.booking_results),
        "parties_fetched": len(session.parties_data),
        "asi_downloaded": len(session.asi_files),
        "errors": session.errors,
        "pending_urls": len(session.pending_urls),
    }
