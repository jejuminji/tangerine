"""카카오톡 주문을 AI로 분석하고 누적 검수하는 Streamlit 앱."""

from __future__ import annotations

import json
import logging
import os
import re
import sys
from copy import copy
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any
from uuid import uuid4

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import openai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Side
from pydantic import BaseModel, Field


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", override=True)

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "tangerine.log"

logger = logging.getLogger("tangerine")
logger.setLevel(logging.INFO)
if not logger.handlers:
    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(file_handler)

logger.info(
    "앱 시작 | Python=%s | OpenAI SDK=%s | 실행파일=%s",
    sys.version.split()[0],
    getattr(openai, "__version__", "unknown"),
    Path(__file__).name,
)
DATA_DIR = BASE_DIR / os.getenv("ORDER_DATA_DIR", "data")
LEGACY_WORK_FILE = DATA_DIR / "current_work.json"
KOREA_TIMEZONE = timezone(timedelta(hours=9))
TEMPLATE_FILE = BASE_DIR / "예시엑셀.xlsx"
PROMPT_FILE = BASE_DIR / "prompts" / "order_extraction.txt"
DEFAULT_SENDER_NAME = "강명순"
DEFAULT_SENDER_PHONE = "010-2689-6055"
DEFAULT_SPEC = "10kg"

TEMPLATE_COLUMNS = ["수하인명", "수하인전화", "수량", "주소", "송하인명", "송하인전화", "규격"]
VISIBLE_COLUMNS = ["삭제 선택", "검수완료", "상태", *TEMPLATE_COLUMNS, "확인할 내용"]
EDITABLE_COLUMNS = TEMPLATE_COLUMNS


class ExtractedOrder(BaseModel):
    sender_explicitly_stated: bool = Field(
        description="원문에 보내는 사람이 명시되어 있으면 true, 전혀 없으면 false"
    )
    sender: str = Field(description="명시된 보내는 사람. 없거나 불명확하면 빈 문자열")
    sender_phone: str = Field(
        description="명시된 보내는 사람 전화번호. 주문 전체에 공통임이 명확한 경우 각 배송지에 사용"
    )
    recipient: str = Field(description="받는 사람. 불명확하면 빈 문자열")
    phone: str = Field(description="수하인 전화번호 원문 표기. 다른 배송지에서 추론하거나 재사용하지 않음")
    base_address: str = Field(description="원문에 적힌 기본 주소. 주소에 붙은 숫자 표기도 그대로 포함")
    detail_address: str = Field(description="동, 호수 및 216~1403 같은 주소 뒤 숫자 표기를 원문 그대로 포함")
    quantity_explicitly_stated: bool = Field(
        description="원문에 해당 배송지의 박스 또는 상자 수량이 명시되었으면 true"
    )
    quantity: int | None = Field(
        description="원문에 해당 배송지별로 명시된 박스 수. 명시되지 않으면 null"
    )
    product_name: str = Field(description="명시된 상품명. 불명확하면 빈 문자열")
    evidence: str = Field(description="이 배송지 추출에 사용한 짧은 원문 근거")
    review_reasons: list[str] = Field(description="사람이 확인하거나 입력해야 할 사항")
    uncertain_candidates: list[str] = Field(description="확정 필드에 넣지 않은 원문 후보")


class ExtractedBatch(BaseModel):
    orders: list[ExtractedOrder] = Field(description="배송지별 주문 목록")


def _now() -> str:
    return datetime.now(KOREA_TIMEZONE).isoformat(timespec="seconds")


def _today() -> str:
    return datetime.now(KOREA_TIMEZONE).date().isoformat()


def _work_file(work_date: str) -> Path:
    # Validate the date before using it as a directory name.
    folder = datetime.strptime(work_date, "%Y-%m-%d").date().isoformat()
    return DATA_DIR / folder / "current_work.json"


def _migrate_legacy_work() -> None:
    if not LEGACY_WORK_FILE.exists():
        return
    contents = LEGACY_WORK_FILE.read_bytes()
    data = json.loads(contents)
    if not isinstance(data, dict) or not isinstance(data.get("messages"), list) or not isinstance(data.get("orders"), list):
        raise ValueError("기존 작업의 messages 또는 orders 형식이 올바르지 않습니다.")
    try:
        updated = datetime.fromisoformat(data["updated_at"])
        if updated.tzinfo is not None:
            updated = updated.astimezone(KOREA_TIMEZONE)
    except (KeyError, TypeError, ValueError):
        updated = datetime.fromtimestamp(LEGACY_WORK_FILE.stat().st_mtime, KOREA_TIMEZONE)
    destination = _work_file(updated.date().isoformat())
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with destination.open("xb") as archive:
            archive.write(contents)
    except FileExistsError:
        pass  # Keep both the existing daily work and the original file intact.


def _load_order_prompt() -> str:
    try:
        prompt = PROMPT_FILE.read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise RuntimeError(f"프롬프트 파일을 읽지 못했습니다: {exc}") from exc
    if not prompt:
        raise RuntimeError("주문 분석 프롬프트 파일이 비어 있습니다.")
    return prompt


def _empty_work(work_date: str | None = None) -> dict[str, Any]:
    now = _now()
    return {
        "version": 1,
        "work_date": work_date or _today(),
        "work_id": uuid4().hex,
        "created_at": now,
        "updated_at": now,
        "messages": [],
        "orders": [],
    }


def _load_work(work_date: str) -> tuple[dict[str, Any], str | None]:
    work_file = _work_file(work_date)
    try:
        _migrate_legacy_work()
        if not work_file.exists():
            return _empty_work(work_date), None
        data = json.loads(work_file.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("messages"), list) or not isinstance(data.get("orders"), list):
            raise ValueError("messages 또는 orders 형식이 올바르지 않습니다.")
        data["work_date"] = work_date
        return data, None
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        return _empty_work(work_date), f"저장된 작업을 읽지 못했습니다. 기존 파일은 덮어쓰지 않습니다: {exc}"


def _save_work(work: dict[str, Any]) -> None:
    work_file = _work_file(work["work_date"])
    work_file.parent.mkdir(parents=True, exist_ok=True)
    work["updated_at"] = _now()
    temporary = work_file.with_suffix(".tmp")
    try:
        temporary.write_text(
            json.dumps(work, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary.replace(work_file)
    except OSError:
        if temporary.exists():
            temporary.unlink(missing_ok=True)
        raise


def _configure_page() -> None:
    st.set_page_config(
        page_title="카카오톡 주문 정리",
        page_icon="🍊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(
        """
        <style>
            .stApp { background: #fffaf4; }
            [data-testid="stHeader"] {
                height: 4rem !important;
                background: #fff !important;
            }
            [data-testid="stMainBlockContainer"], .block-container {
                max-width: 100%;
                padding-top: 7rem !important;
            }
            [data-testid="stSidebarContent"] {
                border-top: 4rem solid #fff;
            }
            h1 { color: #3f3128; letter-spacing: -0.04em; }
            [data-testid="stMetric"] {
                background: white; border: 1px solid #f0dfcf;
                border-radius: 14px; padding: 12px 15px;
            }
            [data-testid="stTextArea"] textarea {
                background: white; border-color: #e4cbb4; line-height: 1.65;
            }
            .privacy-note {
                padding: 10px 13px; border-radius: 10px;
                background: #fff1df; color: #66492f; font-size: 0.9rem;
            }
            .step-label { color: #a45d20; font-weight: 700; font-size: 0.88rem; }
            .st-key-download_approved button,
            .st-key-download_all button,
            .st-key-download_approved_disabled button {
                min-height: 3.7rem;
                font-size: 1.08rem;
                font-weight: 750;
                border-radius: 12px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _initialize_state(work_date: str) -> None:
    if "work" not in st.session_state or st.session_state.work.get("work_date") != work_date:
        st.session_state.pop("orders_editor", None)
        st.session_state.notice = None
        work, error = _load_work(work_date)
        work_updated = False
        for row in work["orders"]:
            old_address = " ".join(
                part.strip()
                for part in (str(row.get("기본 주소", "")), str(row.get("상세 주소", "")))
                if part.strip()
            )
            migrated_values = {
                "수하인명": row.get("수하인명", row.get("받는 사람", "")),
                "수하인전화": row.get("수하인전화", row.get("전화번호", "")),
                "주소": row.get("주소", old_address),
                "송하인명": row.get("송하인명", row.get("보내는 사람", "")),
                "송하인전화": row.get("송하인전화", ""),
            }
            for key, value in migrated_values.items():
                if key not in row:
                    row[key] = value
                    work_updated = True
            quantity = row.get("수량")
            if not isinstance(quantity, int) or quantity < 1:
                row["수량"] = 1
                work_updated = True
            if work_updated:
                _refresh_row_status(row)
        if work_updated and error is None:
            _save_work(work)
        st.session_state.work = work
        st.session_state.load_error = error
    defaults_updated = False
    for row in st.session_state.work["orders"]:
        if "규격" not in row:
            row["규격"] = DEFAULT_SPEC
            defaults_updated = True
        sender_name = str(row.get("송하인명", "")).strip()
        sender_phone = str(row.get("송하인전화", "")).strip()
        if not sender_name:
            row["송하인명"] = DEFAULT_SENDER_NAME
            row["송하인전화"] = DEFAULT_SENDER_PHONE
            _refresh_row_status(row)
            defaults_updated = True
        elif sender_name == DEFAULT_SENDER_NAME and not sender_phone:
            row["송하인전화"] = DEFAULT_SENDER_PHONE
            _refresh_row_status(row)
            defaults_updated = True
    if defaults_updated and st.session_state.load_error is None:
        _save_work(st.session_state.work)
    address_candidates_updated = False
    for row in st.session_state.work["orders"]:
        address, candidates, reasons = _merge_address_candidates(
            str(row.get("주소", "")),
            list(row.get("uncertain_candidates", [])),
            list(row.get("ai_review_reasons", [])),
        )
        if (
            address != str(row.get("주소", ""))
            or candidates != list(row.get("uncertain_candidates", []))
            or reasons != list(row.get("ai_review_reasons", []))
        ):
            row["주소"] = address
            row["uncertain_candidates"] = candidates
            row["ai_review_reasons"] = reasons
            _refresh_row_status(row)
            address_candidates_updated = True
    if address_candidates_updated and st.session_state.load_error is None:
        _save_work(st.session_state.work)
    messages_by_id = {
        str(message.get("message_id", "")): str(message.get("raw_text", ""))
        for message in st.session_state.work["messages"]
    }
    rows_by_message: dict[str, list[dict[str, Any]]] = {}
    for row in st.session_state.work["orders"]:
        rows_by_message.setdefault(str(row.get("message_id", "")), []).append(row)
    global_quantity_updated = False
    for message_id, rows in rows_by_message.items():
        if _apply_global_total_quantity(rows, messages_by_id.get(message_id, "")):
            global_quantity_updated = True
    if global_quantity_updated and st.session_state.load_error is None:
        _save_work(st.session_state.work)
    if "kakao_message" not in st.session_state:
        st.session_state.kakao_message = ""
    if "notice" not in st.session_state:
        st.session_state.notice = None


def _normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return value.strip()


def _looks_like_address_fragment(value: str) -> bool:
    """주소 다음에 적힌 숫자·구분자 조각은 해석하지 않고 주소로 보존한다."""
    compact = re.sub(r"\s+", "", value)
    return bool(
        compact
        and re.fullmatch(r"[0-9A-Za-z가-힣~\-–—/]+", compact)
        and re.search(r"\d", compact)
    )


def _merge_address_candidates(
    address: str,
    candidates: list[str],
    review_reasons: list[str],
) -> tuple[str, list[str], list[str]]:
    address_parts = [address.strip()] if address.strip() else []
    merged_candidates: list[str] = []
    remaining_candidates: list[str] = []
    for candidate in candidates:
        cleaned = str(candidate).strip()
        if cleaned and _looks_like_address_fragment(cleaned):
            if cleaned not in address:
                address_parts.append(cleaned)
            merged_candidates.append(cleaned)
        elif cleaned:
            remaining_candidates.append(cleaned)

    remaining_reasons = [
        reason
        for reason in review_reasons
        if not any(candidate in str(reason) for candidate in merged_candidates)
    ]
    return " ".join(address_parts), remaining_candidates, remaining_reasons


def _apply_global_total_quantity(rows: list[dict[str, Any]], raw_message: str) -> bool:
    """`총 N박스`가 배송지 수와 같으면 특정 행이 아닌 전체 합계로 처리한다."""
    global_pattern = re.compile(r"(?:총|합계)\s*(\d+)\s*(?:박스|상자)")
    totals = [int(match.group(1)) for match in global_pattern.finditer(raw_message)]
    if not totals or len(set(totals)) != 1:
        return False

    total = totals[0]
    if len(rows) <= 1 or total != len(rows):
        return False

    without_global_total = global_pattern.sub("", raw_message)
    per_destination_quantity = re.search(
        r"(?:\d+|한|두|세|네|다섯|여섯|일곱|여덟|아홉)\s*(?:박스|상자)",
        without_global_total,
    )
    if per_destination_quantity:
        return False

    changed = False
    for row in rows:
        if row.get("수량") != 1:
            row["수량"] = 1
            _refresh_row_status(row, edited=bool(row.get("approved")))
            changed = True
    return changed


def _review_reasons(row: dict[str, Any]) -> list[str]:
    issues_resolved = bool(row.get("review_issues_resolved"))
    reasons = [] if issues_resolved else list(row.get("ai_review_reasons", []))
    candidates = row.get("uncertain_candidates", [])
    if not str(row.get("수하인명", "")).strip():
        reasons.append("수하인명이 없어요")
    if not str(row.get("수하인전화", "")).strip():
        reasons.append("수하인 전화번호가 없어요")
    if not str(row.get("주소", "")).strip():
        reasons.append("주소가 없어요")
    if not str(row.get("송하인명", "")).strip():
        reasons.append("송하인명이 없어요")
    if not str(row.get("송하인전화", "")).strip():
        reasons.append("송하인 전화번호가 없어요")
    quantity = row.get("수량")
    if not isinstance(quantity, int) or quantity <= 0:
        reasons.append("수량을 확인해 주세요")
    if candidates and not issues_resolved:
        reasons.append("애매한 원문 후보를 확인해 주세요: " + ", ".join(candidates))
    return list(dict.fromkeys(reason for reason in reasons if reason))


def _refresh_row_status(row: dict[str, Any], edited: bool = False) -> None:
    if edited:
        row["approved"] = False
        row["reviewed_at"] = None
        row["reviewed_against_original"] = False
        row["review_issues_resolved"] = False
    reasons = _review_reasons(row)
    row["확인할 내용"] = " / ".join(reasons)
    if row.get("approved"):
        row["상태"] = "✅ 확인 완료"
    elif reasons:
        row["상태"] = "🔴 확인 필요"
    else:
        row["상태"] = "🟢 검수 대기"


def _make_order_row(order: ExtractedOrder, group_id: str, message_id: str) -> dict[str, Any]:
    ai_values = order.model_dump()
    # 업무 규칙: 배송지별 2박스 이상이 명시되지 않은 주문은 항상 1박스다.
    quantity = (
        order.quantity
        if order.quantity_explicitly_stated
        and isinstance(order.quantity, int)
        and order.quantity >= 2
        else 1
    )
    if order.sender_explicitly_stated:
        sender_name = order.sender.strip()
        sender_phone = _normalize_phone(order.sender_phone)
    else:
        sender_name = DEFAULT_SENDER_NAME
        sender_phone = DEFAULT_SENDER_PHONE
    extracted_address = " ".join(
        part for part in (order.base_address.strip(), order.detail_address.strip()) if part
    )
    extracted_address, uncertain_candidates, review_reasons = _merge_address_candidates(
        extracted_address,
        order.uncertain_candidates,
        order.review_reasons,
    )
    row: dict[str, Any] = {
        "order_id": f"ORD-{uuid4().hex[:8].upper()}",
        "group_id": group_id,
        "message_id": message_id,
        "상태": "",
        "수하인명": order.recipient.strip(),
        "수하인전화": _normalize_phone(order.phone),
        "수량": quantity,
        "주소": extracted_address,
        "송하인명": sender_name,
        "송하인전화": sender_phone,
        "규격": DEFAULT_SPEC,
        "수하인전화 원문": order.phone.strip(),
        "송하인전화 원문": order.sender_phone.strip(),
        "상품명": order.product_name.strip(),
        "확인할 내용": "",
        "evidence": order.evidence,
        "ai_review_reasons": review_reasons,
        "uncertain_candidates": uncertain_candidates,
        "ai_values": ai_values,
        "approved": False,
        "reviewed_at": None,
        "reviewed_against_original": False,
        "review_issues_resolved": False,
        "created_at": _now(),
    }
    _refresh_row_status(row)
    return row


def _usage_dict(response: Any) -> dict[str, Any]:
    usage = getattr(response, "usage", None)
    if usage is None:
        return {"input_tokens": None, "output_tokens": None, "total_tokens": None}
    return {
        "input_tokens": getattr(usage, "input_tokens", None),
        "output_tokens": getattr(usage, "output_tokens", None),
        "total_tokens": getattr(usage, "total_tokens", None),
    }


def _analyze_message(raw_message: str) -> tuple[int, dict[str, Any]]:
    work = st.session_state.work
    existing = next(
        (
            item
            for item in work["messages"]
            if item.get("raw_text") == raw_message and item.get("status") == "success"
        ),
        None,
    )
    if existing:
        logger.info("중복 원문 감지 | 기존 분석 결과 사용")
        return 0, {"kind": "warning", "text": "이미 분석한 같은 원문이라 기존 결과를 사용했어요."}

    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    model = os.getenv("OPENAI_MODEL", "").strip()

    if not api_key:
        logger.error(".env 설정 오류 | OPENAI_API_KEY가 비어 있음")
        return 0, {"kind": "error", "text": ".env에 OPENAI_API_KEY를 입력해 주세요."}
    if not model:
        logger.error(".env 설정 오류 | OPENAI_MODEL이 비어 있음")
        return 0, {"kind": "error", "text": ".env에 OPENAI_MODEL을 입력해 주세요."}

    message_id = uuid4().hex
    request_tag = message_id[:12]
    message_record: dict[str, Any] = {
        "message_id": message_id,
        "raw_text": raw_message,
        "created_at": _now(),
        "status": "requesting",
        "model": model,
        "api_attempts": 1,
        "usage": {"input_tokens": None, "output_tokens": None, "total_tokens": None},
        "order_ids": [],
        "ai_result": None,
        "error": None,
    }
    work["messages"].append(message_record)
    _save_work(work)  # 개인정보 원문을 API 호출 전에 먼저 보존한다.

    logger.info(
        "OpenAI 호출 시작 | request_tag=%s | model=%s | message_chars=%d | SDK=%s",
        request_tag,
        model,
        len(raw_message),
        getattr(openai, "__version__", "unknown"),
    )

    try:
        client = OpenAI(api_key=api_key, max_retries=0, timeout=90.0)

        # responses.parse는 OpenAI Python SDK 2.x 사용을 전제로 한다.
        if not hasattr(client, "responses") or not hasattr(client.responses, "parse"):
            raise RuntimeError(
                "현재 OpenAI SDK에 responses.parse가 없습니다. "
                "openai>=2.0.0,<3.0으로 업그레이드해 주세요."
            )

        prompt = _load_order_prompt()
        logger.info("프롬프트 로드 완료 | request_tag=%s | prompt_chars=%d",
                    request_tag, len(prompt))

        # reasoning 옵션은 추론 모델 전용일 수 있으므로 제거하여
        # gpt-4o 등 일반 모델에서도 요청 파라미터 오류가 나지 않게 한다.
        response = client.responses.parse(
            model=model,
            instructions=prompt,
            input=json.dumps({"kakao_message": raw_message}, ensure_ascii=False),
            text_format=ExtractedBatch,
            store=False,
        )

        response_id = getattr(response, "id", None)
        response_status = getattr(response, "status", None)
        parsed = getattr(response, "output_parsed", None)

        logger.info(
            "OpenAI 응답 수신 | request_tag=%s | response_id=%s | status=%s | parsed=%s",
            request_tag,
            response_id,
            response_status,
            parsed is not None,
        )

        if parsed is None:
            incomplete = getattr(response, "incomplete_details", None)
            refusal_text = None
            try:
                for output_item in getattr(response, "output", []) or []:
                    for content_item in getattr(output_item, "content", []) or []:
                        if getattr(content_item, "type", None) == "refusal":
                            refusal_text = getattr(content_item, "refusal", None)
                            break
            except Exception:
                logger.exception("응답 세부정보 확인 중 오류 | request_tag=%s", request_tag)

            logger.error(
                "구조화 결과 없음 | request_tag=%s | response_id=%s | status=%s | incomplete=%s | refusal=%s",
                request_tag,
                response_id,
                response_status,
                str(incomplete)[:500],
                str(refusal_text)[:300] if refusal_text else None,
            )
            raise ValueError(
                "AI 응답에 주문 분석 결과가 없습니다. "
                f"응답 상태={response_status}, incomplete={incomplete}, refusal={refusal_text}"
            )

        group_id = f"GRP-{uuid4().hex[:8].upper()}"
        new_rows = [_make_order_row(order, group_id, message_id) for order in parsed.orders]
        _apply_global_total_quantity(new_rows, raw_message)
        work["orders"].extend(new_rows)
        message_record["status"] = "success"
        message_record["usage"] = _usage_dict(response)
        message_record["ai_result"] = parsed.model_dump()
        message_record["order_ids"] = [row["order_id"] for row in new_rows]
        message_record["response_id"] = response_id
        _save_work(work)

        usage = message_record["usage"]
        logger.info(
            "OpenAI 분석 성공 | request_tag=%s | response_id=%s | orders=%d | usage=%s",
            request_tag,
            response_id,
            len(new_rows),
            usage,
        )

        if not new_rows:
            return 0, {"kind": "warning", "text": "배송 주문을 찾지 못했어요. 원문을 확인해 주세요."}
        return len(new_rows), {
            "kind": "success",
            "text": f"배송지 {len(new_rows)}건을 읽어 기존 주문표에 추가했어요.",
        }

    except Exception as exc:  # SDK 오류 종류와 무관하게 기존 작업은 보존한다.
        error_type = type(exc).__name__
        error_message = str(exc)
        request_id = getattr(exc, "request_id", None)

        logger.exception(
            "OpenAI 분석 실패 | request_tag=%s | model=%s | error_type=%s | request_id=%s",
            request_tag,
            model,
            error_type,
            request_id,
        )

        message_record["status"] = "failed"
        message_record["error"] = (
            f"{error_type}: {error_message[:700]}"
            + (f" | request_id={request_id}" if request_id else "")
        )
        message_record["request_id"] = request_id
        try:
            _save_work(work)
        except Exception:
            logger.exception("실패 상태 저장 중 추가 오류 | request_tag=%s", request_tag)

        return 0, {
            "kind": "error",
            "text": (
                f"AI 분석에 실패했어요. 기존 주문은 유지했습니다. "
                f"오류: {error_type}: {error_message[:300]} "
                f"(추적번호: {request_tag}) "
                f"상세 로그: logs/tangerine.log"
            ),
        }

def _display_notice() -> None:
    notice = st.session_state.notice
    if not notice:
        return
    getattr(st, notice["kind"])(notice["text"])
    st.session_state.notice = None


def _counts() -> tuple[int, int, int, int]:
    orders = st.session_state.work["orders"]
    attention = sum(row.get("상태") == "🔴 확인 필요" for row in orders)
    waiting = sum(row.get("상태") == "🟢 검수 대기" for row in orders)
    approved = sum(row.get("상태") == "✅ 확인 완료" for row in orders)
    return len(orders), attention, waiting, approved


def _editor_frame(include_total: bool = False) -> pd.DataFrame:
    rows = st.session_state.work["orders"]
    records = [
        {
            column: (
                False
                if column == "삭제 선택"
                else bool(row.get("approved"))
                if column == "검수완료"
                else row.get(column)
            )
            for column in VISIBLE_COLUMNS
        }
        for row in rows
    ]
    if include_total:
        records.append(_total_preview_row(rows))
    return pd.DataFrame(records, columns=VISIBLE_COLUMNS)


def _clean_editor_value(column: str, value: Any) -> Any:
    if pd.isna(value):
        return None if column == "수량" else ""
    if column == "수량":
        number = float(value)
        return int(number) if number.is_integer() else value
    return str(value)


def _apply_editor_changes(edited: pd.DataFrame) -> None:
    changed = False
    rerun_needed = False
    for index, row in edited.iterrows():
        stored = st.session_state.work["orders"][index]
        requested_approval = bool(row["검수완료"])
        row_changed = False
        for column in EDITABLE_COLUMNS:
            value = _clean_editor_value(column, row[column])
            if stored.get(column) != value:
                stored[column] = value
                row_changed = True
        if row_changed:
            _refresh_row_status(stored, edited=True)
            changed = True
            rerun_needed = True
            st.session_state.notice = {
                "kind": "warning",
                "text": "주문 내용을 수정해 검수완료가 자동으로 취소됐어요. 다시 확인해 주세요.",
            }
            continue

        currently_approved = bool(stored.get("approved"))
        if requested_approval and not currently_approved:
            blockers = _approval_blockers(stored)
            if blockers:
                rerun_needed = True
                st.session_state.notice = {
                    "kind": "error",
                    "text": "검수완료할 수 없어요: " + " · ".join(blockers),
                }
                continue
            stored["reviewed_against_original"] = True
            stored["review_issues_resolved"] = True
            stored["approved"] = True
            stored["reviewed_at"] = _now()
            _refresh_row_status(stored)
            changed = True
            rerun_needed = True
        elif not requested_approval and currently_approved:
            stored["approved"] = False
            stored["reviewed_at"] = None
            stored["reviewed_against_original"] = False
            _refresh_row_status(stored)
            changed = True
            rerun_needed = True
    if changed:
        try:
            _save_work(st.session_state.work)
        except OSError as exc:
            st.error(f"수정 내용을 저장하지 못했어요: {exc}")
            return
    if rerun_needed:
        st.session_state.pop("orders_editor", None)
        st.rerun()


def _delete_selected_orders(edited: pd.DataFrame) -> None:
    selected_indexes = {
        int(index)
        for index, row in edited.iterrows()
        if bool(row.get("삭제 선택", False))
    }
    if not selected_indexes:
        return

    orders = st.session_state.work["orders"]
    deleted_count = len(selected_indexes)
    st.session_state.work["orders"] = [
        order for index, order in enumerate(orders) if index not in selected_indexes
    ]
    try:
        _save_work(st.session_state.work)
    except OSError as exc:
        st.session_state.work["orders"] = orders
        st.error(f"선택한 주문을 삭제하지 못했어요: {exc}")
        return

    st.session_state.notice = {
        "kind": "success",
        "text": f"선택한 주문 {deleted_count}건을 삭제했어요.",
    }
    st.session_state.pop("orders_editor", None)
    st.rerun()


def _approval_blockers(row: dict[str, Any]) -> list[str]:
    blockers: list[str] = []
    if not str(row.get("수하인명", "")).strip():
        blockers.append("수하인명을 입력해 주세요")
    if not str(row.get("수하인전화", "")).strip():
        blockers.append("수하인 전화를 입력해 주세요")
    if not str(row.get("주소", "")).strip():
        blockers.append("주소를 입력해 주세요")
    if not str(row.get("송하인명", "")).strip():
        blockers.append("송하인명을 입력해 주세요")
    if not str(row.get("송하인전화", "")).strip():
        blockers.append("송하인 전화를 입력해 주세요")
    quantity = row.get("수량")
    if not isinstance(quantity, int) or quantity <= 0:
        blockers.append("수량을 1 이상으로 입력해 주세요")
    return blockers


def _safe_excel_text(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def _approved_orders() -> list[dict[str, Any]]:
    unique: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in st.session_state.work["orders"]:
        order_id = str(row.get("order_id", ""))
        if row.get("approved") and order_id not in seen:
            unique.append(row)
            seen.add(order_id)
    return unique


def _box_total(rows: list[dict[str, Any]]) -> int:
    return sum(
        quantity
        for row in rows
        if isinstance(quantity := row.get("수량", 1), int) and quantity > 0
    )


def _total_preview_row(rows: list[dict[str, Any]]) -> dict[str, Any]:
    total = {column: "" for column in VISIBLE_COLUMNS}
    total.update({
        "삭제 선택": False,
        "검수완료": False,
        "수하인전화": "총수량",
        "수량": _box_total(rows),
        "확인할 내용": "자동 계산",
    })
    return total


def _build_template_excel(rows: list[dict[str, Any]]) -> bytes:
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"엑셀 양식 파일을 찾을 수 없습니다: {TEMPLATE_FILE.name}")

    workbook = load_workbook(TEMPLATE_FILE)
    if "Sheet2" not in workbook.sheetnames:
        raise ValueError("엑셀 양식에 Sheet2 시트가 없습니다.")
    sheet = workbook["Sheet2"]
    # Column defaults also color empty cells below the last populated row.
    sheet.column_dimensions["D"].fill = PatternFill(fill_type=None)

    # 양식의 첫 데이터 행 서식을 보존한 뒤, 실제 개인정보가 든 기존 행은 모두 제거한다.
    style_source = [sheet.cell(row=2, column=column) for column in range(2, 9)]
    styles = [copy(cell._style) for cell in style_source]
    alignments = [copy(cell.alignment) for cell in style_source]
    protections = [copy(cell.protection) for cell in style_source]
    if sheet.max_row >= 2:
        sheet.delete_rows(2, sheet.max_row - 1)

    for row_number, order in enumerate(rows, start=2):
        values: list[Any] = [
            _safe_excel_text(order.get("수하인명")),
            _safe_excel_text(order.get("수하인전화")),
            int(order.get("수량", 1)),
            _safe_excel_text(order.get("주소")),
            _safe_excel_text(order.get("송하인명")),
            _safe_excel_text(order.get("송하인전화")),
            _safe_excel_text(order.get("규격", DEFAULT_SPEC)),
        ]
        for offset, value in enumerate(values):
            cell = sheet.cell(row=row_number, column=offset + 2, value=value)
            cell._style = copy(styles[offset])
            cell.alignment = copy(alignments[offset])
            cell.protection = copy(protections[offset])
            if offset in (1, 5):
                cell.number_format = "@"

    total_row = len(rows) + 2
    for offset in range(len(TEMPLATE_COLUMNS)):
        cell = sheet.cell(row=total_row, column=offset + 2)
        cell._style = copy(styles[offset])
        font = copy(cell.font)
        font.bold = True
        font.color = "000000"
        cell.font = font
        cell.fill = PatternFill(fill_type=None)
        border = copy(cell.border)
        border.top = Side(style="thin", color="999999")
        cell.border = border
    sheet.cell(total_row, 3, "총수량")
    total_cell = sheet.cell(total_row, 4, _box_total(rows))
    total_cell.number_format = "0"
    total_cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[total_row].height = 24
    sheet.row_dimensions[total_row].hidden = False
    if sheet.print_area:
        sheet.print_area = f"A1:H{total_row}"

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _download_filename(rows: list[dict[str, Any]]) -> str:
    export_date = datetime.fromisoformat(st.session_state.work["work_date"]) + timedelta(days=1)
    return f"{export_date:%Y_%m_%d}_{_box_total(rows)}박스.xlsx"


def _render_preview() -> None:
    orders = st.session_state.work["orders"]
    st.subheader(f"2. 엑셀 미리보기 · 총 {_box_total(orders):,}박스")
    st.caption(
        "새 메시지를 읽을 때마다 아래 표에 계속 추가됩니다. 수하인·전화번호·수량·주소·송하인·규격은 "
        "표에서 직접 입력하거나 지울 수 있고, 수정값이 자동 저장되어 그대로 엑셀에 들어갑니다. "
        "왼쪽 입력창을 접으면 표를 더 넓게 볼 수 있어요."
    )

    if not orders:
        st.info("아직 주문이 없습니다. 왼쪽에 카카오톡 메시지를 붙여넣어 주세요.")
        st.dataframe(
            _editor_frame(include_total=True),
            width="stretch",
            hide_index=True,
            height=600,
            column_config={
                "상태": st.column_config.TextColumn("상태", width=100),
                "수량": st.column_config.NumberColumn("수량", min_value=1, step=1, width=38),
                "수하인전화": st.column_config.TextColumn("수하인전화", width=140),
            },
        )
        return

    column_config = {
        "삭제 선택": st.column_config.CheckboxColumn(
            "삭제",
            help="지울 주문을 선택한 뒤 아래 삭제 버튼을 누르세요.",
            width="small",
        ),
        "검수완료": st.column_config.CheckboxColumn(
            "검수완료",
            help="필수값을 확인한 뒤 체크하세요.",
            width="small",
        ),
        "상태": st.column_config.TextColumn("상태", width=100),
        "수하인명": st.column_config.TextColumn("수하인명", width="small"),
        "수하인전화": st.column_config.TextColumn("수하인전화", width=140),
        "수량": st.column_config.NumberColumn("수량", min_value=1, step=1, width=38),
        "주소": st.column_config.TextColumn("주소", width="large"),
        "송하인명": st.column_config.TextColumn("송하인명", width="small"),
        "송하인전화": st.column_config.TextColumn("송하인전화", width="medium"),
        "규격": st.column_config.TextColumn("규격", width="small", default=DEFAULT_SPEC),
        "확인할 내용": st.column_config.TextColumn("확인할 내용", width="large"),
    }
    edited = st.data_editor(
        _editor_frame(include_total=True),
        width="stretch",
        hide_index=True,
        height=600,
        disabled=["상태", "확인할 내용"],
        column_config=column_config,
        key="orders_editor",
    )
    # Only real orders may be saved, approved, or deleted. Restore edits to the total row.
    edited_orders = edited.iloc[:len(orders)]
    _apply_editor_changes(edited_orders)
    if edited.iloc[-1].to_dict() != _total_preview_row(orders):
        st.session_state.pop("orders_editor", None)
        st.rerun()
    delete_count = int(edited_orders["삭제 선택"].fillna(False).astype(bool).sum())
    if st.button(
        f"선택한 행 삭제 ({delete_count}건)",
        disabled=delete_count == 0,
        width="stretch",
        key="delete_selected_orders",
    ):
        _delete_selected_orders(edited_orders)



def run_app() -> None:
    _configure_page()
    logger.info("화면 렌더링 시작")

    today = _today()
    previous_today = st.session_state.get("calendar_date", today)
    if previous_today != today and st.session_state.get("work_date_selector") == datetime.fromisoformat(previous_today).date():
        st.session_state.work_date_selector = datetime.fromisoformat(today).date()
    st.session_state.calendar_date = today
    if "work_date_selector" not in st.session_state:
        st.session_state.work_date_selector = datetime.fromisoformat(today).date()
    selected_date = st.date_input(
        "작업 날짜",
        value=None,
        max_value=datetime.fromisoformat(today).date(),
        key="work_date_selector",
        help="날짜별로 자동 저장됩니다. 이전 날짜를 선택하면 그날의 주문을 다시 열 수 있습니다.",
    )
    if selected_date is None:
        st.info("작업 날짜를 선택해 주세요.")
        return
    _initialize_state(selected_date.isoformat())

    st.markdown('<div class="step-label">카카오톡 주문 정리</div>', unsafe_allow_html=True)
    st.title("주문표를 확인해 주세요")
    st.caption("왼쪽 입력창에 메시지를 붙여넣고 주문 읽어오기를 누르면 아래 주문표에 계속 누적됩니다.")

    if st.session_state.load_error:
        st.error(st.session_state.load_error)
    _display_notice()

    total, attention, waiting, approved = _counts()
    metric_columns = st.columns(4)
    for column, label, value in zip(
        metric_columns,
        ("전체 주문", "🔴 확인 필요", "🟢 검수 대기", "✅ 확인 완료"),
        (total, attention, waiting, approved),
        strict=True,
    ):
        column.metric(label, value)

    st.write("")
    with st.sidebar:
        st.subheader("1. 카카오톡 메시지")
        st.text_area(
            "주문 메시지 붙여넣기",
            key="kakao_message",
            height=360,
            placeholder=(
                "예시)\n홍길동님 주문입니다.\n받는 분: 김하늘\n"
                "전화번호: 010 1234 5678\n주소: 서울시 해오름로 12\n2박스"
            ),
            label_visibility="collapsed",
        )
        read_column, clear_column = st.columns([1.2, 1])
        with read_column:
            analyze_clicked = st.button("주문 읽어오기", type="primary", width="stretch")
        with clear_column:
            clear_clicked = st.button("입력 지우기", width="stretch")

        if clear_clicked:
            st.session_state.kakao_message = ""
            st.rerun()
        if analyze_clicked:
            raw_message = st.session_state.kakao_message.strip()
            if not raw_message:
                st.warning("카카오톡 주문 내용을 먼저 붙여넣어 주세요.")
            else:
                try:
                    with st.spinner("AI가 배송 정보를 읽고 있어요…"):
                        _, notice = _analyze_message(raw_message)
                    st.session_state.notice = notice
                    st.rerun()
                except OSError as exc:
                    st.error(f"원문을 저장하지 못해 분석을 시작하지 않았어요: {exc}")

    with st.container():
        _render_preview()
        approved_rows = _approved_orders()
        st.info(f"확인 완료한 {len(approved_rows)}건만 양식에 맞춰 엑셀로 저장됩니다.")
        if approved_rows:
            try:
                excel_data = _build_template_excel(approved_rows)
                st.download_button(
                    "엑셀로 다운로드",
                    data=excel_data,
                    file_name=_download_filename(approved_rows),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    width="stretch",
                    key="download_approved",
                )
            except (OSError, ValueError) as exc:
                st.error(f"엑셀 파일을 만들지 못했어요. 작업 내용은 그대로 유지됩니다: {exc}")
        else:
            st.button(
                "엑셀로 다운로드",
                disabled=True,
                width="stretch",
                key="download_approved_disabled",
            )

        all_rows = st.session_state.work["orders"]
        if all_rows:
            try:
                all_excel_data = _build_template_excel(all_rows)
                st.download_button(
                    "확인 필요 포함 다운로드",
                    data=all_excel_data,
                    file_name=_download_filename(all_rows),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="download_all",
                    help="검수 전 주문도 포함하며, 입력되지 않은 값은 빈 셀로 저장됩니다.",
                )
            except (OSError, ValueError) as exc:
                st.error(f"전체 주문 엑셀을 만들지 못했어요: {exc}")
        else:
            st.button(
                "확인 필요 포함 다운로드",
                disabled=True,
                width="stretch",
                key="download_all",
            )
        st.caption(
            f"양식: {TEMPLATE_FILE.name} · {st.session_state.work['work_date']} 작업 자동 저장"
        )


if __name__ == "__main__":
    run_app()
