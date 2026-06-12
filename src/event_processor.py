"""
event_processor.py
이벤트 처리 비즈니스 로직
- 이번주 [실적]: 같은날 병합, 날짜다르고 내용같으면 그룹화
- 다음주 [계획]: (pc+fc+subject) 기준 시간합산, subject 다르면 별도행
- Excel 행 데이터 생성
- 경고 메시지 생성 (날짜 포함 형식)
"""
from __future__ import annotations
import re
from collections import OrderedDict
from datetime import datetime
from config import get_gubun, get_func_name, is_highlight, is_valid_code


# ── subject 정규화 (비교용): [실적/계획] 태그 + 코드 블록 제거 ──────
_NORM_CODE_RE = re.compile(r"\s*\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]\s*")
_NORM_TAG_RE  = re.compile(r"^\s*\[(실적|계획)\]\s*")

def _norm_subj(s: str) -> str:
    """이번주↔다음주 제목 비교용 정규화 (화면 표시와 무관)"""
    s = _NORM_CODE_RE.sub(" ", s)
    s = _NORM_TAG_RE.sub("", s)
    return s.strip()


# ══════════════════════════════════════════════════════════
# 이번주 처리
# ══════════════════════════════════════════════════════════

def _same_day_merge(events: list[dict]) -> list[dict]:
    """
    같은날 + 같은제목 + 같은내용 → 시간 합산, 단일 행
    같은날 + 같은제목 + 다른내용 → 각각 별도 행
    """
    bucket: dict[tuple, dict] = {}
    order:  list[tuple]       = []

    for e in events:
        key = (e["date"].date(), e["subject"], e["body"])
        if key not in bucket:
            bucket[key] = e.copy()
            order.append(key)
        else:
            bucket[key]["hours"] = round(bucket[key]["hours"] + e["hours"], 1)

    return [bucket[k] for k in order]


def process_this_week(events: list[dict]) -> list[dict]:
    """
    [실적] 필터 → 같은날 병합 → func_code 있는 것만
    (날짜순 정렬은 build_rows에서 그룹화 후 처리)
    """
    filtered = [
        e for e in events
        if e.get("tag") == "실적"
    ]
    return _same_day_merge(filtered)


# ══════════════════════════════════════════════════════════
# 다음주 처리
# ══════════════════════════════════════════════════════════

def process_next_week(events: list[dict]) -> list[dict]:
    """
    [계획] 필터
    (project_code, func_code, subject) 동일 → 시간 합산
    subject가 다르면 별도 행
    """
    filtered = [
        e for e in events
        if e.get("tag") == "계획"
    ]
    bucket: dict[tuple, dict] = {}
    order:  list[tuple]       = []

    for e in filtered:
        key = (e.get("project_code") or "", e.get("func_code") or "", e.get("subject",""))
        if key not in bucket:
            bucket[key] = e.copy()
            order.append(key)
        else:
            bucket[key]["hours"] = round(bucket[key]["hours"] + e["hours"], 1)

    return [bucket[k] for k in order]


# ══════════════════════════════════════════════════════════
# Excel 행 데이터 생성
# ══════════════════════════════════════════════════════════

def build_rows(this_week: list[dict], next_week: list[dict]) -> list[dict]:
    """
    이번주 + 다음주 이벤트 → Excel 행 리스트

    이번주 그룹화 규칙:
      같은 (func_code, subject, body) → 그룹으로 묶어 날짜순 배치
      → H열(제목), I열(상세) 그룹 내 수직 병합 처리 (is_group_first/size 활용)

    다음주 연결 규칙:
      (pc, fc, subject) 동일 → 이번주 첫 번째 행에 차주 시간 추가
      subject 다르면 신규 행
    """
    rows:        list[dict]         = []
    key_to_idx:  dict[tuple, int]   = {}   # (pc, fc, subj) → 이번주 첫행 index

    # ── 이번주 그룹화 ──
    groups: OrderedDict = OrderedDict()
    for ev in this_week:
        g_key = (ev.get("func_code") or "", ev.get("subject", ""), ev.get("body", ""))
        if g_key not in groups:
            groups[g_key] = []
        groups[g_key].append(ev)

    for g_key, group_evs in groups.items():
        # 그룹 내 날짜순 정렬
        sorted_evs = sorted(group_evs, key=lambda x: x["date"])
        g_size = len(sorted_evs)

        for ev_idx, ev in enumerate(sorted_evs):
            pc   = ev.get("project_code") or ""
            fc   = ev.get("func_code")    or ""
            subj = ev.get("subject", "")
            body = ev.get("body",    "")

            row = dict(
                gubun          = get_gubun(fc, subj, body, pc),
                project_code   = pc,
                func_code      = fc,
                func_name      = get_func_name(fc),
                date           = ev["date"],
                subject        = subj,
                body           = body,
                this_week_h    = ev["hours"],
                next_week_h    = 0.0,
                highlight      = False,   # 노란 음영 제거
                source         = "this_week",
                # 병합 판단용
                group_key      = g_key,
                is_group_first = (ev_idx == 0),
                is_group_last  = (ev_idx == g_size - 1),
                group_size     = g_size,
            )
            idx = len(rows)
            rows.append(row)

            # 다음주 연결용: 정규화 subject로 비교 (첫 번째 행만 등록)
            nw_key = (pc, fc, _norm_subj(subj))
            if nw_key not in key_to_idx:
                key_to_idx[nw_key] = idx

    # ── 다음주 처리 ──
    for ev in next_week:
        pc   = ev.get("project_code") or ""
        fc   = ev.get("func_code")    or ""
        subj = ev.get("subject", "")
        body = ev.get("body",    "")
        nw_key = (pc, fc, _norm_subj(subj))   # 정규화 비교

        if nw_key in key_to_idx:
            # 이번주 첫 행에 차주 시간 누적
            r = rows[key_to_idx[nw_key]]
            r["next_week_h"] = round(r["next_week_h"] + ev["hours"], 1)
        else:
            # 이번주에 없는 신규 행
            row = dict(
                gubun          = get_gubun(fc, subj, body),
                project_code   = pc,
                func_code      = fc,
                func_name      = get_func_name(fc),
                date           = ev["date"],
                subject        = subj,
                body           = body,
                this_week_h    = 0.0,
                next_week_h    = ev["hours"],
                highlight      = False,
                source         = "next_week_only",
                group_key      = (fc, subj, body),
                is_group_first = True,
                is_group_last  = True,
                group_size     = 1,
            )
            rows.append(row)

    return rows


# ══════════════════════════════════════════════════════════
# 경고 메시지 생성
# ══════════════════════════════════════════════════════════

def get_warnings(
    this_week_raw: list[dict],
    next_week_raw: list[dict],
    this_expected: float = 40.0,   # 주별=40, 일별=8
    next_expected: float = 40.0,
) -> list[str]:
    """사전 검토 결과 경고 문자열 반환"""
    warns: list[str] = []

    this_total = sum(
        e["hours"] for e in this_week_raw if e.get("tag") == "실적"
    )
    next_total = sum(
        e["hours"] for e in next_week_raw if e.get("tag") == "계획"
    )

    # ── 시간 경고 (기준 미만 / 초과 / 자동보정) ─────────────────
    # 허용 오차: ±0.25H (15분) 이내는 반올림 자동보정으로 간주
    AUTO_TOL = 0.25

    def _time_warn(total, expected, label):
        diff = abs(total - expected)
        if diff == 0:
            return None
        if diff <= AUTO_TOL:
            return (f"ℹ️  {label} 합계 {total:.2f}H  →  "
                    f"{expected:.0f}H 자동보정 (30분 단위 반올림 오차)")
        elif total < expected:
            return (f"⚠️  {label} 합계 {total:.1f}H  ←  "
                    f"{expected:.0f}H 미만. 시간 재확인 필요")
        else:
            return (f"⚠️  {label} 합계 {total:.1f}H  ←  "
                    f"{expected:.0f}H 초과. 시간 재확인 필요")

    if this_week_raw:
        w = _time_warn(this_total, this_expected, "이번주 실적")
        if w: warns.append(w)
    if next_week_raw:
        w = _time_warn(next_total, next_expected, "다음주 계획")
        if w: warns.append(w)

    # ── 프로젝트 코드 / Function Code 누락 경고 ─────────────────
    seen_missing: set[tuple] = set()
    for e in this_week_raw + next_week_raw:
        tag = e.get("tag")
        if tag not in ("실적", "계획"):
            continue
        pc = e.get("project_code")
        fc = e.get("func_code")
        if pc and fc:
            continue
        dt = e.get("date")
        date_str = dt.strftime("%m월 %d일") if dt else "날짜 미상"
        key = (date_str, e.get("subject", ""))
        if key in seen_missing:
            continue
        seen_missing.add(key)
        missing = []
        if not pc: missing.append("프로젝트 코드")
        if not fc: missing.append("Function Code")
        subj_short = (e.get("subject") or "")[:35]
        warns.append(
            f"⚠️  {date_str} - {' / '.join(missing)} 없음. 재확인 필요"
            + (f"  ※ {subj_short}" if subj_short else "")
        )

    # ── 미등록 Function Code ─────────────────────────────────────
    seen_fc: set[tuple] = set()
    for e in this_week_raw + next_week_raw:
        fc = e.get("func_code")
        if not fc or is_valid_code(fc):
            continue
        dt = e.get("date")
        date_str = dt.strftime("%m월 %d일") if dt else "날짜 미상"
        key = (date_str, fc)
        if key not in seen_fc:
            seen_fc.add(key)
            warns.append(
                f"⚠️  {date_str} - {fc} - 해당 Code는 존재하지 않음. 재확인 필요"
            )

    # ── FC 개수 집계 및 점검 ─────────────────────────────────────
    GENERAL_PC = "000000"
    gen_fcs  = set()
    proj_fcs = set()
    for e in this_week_raw:
        pc = (e.get("project_code") or "").strip()
        fc = (e.get("func_code") or "").strip()
        if not fc:
            continue
        if pc == GENERAL_PC:
            gen_fcs.add(fc)
        else:
            proj_fcs.add((pc, fc))

    total_fc = len(gen_fcs) + len(proj_fcs)
    if this_week_raw:
        # FC 개수 정보 (항상 표시, 경고보다 먼저)
        warns.insert(0,
            f"ℹ️  Function Code: 총 {total_fc}개  "
            f"(General {len(gen_fcs)}개 / Project {len(proj_fcs)}개)"
        )
        if total_fc <= 2:
            warns.append(
                f"⚠️  Function Code {total_fc}개 (3개 이상 되도록 점검 요망)"
            )

    return warns
