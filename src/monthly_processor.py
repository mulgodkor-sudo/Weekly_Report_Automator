"""
monthly_processor.py  ─  월간 업무 정리 처리 로직 v2
변경사항:
  - 같은 날 + 같은 내용 → 시간 합산 (행 병합)
  - 다른 날 + 같은 내용 → 날짜 추가 (행 추가 X)
  - 별표(*) 및 이후 내용 제거 (개수 무관)
  - FC 엑셀 KPI 컬럼 기반 Non-KPI 항목 하단 배치
  - 표준서/메뉴얼/매뉴얼/절차서 관련 항목 인접 배치
"""
from __future__ import annotations
import re
from datetime import datetime
from collections import OrderedDict

from config import get_gubun as _cfg_gubun

HOLIDAY_FC     = "GE04-02"
MANUAL_KEYWORDS = ["표준서", "메뉴얼", "매뉴얼", "절차서"]
DAY_KO         = ["월","화","수","목","금","토","일"]
ORDER_GUBUN    = {"수행 (KPI)": 0, "입찰": 1, "기타 (KPI)": 2, "기타": 3}

_TRAIL_RE   = re.compile(r"\s*\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]\s*")
_LEADING_RE = re.compile(r"^(\s*\[[^\]]*]\s*)+")


def _to_date(dt):
    return dt.date() if isinstance(dt, datetime) else dt

def _fmt_date(dt) -> str:
    d = _to_date(dt)
    return f"{d.month}/{d.day}({DAY_KO[d.weekday()]})"

def _clean_subj(s: str) -> str:
    s = _TRAIL_RE.sub(" ", s or "")
    s = _LEADING_RE.sub("", s)
    return s.strip()

def _strip_star(text: str) -> str:
    """별표(*) 및 이후 전체 제거 (개수 무관)"""
    m = re.search(r'\*', text)
    if m:
        return text[:m.start()].rstrip('\n').rstrip()
    return text

def _is_manual(subject: str) -> bool:
    return any(kw in subject for kw in MANUAL_KEYWORDS)

def _is_non_kpi(fc: str, fc_data: dict | None) -> bool:
    """KPI 컬럼 첫 글자가 X → Non-KPI (하단 배치)"""
    if not fc_data or not fc:
        return False
    entry = fc_data.get(fc, {})
    return fc_data.get(fc, {}).get("kpi") is False


def process_monthly(events: list[dict], fc_data: dict | None = None) -> dict:
    """
    [실적] 이벤트 → 월간 업무 정리

    반환:
      {
        'holiday': {'dates': [str], 'total_h': float},
        'groups':  [GroupDict]
      }

    GroupDict:
      gubun, project_code, func_code, subject, total_h,
      rows: [BodyRow]

    BodyRow:
      body: str               (별표 제거 후)
      dates: [str]            (중복 날짜 병합 후 날짜 목록)
      date_hours: {str:float} (날짜별 시간)
      total_h: float
    """
    this_events = [e for e in events if e.get("tag") == "실적"]

    holiday_dates: list[str] = []
    holiday_hours = 0.0
    outer: dict[tuple, dict] = OrderedDict()   # (pc,fc,subject) → group

    for e in sorted(this_events, key=lambda x: _to_date(x.get("date", datetime.min))):
        fc = (e.get("func_code") or "").strip()

        if fc == HOLIDAY_FC:
            holiday_dates.append(_fmt_date(e["date"]))
            holiday_hours = round(holiday_hours + e.get("hours", 0.0), 1)
            continue

        pc   = (e.get("project_code") or "").strip()
        subj = _clean_subj(e.get("subject") or "")
        key  = (pc, fc, subj)

        if key not in outer:
            outer[key] = {
                "gubun":        _cfg_gubun(fc, e.get("subject",""), e.get("body",""), pc),
                "project_code": pc,
                "func_code":    fc,
                "subject":      subj,
                "body_map":     OrderedDict(),   # body → {date_str: hours}
                "total_h":      0.0,
            }

        g     = outer[key]
        body  = _strip_star((e.get("body") or "").strip())
        dstr  = _fmt_date(e["date"])
        hours = e.get("hours", 0.0)
        g["total_h"] = round(g["total_h"] + hours, 1)

        if body not in g["body_map"]:
            g["body_map"][body] = {}

        dm = g["body_map"][body]
        # 같은 날 + 같은 내용 → 시간 합산 / 다른 날 + 같은 내용 → 날짜 추가
        dm[dstr] = round(dm.get(dstr, 0.0) + hours, 1)

    # body_map → rows 변환
    result_groups = []
    for g in outer.values():
        rows = []
        for body, dm in g["body_map"].items():
            dates = list(dm.keys())
            rows.append({
                "body":       body,
                "dates":      dates,
                "date_hours": dm,
                "total_h":    round(sum(dm.values()), 1),
            })
        result_groups.append({
            "gubun":        g["gubun"],
            "project_code": g["project_code"],
            "func_code":    g["func_code"],
            "subject":      g["subject"],
            "rows":         rows,
            "total_h":      g["total_h"],
        })

    # ── 정렬 ──────────────────────────────────────────────────────
    # ── 정렬 ──────────────────────────────────────────────────────────────
    # gubun 값 자체로 순서 결정:
    #   수행(KPI)=0 → 기타(KPI)=2 → 기타=3 (GA11-25 포함, conditional도 "기타"로 해결)
    # _is_non_kpi 방식은 "conditional" 코드를 처리 못하므로 gubun 기반으로 변경
    result_groups.sort(key=lambda g: (
        ORDER_GUBUN.get(g["gubun"], 9),      # 구분 순서 (기타=3 → 하단)
        1 if _is_manual(g["subject"]) else 0, # 매뉴얼류 인접
        g["project_code"],
        g["func_code"],
        g["subject"],
    ))

    return {
        "holiday": {"dates": holiday_dates, "total_h": holiday_hours},
        "groups":  result_groups,
    }
