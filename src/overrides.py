"""
overrides.py
되풀이 모임 내용 고정 설정 - 데이터 관리
overrides.json 에 저장/로드, 행 적용
"""
from __future__ import annotations
import json, os, re, sys


# ── 저장 경로: ~/Documents (바탕화면 exe 배포 대응) ─────────────────
def _get_path() -> str:
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    os.makedirs(docs, exist_ok=True)
    return os.path.join(docs, "WeeklyReportAutomaker_overrides.json")


def _bundled_path() -> str:
    """exe 내 번들 overrides (초기 배포값)"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "overrides.json")


# ── subject 정규화 (비교용) ───────────────────────────────────────────
# [#코드#코드] 블록 제거
_CODE_RE = re.compile(r"\s*\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]\s*")
# 앞쪽 모든 [태그] 연속 제거 ([GENERAL], [실적/계획], [메뉴얼] 등 모두)
_LEADING_TAGS_RE = re.compile(r"^(\s*\[[^\]]*\]\s*)+")

def _norm(s: str) -> str:
    """매칭 전용 정규화: 코드블록·앞쪽 태그 모두 제거, 순수 업무명만 남김"""
    s = _CODE_RE.sub(" ", s)
    s = _LEADING_TAGS_RE.sub("", s)
    return s.strip()


# ── CRUD ──────────────────────────────────────────────────────────────
def load() -> list[dict]:
    """
    overrides 로드.
    1순위: ~/Documents/WeeklyReportAutomaker_overrides.json (사용자 수정본)
    2순위: exe 내 번들본 (초기 배포값)
    """
    p = _get_path()
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # Fallback: 번들 초기값
    bundled = _bundled_path()
    if os.path.exists(bundled) and os.path.abspath(bundled) != os.path.abspath(p):
        try:
            with open(bundled, encoding="utf-8") as f:
                data = json.load(f)
            save(data)   # 문서 폴더에 복사
            return data
        except Exception:
            pass
    return []


def save(overrides: list[dict]) -> None:
    """overrides.json 저장."""
    with open(_get_path(), "w", encoding="utf-8") as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)


# ── 행 적용 + 적용 건수 ──────────────────────────────────────────────
def count_applied(rows: list[dict], overrides: list[dict]) -> int:
    """
    실제 매칭된 distinct override 건수.
    next_week_only 행에만 적용 (실적 행은 제외).
    """
    if not overrides:
        return 0
    applied: set[tuple] = set()
    for row in rows:
        if row.get("source") != "next_week_only":   # [계획] 신규 행만
            continue
        pc   = (row.get("project_code") or "").strip()
        fc   = (row.get("func_code")    or "").strip()
        subj = _norm(row.get("subject") or "")
        for ov in overrides:
            if (
                (ov.get("project_code") or "").strip() == pc
                and (ov.get("func_code") or "").strip() == fc
                and _norm(ov.get("subject") or "") == subj
            ):
                applied.add(((ov.get("project_code") or "").strip(),
                              (ov.get("func_code") or "").strip(),
                              _norm(ov.get("subject") or "")))
                break
    return len(applied)


def apply(rows: list[dict], overrides: list[dict]) -> list[dict]:
    """
    override 적용 규칙:
      - source == "next_week_only" 인 행에만 적용
        (이번주 실적에 없고 다음주 계획으로만 새로 추가된 행)
      - source == "this_week" 인 행은 실적 내용 그대로 유지
        (이번주 실적이 있고 차주 시간만 추가된 경우 body 불변)
    """
    if not overrides:
        return rows

    result = []
    for row in rows:
        # ── [계획] 신규 행만 override 적용 ──────────────────────────
        if row.get("source") == "next_week_only":
            pc   = (row.get("project_code") or "").strip()
            fc   = (row.get("func_code")    or "").strip()
            subj = _norm(row.get("subject") or "")
            for ov in overrides:
                if (
                    (ov.get("project_code") or "").strip() == pc
                    and (ov.get("func_code") or "").strip() == fc
                    and _norm(ov.get("subject") or "") == subj
                ):
                    row = dict(row)
                    row["body"] = ov.get("body", "")
                    break
        # ── this_week 행: 실적 body 그대로 ──────────────────────────
        result.append(row)
    return result
