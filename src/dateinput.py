"""
dateinput.py — 날짜 입력칸 공통 처리 (메인 화면 주별/일별, M/H 확인, 월간업무정리)

받는 형식 (앞뒤 공백 무시)
  20261001        YYYYMMDD (숫자만)
  261001          YYMMDD   → 2026-10-01
  1001            MMDD     → 올해 10월 1일
  2026-10-01 / 2026.10.1 / 2026/10/01 / 2026 10 01   (구분자 - . / 공백, 한 자리 월·일 허용)

- 없는 날짜(13월, 9월 31일 등)는 이유와 함께 오류 → 그 칸에 커서를 두고 다시 입력
- 인식한 날짜는 칸에 YYYY-MM-DD 로 정리해서 다시 표시
- 지난달 1일보다 이전 날짜(10월 기준 8월 이전)는 "정말 맞습니까?" 확인
"""
from __future__ import annotations
import re
from datetime import datetime

_SEP_RE = re.compile(r"^(\d{2,4})[-./\s](\d{1,2})[-./\s](\d{1,2})$")


class DateInputError(ValueError):
    pass


def parse(text: str, label: str = "날짜", today: datetime | None = None) -> datetime:
    """입력 문자열 → datetime(자정). 형식·날짜가 잘못되면 DateInputError (한글 안내)"""
    s = (text or "").strip()
    today = today or datetime.today()
    if not s:
        raise DateInputError(f"{label} 칸이 비어 있습니다. 날짜를 입력하세요.\n\n예) 20261001 / 261001 / 1001 / 2026-10-01")

    y = m = d = None
    mt = _SEP_RE.match(s)
    if mt:
        y, m, d = (int(g) for g in mt.groups())
    elif s.isdigit() and len(s) == 8:
        y, m, d = int(s[:4]), int(s[4:6]), int(s[6:])
    elif s.isdigit() and len(s) == 6:
        y, m, d = int(s[:2]), int(s[2:4]), int(s[4:])
    elif s.isdigit() and len(s) == 4:
        y, m, d = today.year, int(s[:2]), int(s[2:])
    else:
        raise DateInputError(
            f"{label} '{s}' 을(를) 날짜로 인식할 수 없습니다.\n\n"
            "예) 20261001 / 261001 / 1001 / 2026-10-01")
    if y < 100:
        y += 2000

    if not 1 <= m <= 12:
        raise DateInputError(f"{label} '{s}' : {m}월은 없는 달입니다. (1~12월)\n다시 입력하세요.")
    try:
        return datetime(y, m, d)
    except ValueError:
        import calendar
        last = calendar.monthrange(y, m)[1] if 1900 <= y <= 9999 else 31
        raise DateInputError(
            f"{label} '{s}' : {y}년 {m}월에는 {d}일이 없습니다. (1~{last}일)\n다시 입력하세요.")


def is_too_old(dt: datetime, today: datetime | None = None) -> bool:
    """지난달 1일보다 이전인지 (10월 기준 8월 이전 → True)"""
    today = today or datetime.today()
    y, m = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
    return dt < datetime(y, m, 1)


def read_fields(fields, parent=None) -> list[datetime] | None:
    """
    fields: [(StringVar, Entry 위젯 또는 None, "시작일"), ...]
    모두 인식하면 [datetime, ...] 반환 (칸 값은 YYYY-MM-DD 로 정리).
    오류면 메시지 → 해당 칸에 커서 → None. 오래된 날짜면 확인 창 → [아니오] 시 None.
    """
    from tkinter import messagebox
    out = []
    for var, entry, label in fields:
        try:
            dt = parse(var.get(), label)
        except DateInputError as e:
            messagebox.showerror("날짜 입력 오류", str(e), parent=parent)
            if entry is not None:
                try:
                    entry.focus_set()
                    entry.select_range(0, "end")
                except Exception:
                    pass
            return None
        if var.get() != dt.strftime("%Y-%m-%d"):
            var.set(dt.strftime("%Y-%m-%d"))
        out.append(dt)

    old = [(label, dt) for (_, _, label), dt in zip(fields, out) if is_too_old(dt)]
    if old:
        lines = "\n".join(f"  {label} : {dt:%Y-%m-%d}" for label, dt in old)
        if not messagebox.askyesno(
                "날짜 확인",
                f"지난달보다 이전 날짜가 입력되었습니다.\n\n{lines}\n\n정말 이 날짜가 맞습니까?",
                parent=parent):
            return None
    return out
