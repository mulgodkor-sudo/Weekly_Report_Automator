"""
outlook_reader.py
아웃룩 캘린더 이벤트 접근 및 파싱

exe 호환 핵심 사항:
  - CoInitialize/CoUninitialize: app.py task 스레드에서만 1회 처리
  - GetActiveObject만 사용: Outlook 미실행 시 즉시 오류 안내
  - GetActiveObject 우선: 실행 중인 Outlook 인스턴스를 직접 사용
  - Restrict + 수동 날짜 검증 이중 적용: Restrict 부정확 시 대비
"""
from __future__ import annotations
import re
from datetime import datetime
from html.parser import HTMLParser


# ══════════════════════════════════════════════════════════
# HTML → 평문 변환
# ══════════════════════════════════════════════════════════

class _HTMLTextParser(HTMLParser):
    SKIP_TAGS  = {"head", "style", "script", "title", "meta"}
    BLOCK_TAGS = {"br", "p", "div", "tr", "h1", "h2", "h3", "h4",
                  "li", "td", "th", "hr", "section", "article"}
    HTML_ENTITIES = {
        "amp": "&", "lt": "<", "gt": ">", "nbsp": " ",
        "quot": '"', "apos": "'", "hellip": "…", "ndash": "–",
    }

    def __init__(self):
        super().__init__()
        self._buf: list[str] = []
        self._skip = False

    def handle_starttag(self, tag: str, attrs):
        t = tag.lower()
        if t in self.SKIP_TAGS:
            self._skip = True
        elif not self._skip and t in self.BLOCK_TAGS:
            self._buf.append("\n")

    def handle_endtag(self, tag: str):
        if tag.lower() in self.SKIP_TAGS:
            self._skip = False

    def handle_data(self, data: str):
        if not self._skip:
            self._buf.append(data)

    def handle_entityref(self, name: str):
        if not self._skip:
            self._buf.append(self.HTML_ENTITIES.get(name, ""))

    def handle_charref(self, name: str):
        if self._skip:
            return
        try:
            ch = chr(int(name[1:], 16) if name.startswith("x") else int(name))
            self._buf.append(ch)
        except Exception:
            pass

    def get_text(self) -> str:
        raw = "".join(self._buf)
        lines = [ln.strip() for ln in raw.splitlines()]
        result: list[str] = []
        prev_blank = False
        for ln in lines:
            if not ln:
                if not prev_blank:
                    result.append("")
                prev_blank = True
            else:
                result.append(ln)
                prev_blank = False
        return "\n".join(result).strip()


def html_to_text(html: str) -> str:
    if not html:
        return ""
    parser = _HTMLTextParser()
    try:
        parser.feed(html)
        return parser.get_text()
    except Exception:
        text = re.sub(r"<[^>]+>", " ", html)
        return re.sub(r"\s+", " ", text).strip()


# ══════════════════════════════════════════════════════════
# 제목 파싱
# ══════════════════════════════════════════════════════════

_TAG_RE   = re.compile(r"^\s*\[(실적|계획)\]")
_CODE_RE  = re.compile(                           # 파싱용 (그룹 캡처)
    r"\[\s*#\s*([A-Za-z0-9]{5,8})\s*#\s*([A-Z]{2}\s*\d{2}\s*-\s*\d{2})\s*\]"
)
_CODE_STRIP = re.compile(                         # 제목 정제용 (블록 전체 제거)
    r"\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]"
)


def parse_subject(subject: str) -> dict:
    tag = None
    m = _TAG_RE.search(subject)
    if m:
        tag = m.group(1)
    project_code = func_code = None
    m2 = _CODE_RE.search(subject)
    if m2:
        project_code = m2.group(1).replace(" ", "")
        func_code    = m2.group(2).replace(" ", "")   # 공백 제거 정규화
    return dict(tag=tag, project_code=project_code, func_code=func_code)


# ══════════════════════════════════════════════════════════
# 이벤트 시간 계산
# ══════════════════════════════════════════════════════════

def _win_to_dt(win_time) -> datetime:
    return datetime(
        win_time.year, win_time.month, win_time.day,
        win_time.hour, win_time.minute, win_time.second,
    )


def _round_half(h: float) -> float:
    """30분(0.5H) 단위 반올림. 55분→1H, 1시간25분→1.5H"""
    return round(h * 2) / 2


def event_hours(item) -> float:
    try:
        if item.AllDayEvent:
            return 8.0
        start = _win_to_dt(item.Start)
        end   = _win_to_dt(item.End)
        h = (end - start).total_seconds() / 3600
        return _round_half(max(h, 0.0))
    except Exception:
        return 0.0


# ══════════════════════════════════════════════════════════
# Outlook 연결 헬퍼
# ══════════════════════════════════════════════════════════

def _get_outlook():
    """
    실행 중인 Outlook 인스턴스만 획득.
    Outlook 미실행 시 즉시 RuntimeError.
    """
    import win32com.client
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        raise RuntimeError(
            "Outlook이 실행 중이지 않습니다.\n"
            "아웃룩을 먼저 실행하고 계정 "
            "로그인 후 다시 시도해 주세요."
        )



# ══════════════════════════════════════════════════════════
# 캘린더 이벤트 조회
# ══════════════════════════════════════════════════════════

def get_events(start: datetime, end: datetime) -> list[dict]:
    """
    지정 기간(start ~ end)의 아웃룩 캘린더 이벤트 반환.
    CoInitialize/CoUninitialize 는 app.py task 스레드에서 처리함.
    """
    try:
        import win32com.client  # noqa: F401 (import 가능 여부 확인)
    except ImportError:
        raise RuntimeError(
            "pywin32 패키지가 없습니다.\n"
            "cmd에서 'pip install pywin32' 실행 후 재시작하세요."
        )

    try:
        outlook = _get_outlook()
        ns  = outlook.GetNamespace("MAPI")
        cal = ns.GetDefaultFolder(9)   # olFolderCalendar

        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Restrict 필터 (US 날짜 포맷 – Outlook MAPI 표준)
        restrict_str = (
            f"[Start] >= '{start.strftime('%m/%d/%Y 12:00 AM')}'"
            f" AND [Start] <= '{end.strftime('%m/%d/%Y 11:59 PM')}'"
        )
        filtered = items.Restrict(restrict_str)

        # 날짜 범위 (수동 2차 검증용)
        start_dt = datetime(start.year, start.month, start.day)
        end_dt   = datetime(end.year,   end.month,   end.day, 23, 59, 59)

        results:     list[dict] = []
        item_errors: list[str]  = []

        for item in filtered:
            try:
                dt_start   = _win_to_dt(item.Start)
                event_date = datetime(dt_start.year, dt_start.month, dt_start.day)

                # Restrict 부정확 대비: 날짜 수동 검증
                if event_date < start_dt or event_date > end_dt:
                    continue

                parsed    = parse_subject(item.Subject or "")
                body_text = ""
                try:
                    body_text = html_to_text(item.HTMLBody or "")
                except Exception:
                    try:
                        body_text = (item.Body or "").strip()
                    except Exception:
                        pass

                results.append(dict(
                    tag          = parsed["tag"],
                    subject      = _CODE_STRIP.sub(
                        "", (item.Subject or "")
                    ).strip(),   # [#PC#FC] 블록 사전 제거 → 병합키 일관성 보장
                    project_code = parsed["project_code"],
                    func_code    = parsed["func_code"],
                    date         = event_date,
                    body         = body_text,
                    hours        = event_hours(item),
                    is_all_day   = bool(item.AllDayEvent),
                ))
            except Exception as item_err:
                # 개별 항목 오류는 수집 (나중에 진단에 활용)
                item_errors.append(str(item_err))

        # 결과가 없고 개별 오류가 있으면 대표 오류를 RuntimeError로 전파
        if not results and item_errors:
            sample = item_errors[0][:120]
            raise RuntimeError(
                f"일정 항목을 읽는 중 오류 발생 ({len(item_errors)}건):\n{sample}\n\n"
                "아웃룩이 실행 중인지, 계정이 로그인되어 있는지 확인하세요."
            )

        return results

    except RuntimeError:
        raise  # 위에서 만든 RuntimeError는 그대로 전파
    except Exception as e:
        raise RuntimeError(
            f"아웃룩 캘린더 접근 실패:\n{type(e).__name__}: {e}\n\n"
            "아웃룩이 실행 중인지, 계정이 로그인되어 있는지 확인하세요.\n"
            "※ exe 실행 시: 반드시 Outlook을 먼저 실행하고 로그인 후 사용하세요."
        )
