"""
fc_rules.py  ─  Function Code 구분 분류 규칙
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
★ 패치 주요 대상 파일 ★
  구분(수행/기타 등) 판단 로직이 변경될 때 이 파일만 수정·배포.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[구분 판단 로직 전체 흐름]

입력값:
  fc_info      : _FC 딕셔너리에서 가져온 FC 정보
                 {group: "Project"|"General", kpi: True|False|"conditional", name: str}
  project_code : 아웃룩 일정의 프로젝트 코드 (예: P2600O, 000000)
  title        : 아웃룩 일정 제목 (conditional 판단에 사용)
  body         : 아웃룩 일정 본문 (conditional 판단에 사용)

판단 순서:
  1. fc_info 없음 (미등록 FC)            → "기타"

  2. kpi == "conditional" (KPI가 O/X 혼재):
     a. GA08-01 (절차서):
        - 제목 또는 본문에 "매뉴얼" 포함  → "기타 (KPI)"
        - 없으면                           → "기타"
     b. 나머지 conditional (GA11-25 등):  → "기타" (사실상 Non-KPI)

  3. kpi == True (KPI=O):
     a. group == "Project"                → "수행 (KPI)"
     b. group == "General":
        - project_code == GENERAL_PC      → "기타 (KPI)"  (사내 일반 업무)
        - project_code != GENERAL_PC      → "수행 (KPI)"  (외부 프로젝트 투입)

  4. kpi == False (KPI=X)               → "기타"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[예외 케이스 요약]

  일반 업무 PC   : GENERAL_PC = "000000"
                   이 코드가 붙은 General 업무 → "기타 (KPI)"
                   다른 PC가 붙은 General 업무 → "수행 (KPI)"
                   (EPC 프로젝트에 투입된 General FC 항목)

  GA08-01 절차서 : conditional로 저장됨
                   매뉴얼 관련이면 KPI, 아니면 Non-KPI

  GA11-25 팀미팅 : conditional로 저장됨
                   사실상 Non-KPI (기타) 취급

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[수정 가이드]

  구분 로직 바꿀 때:
    → get_gubun_from_info() 함수 수정

  GENERAL_PC 바꿀 때:
    → GENERAL_PC 상수만 수정

  새 conditional FC 추가할 때:
    → _CONDITIONAL_RULES 딕셔너리에 규칙 추가

  구분 옵션 추가할 때:
    → GUBUN_OPTIONS 리스트에 추가
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

# ── 상수 ─────────────────────────────────────────────────────────
GENERAL_PC = "000000"       # 사내 일반 업무 프로젝트 코드
                            # 변경 시 여기만 수정

GUBUN_OPTIONS = [           # 드롭다운 선택 목록 (Excel 유효성 검사)
    "수행 (KPI)",
    "입찰",
    "기타 (KPI)",
    "기타",
]

# ── conditional FC 별 개별 규칙 ───────────────────────────────────
# {fc_code: callable(title, body) -> str | None}
# None 반환 시 기본값("기타") 사용
_CONDITIONAL_RULES: dict[str, object] = {
    "GA08-01": lambda title, body: (
        "기타 (KPI)" if "매뉴얼" in (title + body) else "기타"
    ),
    # 새 conditional FC 추가 예시:
    # "GA00-00": lambda title, body: "수행 (KPI)" if "특정키워드" in title else "기타",
}


# ── 핵심 분류 함수 ────────────────────────────────────────────────

def get_gubun_from_info(
    fc_code:      str,
    fc_info:      dict,
    title:        str = "",
    body:         str = "",
    project_code: str = "",
) -> str:
    """
    fc_info 와 project_code 를 받아 구분 문자열 반환.
    config.py 의 get_gubun() 가 _FC 조회 후 이 함수를 호출함.
    """
    group = fc_info.get("group", "General")
    kpi   = fc_info.get("kpi",   False)

    # ── conditional (KPI O/X 혼재) ────────────────────────────────
    if kpi == "conditional":
        rule = _CONDITIONAL_RULES.get(fc_code)
        if rule:
            return rule(title, body)
        return "기타"   # 규칙 없는 conditional → Non-KPI 취급

    # ── KPI=O ─────────────────────────────────────────────────────
    if kpi is True:
        if group == "Project":
            return "수행 (KPI)"

        # General + KPI=O
        pc = (project_code or "").strip()
        if pc and pc != GENERAL_PC:
            # 사내 코드(000000)가 아닌 프로젝트 코드 → 외부 투입
            return "수행 (KPI)"
        return "기타 (KPI)"

    # ── KPI=False (X) ────────────────────────────────────────────
    return "기타"
