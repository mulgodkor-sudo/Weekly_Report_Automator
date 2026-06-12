"""
config.py
설정 관리 + Function Code 테이블 동적 로드
엑셀 파일("기계설계" 포함 파일명)에서 FC 읽기
"""
from __future__ import annotations
import re
import os, sys, json

from fc_rules import get_gubun_from_info, GUBUN_OPTIONS as _GUBUN_OPTIONS

# ── 리소스 경로 (PyInstaller 대응) ──────────────────────────────────
def resource_path(rel: str) -> str:
    """
    번들 리소스(assets 등) 경로 반환.
    frozen: _MEIPASS/assets/...
    개발:   src/assets/...  (config.py가 src/ 안에 있음)
    """
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))  # src/ 폴더
    return os.path.join(base, rel)

def _config_path() -> str:
    """
    config 저장/로드 경로: ~/Documents/WeeklyReportAutomaker_config.json
    exe 위치와 무관하게 사용자 문서 폴더에 저장 (바탕화면 exe 배포 대응)
    """
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    os.makedirs(docs, exist_ok=True)
    return os.path.join(docs, "WeeklyReportAutomaker_config.json")

CONFIG_FILE = _config_path()

_DEFAULT_CONFIG: dict = {
    "func_code_excel_path": "",
    "splash_image": "assets/splash.png",
    "icon_file":    "assets/Schedule_Ico.ico",
}

# ── config.json 읽기/쓰기 ────────────────────────────────────────────
def load_config() -> dict:
    # 1순위: 문서 폴더 (사용자 수정본)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                return {**_DEFAULT_CONFIG, **json.load(f)}
        except Exception:
            pass
    # 2순위: exe 내 번들본 (_MEIPASS) 또는 개발 폴더
    bundled = resource_path("config.json")
    if os.path.exists(bundled) and os.path.abspath(bundled) != os.path.abspath(CONFIG_FILE):
        try:
            with open(bundled, encoding="utf-8") as f:
                data = {**_DEFAULT_CONFIG, **json.load(f)}
            save_config(data)   # 문서 폴더에 복사 (이후 수정 가능)
            return data
        except Exception:
            pass
    return _DEFAULT_CONFIG.copy()

def save_config(cfg: dict) -> None:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ── 구분 선택지 ──────────────────────────────────────────────────────
GUBUN_OPTIONS = _GUBUN_OPTIONS   # fc_rules.py 에서 관리

# 조건부 판단 코드 목록 (KPI 판단이 제목/본문 내용에 따라 달라지는 코드)
CONDITIONAL_CODES = {"GA11-25", "GA08-01"}

# ── FC 테이블 (런타임에 동적 로드) ───────────────────────────────────
_FC: dict[str, dict] = {}

def find_fc_excel_in_folder(folder: str) -> str | None:
    """같은 폴더에서 '기계설계' 포함 엑셀 파일 검색"""
    try:
        for fname in os.listdir(folder):
            if "기계설계" in fname and fname.lower().endswith((".xlsx", ".xlsm")):
                return os.path.join(folder, fname)
    except Exception:
        pass
    return None

def load_fc_from_excel(excel_path: str) -> dict[str, dict]:
    """
    엑셀 파일에서 FC 테이블 로드
    - read_only=True  : 읽기 전용 (동시 접근 허용)
    - data_only=True  : 수식이 아닌 캐시된 결과값 읽기
    - 헤더 자동 감지  : 구분, CODE, KPI, 업무명
    """
    import openpyxl
    fc: dict[str, dict] = {}

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))

        # 헤더 행 탐색 (최대 15행)
        header_idx = None
        col_map: dict[str, int] = {}
        required = {"구분", "CODE", "KPI", "업무명"}

        for r_idx, row in enumerate(all_rows[:15]):
            row_vals = [str(v or "").strip() for v in row]
            if required.issubset(set(row_vals)):
                header_idx = r_idx
                col_map = {v: i for i, v in enumerate(row_vals) if v in required}
                break

        if header_idx is None:
            return fc

        # 데이터 파싱
        for row in all_rows[header_idx + 1:]:
            try:
                group = str(row[col_map["구분"  ]] or "").strip()
                code  = str(row[col_map["CODE"  ]] or "").strip()
                kpi_v = str(row[col_map["KPI"   ]] or "").strip()
                name  = str(row[col_map["업무명"]] or "").strip()

                if not code or not group:
                    continue

                # KPI 판단
                if code in CONDITIONAL_CODES:
                    kpi = "conditional"
                elif "/" in kpi_v and "O" in kpi_v.upper():
                    # "O (LE미팅) / X(그룹리더/팀미팅)" 형식 → conditional
                    kpi = "conditional"
                elif kpi_v.upper() == "O":
                    kpi = True
                else:
                    kpi = False

                fc[code] = dict(group=group, kpi=kpi, name=name)
            except Exception:
                continue
    finally:
        wb.close()

    return fc

def reload_fc(excel_path: str = "") -> None:
    """FC 테이블 재로드 (경로 미입력 시 config.json 참조)"""
    global _FC
    if not excel_path:
        cfg = load_config()
        excel_path = cfg.get("func_code_excel_path", "")
    if excel_path and os.path.exists(excel_path):
        _FC = load_fc_from_excel(excel_path)
    else:
        _FC = {}

def get_fc() -> dict[str, dict]:
    return _FC

# ── 구분(C열) 결정 ───────────────────────────────────────────────────
def get_gubun(func_code: str, title: str = "",
             body: str = "", project_code: str = "") -> str:
    """
    구분 반환. 분류 규칙은 fc_rules.py 에서 관리.
    project_code: General+KPI=O 일 때 수행/기타 구분에 사용.
    """
    info = _FC.get(func_code)
    if not info:
        return "기타"
    return get_gubun_from_info(func_code, info, title, body, project_code)

# ── 편의 함수 ────────────────────────────────────────────────────────
_NAME_PREFIX_RE = re.compile(r"^\[(KPI|Non-KPI)\]\s*")

def get_func_name(func_code: str) -> str:
    name = _FC.get(func_code, {}).get("name", "")
    return _NAME_PREFIX_RE.sub("", name).strip()

def is_valid_code(func_code: str) -> bool:
    return func_code in _FC

def is_highlight(func_code: str) -> bool:
    return False   # 노란 음영 제거됨
