"""
updater.py — 프로그램 시작 시 ECM 배포 폴더에서 새 버전 패치를 확인하고 적용 (v1.23~)

배포 폴더 (팀 공용 ECM, 모든 PC 동일 경로):
  X:\\부서 문서함\\기계팀\\기계설계\\1_부서 업무\\3-9 기술자료\\13. Outlook_Excel_Extractor_Mechanical\\WRA_Patch
  (config.json 의 "update_folder" 로 변경 가능)

패치 파일명 규칙: WRA_Patch_v<버전>.bat   예) WRA_Patch_v1.23.bat
  - patches/make_patch.py msi 가 이 이름으로 만든다. 배포 담당자는 이 파일만 폴더에 올리면 된다.
  - 폴더에서 규칙에 맞는 파일 중 가장 높은 버전을 찾아, 현재 버전보다 높으면 업데이트를 묻는다.

적용 방식:
  - 패치 .bat 을 %TEMP% 로 복사 → "auto <현재 PID>" 인자로 새 창에서 실행 → 프로그램 종료
  - 패치가 프로그램 종료를 기다렸다가 src\\ 교체·검증 후 프로그램을 다시 실행
  - MSI 로 설치된 실행본(exe 옆 msi_installed.txt)에서만 동작 (개발 환경·폴더 복사본 제외)
"""
from __future__ import annotations
# ※ 이 파일은 패치로 "이미 설치된 예전 exe" 위에서도 실행된다.
#   예전 exe 에 번들되지 않았을 수 있는 모듈(subprocess, tempfile, base64 등)은 쓰지 말 것.
#   os / sys / re / binascii(내장) 만 사용.
import binascii
import os
import re
import sys

DEFAULT_UPDATE_FOLDER = (r"X:\부서 문서함\기계팀\기계설계\1_부서 업무\3-9 기술자료"
                         r"\13. Outlook_Excel_Extractor_Mechanical\WRA_Patch")
PATCH_RE = re.compile(r"^WRA_Patch_v(\d+)\.(\d{1,2})\.bat$", re.I)
NOTES_RE = re.compile(r"^REM WRA_NOTES_B64=([A-Za-z0-9+/=]+)\s*$", re.M)


def version_key(v: str) -> int:
    """1.1 < 1.11 < 1.2 < 1.22 < 1.3 순서 (소수부를 두 자리로 맞춰 비교)"""
    major, _, frac = v.partition(".")
    return int(major) * 100 + int((frac or "0").ljust(2, "0")[:2])


def is_msi_install() -> bool:
    if not getattr(sys, "frozen", False):
        return False
    exe_dir = os.path.dirname(sys.executable)
    return os.path.exists(os.path.join(exe_dir, "msi_installed.txt"))


def update_folder() -> str:
    try:
        from config import load_config
        return load_config().get("update_folder") or DEFAULT_UPDATE_FOLDER
    except Exception:
        return DEFAULT_UPDATE_FOLDER


def find_latest(folder: str) -> tuple[str, str] | None:
    """(버전, 파일 경로) — 규칙에 맞는 패치 중 가장 높은 버전. 폴더 접근 불가 시 None"""
    try:
        names = os.listdir(folder)
    except OSError:
        return None
    best = None
    for n in names:
        m = PATCH_RE.match(n)
        if not m:
            continue
        ver = f"{int(m.group(1))}.{m.group(2)}"
        if best is None or version_key(ver) > version_key(best[0]):
            best = (ver, os.path.join(folder, n))
    return best


def read_notes(bat_path: str) -> str:
    """패치 .bat 안의 'REM WRA_NOTES_B64=' 줄에서 패치 내용(UTF-8) 읽기"""
    try:
        with open(bat_path, "rb") as f:
            head = f.read(64 * 1024).decode("ascii", "ignore")
        m = NOTES_RE.search(head)
        return binascii.a2b_base64(m.group(1)).decode("utf-8") if m else ""
    except Exception:
        return ""


def _cleanup_old_copies() -> None:
    """이전 업데이트 때 %TEMP% 에 복사해 둔 패치 정리 (실행 중이던 .bat 은 스스로 못 지움)"""
    base = os.environ.get("TEMP") or ""
    try:
        names = os.listdir(base)
    except OSError:
        return
    for n in names:
        if not n.startswith("wra_update_") or n == f"wra_update_{os.getpid()}":
            continue
        d = os.path.join(base, n)
        try:
            for f in os.listdir(d):
                os.remove(os.path.join(d, f))
            os.rmdir(d)
        except OSError:
            pass


def check(current_version: str) -> dict | None:
    """새 버전이 있으면 {"version", "path", "notes"}, 없거나 확인 불가면 None"""
    _cleanup_old_copies()
    latest = find_latest(update_folder())
    if not latest or version_key(latest[0]) <= version_key(current_version):
        return None
    return dict(version=latest[0], path=latest[1], notes=read_notes(latest[1]))


def launch(patch_path: str) -> None:
    """패치를 %TEMP% 로 복사해 새 창에서 자동 모드로 실행 (호출 후 프로그램은 바로 종료할 것)"""
    tmp_dir = os.path.join(os.environ.get("TEMP") or os.path.expanduser("~"),
                           f"wra_update_{os.getpid()}")
    os.makedirs(tmp_dir, exist_ok=True)
    local = os.path.join(tmp_dir, os.path.basename(patch_path))
    with open(patch_path, "rb") as src, open(local, "wb") as dst:   # 네트워크 폴더에서 복사
        dst.write(src.read())
    # ShellExecute 로 새 콘솔 창에서 실행: WRA_Patch_vX.X.bat auto <PID>
    os.startfile(local, "open", f"auto {os.getpid()}", tmp_dir)
