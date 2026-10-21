"""
updater.py — 프로그램 시작 시 ECM 배포 폴더에서 새 버전 패치를 확인하고 적용 (v1.23~)

배포 폴더 (팀 공용 ECM, 모든 PC 동일 경로):
  X:\\부서 문서함\\기계팀\\기계설계\\1_부서 업무\\3-9 기술자료\\13. Outlook_Excel_Extractor_Mechanical\\WRA_Patch
  (config.json 의 "update_folder" 로 변경 가능)

패치 파일명 규칙: WRA_Patch_v<버전>.bat   예) WRA_Patch_v1.23.bat
  - patches/make_patch.py msi 가 이 이름으로 만든다. 배포 담당자는 이 파일만 폴더에 올리면 된다.
  - 폴더에서 규칙에 맞는 파일 중 가장 높은 버전을 찾아, 현재 버전보다 높으면 업데이트를 묻는다.

적용 방식 (스크립트 실행 없음 — 회사 소프트웨어 제한 정책이 %TEMP% 의 .ps1/.bat 실행을 막음):
  - 프로그램이 패치 .bat 안의 패치 데이터(WRA_PAYLOAD 구간의 base64 zip)를 직접 읽어
    exe 옆 src\\ 의 .py / .json 을 교체·검증 → 완료 안내 → 프로그램 다시 시작
  - 교체는 2단계: 모든 파일을 *.wra_new 로 먼저 쓰고, 모두 성공하면 바꿔치기
    (쓰기 권한 문제 등은 1단계에서 걸리므로 그 경우 기존 파일은 그대로)
  - MSI 로 설치된 실행본(exe 옆 msi_installed.txt)에서만 동작 (개발 환경·폴더 복사본 제외)
"""
from __future__ import annotations
# ※ 이 파일은 패치로 "이미 설치된 예전 exe" 위에서도 실행된다.
#   예전 exe 에 번들되지 않았을 수 있는 모듈(subprocess, tempfile, base64 등)은 쓰지 말 것.
#   os / sys / re / binascii(내장) / io / zipfile(xlsxwriter·openpyxl 이 사용 → 항상 번들) 만 사용.
import binascii
import io
import os
import re
import sys
import zipfile

DEFAULT_UPDATE_FOLDER = (r"X:\부서 문서함\기계팀\기계설계\1_부서 업무\3-9 기술자료"
                         r"\13. Outlook_Excel_Extractor_Mechanical\WRA_Patch")
PATCH_RE = re.compile(r"^WRA_Patch_v(\d+)\.(\d{1,2})\.bat$", re.I)
NOTES_RE = re.compile(r"^REM WRA_NOTES_B64=([A-Za-z0-9+/=]+)\s*$", re.M)
PAYLOAD_RE = re.compile(r"^REM WRA_PAYLOAD_BEGIN\s*$(.*?)^REM WRA_PAYLOAD_END\s*$", re.M | re.S)


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


def check(current_version: str) -> dict | None:
    """새 버전이 있으면 {"version", "path", "notes"}, 없거나 확인 불가면 None"""
    latest = find_latest(update_folder())
    if not latest or version_key(latest[0]) <= version_key(current_version):
        return None
    return dict(version=latest[0], path=latest[1], notes=read_notes(latest[1]))


def src_dir() -> str:
    return os.path.join(os.path.dirname(sys.executable), "src")


def apply(patch_path: str) -> int:
    """
    패치 .bat 의 패치 데이터로 exe 옆 src\\ 를 교체. 교체한 파일 수 반환.
    실패하면 예외. 1단계(새 파일 쓰기)에서 실패하면 기존 파일은 그대로.
    """
    with open(patch_path, "rb") as f:
        text = f.read().decode("ascii", "ignore")
    m = PAYLOAD_RE.search(text)
    if not m:
        raise RuntimeError("패치 파일 형식이 올바르지 않습니다 (패치 데이터 없음)")
    b64 = "".join(ln.strip()[5:] for ln in m.group(1).splitlines()
                  if ln.strip().startswith("echo "))
    zf = zipfile.ZipFile(io.BytesIO(binascii.a2b_base64(b64)))
    entries = {}
    for name in zf.namelist():
        base = name.rsplit("/", 1)[-1]
        if name.startswith("src/") and base.endswith((".py", ".json")) and "/" not in name[4:]:
            entries[base] = zf.read(name)
    if "config.py" not in entries:
        raise RuntimeError("패치 데이터가 비어 있습니다")

    target = src_dir()
    staged = []
    try:
        # 1단계: 새 파일을 *.wra_new 로 먼저 쓰기
        for base, data in entries.items():
            tmp = os.path.join(target, base + ".wra_new")
            with open(tmp, "wb") as f:
                f.write(data)
            staged.append((tmp, os.path.join(target, base), data))
        # 2단계: 바꿔치기
        for tmp, dst, _ in staged:
            os.replace(tmp, dst)
    finally:
        for tmp, _, _ in staged:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    # 검증
    bad = []
    for _, dst, data in staged:
        with open(dst, "rb") as f:
            if f.read() != data:
                bad.append(os.path.basename(dst))
    if bad:
        raise RuntimeError("다음 파일이 제대로 교체되지 않았습니다: " + ", ".join(bad))

    # 이전 버전 컴파일 캐시 정리 (실패해도 무시)
    cache = os.path.join(target, "__pycache__")
    try:
        for n in os.listdir(cache):
            try:
                os.remove(os.path.join(cache, n))
            except OSError:
                pass
    except OSError:
        pass
    return len(staged)


def restart() -> None:
    """프로그램 다시 실행 (호출 후 현재 창은 닫을 것)"""
    os.startfile(sys.executable)
