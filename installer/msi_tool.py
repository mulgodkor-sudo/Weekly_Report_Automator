"""
msi_tool.py — build_msi.bat 보조 스크립트 (WiX v5, installer/wix/Product.wxs 용)

  python installer/msi_tool.py version        → 프로그램 버전 (src/config.py APP_VERSION), 예: 1.2
  python installer/msi_tool.py msiver         → MSI ProductVersion, 예: 1.20.0.0
  python installer/msi_tool.py arch           → 파이썬 비트수에 맞는 WiX arch (x64 / x86)
  python installer/msi_tool.py license <out>  → 버전이 들어간 설치 안내문 RTF 생성
  python installer/msi_tool.py pathcheck      → 빌드 폴더 경로 경고 (바탕화면/OneDrive/한글·공백)

결과는 모두 표준출력으로 내보낸다 (bat 에서 파일로 받아 set /p 로 읽음).
"""
from __future__ import annotations
import re
import struct
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def app_version() -> str:
    text = (ROOT / "src" / "config.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([0-9]+(?:\.[0-9]+)?)"', text, re.M)
    if not m:
        sys.exit("ERROR: APP_VERSION not found in src/config.py")
    return m.group(1)


def msi_version(v: str) -> str:
    """
    프로그램 버전(1.1 < 1.11 < 1.2 < 1.21 < 1.3) → MSI ProductVersion(숫자 비교)
    소수점 이하를 두 자리로 맞춰 순서 유지: 1.1→1.10.0.0, 1.11→1.11.0.0, 1.2→1.20.0.0
    """
    major, _, frac = v.partition(".")
    if len(frac) > 2:
        sys.exit(f"ERROR: version '{v}' has more than 2 decimals - not supported")
    return f"{int(major)}.{int(frac.ljust(2, '0') or 0)}.0.0"


# ── 설치 마법사 안내문 (WixUILicenseRtf) ──────────────────────────────
LICENSE_LINES = [
    ("b", "Weekly Report Automator  Ver.{ver}  설치 안내"),
    ("", ""),
    ("", "DL이앤씨 플랜트본부 기계설계팀 내부 업무용 프로그램입니다."),
    ("", "아웃룩 캘린더의 [실적]/[계획] 일정으로 Weekly Report, Plant M/H(ST/OT),"
         " 월간업무정리를 자동으로 만듭니다."),
    ("", ""),
    ("b", "사용 전 확인"),
    ("", "• Microsoft Outlook(데스크탑)이 실행되어 있고 로그인된 상태여야 합니다."),
    ("", "• 첫 실행 후 [상세 설정]에서 Function Code 엑셀 파일 경로를 지정하세요."),
    ("", ""),
    ("b", "설치 / 업데이트"),
    ("", "• 기본 설치 위치: C:\\Program Files\\Autotools_Mechanical\\Weekly Report Automator"),
    ("", "• 새 버전 설치 파일을 실행하면 이전 버전은 자동으로 제거된 뒤 설치됩니다."),
    ("", "• 설정 파일(문서 폴더의 WeeklyReportAutomaker_*.json)은 재설치·제거 후에도 유지됩니다."),
    ("", "• 패치 파일(apply_patch_vX.X.bat)은 설치 폴더에 복사해 실행하면 적용됩니다."),
    ("", ""),
    ("b", "문의"),
    ("", "불편사항/개선사항은 이수신 차장에게 문의 바랍니다."),
]


def _rtf_escape(text: str) -> str:
    out = []
    for ch in text:
        o = ord(ch)
        if ch in "\\{}":
            out.append("\\" + ch)
        elif o < 128:
            out.append(ch)
        else:
            out.append(f"\\u{o if o < 32768 else o - 65536}?")   # RTF 유니코드 (ASCII 파일 유지)
    return "".join(out)


def write_license(out: Path) -> None:
    ver = app_version()
    body = []
    for style, line in LICENSE_LINES:
        t = _rtf_escape(line.format(ver=ver))
        body.append(("{\\b " + t + "}" if style == "b" else t) + "\\par")
    rtf = ("{\\rtf1\\ansi\\ansicpg949\\deff0\\uc1"
           "{\\fonttbl{\\f0\\fnil\\fcharset129 Malgun Gothic;}}"
           "\\f0\\fs18\n" + "\n".join(body) + "\n}\n")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(rtf.encode("ascii"))


def path_warnings() -> list[str]:
    p = str(ROOT)
    warns = []
    low = p.lower()
    if "\\desktop" in low or "onedrive" in low or "바탕 화면" in p:
        warns.append("WARNING: project is under Desktop/OneDrive. Windows 'Controlled folder "
                     "access' or antivirus may block or damage new files (exe, tools).")
        warns.append("         If the build fails, copy the project to a short local path "
                     "such as C:\\WRA and run again.")
    if any(ord(c) > 127 for c in p):
        warns.append("WARNING: project path contains non-ASCII (Korean) characters. "
                     "If a tool fails, use a short ASCII path such as C:\\WRA.")
    return warns


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "version":
        print(app_version())
    elif cmd == "msiver":
        print(msi_version(app_version()))
    elif cmd == "arch":
        print("x64" if struct.calcsize("P") == 8 else "x86")
    elif cmd == "license" and len(sys.argv) == 3:
        write_license(Path(sys.argv[2]))
    elif cmd == "pathcheck":
        # 콘솔 코드페이지와 무관하게 출력되도록 ASCII 메시지만 사용
        for w in path_warnings():
            print(w)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
