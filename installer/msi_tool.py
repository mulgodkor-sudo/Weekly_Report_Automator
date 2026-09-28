"""
msi_tool.py — build_msi.bat 보조 스크립트 (WiX Toolset v3 용)

  python installer/msi_tool.py version          → 프로그램 버전 출력 (src/config.py APP_VERSION)
  python installer/msi_tool.py arch             → 파이썬 비트수에 맞는 WiX arch (x64 / x86)
  python installer/msi_tool.py wxs <dist> <out> → dist 폴더 전체를 설치하는 .wxs 생성

설치 방식 (관리자 권한 불필요, 사용자별 설치):
  %LOCALAPPDATA%\\Programs\\Weekly_Report_Automator\\
    - src\\ 폴더에 쓰기 가능 → apply_patch_vX.X.bat 패치 그대로 사용 가능
    - 시작 메뉴 / 바탕화면 바로가기
    - 같은 버전 재설치·상위 버전 설치 시 이전 버전 자동 제거 (MajorUpgrade)
    - 제거 시 패치로 추가된 src\\ 파일과 __pycache__ 도 함께 삭제

MSI 안의 문자열은 모두 영문(ASCII) — 코드페이지 문제 방지.
파일 경로(한글 폴더 포함)는 UTF-8 .wxs 로 전달되므로 문제 없음.
"""
from __future__ import annotations
import hashlib
import re
import struct
import sys
import uuid
from pathlib import Path
from xml.sax.saxutils import quoteattr

ROOT = Path(__file__).resolve().parent.parent

PRODUCT_NAME = "Weekly Report Automator"
MANUFACTURER = "DL E&C Plant Mechanical Design Team"
INSTALL_DIR  = "Weekly_Report_Automator"
EXE_NAME     = "Weekly_Report_Automator.exe"      # 설치 후 exe 이름 (버전 표기 없음)
# 절대 바꾸지 말 것: 이 값이 같아야 새 버전 MSI가 이전 버전을 업그레이드함
UPGRADE_CODE = "3E4CD511-D247-48DF-9FA7-1EB0360C3D21"
REG_KEY      = r"Software\WeeklyReportAutomator\Installer"
MARKER_FILE  = ROOT / "installer" / "msi_installed.txt"
ICON_FILE    = ROOT / "src" / "assets" / "Schedule_Ico.ico"
NS           = uuid.UUID("6f1c7a52-0d7e-4b8e-9a51-3b0f2f0c9d11")


def app_version() -> str:
    text = (ROOT / "src" / "config.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([0-9]+(?:\.[0-9]+)?)"', text, re.M)
    if not m:
        sys.exit("ERROR: APP_VERSION not found in src/config.py")
    return m.group(1)


def msi_version(v: str) -> str:
    """
    프로그램 버전(1.1 < 1.11 < 1.2 < 1.21 < 1.3) → MSI 버전(숫자 비교)
    소수점 이하를 두 자리로 맞춰 비교 순서를 유지: 1.1→1.10.0, 1.11→1.11.0, 1.2→1.20.0
    """
    major, _, frac = v.partition(".")
    if len(frac) > 2:
        sys.exit(f"ERROR: version '{v}' has more than 2 decimals - not supported")
    return f"{int(major)}.{int(frac.ljust(2, '0') or 0)}.0"


def _id(prefix: str, key: str) -> str:
    return prefix + hashlib.sha1(key.encode("utf-8")).hexdigest()[:20]


def _guid(key: str) -> str:
    return "{" + str(uuid.uuid5(NS, key)).upper() + "}"


def _reg(comp_id: str) -> str:
    # 사용자별 설치 폴더의 컴포넌트는 HKCU 레지스트리 값을 KeyPath 로 둔다 (ICE38)
    return (f'<RegistryValue Root="HKCU" Key="{REG_KEY}" Name="{comp_id}" '
            f'Type="integer" Value="1" KeyPath="yes" />')


def build_wxs(dist: Path, out: Path) -> None:
    dist = dist.resolve()
    exes = [p for p in dist.glob("*.exe")]
    if len(exes) != 1:
        sys.exit(f"ERROR: expected exactly one .exe in {dist}, found {len(exes)}")
    src_exe = exes[0]
    ver = app_version()

    comp_ids: list[str] = []
    lines: list[str] = []

    def emit_dir(path: Path, rel: str, dir_id: str, indent: str):
        files = sorted(p for p in path.iterdir() if p.is_file())
        subs  = sorted(p for p in path.iterdir() if p.is_dir())
        cid   = _id("c", "dir:" + rel)
        comp_ids.append(cid)
        lines.append(f'{indent}<Component Id="{cid}" Guid="{_guid("dir:" + rel)}">')
        lines.append(f'{indent}  {_reg(cid)}')
        lines.append(f'{indent}  <RemoveFolder Id="{_id("r", rel)}" On="uninstall" />')
        if rel in ("src", "src/__pycache__"):
            # 패치로 추가된 파일 / 실행 중 생성된 .pyc 까지 제거
            lines.append(f'{indent}  <RemoveFile Id="{_id("x", rel)}" Name="*" On="uninstall" />')
        for f in files:
            frel = f"{rel}/{f.name}" if rel else f.name
            name = EXE_NAME if f == src_exe else f.name
            lines.append(f'{indent}  <File Id="{_id("f", frel)}" Name={quoteattr(name)} '
                         f'Source={quoteattr(str(f))} />')
        if rel == "":
            lines.append(f'{indent}  <File Id="MarkerFile" Name="msi_installed.txt" '
                         f'Source={quoteattr(str(MARKER_FILE))} />')
        lines.append(f'{indent}</Component>')

        names = {p.name for p in subs}
        for s in subs:
            srel = f"{rel}/{s.name}" if rel else s.name
            sid  = _id("d", srel)
            lines.append(f'{indent}<Directory Id="{sid}" Name={quoteattr(s.name)}>')
            emit_dir(s, srel, sid, indent + "  ")
            lines.append(f'{indent}</Directory>')
        if rel == "src" and "__pycache__" not in names:
            # 실행 중 생기는 src\__pycache__ 정리용 (빈 디렉터리 항목)
            prel = "src/__pycache__"
            cid2 = _id("c", "dir:" + prel)
            comp_ids.append(cid2)
            lines.extend([
                f'{indent}<Directory Id="{_id("d", prel)}" Name="__pycache__">',
                f'{indent}  <Component Id="{cid2}" Guid="{_guid("dir:" + prel)}">',
                f'{indent}    {_reg(cid2)}',
                f'{indent}    <RemoveFile Id="{_id("x", prel)}" Name="*" On="uninstall" />',
                f'{indent}    <RemoveFolder Id="{_id("r", prel)}" On="uninstall" />',
                f'{indent}  </Component>',
                f'{indent}</Directory>',
            ])

    emit_dir(dist, "", "INSTALLFOLDER", "          ")
    tree = "\n".join(lines)
    comp_refs = "\n".join(f'      <ComponentRef Id="{c}" />' for c in comp_ids)

    wxs = f"""<?xml version="1.0" encoding="utf-8"?>
<!-- Generated by installer/msi_tool.py - do not edit -->
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="{PRODUCT_NAME}" Language="1033" Codepage="1252"
           Version="{msi_version(ver)}" Manufacturer={quoteattr(MANUFACTURER)}
           UpgradeCode="{UPGRADE_CODE}">
    <Package InstallerVersion="500" Compressed="yes" InstallScope="perUser"
             InstallPrivileges="limited" Description="{PRODUCT_NAME} {ver}" />
    <MajorUpgrade AllowSameVersionUpgrades="yes"
                  DowngradeErrorMessage="A newer version of [ProductName] is already installed." />
    <MediaTemplate EmbedCab="yes" CompressionLevel="high" />

    <Icon Id="AppIcon.ico" SourceFile={quoteattr(str(ICON_FILE))} />
    <Property Id="ARPPRODUCTICON" Value="AppIcon.ico" />
    <Property Id="ARPNOMODIFY" Value="1" />
    <Property Id="ARPCOMMENTS" Value="Ver.{ver}" />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="LocalAppDataFolder">
        <Directory Id="ProgramsDir" Name="Programs">
          <Component Id="cProgramsDir" Guid="{_guid('programs-dir')}">
            {_reg('cProgramsDir')}
            <RemoveFolder Id="rProgramsDir" On="uninstall" />
          </Component>
          <Directory Id="INSTALLFOLDER" Name="{INSTALL_DIR}">
{tree}
          </Directory>
        </Directory>
      </Directory>
      <Directory Id="ProgramMenuFolder" />
      <Directory Id="DesktopFolder" />
    </Directory>

    <DirectoryRef Id="ProgramMenuFolder">
      <Component Id="cStartMenu" Guid="{_guid('start-menu-shortcut')}">
        <Shortcut Id="StartMenuShortcut" Name="{PRODUCT_NAME}"
                  Target="[INSTALLFOLDER]{EXE_NAME}" WorkingDirectory="INSTALLFOLDER"
                  Icon="AppIcon.ico" />
        {_reg('cStartMenu')}
      </Component>
    </DirectoryRef>
    <DirectoryRef Id="DesktopFolder">
      <Component Id="cDesktop" Guid="{_guid('desktop-shortcut')}">
        <Shortcut Id="DesktopShortcut" Name="{PRODUCT_NAME}"
                  Target="[INSTALLFOLDER]{EXE_NAME}" WorkingDirectory="INSTALLFOLDER"
                  Icon="AppIcon.ico" />
        {_reg('cDesktop')}
      </Component>
    </DirectoryRef>

    <Feature Id="Main" Title="{PRODUCT_NAME}" Level="1">
      <ComponentRef Id="cProgramsDir" />
      <ComponentRef Id="cStartMenu" />
      <ComponentRef Id="cDesktop" />
{comp_refs}
    </Feature>
  </Product>
</Wix>
"""
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(wxs, encoding="utf-8")
    print(f"wxs: {len(comp_ids)} folders, version {ver} (MSI {msi_version(ver)}) -> {out}")


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "version":
        print(app_version())
    elif cmd == "arch":
        print("x64" if struct.calcsize("P") == 8 else "x86")
    elif cmd == "wxs" and len(sys.argv) == 4:
        build_wxs(Path(sys.argv[2]), Path(sys.argv[3]))
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
