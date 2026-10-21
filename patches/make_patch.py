"""
make_patch.py — 단일 .bat 패치 파일 생성기

사용법 (저장소 루트에서):
  python patches/make_patch.py 1.2 1.11 1.1     (onedir 폴더용, ~v1.2)
  python patches/make_patch.py msi               (MSI 설치본용, v1.22~)

  첫 번째 인자 : 새 버전 (예: 1.2)
  나머지 인자  : 이 패치를 적용할 수 있는 이전 버전들 (exe/폴더명 V표기 변경용)
                 ※ "V1.1" 은 "V1.11" 의 앞부분이므로 긴 버전부터 자동 정렬해 처리

src/ 의 *.py, *.json 전체를 zip → base64 로 .bat 안에 넣는다.
.bat 실행 시 exe 옆 src/ 폴더를 통째로 덮어쓰고, exe 파일명과 폴더명의
버전 표기를 새 버전으로 바꾼다. 백업은 만들지 않고 임시파일은 항상 삭제한다.
MSI 로 설치된 폴더(msi_installed.txt 존재)에서는 이름을 바꾸지 않는다 (바로가기 유지).
"""
from __future__ import annotations
import base64
import io
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC  = ROOT / "src"

TEMPLATE = r"""@echo off
setlocal EnableDelayedExpansion

set "PATCH_VERSION=@NEW@"
set "OLD_TAGS=@OLD_TAGS@"
set "TARGET_DIR=%~dp0src"
set "TMP_B64=%TEMP%\wra_patch_%RANDOM%.b64"
set "TMP_ZIP=%TEMP%\wra_patch_%RANDOM%.zip"
set "NEW_TAG=V%PATCH_VERSION%"

echo ============================================================
echo  Weekly Report Automator - Patch to v%PATCH_VERSION%
echo  (from: @OLD_LIST@)
echo ============================================================
echo.

if not exist "%TARGET_DIR%" (
    echo ERROR: "src" folder not found next to this file.
    echo Place this .bat in the same folder as Weekly_Report_Automator*.exe and run it again.
    pause
    exit /b 1
)

echo Writing patch data...
> "%TMP_B64%" (
@DATA@
)

echo Decoding patch data...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$b64 = (Get-Content -Raw '%TMP_B64%') -replace '\s',''; [IO.File]::WriteAllBytes('%TMP_ZIP%', [Convert]::FromBase64String($b64))"
if errorlevel 1 (
    echo ERROR: decode failed. Your files were not changed.
    del "%TMP_B64%" "%TMP_ZIP%" >nul 2>&1
    pause
    exit /b 1
)

echo Applying patch to "%TARGET_DIR%"...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%TMP_ZIP%' -DestinationPath '%TARGET_DIR%' -Force"
if errorlevel 1 (
    echo ERROR: patch apply failed.
    del "%TMP_B64%" "%TMP_ZIP%" >nul 2>&1
    pause
    exit /b 1
)

del "%TMP_B64%" "%TMP_ZIP%" >nul 2>&1

REM MSI install: keep exe/folder names (Start menu and desktop shortcuts point to them)
if exist "%~dp0msi_installed.txt" (
    echo.
    echo Installed by MSI - exe and folder names are kept unchanged.
    goto :patch_done
)

echo.
echo Renaming exe file to v%PATCH_VERSION%...
set "OLD_EXE="
for %%F in ("%~dp0*.exe") do set "OLD_EXE=%%~nxF"
if not defined OLD_EXE (
    echo WARNING: no .exe found next to this file. Skipping exe rename.
) else (
    set "CHECK_NEW=!OLD_EXE:%NEW_TAG%=!"
    if not "!CHECK_NEW!"=="!OLD_EXE!" (
        echo Exe is already named for v%PATCH_VERSION%: !OLD_EXE!
    ) else (
        set "NEW_EXE="
        for %%T in (%OLD_TAGS%) do if not defined NEW_EXE (
            set "TRY=!OLD_EXE:%%T=%NEW_TAG%!"
            if not "!TRY!"=="!OLD_EXE!" set "NEW_EXE=!TRY!"
        )
        if not defined NEW_EXE (
            echo WARNING: exe filename has no old version pattern - rename it manually: !OLD_EXE!
        ) else (
            ren "%~dp0!OLD_EXE!" "!NEW_EXE!"
            if errorlevel 1 (
                echo WARNING: could not rename exe. Close the program if running, then rename manually to: !NEW_EXE!
            ) else (
                echo Renamed exe: !OLD_EXE! -^> !NEW_EXE!
            )
        )
    )
)

echo.
echo Renaming program folder to v%PATCH_VERSION%...
for %%D in ("%~dp0.") do set "OLD_DIRNAME=%%~nxD"
for %%D in ("%~dp0..") do set "PARENT_DIR=%%~fD"
set "CHECK_DNEW=!OLD_DIRNAME:%NEW_TAG%=!"
if not "!CHECK_DNEW!"=="!OLD_DIRNAME!" (
    echo Folder is already named for v%PATCH_VERSION%: !OLD_DIRNAME!
) else (
    set "NEW_DIRNAME="
    for %%T in (%OLD_TAGS%) do if not defined NEW_DIRNAME (
        set "TRY=!OLD_DIRNAME:%%T=%NEW_TAG%!"
        if not "!TRY!"=="!OLD_DIRNAME!" set "NEW_DIRNAME=!TRY!"
    )
    if not defined NEW_DIRNAME (
        echo WARNING: folder name has no old version pattern - rename it manually: !OLD_DIRNAME!
    ) else (
        cd /d "%PARENT_DIR%"
        ren "!OLD_DIRNAME!" "!NEW_DIRNAME!"
        if exist "%PARENT_DIR%\!NEW_DIRNAME!" (
            echo Renamed folder: !OLD_DIRNAME! -^> !NEW_DIRNAME!
        ) else (
            echo WARNING: could not rename folder. Close this window, the program, and any
            echo Explorer windows open inside it, then rename the folder manually to: !NEW_DIRNAME!
        )
    )
)

:patch_done
echo.
echo  PATCH SUCCESS - now at v%PATCH_VERSION%
echo  Please restart the program for the changes to take effect.
pause
"""


def build(new: str, olds: list[str]) -> Path:
    files = sorted(p for p in SRC.iterdir()
                   if p.is_file() and p.suffix in (".py", ".json"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in files:
            zf.write(p, p.name)
    b64  = base64.b64encode(buf.getvalue()).decode("ascii")
    data = "\n".join(f"echo {b64[i:i + 76]}" for i in range(0, len(b64), 76))

    # 긴 버전 먼저 ("V1.11" 을 "V1.1" 보다 먼저 치환해야 V1.21 같은 오류가 없음)
    olds = sorted(olds, key=len, reverse=True)
    text = (TEMPLATE
            .replace("@NEW@", new)
            .replace("@OLD_TAGS@", " ".join(f"V{v}" for v in olds))
            .replace("@OLD_LIST@", " / ".join(f"v{v}" for v in olds))
            .replace("@DATA@", data))
    text.encode("ascii")   # 비ASCII 문자가 있으면 여기서 오류 (cmd 코드페이지 문제 방지)

    out = Path(__file__).resolve().parent / f"apply_patch_v{new}.bat"
    out.write_bytes(text.replace("\n", "\r\n").encode("ascii"))
    print(f"{out.name}: {len(files)} files -> {', '.join(p.name for p in files)}")
    return out


# ════════════════════════════════════════════════════════════════════
# MSI 설치본용 패치 (v1.22~)
#   python patches/make_patch.py msi
#   → patches/apply_patch_v<APP_VERSION>.bat
#   - 아무 폴더에서나 실행 → MSI 고정 설치 경로
#     (C:\Program Files\Autotools_Mechanical\Weekly Report Automator) 를 찾아 src\ 교체
#   - 버전은 src/config.py APP_VERSION, 패치 내용은 patches/notes_v<버전>.txt (UTF-8)
#   - .bat 은 ASCII 실행기, 실제 작업과 한글 안내는 함께 넣은 PowerShell 스크립트가 처리
#     (bat 에 한글을 넣으면 CP949 코드페이지 문제로 깨짐)
# ════════════════════════════════════════════════════════════════════
MSI_TEMPLATE = r"""@echo off
setlocal
title Weekly Report Automator - Patch v@NEW@ (MSI install)
REM ASCII-only launcher. The patch logic and Korean messages are in the
REM PowerShell script embedded below (patches/msi_patch_template.ps1).

set "WD=%TEMP%\wra_patch_%RANDOM%%RANDOM%"
mkdir "%WD%" >nul 2>&1
if not exist "%WD%" goto err_tmp

> "%WD%\p.b64" (
@DATA@
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $d=$env:WD; $b=(Get-Content -Raw -LiteralPath (Join-Path $d 'p.b64')) -replace '\s',''; [IO.File]::WriteAllBytes((Join-Path $d 'p.zip'), [Convert]::FromBase64String($b)); Expand-Archive -LiteralPath (Join-Path $d 'p.zip') -DestinationPath (Join-Path $d 'x') -Force"
if errorlevel 1 goto err_decode

powershell -NoProfile -ExecutionPolicy Bypass -File "%WD%\x\patch.ps1" -PayloadDir "%WD%\x\src"
set "RC=%errorlevel%"
rmdir /s /q "%WD%" >nul 2>&1
pause >nul
exit /b %RC%

:err_tmp
echo.
echo [ERROR] Could not create a temporary folder: "%WD%"
goto fail

:err_decode
echo.
echo [ERROR] Patch data could not be decoded. Nothing was changed.
echo         Download the patch file again and retry.
goto fail

:fail
rmdir /s /q "%WD%" >nul 2>&1
echo.
echo Press any key to close this window...
pause >nul
exit /b 1
"""


def _app_version() -> str:
    import re
    text = (SRC / "config.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([0-9.]+)"', text, re.M)
    if not m:
        sys.exit("ERROR: APP_VERSION not found in src/config.py")
    return m.group(1)


def build_msi_patch() -> Path:
    here  = Path(__file__).resolve().parent
    ver   = _app_version()
    notes = here / f"notes_v{ver}.txt"
    if not notes.exists():
        sys.exit(f"ERROR: patch notes not found: {notes}")
    note_text = notes.read_text(encoding="utf-8").strip()
    if "'@" in note_text:
        sys.exit("ERROR: notes must not contain the sequence '@ (PowerShell here-string end)")
    ps1 = ((here / "msi_patch_template.ps1").read_text(encoding="utf-8")
           .replace("@VERSION@", ver).replace("@NOTES@", note_text))

    files = sorted(p for p in SRC.iterdir()
                   if p.is_file() and p.suffix in (".py", ".json"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Windows PowerShell 5.1 은 BOM 없는 UTF-8 을 ANSI(CP949)로 읽으므로 BOM 필수
        zf.writestr("patch.ps1", b"\xef\xbb\xbf" + ps1.replace("\n", "\r\n").encode("utf-8"))
        for p in files:
            zf.write(p, "src/" + p.name)
    b64  = base64.b64encode(buf.getvalue()).decode("ascii")
    data = "\n".join(f"echo {b64[i:i + 76]}" for i in range(0, len(b64), 76))

    text = MSI_TEMPLATE.replace("@NEW@", ver).replace("@DATA@", data)
    text.encode("ascii")
    out = here / f"apply_patch_v{ver}.bat"
    out.write_bytes(text.replace("\n", "\r\n").encode("ascii"))
    print(f"{out.name} (MSI install): {len(files)} files -> {', '.join(p.name for p in files)}")
    return out


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "msi":
        build_msi_patch()
    elif len(sys.argv) >= 3:
        build(sys.argv[1], sys.argv[2:])
    else:
        sys.exit(__doc__)
