"""
make_patch.py — 단일 .bat 패치 파일 생성기

사용법 (저장소 루트에서):
    python patches/make_patch.py 1.2 1.11 1.1

  첫 번째 인자 : 새 버전 (예: 1.2)
  나머지 인자  : 이 패치를 적용할 수 있는 이전 버전들 (exe/폴더명 V표기 변경용)
                 ※ "V1.1" 은 "V1.11" 의 앞부분이므로 긴 버전부터 자동 정렬해 처리

src/ 의 *.py, *.json 전체를 zip → base64 로 .bat 안에 넣는다.
.bat 실행 시 exe 옆 src/ 폴더를 통째로 덮어쓰고, exe 파일명과 폴더명의
버전 표기를 새 버전으로 바꾼다. 백업은 만들지 않고 임시파일은 항상 삭제한다.
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


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    build(sys.argv[1], sys.argv[2:])
