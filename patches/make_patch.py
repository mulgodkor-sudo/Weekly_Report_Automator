"""
make_patch.py — 단일 .bat 패치 파일 생성기

사용법 (저장소 루트에서):
  python patches/make_patch.py 1.2 1.11 1.1     (onedir 폴더용, ~v1.2)
  python patches/make_patch.py msi               (MSI 설치본용 → WRA_Patch_v<버전>.bat, v1.23~)

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
# MSI 설치본용 패치 (v1.23~)
#   python patches/make_patch.py msi
#   → patches/WRA_Patch_v<APP_VERSION>.bat  (파일명 규칙 — 이 파일만 ECM WRA_Patch 폴더에 올리면
#     프로그램이 시작할 때 찾아서 업데이트를 묻는다, src/updater.py)
#
#   회사 보안 정책(소프트웨어 제한 정책)이 %TEMP% 의 스크립트 파일(.ps1/.bat) 실행을 막으므로
#   - 수동 실행: .bat 이 직접 처리 (xcopy / fc). .ps1 파일은 만들지도 실행하지도 않음.
#     PowerShell 은 한 줄 명령(-Command)으로 패치 데이터 풀기와 한글 안내 출력에만 사용
#     (bat 에 한글을 직접 넣으면 CP949 코드페이지 문제로 깨지므로 한글은 base64 로 넣어 출력)
#   - 자동 업데이트: 설치된 프로그램이 이 .bat 의 패치 데이터를 직접 읽어 src\ 교체 (스크립트 실행 없음)
#   - 버전은 src/config.py APP_VERSION, 패치 내용은 patches/notes_v<버전>.txt (UTF-8)
# ════════════════════════════════════════════════════════════════════
MSI_TEMPLATE = r"""@echo off
setlocal EnableExtensions
title Weekly Report Automator - Patch v@NEW@ (MSI install)
REM ASCII only. Korean messages are base64 UTF-8 strings printed by one-line
REM PowerShell commands. No script file (.ps1) is created or executed.
REM The program's auto update reads the payload between WRA_PAYLOAD markers.
REM WRA_NOTES_B64=@NOTES_B64@

set "PV=@NEW@"
set "SUB=Autotools_Mechanical\Weekly Report Automator"
set "WD=%TEMP%\wra_patch_%RANDOM%%RANDOM%"

call :say "@M_TITLE@" Cyan
echo.

REM --- 1) find the MSI install folder ------------------------------------
set "INST="
if exist "%ProgramFiles%\%SUB%\src\" set "INST=%ProgramFiles%\%SUB%"
if not defined INST if defined ProgramW6432 if exist "%ProgramW6432%\%SUB%\src\" set "INST=%ProgramW6432%\%SUB%"
if not defined INST if exist "%ProgramFiles(x86)%\%SUB%\src\" set "INST=%ProgramFiles(x86)%\%SUB%"
if not defined INST goto err_notfound
call :sayn "@M_PATH@" Gray
echo  %INST%

REM --- 2) installed version (src\config.py APP_VERSION) --------------------
set "IV="
for /f "tokens=2 delims==" %%A in ('findstr /b /c:"APP_VERSION" "%INST%\src\config.py" 2^>nul') do set "IV=%%A"
if defined IV set "IV=%IV: =%"
if defined IV set "IV=%IV:"=%"
call :sayn "@M_VER@" Gray
if defined IV (echo  v%IV% -^> v%PV%) else (echo  ? -^> v%PV%)
if not defined IV goto ver_ok
call :vkey "%IV%" IK
call :vkey "%PV%" PK
if %IK% GTR %PK% goto err_newer
:ver_ok
echo.

REM --- 3) unpack the patch data -----------------------------------------
mkdir "%WD%" >nul 2>&1
if not exist "%WD%" goto err_decode
REM WRA_PAYLOAD_BEGIN
> "%WD%\p.b64" (
@DATA@
)
REM WRA_PAYLOAD_END
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $d=$env:WD; $b=(Get-Content -Raw -LiteralPath (Join-Path $d 'p.b64')) -replace '\s',''; [IO.File]::WriteAllBytes((Join-Path $d 'p.zip'), [Convert]::FromBase64String($b)); Expand-Archive -LiteralPath (Join-Path $d 'p.zip') -DestinationPath (Join-Path $d 'x') -Force"
if errorlevel 1 goto err_decode
if not exist "%WD%\x\src\config.py" goto err_decode

REM --- 4) replace src\ files and verify --------------------------------
xcopy /Y /Q /R "%WD%\x\src\*.*" "%INST%\src\" >nul 2>&1
if errorlevel 1 goto err_denied
set "BAD="
set "N=0"
for %%F in ("%WD%\x\src\*.*") do (
    set /a N+=1
    fc /b "%%~fF" "%INST%\src\%%~nxF" >nul 2>&1 || set "BAD=1"
)
if defined BAD goto err_verify
if exist "%INST%\src\__pycache__" rmdir /s /q "%INST%\src\__pycache__" >nul 2>&1
rmdir /s /q "%WD%" >nul 2>&1

call :say "@M_OK@" Green
echo.
call :say "@M_NOTES@" Gray
echo.
call :say "@M_AFTER@" Yellow
goto done

REM --- errors -------------------------------------------------------------
:err_notfound
call :say "@M_NOTFOUND@" Red
echo   %ProgramFiles%\%SUB%
if defined ProgramW6432 echo   %ProgramW6432%\%SUB%
echo   %ProgramFiles(x86)%\%SUB%
call :say "@M_NOTFOUND2@" Red
goto done

:err_newer
call :say "@M_NEWER@" Red
goto done

:err_decode
call :say "@M_DECODE@" Red
goto cleanup_done

:err_denied
call :say "@M_DENIED@" Red
echo   %INST%\src
goto cleanup_done

:err_verify
call :say "@M_VERIFY@" Red
goto cleanup_done

:cleanup_done
rmdir /s /q "%WD%" >nul 2>&1

:done
echo.
call :say "@M_CLOSE@" Gray
pause >nul
exit /b 0

REM --- helpers ------------------------------------------------------------
REM :say "<base64 utf-8 text>" <color>   print Korean text (no script file)
:say
powershell -NoProfile -Command "Write-Host ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('%~1'))) -ForegroundColor %~2"
exit /b 0

:sayn
powershell -NoProfile -Command "Write-Host ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('%~1'))) -ForegroundColor %~2 -NoNewline"
exit /b 0

REM :vkey "1.23" VAR   ->  VAR = 123  (1.1 < 1.11 < 1.2 < 1.22 < 1.3)
:vkey
set "_ma=0"
set "_fr="
for /f "tokens=1,2 delims=." %%a in ("%~1") do (
    set "_ma=%%a"
    set "_fr=%%b"
)
set "_fr=%_fr%00"
set "_fr=%_fr:~0,2%"
set /a "%~2=_ma*100 + 1%_fr% - 100"
exit /b 0
"""


def _app_version() -> str:
    import re
    text = (SRC / "config.py").read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([0-9.]+)"', text, re.M)
    if not m:
        sys.exit("ERROR: APP_VERSION not found in src/config.py")
    return m.group(1)


def _b64(text: str) -> str:
    return base64.b64encode(text.encode("utf-8")).decode("ascii")


def build_msi_patch() -> Path:
    here  = Path(__file__).resolve().parent
    ver   = _app_version()
    notes = here / f"notes_v{ver}.txt"
    if not notes.exists():
        sys.exit(f"ERROR: patch notes not found: {notes}")
    note_text = notes.read_text(encoding="utf-8").strip()

    files = sorted(p for p in SRC.iterdir()
                   if p.is_file() and p.suffix in (".py", ".json"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in files:
            zf.write(p, "src/" + p.name)
    b64  = base64.b64encode(buf.getvalue()).decode("ascii")
    data = "\n".join(f"echo {b64[i:i + 76]}" for i in range(0, len(b64), 76))

    line = "=" * 60
    msgs = {
        "M_TITLE":     f"{line}\n Weekly Report Automator 패치 (v{ver}, MSI 설치본용)\n{line}",
        "M_PATH":      "설치 경로 :",
        "M_VER":       "버전      :",
        "M_OK":        f"{line}\n 패치가 성공적으로 완료되었습니다.  "
                       f"(v{ver}, 파일 {len(files)}개 교체·검증 완료)\n{line}",
        "M_NOTES":     "패치 내용\n" + "\n".join("  " + l for l in note_text.splitlines()),
        "M_AFTER":     "※ 프로그램을 다시 시작하면 적용됩니다. (프로그램 화면에 새 버전이 표시됩니다)\n"
                       "※ 이후 새 버전은 프로그램을 켤 때 자동으로 확인합니다.\n"
                       "※ 시작 메뉴 / 앱 목록의 버전 표시는 다음에 MSI로 설치할 때 바뀝니다.",
        "M_NOTFOUND":  "[오류] 다음 경로에서 Weekly Report Automator 설치 파일을 찾을 수 없습니다:",
        "M_NOTFOUND2": "        MSI 설치 파일(Weekly_Report_Automator_Vx.x.msi)로 먼저 설치한 뒤 다시 실행하세요.",
        "M_NEWER":     f"[오류] 이미 더 높은 버전이 설치되어 있습니다. 이 패치(v{ver})는 적용하지 않았습니다.",
        "M_DECODE":    "[오류] 패치 데이터를 풀지 못했습니다. 아무것도 바뀌지 않았습니다.\n"
                       "        패치 파일을 다시 받아 실행하세요.",
        "M_DENIED":    "[오류] 다음 폴더에 파일을 쓸 수 없습니다 (권한 없음 또는 파일 사용 중).\n"
                       "        프로그램을 닫고, 패치 파일을 마우스 오른쪽 클릭 → [관리자 권한으로 실행] 하세요:",
        "M_VERIFY":    "[오류] 일부 파일이 제대로 교체되지 않았습니다. 프로그램을 닫고 다시 실행하세요.",
        "M_CLOSE":     "창을 닫으려면 아무 키나 누르세요...",
    }
    text = (MSI_TEMPLATE.replace("@NEW@", ver).replace("@DATA@", data)
            .replace("@NOTES_B64@", _b64(note_text)))
    for k, v in msgs.items():
        text = text.replace(f"@{k}@", _b64(v))
    text.encode("ascii")
    if len(max(text.splitlines(), key=len)) > 8000:
        sys.exit("ERROR: a line is too long for cmd.exe (8191 chars)")
    # 파일명 규칙: WRA_Patch_v<버전>.bat — 프로그램 자동 업데이트가 이 이름으로 찾는다
    out = here / f"WRA_Patch_v{ver}.bat"
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
