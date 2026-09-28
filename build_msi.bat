@echo off
setlocal
REM ==================================================================
REM  Weekly Report Automator - one-click build: onedir folder + MSI
REM
REM    build_msi.bat            build onedir folder and MSI installer
REM    build_msi.bat onedir     build onedir folder only (no MSI)
REM    add "nopause" to skip the final pause (for automation)
REM
REM  Output:
REM    dist\Weekly_Report_Automator_V<ver>\       onedir program folder
REM    dist\Weekly_Report_Automator_V<ver>.msi    installer (per-user, no admin)
REM
REM  Version comes from APP_VERSION in src\config.py (no manual edits here).
REM  MSI tool: WiX Toolset v3.14 (free, MS-RL license). Downloaded once into
REM  tools\wix314\ if missing. If download is blocked, get wix314-binaries.zip
REM  from https://github.com/wixtoolset/wix3/releases and unzip it there.
REM
REM  Keep this file ASCII-only (non-ASCII text breaks cmd code pages).
REM ==================================================================

set "MODE=msi"
set "NOPAUSE="
for %%A in (%*) do (
    if /i "%%~A"=="onedir"  set "MODE=onedir"
    if /i "%%~A"=="nopause" set "NOPAUSE=1"
)

REM pushd also works when this folder is on a UNC network path
pushd "%~dp0" || (echo ERROR: cannot enter build folder "%~dp0" & goto :fail_nopopd)

set "WIX_URL=https://github.com/wixtoolset/wix3/releases/download/wix3141rtm/wix314-binaries.zip"
set "WIX_SHA256=6ac824e1642d6f7277d0ed7ea09411a508f6116ba6fae0aa5f2c7daa2ff43d31"
set "WIX_DIR=%CD%\tools\wix314"

REM ------------------------------------------------------------------
echo.
echo [1/6] Checking Python and packages...
set "PY="
python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY (
    echo ERROR: Python not found. Install Python 3 and check "Add python.exe to PATH".
    goto :fail
)
%PY% -c "import PyInstaller, win32com.client, xlsxwriter, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing missing packages: pyinstaller pywin32 xlsxwriter openpyxl
    %PY% -m pip install --upgrade pyinstaller pywin32 xlsxwriter openpyxl
    %PY% -c "import PyInstaller, win32com.client, xlsxwriter, openpyxl" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: package install failed. Run manually:
        echo        %PY% -m pip install pyinstaller pywin32 xlsxwriter openpyxl
        goto :fail
    )
)

set "VER="
for /f "usebackq delims=" %%V in (`%PY% installer\msi_tool.py version`) do set "VER=%%V"
if not defined VER (
    echo ERROR: could not read APP_VERSION from src\config.py
    goto :fail
)
set "NAME=Weekly_Report_Automator_V%VER%"
echo       Python : %PY%
echo       Version: %VER%   Output name: %NAME%

REM ------------------------------------------------------------------
echo.
echo [2/6] Cleaning previous build output...
REM main.py loads dist\NAME\src\*.py (live patch) before the bundled code,
REM so leftovers from a previous build can hide the new code. Always wipe.
if exist "dist\%NAME%" rmdir /s /q "dist\%NAME%"
if exist "dist\%NAME%" (
    echo ERROR: could not delete dist\%NAME% - close the program and any
    echo        Explorer windows open inside it, then run again.
    goto :fail
)
if exist "dist\%NAME%.msi" del /q "dist\%NAME%.msi"
if exist "dist\%NAME%.msi" (
    echo ERROR: could not delete dist\%NAME%.msi - is it open or running?
    goto :fail
)
if exist build rmdir /s /q build
if exist "%NAME%.spec" del /q "%NAME%.spec"

REM ------------------------------------------------------------------
echo.
echo [3/6] Building onedir program with PyInstaller...
%PY% -m PyInstaller --noconfirm --clean --onedir --noconsole --paths src ^
    --icon=src/assets/Schedule_Ico.ico ^
    --add-data "src/assets;assets" --add-data "src/config.json;." ^
    --add-data "src/overrides.json;." --add-data "src;src" ^
    --hidden-import win32com --hidden-import win32com.client ^
    --hidden-import win32com.client.dynamic --hidden-import win32api ^
    --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32timezone ^
    --exclude-module tkcalendar --exclude-module babel --exclude-module numpy ^
    --exclude-module pandas --exclude-module matplotlib ^
    --name %NAME% main.py
if errorlevel 1 (
    echo ERROR: PyInstaller build failed. See messages above.
    goto :fail
)
if not exist "dist\%NAME%\%NAME%.exe" (
    echo ERROR: dist\%NAME%\%NAME%.exe was not created.
    goto :fail
)

echo Copying src files next to the exe (live patch folder)...
if not exist "dist\%NAME%\src" mkdir "dist\%NAME%\src"
xcopy /Y /Q "src\*.py" "dist\%NAME%\src\" >nul
if errorlevel 1 (
    echo ERROR: copying src\*.py failed.
    goto :fail
)
xcopy /Y /Q "src\*.json" "dist\%NAME%\src\" >nul
if errorlevel 1 (
    echo ERROR: copying src\*.json failed.
    goto :fail
)
if exist build rmdir /s /q build
if exist "%NAME%.spec" del /q "%NAME%.spec"
echo       OK: dist\%NAME%\

if /i "%MODE%"=="onedir" goto :done

REM ------------------------------------------------------------------
echo.
echo [4/6] Preparing WiX Toolset v3.14 (free)...
if exist "%WIX_DIR%\candle.exe" if exist "%WIX_DIR%\light.exe" goto :wix_ready
echo       Downloading WiX binaries (about 40 MB, first time only)...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" ^
  "$z=Join-Path $env:TEMP ('wix314_'+[guid]::NewGuid().ToString()+'.zip');" ^
  "try {" ^
  "  Invoke-WebRequest -UseBasicParsing -Uri $env:WIX_URL -OutFile $z;" ^
  "  $h=(Get-FileHash -Algorithm SHA256 -Path $z).Hash;" ^
  "  if ($h -ne $env:WIX_SHA256) { throw ('SHA256 mismatch: '+$h) };" ^
  "  New-Item -ItemType Directory -Force -Path $env:WIX_DIR | Out-Null;" ^
  "  Expand-Archive -Path $z -DestinationPath $env:WIX_DIR -Force;" ^
  "  Get-ChildItem -Path $env:WIX_DIR -Recurse | Unblock-File" ^
  "} finally { Remove-Item -Force -ErrorAction SilentlyContinue $z }"
if errorlevel 1 (
    echo ERROR: WiX download failed. Download manually:
    echo        %WIX_URL%
    echo        and unzip it into: "%WIX_DIR%"
    goto :fail
)
if not exist "%WIX_DIR%\candle.exe" (
    echo ERROR: candle.exe not found in "%WIX_DIR%"
    goto :fail
)
:wix_ready
echo       OK: %WIX_DIR%

REM ------------------------------------------------------------------
echo.
echo [5/6] Generating installer definition...
set "ARCH="
for /f "usebackq delims=" %%A in (`%PY% installer\msi_tool.py arch`) do set "ARCH=%%A"
if not defined ARCH set "ARCH=x64"
if not exist "build\msi" mkdir "build\msi"
%PY% installer\msi_tool.py wxs "dist\%NAME%" "build\msi\product.wxs"
if errorlevel 1 (
    echo ERROR: could not generate build\msi\product.wxs
    goto :fail
)

REM ------------------------------------------------------------------
echo.
echo [6/6] Building MSI with WiX (%ARCH%)...
"%WIX_DIR%\candle.exe" -nologo -arch %ARCH% -out "build\msi\product.wixobj" "build\msi\product.wxs"
if errorlevel 1 (
    echo ERROR: WiX candle failed. See messages above.
    goto :fail
)
set "LIGHT_OPTS=-nologo -spdb -sice:ICE61 -sice:ICE91"
"%WIX_DIR%\light.exe" %LIGHT_OPTS% -out "dist\%NAME%.msi" "build\msi\product.wixobj"
if errorlevel 1 (
    echo.
    echo WARNING: MSI validation step failed on this PC ^(often a Windows
    echo          Installer policy issue^). Retrying without ICE validation...
    if exist "dist\%NAME%.msi" del /q "dist\%NAME%.msi"
    "%WIX_DIR%\light.exe" %LIGHT_OPTS% -sval -out "dist\%NAME%.msi" "build\msi\product.wixobj"
    if errorlevel 1 (
        echo ERROR: WiX light failed. See messages above.
        goto :fail
    )
)
if not exist "dist\%NAME%.msi" (
    echo ERROR: dist\%NAME%.msi was not created.
    goto :fail
)
if exist build rmdir /s /q build

:done
echo.
echo ============================================================
echo  BUILD SUCCESS  (v%VER%)
echo    Program folder : %CD%\dist\%NAME%\
if /i not "%MODE%"=="onedir" echo    MSI installer  : %CD%\dist\%NAME%.msi
echo ============================================================
popd
if not defined NOPAUSE pause
exit /b 0

:fail
popd
:fail_nopopd
echo.
echo  BUILD FAILED
if not defined NOPAUSE pause
exit /b 1
