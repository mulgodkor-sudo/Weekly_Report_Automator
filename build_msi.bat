@echo off
REM ==========================================================================
REM  build_msi.bat - Weekly Report Automator one-click build
REM
REM    build_msi.bat            onedir program folder + MSI installer
REM    build_msi.bat onedir     onedir program folder only (build.bat uses this)
REM    add "nopause" to skip the final pause
REM
REM  Output:
REM    dist\Weekly_Report_Automator_V<ver>\       onedir program folder
REM    dist\Weekly_Report_Automator_V<ver>.msi    installer (Korean wizard)
REM
REM  Installer layout follows MHS_Searcher (Equipment Item DB):
REM    WiX v5.0.2 (free, MS-RL; v6+ needs the paid OSMF EULA), Korean wizard
REM    (installer\wix\ui.ko-KR.wxl), branding bitmaps, install folder
REM    C:\Program Files\Autotools_Mechanical\Weekly Report Automator\
REM    WiX is installed with "dotnet tool" (needs the .NET SDK) - no zip
REM    download into the project folder.
REM
REM  Version comes ONLY from APP_VERSION in src\config.py.
REM
REM  Keep this file PURE ASCII. Korean Windows reads .bat files as CP949 and
REM  non-ASCII bytes break parsing. Korean text lives in the .wxs/.wxl/.py
REM  files, which are read as UTF-8.
REM ==========================================================================
setlocal EnableExtensions
chcp 65001 >nul
title Build Weekly Report Automator

set "MODE=msi"
set "NOPAUSE="
for %%A in (%*) do (
    if /i "%%~A"=="onedir"  set "MODE=onedir"
    if /i "%%~A"=="nopause" set "NOPAUSE=1"
)

REM pushd also works when the project is on a UNC network path
pushd "%~dp0" || goto fail_nopopd
set "ROOT=%CD%"
if not defined WIXVER set "WIXVER=5.0.2"
set "TMPF=%TEMP%\wra_build_%RANDOM%%RANDOM%.txt"
echo Project root: %ROOT%

REM --- 1) Python (prefer project .venv) --------------------------------------
echo.
echo [1/7] Python and packages ...
set "PY=%ROOT%\.venv\Scripts\python.exe"
if exist "%PY%" goto py_ready
set "PY="
for /f "delims=" %%p in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if defined PY goto py_ready
for /f "delims=" %%p in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if defined PY goto py_ready
goto err_py
:py_ready
"%PY%" -c "import sys;print('  python',sys.version.split()[0],sys.executable)"
if errorlevel 1 goto err_py

"%PY%" -c "import PyInstaller, win32com.client, xlsxwriter, openpyxl" >nul 2>&1
if not errorlevel 1 goto pkgs_ready
echo   Installing missing packages: pyinstaller pywin32 xlsxwriter openpyxl
"%PY%" -m pip install --upgrade pyinstaller pywin32 xlsxwriter openpyxl
"%PY%" -c "import PyInstaller, win32com.client, xlsxwriter, openpyxl" >nul 2>&1
if errorlevel 1 goto err_pip
:pkgs_ready

REM --- 2) version ------------------------------------------------------------
set "VER="
set "MSIVER="
set "ARCH="
"%PY%" installer\msi_tool.py version > "%TMPF%"
if errorlevel 1 goto err_ver
set /p VER=<"%TMPF%"
"%PY%" installer\msi_tool.py msiver > "%TMPF%"
if errorlevel 1 goto err_ver
set /p MSIVER=<"%TMPF%"
"%PY%" installer\msi_tool.py arch > "%TMPF%"
set /p ARCH=<"%TMPF%"
del /q "%TMPF%" >nul 2>&1
if not defined VER goto err_ver
if not defined MSIVER goto err_ver
if not defined ARCH set "ARCH=x64"
set "NAME=Weekly_Report_Automator_V%VER%"
set "EXE=dist\%NAME%\%NAME%.exe"
echo   Version : %VER%   (MSI ProductVersion %MSIVER%, %ARCH%)
echo   Output  : dist\%NAME%\
"%PY%" installer\msi_tool.py pathcheck

REM --- 3) clean previous output ----------------------------------------------
REM  main.py loads dist\NAME\src\*.py (live patch) before the bundled code,
REM  so leftovers from a previous build can hide the new code. Always wipe.
echo.
echo [2/7] Cleaning previous build output ...
if exist "dist\%NAME%" rmdir /s /q "dist\%NAME%"
if exist "dist\%NAME%" goto err_locked_dir
if exist "dist\%NAME%.msi" del /q "dist\%NAME%.msi"
if exist "dist\%NAME%.msi" goto err_locked_msi
if exist "dist\%NAME%.wixpdb" del /q "dist\%NAME%.wixpdb"
if exist build rmdir /s /q build
if exist "%NAME%.spec" del /q "%NAME%.spec"

REM --- 4) PyInstaller onedir -------------------------------------------------
echo.
echo [3/7] PyInstaller onedir build ...
"%PY%" -m PyInstaller --noconfirm --clean --onedir --noconsole --paths src ^
    --icon=src/assets/Schedule_Ico.ico ^
    --add-data "src/assets;assets" --add-data "src/config.json;." ^
    --add-data "src/overrides.json;." --add-data "src;src" ^
    --hidden-import win32com --hidden-import win32com.client ^
    --hidden-import win32com.client.dynamic --hidden-import win32api ^
    --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32timezone ^
    --hidden-import tkinter.font --hidden-import tkinter.simpledialog ^
    --exclude-module tkcalendar --exclude-module babel --exclude-module numpy ^
    --exclude-module pandas --exclude-module matplotlib ^
    --name %NAME% main.py
if errorlevel 1 goto err_build
if not exist "%EXE%" goto err_no_exe

REM  exe destroyed by antivirus leaves a stub of a few bytes/KB
set "EXESIZE=0"
for %%A in ("%EXE%") do set "EXESIZE=%%~zA"
echo   exe size: %EXESIZE% bytes
if %EXESIZE% LSS 100000 goto err_tiny

echo   Copying src\*.py, src\*.json next to the exe (live patch folder) ...
if not exist "dist\%NAME%\src" mkdir "dist\%NAME%\src"
xcopy /Y /Q "src\*.py" "dist\%NAME%\src\" >nul
if errorlevel 1 goto err_copy
xcopy /Y /Q "src\*.json" "dist\%NAME%\src\" >nul
if errorlevel 1 goto err_copy
if exist build rmdir /s /q build
if exist "%NAME%.spec" del /q "%NAME%.spec"
echo   OK: dist\%NAME%\

if /i "%MODE%"=="onedir" goto done

REM --- 5) WiX v5 CLI + extensions --------------------------------------------
echo.
echo [4/7] WiX Toolset v%WIXVER% (dotnet tool) ...
set "PATH=%PATH%;%USERPROFILE%\.dotnet\tools"
where dotnet >nul 2>&1
if errorlevel 1 goto err_dotnet
set "WIXCUR="
wix --version > "%TMPF%" 2>nul
set /p WIXCUR=<"%TMPF%"
del /q "%TMPF%" >nul 2>&1
echo %WIXCUR%| findstr /b /c:"%WIXVER%" >nul
if not errorlevel 1 goto wix_ready
echo   Installing WiX %WIXVER% (current: %WIXCUR%) ...
dotnet tool uninstall --global wix >nul 2>&1
dotnet tool install --global wix --version %WIXVER%
if errorlevel 1 goto err_wix
:wix_ready
echo   WiX OK
wix extension add -g WixToolset.UI.wixext/%WIXVER%
if errorlevel 1 goto err_wix
wix extension add -g WixToolset.Util.wixext/%WIXVER%
if errorlevel 1 goto err_wix

REM --- 6) installer inputs ---------------------------------------------------
echo.
echo [5/7] Checking installer inputs ...
if not exist "installer\wix\Product.wxs"         goto err_input
if not exist "installer\wix\Variables.wxi"       goto err_input
if not exist "installer\wix\ui.ko-KR.wxl"        goto err_input
if not exist "installer\wix\WixUIDialogBmp.bmp"  goto err_input
if not exist "installer\wix\WixUIBannerBmp.bmp"  goto err_input
if not exist "installer\msi_installed.txt"       goto err_input
if not exist "src\assets\Schedule_Ico.ico"       goto err_input
if not exist "build\msi" mkdir "build\msi"
echo [6/7] Writing install notice (License.rtf) with version %VER% ...
"%PY%" installer\msi_tool.py license "%ROOT%\build\msi\License.rtf"
if errorlevel 1 goto err_license

REM --- 7) build MSI ----------------------------------------------------------
echo.
echo [7/7] Building MSI (WiX v%WIXVER%, Korean install wizard) ...
wix build "installer\wix\Product.wxs" ^
    -ext WixToolset.UI.wixext/%WIXVER% ^
    -ext WixToolset.Util.wixext/%WIXVER% ^
    -culture ko-KR ^
    -loc "installer\wix\ui.ko-KR.wxl" ^
    -d "PublishDir=%ROOT%\dist\%NAME%" ^
    -d "ProjectRoot=%ROOT%" ^
    -d "ExeName=%NAME%.exe" ^
    -d "ProductVersion=%MSIVER%" ^
    -d "DisplayVersion=%VER%" ^
    -d "LicenseRtf=%ROOT%\build\msi\License.rtf" ^
    -arch %ARCH% ^
    -sw1077 ^
    -pdbtype none ^
    -intermediatefolder "%ROOT%\build\msi\obj" ^
    -o "dist\%NAME%.msi"
if errorlevel 1 goto err_msi
if not exist "dist\%NAME%.msi" goto err_msi
if exist build rmdir /s /q build
set "MSISIZE=0"
for %%A in ("dist\%NAME%.msi") do set "MSISIZE=%%~zA"

:done
echo.
echo ============================================================
echo   BUILD SUCCESS  (v%VER%)
echo     one-dir : dist\%NAME%\%NAME%.exe
if /i "%MODE%"=="msi" echo     MSI     : dist\%NAME%.msi  (%MSISIZE% bytes)
if /i "%MODE%"=="msi" echo     Install : C:\Program Files\Autotools_Mechanical\Weekly Report Automator\
echo ============================================================
popd
if not defined NOPAUSE pause
exit /b 0

REM --- error handlers --------------------------------------------------------
:err_py
echo.
echo [ERROR] No working Python found.
echo         Install Python 3 and check "Add python.exe to PATH", or create .venv\ here.
goto fail

:err_pip
echo.
echo [ERROR] Package install failed. Run manually to see the real error:
echo           "%PY%" -m pip install pyinstaller pywin32 xlsxwriter openpyxl
goto fail

:err_ver
echo.
echo [ERROR] Could not read APP_VERSION from src\config.py (e.g. APP_VERSION = "1.2").
goto fail

:err_locked_dir
echo.
echo [ERROR] Could not delete dist\%NAME%\ - close the program and any Explorer
echo         windows open inside it, then run again.
goto fail

:err_locked_msi
echo.
echo [ERROR] Could not delete dist\%NAME%.msi - is it open or running?
goto fail

:err_build
echo.
echo [ERROR] PyInstaller build failed. See the messages above.
goto fail

:err_no_exe
echo.
echo [ERROR] %EXE% was not created.
echo         Antivirus may have removed it - check Windows Security protection history.
goto fail

:err_tiny
echo.
echo [ERROR] %EXE% is only %EXESIZE% bytes - destroyed by antivirus.
echo         Windows Security ^> Protection history: restore + allow, add the project
echo         folder to exclusions (or move it to C:\WRA), then run again.
goto fail

:err_copy
echo.
echo [ERROR] Copying src files into dist\%NAME%\src failed.
goto fail

:err_dotnet
echo.
echo [ERROR] .NET SDK (dotnet) not found - needed to install WiX v%WIXVER%.
echo         Install ".NET 8 SDK" from https://dotnet.microsoft.com/download
echo         (or: winget install Microsoft.DotNet.SDK.8), open a new window, run again.
echo         The onedir folder is already built: dist\%NAME%\
goto fail

:err_wix
echo.
echo [ERROR] WiX CLI or an extension is not available. Run manually:
echo           dotnet tool uninstall --global wix
echo           dotnet tool install --global wix --version %WIXVER%
echo           wix extension add -g WixToolset.UI.wixext/%WIXVER%
echo           wix extension add -g WixToolset.Util.wixext/%WIXVER%
echo         If WIX0094 (Wix4UtilCA) appears, remove old cached versions:
echo           wix extension list -g
echo           wix extension remove -g WixToolset.Util.wixext/OLD_VERSION
goto fail

:err_input
echo.
echo [ERROR] An installer input file is missing (installer\wix\*, installer\msi_installed.txt,
echo         src\assets\Schedule_Ico.ico). Get the latest files from GitHub (git pull).
goto fail

:err_license
echo.
echo [ERROR] Could not write build\msi\License.rtf
goto fail

:err_msi
echo.
echo [ERROR] WiX MSI build failed. Read the messages above.
goto fail

:fail
popd
:fail_nopopd
del /q "%TMPF%" >nul 2>&1
echo.
echo ============================================================
echo   BUILD FAILED. Read the error above.
echo ============================================================
if not defined NOPAUSE pause
exit /b 1
