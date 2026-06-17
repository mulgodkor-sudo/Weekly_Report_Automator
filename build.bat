@echo off
set NAME=Weekly_Report_Automator_V1.1

REM ── 이전 빌드 잔재 제거 ──────────────────────────────────────
REM main.py의 _setup_src()는 dist\%NAME%\src\*.py (라이브 패치)를
REM 번들된 코드보다 우선해서 로드한다. 이 폴더에 구버전 파일이 남아있으면
REM 재빌드해도 새 코드가 적용되지 않은 것처럼 보이므로, 매 빌드마다 통째로 삭제한다.
if exist "dist\%NAME%" rmdir /s /q "dist\%NAME%"

pyinstaller --onedir --noconsole --paths src --icon=src/assets/Schedule_Ico.ico --add-data "src/assets;assets" --add-data "src/config.json;." --add-data "src/overrides.json;." --add-data "src;src" --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.client.dynamic --hidden-import win32api --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32timezone --exclude-module tkcalendar --exclude-module babel --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --name %NAME% main.py

if errorlevel 1 (
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo Copying src files...
if not exist "dist\%NAME%\src" mkdir "dist\%NAME%\src"
xcopy /Y src\*.py "dist\%NAME%\src\"
if errorlevel 1 (
    echo SRC PY COPY FAILED - 작업 디렉터리가 저장소 루트인지 확인하세요.
    pause
    exit /b 1
)
xcopy /Y src\*.json "dist\%NAME%\src\"
if errorlevel 1 (
    echo SRC JSON COPY FAILED - 작업 디렉터리가 저장소 루트인지 확인하세요.
    pause
    exit /b 1
)

echo Cleaning up...
if exist build rmdir /s /q build
if exist %NAME%.spec del %NAME%.spec

echo BUILD SUCCESS. Output: dist\%NAME%\
pause
