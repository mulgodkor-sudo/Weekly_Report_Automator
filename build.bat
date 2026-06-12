@echo off
set NAME=Weekly_Report_Automator_V1.1

pyinstaller --onedir --noconsole --paths src --icon=src/assets/Schedule_Ico.ico --add-data "src/assets;assets" --add-data "src/config.json;." --add-data "src/overrides.json;." --add-data "src;src" --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.client.dynamic --hidden-import win32api --hidden-import pywintypes --hidden-import pythoncom --hidden-import win32timezone --exclude-module tkcalendar --exclude-module babel --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --name %NAME% main.py

if errorlevel 1 (
    echo BUILD FAILED.
    pause
    exit /b 1
)

echo Copying src files...
if not exist "dist\%NAME%\src" mkdir "dist\%NAME%\src"
xcopy /Y src\*.py "dist\%NAME%\src\"
xcopy /Y src\*.json "dist\%NAME%\src\"

echo Cleaning up...
if exist build rmdir /s /q build
if exist %NAME%.spec del %NAME%.spec

echo BUILD SUCCESS. Output: dist\%NAME%\
pause
