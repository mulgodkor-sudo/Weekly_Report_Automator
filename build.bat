@echo off
REM Build the onedir program folder only (no MSI).
REM All build logic lives in build_msi.bat so the two never drift apart.
REM Keep this file ASCII-only.
call "%~dp0build_msi.bat" onedir %*
exit /b %errorlevel%
