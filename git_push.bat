@echo off
:: =====================================================
:: Weekly Report Automator - GitHub 초기 업로드
:: 실행 전: 프로젝트 소스 폴더에서 실행할 것
:: =====================================================

set REPO_URL=https://github.com/YOUR_USERNAME/Weekly_Report_Automator.git

echo.
echo [1] Git 초기화...
git init
git branch -m main

echo.
echo [2] .gitignore, README 복사 확인...
if not exist .gitignore echo .gitignore 없음 - 수동으로 추가하세요
if not exist README.md echo README.md 없음 - 수동으로 추가하세요

echo.
echo [3] 파일 스테이징...
git add .

echo.
echo [4] 첫 커밋...
git commit -m "feat: Weekly Report Automator Ver.1.1 초기 업로드"

echo.
echo [5] 원격 저장소 연결...
git remote add origin %REPO_URL%

echo.
echo [6] GitHub 푸시...
git push -u origin main

echo.
echo =====================================================
echo  완료! GitHub에서 확인하세요.
echo  %REPO_URL%
echo =====================================================
pause
