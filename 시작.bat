@echo off
chcp 65001 > nul
title HR 근태정리 자동화 도구

echo ========================================
echo  HR 근태정리 자동화 도구 시작
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] Python 설치 확인 중...
python --version > nul 2>&1
if errorlevel 1 (
    echo.
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    https://python.org 에서 Python 3.10 이상 설치 후 다시 실행하세요.
    echo.
    pause
    exit /b
)

echo [2/3] 필요한 라이브러리 설치 확인 중... (최초 1회만 소요)
pip install -q -r requirements.txt

echo [3/3] 앱 실행 중... 잠시 후 브라우저가 자동으로 열립니다.
echo.
echo 종료하려면 이 창에서 Ctrl+C 입력 또는 창 닫기
echo ========================================
echo.

streamlit run app.py --server.headless true --browser.gatherUsageStats false

pause
