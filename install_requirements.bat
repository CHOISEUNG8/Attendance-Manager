@echo off
chcp 65001 >nul
echo ========================================
echo 필요한 패키지 설치 스크립트
echo ========================================
echo.

REM pip 업그레이드
echo pip 업그레이드 중...
python -m pip install --upgrade pip

echo.
echo 필요한 패키지 설치 중...
echo.

REM requirements.txt에서 패키지 설치
python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ========================================
    echo 패키지 설치 중 오류가 발생했습니다!
    echo ========================================
    echo.
    echo 수동 설치를 시도하시겠습니까?
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo 패키지 설치 완료!
echo ========================================
echo.
echo 설치된 패키지:
python -m pip list | findstr /i "pandas openpyxl PySide6 pyinstaller"
echo.
pause

