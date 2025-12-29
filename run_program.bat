@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 출퇴근 관리 프로그램 실행 중...
echo ========================================
echo.
python "Attendance and Leave Management Program.py"
if errorlevel 1 (
    echo.
    echo ========================================
    echo 오류가 발생했습니다. 위의 메시지를 확인하세요.
    echo ========================================
    pause
)

