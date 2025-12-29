@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
echo ========================================
echo EXE 파일 빌드 스크립트 (Spec 파일 사용)
echo ========================================
echo.
echo 작업 디렉토리: %CD%
echo.

REM PyInstaller 설치 확인 및 설치
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다. 설치 중...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo PyInstaller 설치 실패!
        pause
        exit /b 1
    )
    echo PyInstaller 설치 완료!
    echo.
)

REM 필요한 패키지 설치 확인
echo 필요한 패키지 확인 중...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo 패키지 설치 중 오류 발생!
    pause
    exit /b 1
)
echo 패키지 설치 완료!
echo.

REM 빌드 디렉토리 정리
echo 기존 빌드 파일 정리 중...
if exist "dist" (
    rmdir /s /q dist 2>nul
)
if exist "build" (
    rmdir /s /q build 2>nul
)
echo 정리 완료!
echo.

REM favicon.ico 파일 확인
if not exist "favicon.ico" (
    echo 경고: favicon.ico 파일이 없습니다. 아이콘 없이 빌드합니다.
    echo.
)

REM 데이터베이스 파일 확인
if not exist "leave_attendance.db" (
    echo 경고: leave_attendance.db 파일이 없습니다. 빈 데이터베이스로 시작합니다.
    echo.
)

echo.
echo ========================================
echo EXE 파일 빌드 시작...
echo ========================================
echo.
echo 이 작업은 몇 분 정도 소요될 수 있습니다...
echo.

REM Spec 파일을 사용하여 빌드
python -m PyInstaller wb_attendance_v4_1.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo ========================================
    echo 빌드 실패!
    echo ========================================
    echo.
    echo 오류가 발생했습니다. 위의 오류 메시지를 확인하세요.
    echo.
    if not defined NO_PAUSE pause
    exit /b 1
)

echo.
echo ========================================
echo 빌드 완료!
echo ========================================
echo.
echo EXE 파일 위치: dist\WB_Attendance_Manager_v4_1.exe
echo.
echo 파일 크기 확인:
dir "dist\WB_Attendance_Manager_v4_1.exe" | findstr "WB_Attendance_Manager_v4_1.exe"
echo.
echo 빌드된 파일을 실행하려면 dist 폴더의 exe 파일을 실행하세요.
echo.
echo 배포 시 주의사항:
echo - EXE 파일과 같은 폴더에 leave_attendance.db 파일이 생성됩니다.
echo - 데이터를 보존하려면 이 파일을 함께 백업하세요.
echo.
if not defined NO_PAUSE pause

