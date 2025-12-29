@echo off
chcp 65001 >nul
echo ========================================
echo EXE 파일 빌드 스크립트 (고급 버전)
echo ========================================
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

REM 빌드 디렉토리 정리
if exist "dist" (
    echo 기존 빌드 파일 삭제 중...
    rmdir /s /q dist
)
if exist "build" (
    rmdir /s /q build
)
REM NOTE: do not delete *.spec automatically (we keep a maintained spec file: wb_attendance_v4_1.spec)

echo.
echo EXE 파일 빌드 시작...
echo.

REM favicon.ico 파일 확인
set ICON_OPTION=
if exist "favicon.ico" (
    set ICON_OPTION=--icon="favicon.ico" --add-data="favicon.ico;."
    echo favicon.ico 파일을 포함합니다.
) else (
    echo favicon.ico 파일이 없습니다. 아이콘 없이 빌드합니다.
)

echo.
echo PyInstaller spec 파일 생성 중...
python -m PyInstaller --name="WB_Attendance_Manager_v4_1" --onefile --windowed %ICON_OPTION% --hidden-import=pandas --hidden-import=openpyxl --hidden-import=sqlite3 --hidden-import=PySide6.QtCore --hidden-import=PySide6.QtGui --hidden-import=PySide6.QtWidgets --collect-all=pandas --collect-all=openpyxl --collect-all=PySide6 "Attendance and Leave Management Program.py" --noconfirm

if errorlevel 1 (
    echo.
    echo ========================================
    echo 빌드 실패!
    echo ========================================
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
echo 빌드된 파일을 실행하려면 dist 폴더의 exe 파일을 실행하세요.
echo.
if not defined NO_PAUSE pause

