@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
echo ========================================
echo EXE 파일 빌드 스크립트
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

REM 데이터베이스 파일 확인 및 포함
set DB_OPTION=
REM EXE 실행 시 기본 DB는 AppData(%APPDATA%\근태관리프로그램)에 저장됩니다.
REM "현재 사용 중인 DB(데이터 포함)"로 빌드하려면, 빌드 전에 AppData DB를 프로젝트 루트로 동기화합니다.
set APPDATA_DB=%APPDATA%\근태관리프로그램\leave_attendance.db
if exist "%APPDATA_DB%" (
    echo.
    echo [DB 동기화] AppData DB 발견: "%APPDATA_DB%"
    if exist "leave_attendance.db" (
        echo [DB 동기화] 기존 로컬 DB 백업 생성: leave_attendance.db.bak
        copy /Y "leave_attendance.db" "leave_attendance.db.bak" >nul
    )
    echo [DB 동기화] AppData DB -> 로컬 leave_attendance.db 복사 중...
    copy /Y "%APPDATA_DB%" "leave_attendance.db" >nul
    echo [DB 동기화] 완료
) else (
    echo.
    echo [DB 동기화] AppData DB 없음: "%APPDATA_DB%"
)

if exist "leave_attendance.db" (
    set DB_OPTION=--add-data="leave_attendance.db;."
    echo leave_attendance.db 파일을 포함합니다. (현재 입력된 데이터 포함)
) else (
    echo leave_attendance.db 파일이 없습니다. 빈 데이터베이스로 시작합니다.
)

REM PyInstaller 실행
python -m PyInstaller ^
    --name="WB_Attendance_Manager_v4_1" ^
    --onefile ^
    --windowed ^
    %ICON_OPTION% ^
    %DB_OPTION% ^
    --hidden-import=pandas ^
    --hidden-import=pandas._libs.tslibs.timedeltas ^
    --hidden-import=pandas._libs.tslibs.nattype ^
    --hidden-import=pandas._libs.tslibs.np_datetime ^
    --hidden-import=pandas._libs.skiplist ^
    --hidden-import=pandas._libs.algos ^
    --hidden-import=pandas._libs.window.aggregations ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --hidden-import=openpyxl.workbook.external_link.external ^
    --hidden-import=openpyxl.packaging.workbook ^
    --hidden-import=sqlite3 ^
    --hidden-import=PySide6.QtCore ^
    --hidden-import=PySide6.QtGui ^
    --hidden-import=PySide6.QtWidgets ^
    --collect-all=pandas ^
    --collect-all=openpyxl ^
    --collect-all=PySide6 ^
    --collect-submodules=pandas ^
    --collect-submodules=openpyxl ^
    "Attendance and Leave Management Program.py"

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

