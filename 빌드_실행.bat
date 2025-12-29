@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo 근태관리 프로그램 EXE 빌드
echo ========================================
echo.
echo [1] Spec 파일로 빌드 (권장)
echo [2] 기본 빌드
echo [3] 빌드 취소
echo.
set /p choice="선택하세요 (1-3): "

if "%choice%"=="1" (
    call build_exe_spec.bat
) else if "%choice%"=="2" (
    call build_exe.bat
) else (
    echo 빌드를 취소했습니다.
    pause
    exit /b 0
)

