# 필요한 패키지 설치 스크립트 (PowerShell)
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "필요한 패키지 설치 스크립트" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# pip 업그레이드
Write-Host "pip 업그레이드 중..." -ForegroundColor Yellow
python -m pip install --upgrade pip

Write-Host ""
Write-Host "필요한 패키지 설치 중..." -ForegroundColor Yellow
Write-Host ""

# requirements.txt에서 패키지 설치
python -m pip install -r requirements.txt

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "패키지 설치 중 오류가 발생했습니다!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "수동 설치를 시도하시겠습니까?" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "아무 키나 누르면 종료됩니다"
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "패키지 설치 완료!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "설치된 패키지:" -ForegroundColor Cyan
python -m pip list | Select-String -Pattern "pandas|openpyxl|PySide6|pyinstaller"
Write-Host ""
Read-Host "아무 키나 누르면 종료됩니다"

