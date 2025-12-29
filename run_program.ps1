# PowerShell 실행 스크립트
# UTF-8 인코딩 설정
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 현재 스크립트가 있는 디렉토리로 이동
Set-Location -Path $PSScriptRoot

# Python 프로그램 실행
python "Attendance and Leave Management Program.py"

# 오류 발생 시 일시 정지
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "오류가 발생했습니다. 위의 메시지를 확인하세요." -ForegroundColor Red
    Read-Host "아무 키나 누르면 종료됩니다"
}

