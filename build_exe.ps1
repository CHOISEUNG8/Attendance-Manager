# EXE build script (PowerShell) - ASCII only for maximum compatibility
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "EXE Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# PyInstaller 설치 확인 및 설치
python -m pip show pyinstaller 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller is not installed. Installing..." -ForegroundColor Yellow
    python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "PyInstaller install failed." -ForegroundColor Red
        if (-not $env:NO_PAUSE) { Read-Host "Press Enter to exit" }
        exit 1
    }
    Write-Host "PyInstaller installed." -ForegroundColor Green
    Write-Host ""
}

# 빌드 디렉토리 정리
if (Test-Path "dist") {
    Write-Host "Cleaning dist/ ..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force "dist"
}
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
}
 # NOTE: do not delete *.spec automatically (we keep a maintained spec file: wb_attendance_v4_1.spec)

Write-Host ""
Write-Host "Starting EXE build..." -ForegroundColor Cyan
Write-Host ""

# favicon.ico 파일 확인
$iconOptions = ""
if (Test-Path "favicon.ico") {
    $iconOptions = "--icon=`"favicon.ico`" --add-data=`"favicon.ico;.`""
    Write-Host "Including favicon.ico" -ForegroundColor Green
} else {
    Write-Host "favicon.ico not found (building without icon)" -ForegroundColor Yellow
}

# 데이터베이스 파일 확인 및 포함
$dbOptions = ""

# The app stores DB under AppData\Roaming\근태관리프로그램 when running as EXE.
# To bundle "current" data into the EXE, sync that DB into project root before building.
try {
    $appDataDb = Join-Path $env:APPDATA "근태관리프로그램\\leave_attendance.db"
    if (Test-Path $appDataDb) {
        Write-Host "" 
        Write-Host "[DB Sync] Found AppData DB: $appDataDb" -ForegroundColor Green
        if (Test-Path "leave_attendance.db") {
            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            $bak = "leave_attendance.db.bak_$ts"
            Copy-Item -Force "leave_attendance.db" $bak
            Write-Host "[DB Sync] Backup created: $bak" -ForegroundColor Yellow
        }
        Copy-Item -Force $appDataDb "leave_attendance.db"
        Write-Host "[DB Sync] Copied AppData DB -> ./leave_attendance.db" -ForegroundColor Green
    } else {
        Write-Host "" 
        Write-Host "[DB Sync] AppData DB not found: $appDataDb" -ForegroundColor Yellow
    }
} catch {
    Write-Host "[DB Sync] Failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

if (Test-Path "leave_attendance.db") {
    $dbOptions = "--add-data=`"leave_attendance.db;.`""
    Write-Host "Including leave_attendance.db (with current data)" -ForegroundColor Green
} else {
    Write-Host "leave_attendance.db not found (will start empty)" -ForegroundColor Yellow
}

# PyInstaller 실행
$buildCommand = "python -m PyInstaller --name=`"WB_Attendance_Manager_v4_1`" --onefile --windowed $iconOptions $dbOptions --hidden-import=pandas --hidden-import=pandas._libs.tslibs.timedeltas --hidden-import=pandas._libs.tslibs.nattype --hidden-import=pandas._libs.tslibs.np_datetime --hidden-import=pandas._libs.skiplist --hidden-import=pandas._libs.algos --hidden-import=pandas._libs.window.aggregations --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --hidden-import=openpyxl.workbook.external_link.external --hidden-import=openpyxl.packaging.workbook --hidden-import=sqlite3 --hidden-import=PySide6.QtCore --hidden-import=PySide6.QtGui --hidden-import=PySide6.QtWidgets --collect-all=pandas --collect-all=openpyxl --collect-all=PySide6 --collect-submodules=pandas --collect-submodules=openpyxl `"Attendance and Leave Management Program.py`""

Invoke-Expression $buildCommand

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "BUILD FAILED" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    if (-not $env:NO_PAUSE) { Read-Host "Press Enter to exit" }
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "BUILD COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "EXE output: dist\\WB_Attendance_Manager_v4_1.exe" -ForegroundColor Cyan
Write-Host ""
Write-Host "Run the exe from the dist folder." -ForegroundColor Yellow
Write-Host ""
if (-not $env:NO_PAUSE) { Read-Host "Press Enter to exit" }

