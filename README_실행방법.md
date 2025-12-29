# 프로그램 실행 방법

## 방법 1: 배치 파일 사용 (권장)
`run_program.bat` 파일을 더블클릭하여 실행

## 방법 2: PowerShell에서 실행
```powershell
# 방법 A: 작은따옴표 사용
python '.\Attendance and Leave Management Program.py'

# 방법 B: 큰따옴표 사용
python ".\Attendance and Leave Management Program.py"

# 방법 C: PowerShell 스크립트 사용
.\run_program.ps1
```

## 방법 3: 명령 프롬프트(CMD)에서 실행
```cmd
python "Attendance and Leave Management Program.py"
```

## 방법 4: VS Code에서 실행
1. 파일을 열고 F5 키를 누르거나
2. "Run Python File" 버튼 클릭 또는
3. 우클릭 > "Run Python File in Terminal"

## 주의사항
- 파일 이름에 공백이 있으므로 반드시 따옴표로 감싸야 합니다
- PowerShell에서 백틱(`)을 사용하지 마세요
- PySide6가 설치되어 있어야 합니다 (자동 설치 기능 제공)

