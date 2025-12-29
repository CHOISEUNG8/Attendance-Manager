# EXE 파일 빌드 가이드

이 프로그램을 Windows 실행 파일(.exe)로 빌드하는 방법입니다.

## 사전 요구사항

1. Python 3.8 이상이 설치되어 있어야 합니다.
2. 필요한 패키지들이 설치되어 있어야 합니다:
   ```bash
   pip install -r requirements.txt
   ```

## 빌드 방법

### 방법 1: 배치 파일 사용 (권장)

1. `build_exe.bat` 파일을 더블클릭하여 실행합니다.
2. 빌드가 완료되면 `dist` 폴더에 `WB_Attendance_Manager_v4_1.exe` 파일이 생성됩니다.

### 방법 2: PowerShell 스크립트 사용

1. PowerShell에서 `build_exe.ps1` 파일을 실행합니다:
   ```powershell
   .\build_exe.ps1
   ```
2. 빌드가 완료되면 `dist` 폴더에 `WB_Attendance_Manager_v4_1.exe` 파일이 생성됩니다.

### 방법 3: 수동 빌드

명령 프롬프트에서 다음 명령을 실행합니다:

```bash
python -m pip install pyinstaller
python -m PyInstaller --name="WB_Attendance_Manager_v4_1" --onefile --windowed --icon="favicon.ico" --add-data="favicon.ico;." "Attendance and Leave Management Program.py"
```

## 빌드 옵션 설명

- `--name="WB_Attendance_Manager_v4_1"`: 생성될 exe 파일의 이름
- `--onefile`: 단일 exe 파일로 생성 (폴더 대신)
- `--windowed`: 콘솔 창 없이 GUI만 표시
- `--icon="favicon.ico"`: exe 파일의 아이콘 설정
- `--add-data="favicon.ico;."`: favicon.ico 파일을 exe에 포함

## 빌드 후 파일 위치

빌드가 완료되면 다음 위치에 파일이 생성됩니다:

```
dist/
  └── WB_Attendance_Manager_v4_1.exe
```

## 주의사항

1. **첫 실행 시 시간**: exe 파일을 처음 실행할 때는 압축 해제 과정으로 인해 시간이 걸릴 수 있습니다.

2. **바이러스 백신 프로그램**: 일부 바이러스 백신 프로그램이 PyInstaller로 만든 exe 파일을 의심스러운 파일로 감지할 수 있습니다. 이는 정상적인 현상이며, 프로그램에 문제가 없습니다.

3. **데이터베이스 파일**: exe 파일과 같은 폴더에 `leave_attendance.db` 파일이 생성됩니다. 데이터를 보존하려면 이 파일을 함께 배포해야 합니다.

4. **파일 크기**: PySide6와 pandas를 포함하므로 exe 파일 크기가 약 100-200MB 정도 될 수 있습니다.

## 배포 방법

다른 컴퓨터에서 실행하려면:

1. `WB_Attendance_Manager_v4_1.exe` 파일을 복사합니다.
2. 필요시 `favicon.ico` 파일도 함께 복사합니다 (선택사항).
3. exe 파일을 실행합니다.
4. 프로그램이 자동으로 `leave_attendance.db` 데이터베이스 파일을 생성합니다.

## 문제 해결

### 빌드 오류 발생 시

1. 모든 패키지가 최신 버전인지 확인:
   ```bash
   pip install --upgrade -r requirements.txt
   ```

2. PyInstaller를 최신 버전으로 업그레이드:
   ```bash
   pip install --upgrade pyinstaller
   ```

3. 빌드 캐시 삭제 후 재빌드:
   ```bash
   rmdir /s /q build dist
   del *.spec
   ```

### 실행 오류 발생 시

1. exe 파일을 관리자 권한으로 실행해보세요.
2. Windows Defender나 바이러스 백신 프로그램에서 예외 처리하세요.
3. 필요한 Visual C++ 재배포 가능 패키지가 설치되어 있는지 확인하세요.

