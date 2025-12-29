# EXE 빌드 체크리스트

## 최근 수정 사항 확인

### 1. 데이터베이스 연결 관리 (database is locked 오류 수정)
- [ ] `process_attendance_record` 함수에 `conn` 파라미터가 있는지 확인
- [ ] `save_changes` 함수에서 `conn`을 전달하는지 확인
- 위치: `AttendanceCalculator.process_attendance_record` (line 401)

### 2. 야근(20시이후) 가운데 정렬
- [ ] 야근 컬럼에 `setTextAlignment(Qt.AlignCenter)` 적용 확인
- 위치: `_refresh_month_data` 함수 내

### 3. 반차 음양처리
- [ ] 반차가 `['연차', '반차', ...]` 리스트에 포함되어 있는지 확인
- [ ] `#F8CBAD` 배경색이 적용되는지 확인
- 위치: `_refresh_month_data` 함수 (line 3898, 3930 등)

### 4. 입사일 대각선 처리
- [ ] 입사일 이전 날짜에 대각선이 그어지는지 확인
- 위치: `_refresh_month_data` 함수 내

### 5. 반차 삭제 처리
- [ ] 빈 값 처리 시 `is_delete: True` 플래그가 설정되는지 확인
- [ ] `save_changes`에서 `is_delete` 처리 로직 확인
- 위치: `on_cell_changed` (line 4450), `save_changes` (line 3355)

### 6. 토요일/일요일 평균 계산 제외
- [ ] 평균 출근시간 계산 시 토요일(5), 일요일(6) 제외 확인
- [ ] 평균 퇴근시간 계산 시 토요일(5), 일요일(6) 제외 확인
- 위치: `_refresh_month_data` (line 3908-3911, 3938-3941, 4084-4087, 4115-4118)

### 7. 매월 셋째주 수요일 평균 퇴근시간 제외
- [ ] `is_third_wednesday_17_00` 함수 확인
- [ ] 평균 퇴근시간 계산 시 셋째주 수요일 17:00 제외 확인
- 위치: `is_third_wednesday_17_00` (line 3197), `_refresh_month_data` (line 4087, 4118)

### 8. 엑셀 다운로드 기능
- [ ] 엑셀 다운로드 시에도 동일한 평균 계산 로직 적용 확인
- 위치: `download_excel` 함수 (line 5962 등)

## 빌드 확인 사항

### Spec 파일 확인
- [ ] `wb_attendance_v4_1.spec` 파일이 최신인지 확인
- [ ] `Attendance and Leave Management Program.py` 파일이 올바르게 지정되어 있는지 확인 (line 19)

### 빌드 시간 확인
- EXE 빌드 시간: 2025-12-03 오후 5:54:48
- Python 파일 수정 시간: 2025-12-03 오후 4:48:40
- ✅ EXE가 Python 파일보다 나중에 빌드되었으므로 최신 코드 반영됨

### 빌드 후 추가 수정 확인
- [ ] EXE 빌드 후 코드 수정이 있었는지 확인
- [ ] 수정이 있었다면 재빌드 필요

## 재빌드 필요 시

1. `build_exe_spec.bat` 실행
2. `dist\WB_Attendance_Manager_v4_1.exe` 파일 확인
3. 테스트 실행

## 주요 함수 위치

- `process_attendance_record`: line 401
- `is_third_wednesday_17_00`: line 3197
- `_refresh_month_data`: line 3708
- `on_cell_changed`: line 4400
- `save_changes`: line 3339
- `download_excel`: line 5806

