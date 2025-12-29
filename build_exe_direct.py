# -*- coding: utf-8 -*-
"""
EXE 파일 빌드 스크립트 (직접 실행용)
"""
import os
import sys
import subprocess
import shutil

def main():
    print("=" * 50)
    print("EXE 파일 빌드 스크립트")
    print("=" * 50)
    print()
    
    # 현재 스크립트의 디렉토리로 이동
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(f"작업 디렉토리: {os.getcwd()}")
    print()
    
    # PyInstaller 설치 확인
    try:
        import PyInstaller
        print(f"PyInstaller 버전: {PyInstaller.__version__}")
    except ImportError:
        print("PyInstaller가 설치되어 있지 않습니다. 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller 설치 완료!")
        print()
    
    # 빌드 디렉토리 정리
    if os.path.exists("dist"):
        print("기존 빌드 파일 삭제 중...")
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")
    
    print()
    print("EXE 파일 빌드 시작...")
    print()
    
    # 빌드 옵션 준비
    build_args = [
        sys.executable, "-m", "PyInstaller",
        "--name=WB_Attendance_Manager_v4_1",
        "--onefile",
        "--windowed",
        "--hidden-import=pandas",
        "--hidden-import=pandas._libs.tslibs.timedeltas",
        "--hidden-import=pandas._libs.tslibs.nattype",
        "--hidden-import=pandas._libs.tslibs.np_datetime",
        "--hidden-import=pandas._libs.skiplist",
        "--hidden-import=pandas._libs.algos",
        "--hidden-import=pandas._libs.window.aggregations",
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=openpyxl.workbook.external_link.external",
        "--hidden-import=openpyxl.packaging.workbook",
        "--hidden-import=sqlite3",
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=PySide6.QtWidgets",
        "--collect-all=pandas",
        "--collect-all=openpyxl",
        "--collect-all=PySide6",
        "--collect-submodules=pandas",
        "--collect-submodules=openpyxl",
    ]
    
    # favicon.ico 파일 확인
    if os.path.exists("favicon.ico"):
        build_args.extend(["--icon=favicon.ico", "--add-data=favicon.ico;."])
        print("favicon.ico 파일을 포함합니다.")
    else:
        print("favicon.ico 파일이 없습니다. 아이콘 없이 빌드합니다.")
    
    # 데이터베이스 파일 확인 및 포함
    if os.path.exists("leave_attendance.db"):
        build_args.extend(["--add-data=leave_attendance.db;."])
        print("leave_attendance.db 파일을 포함합니다. (현재 입력된 데이터 포함)")
    else:
        print("leave_attendance.db 파일이 없습니다. 빈 데이터베이스로 시작합니다.")
    
    # 메인 스크립트 파일 추가
    main_script = "Attendance and Leave Management Program.py"
    if not os.path.exists(main_script):
        print(f"오류: {main_script} 파일을 찾을 수 없습니다!")
        return 1
    
    build_args.append(main_script)
    
    # PyInstaller 실행
    try:
        result = subprocess.run(build_args, check=True, encoding='utf-8', errors='replace')
        print()
        print("=" * 50)
        print("빌드 완료!")
        print("=" * 50)
        print()
        print(f"EXE 파일 위치: {os.path.join(os.getcwd(), 'dist', 'WB_Attendance_Manager_v4_1.exe')}")
        print()
        print("빌드된 파일을 실행하려면 dist 폴더의 exe 파일을 실행하세요.")
        print()
        return 0
    except subprocess.CalledProcessError as e:
        print()
        print("=" * 50)
        print("빌드 실패!")
        print("=" * 50)
        print(f"오류: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())

