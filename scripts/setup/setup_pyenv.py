"""
파일명: scripts/setup/setup_pyenv.py
목적: 프로젝트 파이썬 버전을 인식하여 버전 설정
설명: .python-version 파일을 읽어 pyenv local 명령으로 버전 설정
변경이력:
  - 2025-09-24: print(f"[setup_pyenv] ...") 표준 출력 포맷 적용
  - 2025-10-20: Windows 환경에서 pyenv 명령 실행 문제 해결
""" 

import subprocess
import shutil
import os
from pathlib import Path


def run_pyenv_command(version):
    """pyenv 명령을 실행하는 함수"""
    try:
        # 첫 번째 시도: 직접 pyenv 명령 실행
        subprocess.run(["pyenv", "local", version], check=True, shell=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        try:
            # 두 번째 시도: PowerShell을 통해 실행
            subprocess.run(["powershell", "-Command", f"pyenv local {version}"], check=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            try:
                # 세 번째 시도: pyenv 경로를 직접 찾아서 실행
                pyenv_path = shutil.which("pyenv")
                if pyenv_path:
                    subprocess.run([pyenv_path, "local", version], check=True)
                    return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                pass
    return False


version_file = Path(".python-version")
if version_file.exists():
    version = version_file.read_text(encoding="utf-8").strip()
    if version:
        if run_pyenv_command(version):
            print(f"[setup-pyenv] pyenv local {version} 실행 완료")
        else:
            print(f"[setup-pyenv] 경고: pyenv 명령 실행 실패. Python {version}이 설치되어 있는지 확인하세요.")
            print(f"[setup-pyenv] .python-version 파일은 이미 {version}으로 설정되어 있습니다.")
    else:
        print("[setup-pyenv] .python-version 파일이 비어 있습니다.")
else:
    print("[setup-pyenv] .python-version 파일이 없습니다.")
