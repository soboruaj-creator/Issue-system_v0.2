#!/usr/bin/env python3
"""
VOC 관리 시스템 - Python 직접 실행 스크립트

사전 요구사항:
  - Python 3.12+
  - Node.js 20+  (프론트엔드 최초 빌드 시 필요)
  - MongoDB 7.0  (로컬 설치 또는 원격 접속)

사용법:
  python start.py              # 의존성 설치 + 프론트엔드 빌드 + 서버 시작
  python start.py --skip-build # 이미 빌드된 경우 빌드 단계 건너뜀
  python start.py --dev        # 백엔드만 실행 (프론트엔드는 별도로 npm run dev)
"""
import os
import sys
import subprocess
import shutil
import argparse

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_DIR = os.path.join(ROOT_DIR, 'backend')
FRONTEND_DIR = os.path.join(ROOT_DIR, 'frontend')
FRONTEND_DIST = os.path.join(FRONTEND_DIR, 'dist')


def run_cmd(cmd, cwd=None):
    print(f"  $ {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd)
    if result.returncode != 0:
        print(f"\n  오류: 명령 실패 (종료 코드 {result.returncode})")
        sys.exit(result.returncode)


def install_backend_deps():
    print("\n[1/3] 백엔드 Python 패키지 설치 중...")
    run_cmd([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'], cwd=BACKEND_DIR)
    print("  완료.")


def get_npm():
    """Windows에서는 npm.cmd, 그 외에는 npm 사용"""
    if sys.platform == 'win32':
        return shutil.which('npm.cmd') or shutil.which('npm')
    return shutil.which('npm')


def build_frontend():
    print("\n[2/3] 프론트엔드 빌드 중...")

    npm = get_npm()
    if not npm:
        print("  경고: npm을 찾을 수 없습니다.")
        if os.path.exists(FRONTEND_DIST):
            print("  기존 빌드(frontend/dist)를 사용합니다.")
        else:
            print("  Node.js를 설치하거나 frontend/dist 디렉토리가 있어야 합니다.")
            print("  프론트엔드 없이 API 서버만 실행됩니다.")
        return

    run_cmd([npm, 'install'], cwd=FRONTEND_DIR)
    run_cmd([npm, 'run', 'build'], cwd=FRONTEND_DIR)
    print("  완료. (frontend/dist 생성됨)")


def start_server(dev_mode=False):
    print("\n[3/3] 서버 시작 중...")

    uploads_dir = os.path.join(BACKEND_DIR, 'uploads')
    os.makedirs(uploads_dir, exist_ok=True)

    if dev_mode:
        print("  [개발 모드] 백엔드만 실행 (프론트엔드: cd frontend && npm run dev)")
        print("  백엔드: http://localhost:8001")
        print("  API 문서: http://localhost:8001/docs")
    else:
        if os.path.exists(FRONTEND_DIST):
            print("  접속 주소: http://localhost:8001  (프론트엔드 + API 통합)")
        else:
            print("  접속 주소: http://localhost:8001  (API 서버만)")
        print("  API 문서: http://localhost:8001/docs")

    print("  종료: Ctrl+C\n")

    run_cmd([
        sys.executable, '-m', 'uvicorn', 'main:app',
        '--host', '0.0.0.0', '--port', '8001', '--reload'
    ], cwd=BACKEND_DIR)


def main():
    parser = argparse.ArgumentParser(description='VOC 관리 시스템 실행')
    parser.add_argument('--skip-build', action='store_true', help='프론트엔드 빌드 건너뜀')
    parser.add_argument('--skip-install', action='store_true', help='pip 패키지 설치 건너뜀')
    parser.add_argument('--dev', action='store_true', help='개발 모드 (백엔드만 실행)')
    args = parser.parse_args()

    print("=" * 55)
    print("  VOC 관리 시스템")
    print("=" * 55)

    if not args.skip_install:
        install_backend_deps()
    else:
        print("\n[1/3] 패키지 설치 건너뜀 (--skip-install)")

    if not args.skip_build and not args.dev:
        build_frontend()
    else:
        print("\n[2/3] 프론트엔드 빌드 건너뜀")

    start_server(dev_mode=args.dev)


if __name__ == '__main__':
    main()
