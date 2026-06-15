#!/usr/bin/env bash
# Mac / Linux 실행 스크립트. 터미널에서:  bash run.sh
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "[오류] python3가 필요합니다. https://www.python.org/downloads/ 에서 설치하세요."
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "[준비] 가상환경 생성 중..."
  python3 -m venv .venv
fi

if [ ! -f ".venv/deps.ok" ]; then
  echo "[준비] 라이브러리 설치 중... (처음 한 번, 1~2분)"
  ./.venv/bin/python -m pip install --upgrade pip
  ./.venv/bin/python -m pip install -r requirements.txt || { echo "[오류] 설치 실패"; exit 1; }
  touch .venv/deps.ok
fi

echo "[실행] Mock-deck Generator 시작 — 브라우저가 자동으로 열립니다."
# echo "" 로 첫 실행 시 이메일 입력 프롬프트를 빈 값으로 건너뜀
echo "" | ./.venv/bin/python -m streamlit run app.py
