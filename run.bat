@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Mock-deck Generator

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo [오류] Python이 설치되어 있지 않습니다.
  echo   https://www.python.org/downloads/ 에서 Python을 설치하세요.
  echo   설치 화면에서 "Add Python to PATH" 체크가 필수입니다.
  echo.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [준비] 가상환경 생성 중...
  python -m venv .venv
)

if not exist ".venv\deps.ok" (
  echo [준비] 라이브러리 설치 중... 처음 한 번만, 1~2분 걸립니다.
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo [오류] 설치 실패. 인터넷 연결을 확인한 뒤 다시 실행하세요.
    pause
    exit /b 1
  )
  echo ok> ".venv\deps.ok"
)

echo.
echo [실행] Mock-deck Generator 시작 — 잠시 후 브라우저가 자동으로 열립니다.
echo   (종료하려면 이 검은 창에서 Ctrl+C 를 누르거나 창을 닫으세요.)
echo.
".venv\Scripts\python.exe" -m streamlit run app.py
pause
