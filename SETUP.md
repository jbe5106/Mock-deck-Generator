# 다른 컴퓨터에서 실행하기 (폴더 복사 + 로컬 실행)

데이터가 전부 그 PC에 머무는 방식입니다. 추가 비용 없음(Claude API 사용료만).

## 받는 PC에서 준비 (한 번만)

1. **Python 설치** — https://www.python.org/downloads/
   - 설치 화면에서 **"Add Python to PATH"** 체크 (중요!)

## 폴더 복사

2. 이 **`mock_deck_generator` 폴더 전체**를 받는 PC로 복사 (USB·공유드라이브·압축 등).
   - `.venv`, `__pycache__`, `output.pptx` 같은 건 **없어도 됩니다**(자동 생성). 복사돼 있어도 무방.
   - 반드시 같이 가야 하는 것: `app.py`, `*.py` 전부, `requirements.txt`, `template/template.pptx`, `.streamlit/`, `run.bat`

## 실행

3. **Windows**: 폴더 안의 **`run.bat` 더블클릭**
   - **Mac/Linux**: 터미널에서 `bash run.sh`
   - 처음 실행 시 라이브러리를 자동 설치(1~2분), 이후엔 바로 켜집니다.
   - 잠시 후 브라우저에 앱이 자동으로 열립니다. (안 열리면 http://localhost:8501)

4. 사이드바에 **본인 Anthropic API 키**를 입력 → 그 PC에 저장되어 다음부터 재입력 불필요.
   - 키 발급: https://console.anthropic.com/ → API Keys
   - **각자 본인 키를 쓰면 API 사용료도 각자 청구**됩니다(공용 키 비추천).

## 자주 묻는 것

- **"python을 찾을 수 없습니다"** → Python 미설치 또는 PATH 미체크. 1번 재설치(“Add to PATH” 체크).
- **느리게 켜짐 / 처음만 오래** → 최초 1회 라이브러리 설치 때문. 이후엔 빠릅니다.
- **포트 충돌(8501 사용 중)** → `run.bat` 마지막 줄을
  `... -m streamlit run app.py --server.port 8530` 으로 바꾸세요.
- **한글 폰트가 깨진 PPT** → 받는 PC에 맑은 고딕 등 기본 한글 폰트가 있으면 정상입니다(대부분 Windows 기본 탑재).
