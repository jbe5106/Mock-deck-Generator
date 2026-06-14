# Streamlit Community Cloud 무료 배포

설치 0·무료 호스팅·항상 접속. 입력 내용이 Streamlit 서버를 경유하므로 펌 보안정책을
먼저 확인하세요. 앱은 **특정 이메일만 접근**하도록 비공개로 잠글 수 있습니다.

## 1. GitHub에 코드 올리기 (무료)
1. GitHub 계정(무료): https://github.com/join
2. 새 레포 생성 — **Private 권장**, README 등 추가 없이 빈 레포로
3. 이 폴더를 푸시 (git init·첫 commit은 이미 해뒀습니다 → remote 연결 + push만):
   ```
   cd "Desktop/AI Test/mock_deck_generator"
   git remote add origin https://github.com/<아이디>/<레포이름>.git
   git branch -M main
   git push -u origin main
   ```
   (GitHub Desktop 앱을 쓰면 클릭만으로도 가능합니다.)

## 2. Streamlit Cloud에서 배포
1. https://share.streamlit.io → **Continue with GitHub** 로 로그인
2. **Create app → Deploy from GitHub**
3. Repository = `<아이디>/<레포이름>`, Branch = `main`, **Main file path = `app.py`**
4. **Deploy** → 1~2분 후 `https://<...>.streamlit.app` URL 생성

## 3. 보안 · 키 처리
- 앱 **Settings → Sharing**에서 **Viewer를 특정 이메일만 허용**(비공개)
- API 키: 각 사용자가 사이드바에 **본인 키** 입력 → 클라우드에선 **세션에만 저장**(보안상 의도된 동작, 사용자 간 유출 없음)
- 본인 키를 쓰면 **API 요금도 각자 청구** (공용 키 비추천)
- (선택) Settings → **Secrets**에 `cloud = true` 추가 — 코드가 경로로도 자동 감지하지만 명시적 안전장치

## 업데이트
- 코드 수정 후 `git push` 하면 Streamlit Cloud가 **자동 재배포**.

## 참고
- 무료 티어: 일정 시간 미사용 시 휴면 → 접속하면 수십 초 내 깨어남. RAM ~1GB(이 앱엔 충분).
- **버전 고정 권장**: 재배포 때 최신 streamlit이 설치됩니다. 안정성을 위해 `requirements.txt`의
  `streamlit>=1.36` 을 `streamlit==1.58.0` 으로 바꿔두면 항상 동일 버전으로 배포됩니다.
