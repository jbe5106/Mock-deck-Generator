# Mock-deck Generator

헤드메시지만 적으면 → 본문에 들어갈 **리서치 항목**과 **구조화 방안**을 추천하고 →
컨펌한 구조로 **.pptx 목덱**을 자동 생성하는 도구.

대상: 스토리라인을 짜는 PM, 헤드메시지 장표를 받고 막막한 주니어 컨설턴트.

## 서비스 플로우 ↔ 모듈

| 단계 | 내용 | 구현 |
|---|---|---|
| 1 | 프로젝트 주제 + 장표별 헤드메시지·중제목 입력 | `app.py` (Streamlit) |
| 2 | **콘텐츠 에이전트**: 칸별 Requirements(질문)·Hypothesis(가설 답=본문)·Visualization(도식화) | `content_agent.py` (Claude API) |
| 3 | 장표별 컨펌 · 피드백 재생성 · 직접 수정 | `app.py` |
| 4 | **PPT 작성**: 템플릿 복제·채우기 (위치·폰트·크기 불변 보장) | `pptx_engine.py` (코드) |
| 5 | **QA**: 내용 완전성 + 포맷 무결성 + 오버플로우 점검 | `pptx_engine.py` (코드) |

2·4·5단계는 각각 독립 모듈. 4·5단계는 LLM이 아니라 **결정론적 코드**라
"템플릿 박스의 위치·글씨 크기·폰트는 절대 안 바뀐다"는 원칙이 구조적으로 보장된다
(런의 텍스트만 교체).

## 설치 & 실행

```bash
cd "Desktop/AI Test/mock_deck_generator"
pip install -r requirements.txt
export ANTHROPIC_API_KEY=<key>      # 또는 앱 사이드바에 입력

streamlit run app.py
```

## 템플릿 규칙 (`template/template.pptx`)

3종 슬라이드(소제목 1·2·3개). 도형 이름으로 텍스트박스를 매칭한다:

| 역할 | 도형 이름 |
|---|---|
| 헤드메시지 | `head_msg` |
| 중제목 | `subtitle` |
| 소제목 | `sub1 text` / `sub2 text` / `sub3 text` (그룹 `subN group` 내부) |
| 본문 | `body1` / `body2` / `body3` |

- 소제목 개수(1~3) = 좌/우/중 분할 수 = 사용할 템플릿 (**규칙 A: 칸 1개 = 소제목 1개**).
- 엔진은 이름 불일치를 자동 흡수한다(예: 1소제목 슬라이드의 `sub1 line`/`body`).
- think-cell OLE 객체는 복제 시 자동 제외(슬라이드 복제 안정성 확보).

## 비-LLM 코드 단독 사용 (4·5단계만)

```python
import pptx_engine as PE
specs = [{"head": "...", "subtitle": "...", "subs": ["좌", "우"], "bodies": ["...", "..."]}]
path, report = PE.build_deck(specs, "out.pptx")
print(PE.qa_report(path, specs))
```

## 콘텐츠 구조 (2단계 산출)

각 소제목(칸)마다:
- **items** — `질문(question) ↔ 가설(hypothesis)`을 1:1로 짝지은 배열. 각 질문에 **조사 방법**(`인터뷰 필요`/`리서치 필요`/`직접 검토 분석 필요`)이 붙는다. 가설들이 모여 .pptx 본문이 되며, 미확인 수치·비율·고유명사·날짜는 `[ ]` 플레이스홀더.
- **viz_type + viz_note** — 이 칸을 어떤 도식으로 그릴지(불릿/비교표/막대/도넛/워터폴/2x2/프로세스/강조박스) + 한 줄 메모.

2단계 UI는 칸별 `질문 → 가설` 카드와, 칸들을 박스·화살표로 그린 **슬라이드 와이어프레임 미리보기**를 보여준다(앱: `slide_wireframe_svg`).

## 메모

- 콘텐츠 에이전트 기본 모델: `claude-opus-4-8` (최신 Opus, 사이드바에서 변경 가능).
- **API 키는 처음 한 번만 입력**하면 `~/.mockdeck_config.json`(평문)에 저장돼 재입력 불필요.
- 밝은 테마 고정: `.streamlit/config.toml`. 시스템 프롬프트는 `agent_prompt.py`.
- 단계명: 1단계 Head Message 초안 → 2단계 본문 방향 잡기 → 3단계 PPT 생성.
