# -*- coding: utf-8 -*-
"""
Mock-deck Generator — 2단계 콘텐츠 에이전트

입력(프로젝트 주제 + 장표별 헤드메시지 + 중제목)을 받아, Claude API 로
  (a) research_checklist : 헤드메시지를 입증하려면 추가로 파악·확보해야 할 것
  (b) 구조화 방안          : 좌/우(2) · 좌/중/우(3) · 단일(1) 분할 (규칙 A: 칸=소제목)
  (c) 소제목/본문 초안     : 가짜 단정 없이 [ ] 플레이스홀더 활용
을 생성한다.

반환 스키마(SLIDE_SCHEMA):
{
  "n_sub": 1|2|3,
  "split": "single"|"left-right"|"left-center-right",
  "structure_rationale": str,
  "research_checklist": [str, ...],
  "sections": [
    {"position": "full|left|center|right",
     "subtitle": str, "body": str(불릿은 \\n 구분),
     "research_points": [str, ...]}
  ]   # len == n_sub
}

NOTE: SYSTEM_PROMPT 은 별도 워크플로(content-agent-prompt-design)로 최적화한
      최종본으로 교체될 예정. 교체 시 이 상수만 바꾸면 된다.
"""

import os
import json
import re

DEFAULT_MODEL = "claude-opus-4-8"

# --------------------------------------------------------------------------
# 시스템 프롬프트 + few-shot
# content-agent-prompt-design 워크플로 합성본(후보3 베이스, 점수 137/150).
# agent_prompt.py 는 워크플로 출력에서 자동 생성된다. 프롬프트 개선 시
# 워크플로를 재실행하고 agent_prompt.py 만 교체하면 된다.
# --------------------------------------------------------------------------
from agent_prompt import SYSTEM_PROMPT, FEWSHOT_USER, FEWSHOT_ASSISTANT, METHODS, VIZ_TYPES

FEWSHOT = [
    {"role": "user", "content": FEWSHOT_USER},
    {"role": "assistant", "content": FEWSHOT_ASSISTANT},
]

DEFAULT_METHOD = "리서치 필요"
DEFAULT_VIZ = "불릿 리스트"

SPLIT_BY_N = {1: "single", 2: "left-right", 3: "left-center-right"}
POS_BY_N = {1: ["full"], 2: ["left", "right"], 3: ["left", "center", "right"]}


# --------------------------------------------------------------------------
# 메인 진입점
# --------------------------------------------------------------------------
def generate_slide_content(
    project_topic,
    head_message,
    middle_title,
    project_context="",
    api_key=None,
    model=DEFAULT_MODEL,
    feedback=None,
    prior=None,
    max_tokens=8000,
):
    """한 장표에 대한 콘텐츠 초안 생성.

    feedback/prior 가 주어지면(3단계 재생성) 이전 결과를 피드백에 맞춰 수정한다.
    반환: SLIDE_SCHEMA dict.  실패 시 ContentAgentError 발생.
    """
    from anthropic import Anthropic

    client = Anthropic(api_key=api_key or os.environ.get("ANTHROPIC_API_KEY"))

    user_msg = (
        "프로젝트 주제: %s\n헤드메시지: %s\n중제목: %s"
        % (project_topic, head_message, middle_title)
    )
    if project_context and project_context.strip():
        user_msg += "\n프로젝트 Context(참고 맥락): %s" % project_context.strip()
    if feedback and prior:
        user_msg += (
            "\n\n# 이전 제안\n%s\n\n# 사용자 피드백(이에 맞춰 다시 제안하라)\n%s"
            % (json.dumps(prior, ensure_ascii=False), feedback)
        )

    messages = FEWSHOT + [{"role": "user", "content": user_msg}]

    resp = _create_message(client, model, max_tokens, SYSTEM_PROMPT, messages)
    # adaptive thinking 사용 시 thinking 블록이 섞이므로 text 블록만 취한다
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    data = _parse_json(text)
    return _normalize(data)


def _create_message(client, model, max_tokens, system, messages):
    """Opus 확장 사고(adaptive thinking) + effort=high 로 깊이 있게 생성.
    Opus 4.8은 thinking이 기본 OFF라, 켜야 길게 고민한 답이 나온다.
    설치된 anthropic SDK가 해당 파라미터를 모르면(구버전) 자동으로 일반 호출로 폴백."""
    base = dict(model=model, max_tokens=max_tokens, system=system, messages=messages)
    try:
        return client.messages.create(
            thinking={"type": "adaptive"},
            output_config={"effort": "high"},
            **base,
        )
    except TypeError:
        # 구버전 SDK: thinking/output_config 미지원 → 일반 호출
        return client.messages.create(**base)


class ContentAgentError(Exception):
    pass


# --------------------------------------------------------------------------
# 파싱 / 정규화
# --------------------------------------------------------------------------
def _parse_json(text):
    text = text.strip()
    # 코드펜스 제거
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    # strict=False: 모델이 body 문자열 안에 실제 줄바꿈을 넣어도 허용
    try:
        return json.loads(text, strict=False)
    except Exception:
        # 본문에서 최외곽 중괄호 추출
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0), strict=False)
            except Exception as e:
                raise ContentAgentError("JSON 파싱 실패: %s" % e)
        raise ContentAgentError("응답에서 JSON을 찾지 못함")


def _normalize(data):
    """스키마 정합성 보정: n_sub(1~3) clamp, sections 길이 일치, split/position 정합.
    각 섹션은 items[{question, method, hypothesis}] (질문↔가설 1:1) + viz_type + viz_note 를 가진다."""
    sections = data.get("sections") or []
    n = data.get("n_sub") or len(sections) or 1
    n = max(1, min(3, int(n)))
    if len(sections) > n:
        sections = sections[:n]
    while len(sections) < n:
        sections.append({})

    norm = []
    for i, sec in enumerate(sections):
        items = []
        for it in (sec.get("items") or []):
            q = (it.get("question") or "").strip()
            h = (it.get("hypothesis") or "").strip()
            m = it.get("method") or DEFAULT_METHOD
            if m not in METHODS:
                m = DEFAULT_METHOD
            if q or h:
                items.append({"question": q, "method": m, "hypothesis": h})
        vt = sec.get("viz_type") or DEFAULT_VIZ
        if vt not in VIZ_TYPES:
            vt = DEFAULT_VIZ
        norm.append({
            "position": POS_BY_N[n][i],
            "subtitle": (sec.get("subtitle") or "").strip(),
            "items": items,
            "viz_type": vt,
            "viz_note": (sec.get("viz_note") or "").strip(),
        })
    # 한 장표에 그래프(막대/도넛/워터폴)는 하나만 — 2개째부터 비-그래프로 강등(백스톱)
    _seen_graph = False
    for sec in norm:
        if sec["viz_type"] in ("막대 차트", "도넛/파이", "워터폴"):
            if _seen_graph:
                sec["viz_type"] = "비교표"
            else:
                _seen_graph = True
    return {
        "n_sub": n,
        "split": SPLIT_BY_N[n],
        "structure_rationale": (data.get("structure_rationale") or "").strip(),
        "sections": norm,
    }


def to_build_spec(head_message, middle_title, slide_content):
    """콘텐츠 결과(또는 사용자 수정본)를 pptx_engine.build_deck 의 spec 으로 변환.
    칸(섹션)마다 질문→KQ, 가설→Hypothesis, 도식유형+메모→Structure 로 매핑한다."""
    sections = []
    for sec in slide_content["sections"]:
        items = sec.get("items", [])
        # PPT 출력 시 각 줄 맨 앞에 라벨을 단다
        kq = "\n".join("[KQ] " + it["question"] for it in items if it.get("question"))
        hyp = "\n".join("[가설] " + it["hypothesis"] for it in items if it.get("hypothesis"))
        struct = (sec.get("viz_type") or "").strip()
        note = (sec.get("viz_note") or "").strip()
        if note:
            struct = (struct + " — " + note).strip(" —")
        struct = "[구조화 방향 제안] " + struct if struct else ""
        sections.append({
            "sub": sec.get("subtitle", ""),
            "kq": kq,
            "hyp": hyp,
            "struct": struct,
        })
    return {"head": head_message, "subtitle": middle_title, "sections": sections}
