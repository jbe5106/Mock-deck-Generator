# -*- coding: utf-8 -*-
"""
Mock-deck Generator — Streamlit 앱

1단계: Head Message 초안 작성하기 (주제 + 장표별 헤드메시지·중제목)
2단계: 본문 작성 방향 잡기 (칸별 질문↔가설 1:1 + 조사방법 + 도식 제안, 검토·수정·컨펌)
3단계: PPT로 생성 (템플릿 복제·채우기 + QA)

실행:  streamlit run app.py
"""

import os
import json
import datetime
import html as _html
import streamlit as st
import streamlit.components.v1 as components

import content_agent as CA
import pptx_engine as PE
import emailer as EM

st.set_page_config(page_title="Mock-deck Generator", page_icon="📊", layout="wide")

# 입력칸 글자(볼드+살짝 크게) + 버튼(크게+볼드)
st.markdown("""<style>
.stTextInput input, [data-testid="stTextInput"] input,
.stTextArea textarea, [data-testid="stTextArea"] textarea {
    font-weight: 700 !important;
    font-size: 1.05rem !important;
}
/* 버튼: 라벨 텍스트는 내부 <p>/<div>에 있어 그걸 직접 타겟해야 적용됨 */
.stButton button, .stDownloadButton button {
    padding: 0.65rem 1.2rem !important;
    min-height: 3rem !important;
}
.stButton button p, .stButton button div,
.stDownloadButton button p, .stDownloadButton button div {
    font-size: 1.15rem !important;
    font-weight: 800 !important;
}
/* 셀렉트박스(추천 조사 방법·도식 유형)의 선택된 값도 크게+볼드 */
.stSelectbox div[data-baseweb="select"] div {
    font-size: 1.05rem !important;
    font-weight: 700 !important;
}
</style>""", unsafe_allow_html=True)

OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output.pptx")
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".mockdeck_config.json")

# (도식 유형은 slide_wireframe_svg 의 _schematic 으로 미니 모식도를 그린다)


# --------------------------------------------------------------------------
# 설정 영속화 (로컬: 파일 저장 → 재입력 불필요 / 클라우드: 세션 전용)
# 클라우드(공유 컨테이너)에서 키를 파일에 저장하면 다른 사용자에게 노출되므로,
# 클라우드로 감지되면 디스크 저장을 끄고 각 세션에만 키를 둔다.
# --------------------------------------------------------------------------
def _is_cloud():
    here = os.path.abspath(__file__)
    if here.startswith("/mount") or here.startswith("/app"):  # Streamlit Cloud 경로
        return True
    try:
        return bool(st.secrets.get("cloud", False))
    except Exception:
        return False


def _load_config():
    if _is_cloud():
        return {}
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config():
    if _is_cloud():
        return  # 클라우드에서는 디스크에 저장하지 않음(세션 전용)
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"api_key": st.session_state.get("api_key", ""),
                       "model": st.session_state.get("model", CA.DEFAULT_MODEL),
                       "smtp_sender": st.session_state.get("smtp_sender", ""),
                       "smtp_password": st.session_state.get("smtp_password", "")},
                      f, ensure_ascii=False)
    except Exception:
        pass


# --------------------------------------------------------------------------
# 슬라이드 와이어프레임 (제안 레이아웃 미니 도식)
# --------------------------------------------------------------------------
def _esc(s):
    return _html.escape(str(s), quote=True)


def _wrap(text, max_chars, max_lines):
    s = (text or "").strip()
    lines = []
    while s and len(lines) < max_lines:
        lines.append(s[:max_chars])
        s = s[max_chars:]
    if s and lines:
        lines[-1] = lines[-1][:max(1, max_chars - 1)] + "…"
    return lines


def _schematic(vt, cx, cy):
    """viz_type 별 미니 모식도 SVG 요소 (중심 cx,cy, 대략 ±37 x ±24)."""
    B, L, G = "#2563EB", "#93C5FD", "#94A3B8"
    e = []
    if vt == "막대 차트":
        base = cy + 22
        for i, h in enumerate((16, 30, 22, 38)):
            e.append('<rect x="%.1f" y="%.1f" width="12" height="%d" fill="%s"/>'
                     % (cx - 33 + i * 18, base - h, h, B if i == 3 else L))
        e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1.5"/>'
                 % (cx - 37, base, cx + 37, base, G))
    elif vt == "비교표":
        x0, y0, w, h = cx - 36, cy - 22, 72, 44
        e.append('<rect x="%.1f" y="%.1f" width="%d" height="%d" fill="#FFFFFF" stroke="%s"/>' % (x0, y0, w, h, G))
        e.append('<rect x="%.1f" y="%.1f" width="%d" height="14" fill="%s"/>' % (x0, y0, w, L))
        for xc in (x0 + 24, x0 + 48):
            e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1"/>' % (xc, y0, xc, y0 + h, G))
        for yc in (y0 + 14, y0 + 29):
            e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1"/>' % (x0, yc, x0 + w, yc, G))
    elif vt == "도넛/파이":
        e.append('<circle cx="%.1f" cy="%.1f" r="18" fill="none" stroke="%s" stroke-width="10"/>' % (cx, cy, L))
        e.append('<circle cx="%.1f" cy="%.1f" r="18" fill="none" stroke="%s" stroke-width="10" '
                 'stroke-dasharray="45 68" transform="rotate(-90 %.1f %.1f)"/>' % (cx, cy, B, cx, cy))
    elif vt == "워터폴":
        base = cy + 22
        for x, yy, h, f in [(cx - 34, base - 16, 16, L), (cx - 16, base - 26, 8, B),
                            (cx + 2, base - 34, 8, B), (cx + 20, base - 34, 34, L)]:
            e.append('<rect x="%.1f" y="%.1f" width="14" height="%d" fill="%s"/>' % (x, yy, h, f))
    elif vt == "2x2 매트릭스":
        s = 44
        x0, y0 = cx - s / 2, cy - s / 2
        e.append('<rect x="%.1f" y="%.1f" width="%d" height="%d" fill="#FFFFFF" stroke="%s"/>' % (x0, y0, s, s, G))
        e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1"/>' % (cx, y0, cx, y0 + s, G))
        e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1"/>' % (x0, cy, x0 + s, cy, G))
        e.append('<circle cx="%.1f" cy="%.1f" r="4.5" fill="%s"/>' % (cx + 11, cy - 11, B))
    elif vt == "프로세스(화살표)":
        for i in range(3):
            bx = cx - 36 + i * 27
            e.append('<rect x="%.1f" y="%.1f" width="18" height="16" rx="2" fill="%s"/>' % (bx, cy - 8, B if i == 2 else L))
            if i < 2:
                ax = bx + 18
                e.append('<path d="M%.1f,%.1f L%.1f,%.1f L%.1f,%.1f Z" fill="%s"/>'
                         % (ax + 1, cy - 4, ax + 7, cy, ax + 1, cy + 4, G))
    elif vt == "강조 박스":
        e.append('<rect x="%.1f" y="%.1f" width="72" height="34" rx="5" fill="#EFF6FF" stroke="%s" stroke-width="2"/>' % (cx - 36, cy - 17, B))
        e.append('<rect x="%.1f" y="%.1f" width="5" height="34" rx="2" fill="%s"/>' % (cx - 36, cy - 17, B))
        e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="2"/>' % (cx - 22, cy - 5, cx + 28, cy - 5, G))
        e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="2"/>' % (cx - 22, cy + 6, cx + 14, cy + 6, G))
    else:  # 불릿 리스트
        for i in range(3):
            yy = cy - 14 + i * 14
            e.append('<circle cx="%.1f" cy="%.1f" r="2.6" fill="%s"/>' % (cx - 30, yy, B))
            e.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="2"/>' % (cx - 23, yy, cx + 30, yy, L))
    return "".join(e)


def slide_wireframe_svg(sections):
    n = max(1, len(sections))
    W, H, pad, gap = 760, 226, 12, 46
    bw = (W - 2 * pad - (n - 1) * gap) / n
    y, bh = 14, 196
    p = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">' % (W, H, W, H)]
    p.append('<rect x="0" y="0" width="%d" height="%d" rx="8" fill="#FFFFFF"/>' % (W, H))
    p.append('<defs><marker id="ah" markerWidth="10" markerHeight="10" refX="7" refY="3" '
             'orient="auto"><path d="M0,0 L7,3 L0,6 Z" fill="#94A3B8"/></marker></defs>')
    for i, sec in enumerate(sections):
        x = pad + i * (bw + gap)
        cx = x + bw / 2
        p.append('<rect x="%.1f" y="%d" width="%.1f" height="%d" rx="10" fill="#F8FAFC" stroke="#CBD5E1"/>'
                 % (x, y, bw, bh))
        for li, line in enumerate(_wrap(sec.get("subtitle") or "(소제목)", int(bw / 12), 2)):
            p.append('<text x="%.1f" y="%d" text-anchor="middle" font-family="sans-serif" '
                     'font-size="13.5" font-weight="700" fill="#1E293B">%s</text>'
                     % (cx, y + 22 + li * 17, _esc(line)))
        p.append(_schematic(sec.get("viz_type") or "불릿 리스트", cx, y + 96))
        p.append('<text x="%.1f" y="%d" text-anchor="middle" font-family="sans-serif" '
                 'font-size="12" font-weight="600" fill="#2563EB">%s</text>'
                 % (cx, y + 140, _esc(sec.get("viz_type") or "")))
        for li, line in enumerate(_wrap(sec.get("viz_note") or "", int(bw / 8), 3)):
            p.append('<text x="%.1f" y="%d" text-anchor="middle" font-family="sans-serif" '
                     'font-size="10.5" fill="#64748B">%s</text>' % (cx, y + 156 + li * 14, _esc(line)))
        if i < n - 1:
            ay = y + bh / 2
            p.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#94A3B8" '
                     'stroke-width="2" marker-end="url(#ah)"/>' % (x + bw + 6, ay, x + bw + gap - 8, ay))
    p.append('</svg>')
    return "".join(p)


def single_schematic_svg(vt):
    """도식 유형 하나의 미니 모식도(셀렉트박스 옆에 보여줄 그림)."""
    W, H = 170, 92
    s = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">' % (W, H, W, H)]
    s.append('<rect x="1" y="1" width="%d" height="%d" rx="8" fill="#F8FAFC" stroke="#CBD5E1"/>' % (W - 2, H - 2))
    s.append(_schematic(vt, W / 2, H / 2 - 2))
    s.append('</svg>')
    return "".join(s)


def _idx(options, value, default=0):
    try:
        return options.index(value)
    except ValueError:
        return default


# --------------------------------------------------------------------------
# 세션 상태
# --------------------------------------------------------------------------
def _new_slide():
    sid = st.session_state._next_id
    st.session_state._next_id += 1
    return {"id": sid, "head": "", "middle": "", "content": None, "confirmed": False}


def init_state():
    if "_init" in st.session_state:
        return
    st.session_state._init = True
    st.session_state._next_id = 0
    st.session_state.topic = ""
    st.session_state.context = ""
    st.session_state.slides = [_new_slide() for _ in range(3)]
    st.session_state.built_path = None
    st.session_state.qa = None
    cfg = _load_config()
    st.session_state.api_key = cfg.get("api_key") or os.environ.get("ANTHROPIC_API_KEY", "")
    st.session_state.model = cfg.get("model") or CA.DEFAULT_MODEL
    st.session_state.smtp_sender = cfg.get("smtp_sender") or ""
    st.session_state.smtp_password = cfg.get("smtp_password") or ""


init_state()


# --------------------------------------------------------------------------
# 사이드바
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 설정")
    st.text_input("Anthropic API Key", type="password", key="api_key", on_change=_save_config,
                  help="한 번 입력하면 로컬(~/.mockdeck_config.json)에 저장돼 다시 입력할 필요가 없습니다.")
    st.text_input("모델", key="model", on_change=_save_config)
    if st.session_state.api_key:
        st.caption("🔑 키 입력됨 — 이 세션에만 유지(클라우드)" if _is_cloud()
                   else "🔑 키 저장됨 — 재입력 불필요")
    with st.expander("✉️ 이메일(SMTP) 설정"):
        st.text_input("보내는 Gmail 주소", key="smtp_sender", on_change=_save_config,
                      placeholder="you@gmail.com")
        st.text_input("Gmail 앱 비밀번호", type="password", key="smtp_password", on_change=_save_config,
                      help="2단계 인증 후 Google 계정 > 보안 > 앱 비밀번호에서 발급한 16자리")
        st.caption("Gmail 기준. 회사 SMTP면 emailer.py 의 host/port를 바꾸세요.")
    st.caption("템플릿: template/template.pptx (소제목 1/2/3종)")


st.title("📊 Mock-deck Generator")


# ==========================================================================
# 1단계: Head Message 초안 작성하기
# ==========================================================================
st.header("1단계: Head Message 초안 작성하기")
st.session_state.topic = st.text_input(
    "프로젝트 주제", value=st.session_state.topic,
    placeholder="예: A사·B사 합병 후 영업조직 통합(PMI) 전략 수립")
st.session_state.context = st.text_area(
    "프로젝트 Context (Optional)", value=st.session_state.get("context", ""),
    placeholder="프로젝트 배경·맥락을 적으면 2단계 생성에 반영됩니다. (비워도 됩니다)", height=90)

st.markdown("**장표별 헤드메시지 + 중제목**")
remove_id = None
for idx, slide in enumerate(st.session_state.slides):
    sid = slide["id"]
    with st.container(border=True):
        st.markdown("**%d번 장표**" % (idx + 1))
        c1, c2, c3 = st.columns([6, 5, 1])
        slide["head"] = c1.text_input("헤드메시지", value=slide["head"], key="head_%d" % sid,
                                       placeholder="이 장표가 말하려는 핵심 메시지 한 문장")
        slide["middle"] = c2.text_input("중제목", value=slide["middle"], key="mid_%d" % sid,
                                        placeholder="예: 양사 채널 정책 비교")
        c3.markdown("<br>", unsafe_allow_html=True)
        if c3.button("🗑️", key="del_%d" % sid, help="이 장표 삭제"):
            remove_id = sid

if remove_id is not None and len(st.session_state.slides) > 1:
    st.session_state.slides = [s for s in st.session_state.slides if s["id"] != remove_id]
    st.rerun()

cadd, cgen = st.columns([1, 2])
if cadd.button("➕ 장표 추가"):
    st.session_state.slides.append(_new_slide())
    st.rerun()

if cgen.button("🤖 본문 작성 방향 잡기", type="primary"):
    if not st.session_state.topic.strip():
        st.error("프로젝트 주제를 입력하세요.")
    elif not any(s["head"].strip() for s in st.session_state.slides):
        st.error("최소 한 장표의 헤드메시지를 입력하세요.")
    elif not st.session_state.api_key:
        st.error("사이드바에 API Key를 입력하세요.")
    else:
        prog = st.progress(0.0, "생성 중...")
        targets = [s for s in st.session_state.slides if s["head"].strip()]
        for k, slide in enumerate(targets):
            try:
                slide["content"] = CA.generate_slide_content(
                    st.session_state.topic, slide["head"], slide["middle"],
                    project_context=st.session_state.get("context", ""),
                    api_key=st.session_state.api_key, model=st.session_state.model)
                slide["confirmed"] = False
            except Exception as e:
                st.error("장표 '%s' 생성 실패: %s" % (slide["head"][:20], e))
            prog.progress((k + 1) / len(targets), "생성 중... (%d/%d)" % (k + 1, len(targets)))
        prog.empty()
        st.rerun()


# ==========================================================================
# 2단계: 본문 작성 방향 잡기
# ==========================================================================
generated = [s for s in st.session_state.slides if s.get("content")]
if generated:
    st.divider()
    st.header("2단계: 본문 작성 방향 잡기")
    st.caption("Head Message 내용을 뒷받침하기 위해 조사 필요한 Key Question과 가설을 도출합니다")

    split_label = {"single": "단일", "left-right": "좌·우 2분할", "left-center-right": "좌·중·우 3분할"}

    for slide in generated:
        sid = slide["id"]
        content = slide["content"]
        n = content["n_sub"]
        status = "✅ 컨펌됨" if slide["confirmed"] else "⏳ 미확정"
        title = "📄 %s  —  [%s] %s" % (slide["head"][:40] or "(제목없음)", split_label[content["split"]], status)
        with st.expander(title, expanded=not slide["confirmed"]):

            st.caption("구조화 근거: " + (content["structure_rationale"] or "—"))

            cols = st.columns(n)
            for j, sec in enumerate(content["sections"]):
                with cols[j]:
                    st.markdown("##### [%s] 소제목 %d" % (sec["position"], j + 1))
                    sec["subtitle"] = st.text_input("소제목", value=sec["subtitle"],
                                                    key="sub_%d_%d" % (sid, j), label_visibility="collapsed")

                    st.markdown("<p style='font-size:1.18rem;font-weight:800;margin:8px 0 2px'>1. Key Question 및 초기 가설 도출</p>", unsafe_allow_html=True)
                    remove_item = None
                    for k, item in enumerate(sec["items"]):
                        with st.container(border=True):
                            item["question"] = st.text_area(
                                "질문 %d" % (k + 1), value=item["question"],
                                key="q_%d_%d_%d" % (sid, j, k), height=90,
                                placeholder="조사 필요 질문?")
                            item["hypothesis"] = st.text_area(
                                "⤷ 가설", value=item["hypothesis"],
                                key="h_%d_%d_%d" % (sid, j, k), height=180,
                                placeholder="이 질문에 대한 2~3문장 가설")
                            mc, dc = st.columns([4, 1])
                            item["method"] = mc.selectbox(
                                "추천 조사 방법", CA.METHODS, index=_idx(CA.METHODS, item.get("method")),
                                key="m_%d_%d_%d" % (sid, j, k))
                            dc.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                            if dc.button("🗑", key="dq_%d_%d_%d" % (sid, j, k), help="이 질문 삭제"):
                                remove_item = k
                    if remove_item is not None and len(sec["items"]) > 1:
                        sec["items"].pop(remove_item)
                        st.rerun()
                    if st.button("➕ 질문·가설 추가", key="addq_%d_%d" % (sid, j)):
                        sec["items"].append({"question": "", "method": CA.DEFAULT_METHOD, "hypothesis": ""})
                        st.rerun()

                    st.markdown("<p style='font-size:1.18rem;font-weight:800;margin:8px 0 2px'>2. 구조화 방향 제안</p>", unsafe_allow_html=True)
                    sec["viz_type"] = st.selectbox("도식 유형", CA.VIZ_TYPES,
                                                   index=_idx(CA.VIZ_TYPES, sec.get("viz_type")),
                                                   key="vt_%d_%d" % (sid, j), label_visibility="collapsed")
                    components.html(single_schematic_svg(sec["viz_type"]), height=104)
                    sec["viz_note"] = st.text_area("구조화 방향 제안", value=sec.get("viz_note", ""),
                                                   key="vn_%d_%d" % (sid, j), height=90,
                                                   placeholder="어떻게 구조화할지 (강조·화살표·배치 등)")

            # 컨펌
            if st.button(("✅ 이 장표 컨펌" if not slide["confirmed"] else "↩️ 컨펌 해제"),
                         key="cfm_%d" % sid, type="primary" if not slide["confirmed"] else "secondary"):
                slide["confirmed"] = not slide["confirmed"]
                st.rerun()


# ==========================================================================
# 3단계: PPT로 생성
# ==========================================================================
confirmed = [s for s in generated if s["confirmed"]]
if generated:
    st.divider()
    st.header("3단계: PPT로 생성")
    st.caption("컨펌된 장표 %d / %d 장이 .pptx로 생성됩니다. (각 칸의 질문→KQ, 가설→Hypothesis, 구조화 방향→Structure 박스)"
               % (len(confirmed), len(generated)))

    if st.button("📦 PPT 생성", type="primary", disabled=not confirmed):
        specs = [CA.to_build_spec(s["head"], s["middle"], s["content"]) for s in confirmed]
        try:
            path, report = PE.build_deck(specs, OUT_PATH)
            qa = PE.qa_report(path, specs)
            st.session_state.built_path = path
            st.session_state.qa = qa
        except Exception as e:
            st.error("생성 실패: %s" % e)
        st.rerun()

    if st.session_state.built_path:
        qa = st.session_state.qa
        if qa["ok"]:
            st.success("✅ QA 통과 — 위치·폰트·크기 무결성 + 내용 완전성 확인됨")
        else:
            st.error("❌ QA 이슈:")
            for i in qa["issues"]:
                st.markdown("- " + i)
        if qa["warnings"]:
            st.warning("⚠️ 경고(오버플로우 가능):")
            for w in qa["warnings"]:
                st.markdown("- " + w)

        _fname = "Mock-deck 초안_%s.pptx" % datetime.datetime.now().strftime("%y%m%d")
        with open(st.session_state.built_path, "rb") as f:
            st.download_button("⬇️ PPT 다운로드 (%s)" % _fname, f, file_name=_fname,
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

        with st.expander("📧 팀원에게 이메일로 보내기"):
            if not st.session_state.smtp_sender or not st.session_state.smtp_password:
                st.info("사이드바 ‘✉️ 이메일(SMTP) 설정’에 보내는 Gmail 주소와 앱 비밀번호를 먼저 입력하세요.")
            to = st.text_input("받는 사람 (쉼표로 구분)", key="email_to",
                               placeholder="teammate1@company.com, teammate2@company.com")
            subj = st.text_input("제목", value="[Mock-deck] 목덱 초안 공유", key="email_subj")
            body = st.text_area("내용", height=120, key="email_body",
                                value="안녕하세요,\n첨부된 목덱 초안을 공유드립니다. 검토 부탁드립니다.")
            if st.button("📨 이메일 전송", disabled=not st.session_state.smtp_sender):
                try:
                    rcpts = EM.parse_recipients(to)
                    with st.spinner("전송 중..."):
                        EM.send_email(st.session_state.smtp_sender, st.session_state.smtp_password,
                                      rcpts, subj, body, st.session_state.built_path)
                    st.success("✅ 전송 완료 → %s" % ", ".join(rcpts))
                except Exception as e:
                    st.error("전송 실패: %s" % e)
