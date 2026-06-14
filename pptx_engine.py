# -*- coding: utf-8 -*-
"""
Mock-deck Generator — PPT 작성/QA 엔진 (4단계 + 5단계)

핵심 원칙
---------
- 템플릿의 각 텍스트박스 "위치·글씨 크기·폰트·색"은 절대 바뀌지 않는다.
  → run 의 텍스트(<a:t>)만 교체하고, 도형/문단/run 의 서식 노드는 보존한다.
- 사용자가 장표를 N개 추가하면, 3종 템플릿 슬라이드(소제목 1/2/3개)를 골라
  복제해서 N장짜리 덱을 만든다.  복제는 think-cell OLE 객체만 제외하면
  외부 relationship 참조가 없으므로 rId 재매핑 없이 안전하게 deep-copy 된다.

역할(role) ↔ 도형 이름 매핑 (템플릿 이름 불일치 자동 흡수)
- head     : "head_msg"
- subtitle : "subtitle"
- sub{i}   : "sub{i} text"  또는  "sub{i} group" 그룹 안의 텍스트박스
             (슬라이드1은 텍스트박스가 "sub1 line"으로 잘못 명명돼 있어 그룹 스캔으로 흡수)
- body{i}  : "body{i}"  (소제목 1개 템플릿은 "body" 도 허용)
"""

import os
import copy
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

# --------------------------------------------------------------------------
# 경로 / 상수
# --------------------------------------------------------------------------
_HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(_HERE, "template", "template.pptx")

# 소제목 개수 -> 템플릿 슬라이드 인덱스 (0-based)
TEMPLATE_SLIDE_BY_NSUB = {1: 0, 2: 1, 3: 2}

_THINKCELL = "think-cell"

GROUP_TYPE = 6  # MSO_SHAPE_TYPE.GROUP
OLE_TYPE = 7    # MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT


# --------------------------------------------------------------------------
# 도형 순회 / 식별
# --------------------------------------------------------------------------
def iter_shapes(shapes):
    """그룹 내부까지 재귀적으로 모든 도형을 순회."""
    for sh in shapes:
        yield sh
        if sh.shape_type == GROUP_TYPE:
            yield from iter_shapes(sh.shapes)


def _is_ole(sh):
    name = (sh.name or "").lower()
    return sh.shape_type == OLE_TYPE or _THINKCELL in name


# --------------------------------------------------------------------------
# 역할 -> 도형 찾기 (이름 불일치 흡수)
# --------------------------------------------------------------------------
def _find_named_text(slide, name):
    for s in iter_shapes(slide.shapes):
        if s.name == name and s.has_text_frame:
            return s
    return None


def find_head(slide):
    return _find_named_text(slide, "head_msg")


def find_subtitle(slide):
    return _find_named_text(slide, "subtitle")


def find_sub(slide, i):
    """소제목 i (1-based) 텍스트박스. 'sub{i} text' 우선, 없으면 'sub{i} group' 내부 텍스트박스."""
    direct = _find_named_text(slide, "sub%d text" % i)
    if direct is not None:
        return direct
    for s in slide.shapes:
        if s.shape_type == GROUP_TYPE and s.name == "sub%d group" % i:
            cands = [c for c in s.shapes if c.has_text_frame]
            if cands:
                return cands[0]
    return None


def find_body(slide, i, n_sub):
    """본문 i (1-based). 'body{i}' 우선, 소제목 1개 템플릿은 'body' 도 허용."""
    s = _find_named_text(slide, "body%d" % i)
    if s is not None:
        return s
    if n_sub == 1:
        return _find_named_text(slide, "body")
    return None


# --------------------------------------------------------------------------
# 텍스트 교체 (서식 보존, 멀티라인 지원)
# --------------------------------------------------------------------------
def _set_para_text(p_elem, text):
    """문단 element 의 첫 run 텍스트만 교체하고 나머지 run 제거. 첫 run 의 rPr(서식) 보존."""
    runs = p_elem.findall(qn("a:r"))
    if runs:
        first = runs[0]
        t = first.find(qn("a:t"))
        if t is None:
            t = first.makeelement(qn("a:t"), {})
            first.append(t)
        t.text = text
        for r in runs[1:]:
            p_elem.remove(r)
    else:
        # run 이 없는 문단: endParaRPr 의 서식을 빌려 run 생성
        r = p_elem.makeelement(qn("a:r"), {})
        end_rpr = p_elem.find(qn("a:endParaRPr"))
        if end_rpr is not None:
            rpr = copy.deepcopy(end_rpr)
            rpr.tag = qn("a:rPr")
            r.append(rpr)
        t = p_elem.makeelement(qn("a:t"), {})
        t.text = text
        r.append(t)
        # pPr 다음에 삽입
        ppr = p_elem.find(qn("a:pPr"))
        if ppr is not None:
            ppr.addnext(r)
        else:
            p_elem.insert(0, r)


def set_text(shape, text):
    """도형의 텍스트를 교체. '\n' 은 문단 분리. 첫 문단을 복제해 추가 문단 서식 보존.

    위치/크기/폰트 서식은 일절 건드리지 않는다(런의 <a:t> 만 변경)."""
    tf = shape.text_frame
    txbody = tf._txBody
    paras = txbody.findall(qn("a:p"))
    if not paras:
        # 텍스트가 비어 있던 도형 — 단순 처리
        tf.text = text
        return
    p0 = paras[0]
    lines = text.split("\n") if text != "" else [""]
    _set_para_text(p0, lines[0])
    # 기존 추가 문단 제거
    for p in paras[1:]:
        txbody.remove(p)
    # 남은 줄은 p0 를 복제(서식 유지)해서 추가
    anchor = p0
    for line in lines[1:]:
        newp = copy.deepcopy(p0)
        _set_para_text(newp, line)
        anchor.addnext(newp)
        anchor = newp


# --------------------------------------------------------------------------
# 슬라이드 복제 (think-cell OLE 제외 → rId 재매핑 불필요)
# --------------------------------------------------------------------------
def clone_slide(prs, src_slide):
    """src_slide 의 콘텐츠 도형을 복제해 prs 에 새 슬라이드로 추가하고 반환."""
    layout = src_slide.slide_layout
    dst = prs.slides.add_slide(layout)
    # add_slide 가 레이아웃에서 만든 placeholder 전부 제거
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    sptree = dst.shapes._spTree
    for sh in src_slide.shapes:
        if _is_ole(sh):
            continue  # think-cell OLE 는 복제하지 않음
        sptree.append(copy.deepcopy(sh._element))
    return dst


def _remove_leading_slides(prs, count):
    """sldIdLst 의 앞쪽 count개 참조를 제거(원본 템플릿 슬라이드 정리).
    원본 3종 템플릿은 항상 맨 앞에 있고, 복제본은 그 뒤에 추가되므로 인덱스 기반이 안전."""
    sld_id_lst = prs.slides._sldIdLst
    for sid in list(sld_id_lst)[:count]:
        sld_id_lst.remove(sid)


# --------------------------------------------------------------------------
# 슬라이드 채우기
# --------------------------------------------------------------------------
def fill_slide(slide, spec):
    """spec = {head, subtitle, subs:[...], bodies:[...]}  (len(subs)==len(bodies), 1..3)
    반환: 채운 역할 목록 / 못 찾은 역할 목록."""
    n = len(spec["subs"])
    filled, missing = [], []

    def _do(shape, role, text):
        if shape is None:
            missing.append(role)
        else:
            set_text(shape, text)
            filled.append(role)

    _do(find_head(slide), "head", spec.get("head", ""))
    _do(find_subtitle(slide), "subtitle", spec.get("subtitle", ""))
    for i in range(1, n + 1):
        _do(find_sub(slide, i), "sub%d" % i, spec["subs"][i - 1])
        _do(find_body(slide, i, n), "body%d" % i, spec["bodies"][i - 1])
    return filled, missing


# --------------------------------------------------------------------------
# 덱 빌드 (4단계 메인 진입점)
# --------------------------------------------------------------------------
def build_deck(slides_spec, out_path, template_path=TEMPLATE_PATH):
    """slides_spec: list of {head, subtitle, subs:[1..3], bodies:[1..3]}
    각 spec 의 소제목 개수로 템플릿을 골라 복제·채워 N장 덱 생성."""
    prs = Presentation(template_path)
    template_slides = list(prs.slides)  # [1sub, 2sub, 3sub]

    report = []
    for idx, spec in enumerate(slides_spec, 1):
        n = len(spec["subs"])
        if n not in TEMPLATE_SLIDE_BY_NSUB:
            raise ValueError("슬라이드 %d: 소제목 개수는 1~3개여야 합니다 (현재 %d)" % (idx, n))
        src = template_slides[TEMPLATE_SLIDE_BY_NSUB[n]]
        new_slide = clone_slide(prs, src)
        filled, missing = fill_slide(new_slide, spec)
        report.append({"slide": idx, "n_sub": n, "filled": filled, "missing": missing})

    _remove_leading_slides(prs, len(template_slides))
    prs.save(out_path)
    return out_path, report


# --------------------------------------------------------------------------
# QA (5단계)
# --------------------------------------------------------------------------
def _role_signature(slide, n_sub):
    """역할별 (geometry, font) 시그니처 — 텍스트 제외."""
    sig = {}
    mapping = {"head": find_head(slide), "subtitle": find_subtitle(slide)}
    for i in range(1, n_sub + 1):
        mapping["sub%d" % i] = find_sub(slide, i)
        mapping["body%d" % i] = find_body(slide, i, n_sub)
    for role, sh in mapping.items():
        if sh is None:
            sig[role] = None
            continue
        geo = (sh.left, sh.top, sh.width, sh.height)
        font = None
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                f = r.font
                font = (f.name, (f.size.pt if f.size else None), f.bold)
                break
            break
        sig[role] = (geo, font)
    return sig


def _role_text(slide, role, n_sub):
    if role == "head":
        sh = find_head(slide)
    elif role == "subtitle":
        sh = find_subtitle(slide)
    elif role.startswith("sub"):
        sh = find_sub(slide, int(role[3:]))
    elif role.startswith("body"):
        sh = find_body(slide, int(role[4:]), n_sub)
    else:
        sh = None
    if sh is None:
        return None
    return "\n".join(p.text for p in sh.text_frame.paragraphs)


def qa_report(output_path, slides_spec, template_path=TEMPLATE_PATH):
    """5단계 QA:
       (1) 컨펌 내용이 모두 박스에 들어갔는가 (content completeness)
       (2) 위치·폰트·크기가 템플릿과 동일한가 (format integrity)
       (3) 오버플로우 가능성 (heuristic warning)
    반환: {ok: bool, issues: [...], warnings: [...]}
    """
    tmpl = Presentation(template_path)
    tmpl_slides = list(tmpl.slides)
    out = Presentation(output_path)
    out_slides = list(out.slides)

    issues, warnings = [], []

    if len(out_slides) != len(slides_spec):
        issues.append("출력 슬라이드 수(%d)가 컨펌 슬라이드 수(%d)와 다릅니다."
                      % (len(out_slides), len(slides_spec)))

    for idx, (spec, oslide) in enumerate(zip(slides_spec, out_slides), 1):
        n = len(spec["subs"])
        tslide = tmpl_slides[TEMPLATE_SLIDE_BY_NSUB[n]]

        # (2) 포맷 무결성
        t_sig = _role_signature(tslide, n)
        o_sig = _role_signature(oslide, n)
        for role, ts in t_sig.items():
            os_ = o_sig.get(role)
            if ts is None or os_ is None:
                issues.append("슬라이드 %d: 역할 '%s' 박스를 찾지 못함." % (idx, role))
                continue
            if ts[0] != os_[0]:
                issues.append("슬라이드 %d: '%s' 위치/크기 변경됨 (템플릿 %s ≠ 출력 %s)."
                              % (idx, role, ts[0], os_[0]))
            if ts[1] != os_[1]:
                issues.append("슬라이드 %d: '%s' 폰트 변경됨 (템플릿 %s ≠ 출력 %s)."
                              % (idx, role, ts[1], os_[1]))

        # (1) 내용 완전성
        expected = {"head": spec.get("head", ""), "subtitle": spec.get("subtitle", "")}
        for i in range(1, n + 1):
            expected["sub%d" % i] = spec["subs"][i - 1]
            expected["body%d" % i] = spec["bodies"][i - 1]
        for role, exp in expected.items():
            got = _role_text(oslide, role, n)
            if got is None:
                issues.append("슬라이드 %d: '%s' 박스 없음." % (idx, role))
            elif exp.strip() and exp.strip() not in got.replace("\n", " ").strip() \
                    and exp.strip() not in got:
                # 멀티라인 본문은 라인 단위로 재확인
                exp_lines = [l.strip() for l in exp.split("\n") if l.strip()]
                if not all(any(el in gl for gl in got.split("\n")) for el in exp_lines):
                    issues.append("슬라이드 %d: '%s' 내용 누락/불일치." % (idx, role))

        # (3) 오버플로우 휴리스틱 (본문 박스 기준)
        for i in range(1, n + 1):
            sh = find_body(oslide, i, n)
            if sh is None:
                continue
            body_text = spec["bodies"][i - 1]
            warn = _overflow_warn(sh, body_text)
            if warn:
                warnings.append("슬라이드 %d: body%d %s" % (idx, i, warn))

    return {"ok": len(issues) == 0, "issues": issues, "warnings": warnings}


def _overflow_warn(shape, text):
    """매우 러프한 오버플로우 추정. 박스 폭/높이와 폰트 크기로 수용 글자수를 가늠."""
    try:
        size_pt = None
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size:
                    size_pt = r.font.size.pt
                break
            if size_pt:
                break
        if not size_pt:
            size_pt = 14
        w_cm = Emu(shape.width).cm
        h_cm = Emu(shape.height).cm
        # 한글 기준 대략: 글자폭 ≈ size_pt*0.9pt ≈ size_pt*0.032cm, 줄높이 ≈ size_pt*0.045cm
        char_w = size_pt * 0.032
        line_h = size_pt * 0.045
        chars_per_line = max(1, int(w_cm / char_w))
        max_lines = max(1, int(h_cm / line_h))
        capacity = chars_per_line * max_lines
        # 실제 길이(줄바꿈은 줄 소비)
        lines = text.split("\n")
        used = sum(max(1, (len(l) + chars_per_line - 1) // chars_per_line) for l in lines)
        if used > max_lines:
            return "오버플로우 가능 (예상 %d줄 > 박스 %d줄)" % (used, max_lines)
        if len(text) > capacity:
            return "오버플로우 가능 (글자수 %d > 수용 ~%d)" % (len(text), capacity)
    except Exception:
        return None
    return None
