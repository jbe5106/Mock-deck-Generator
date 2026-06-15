# -*- coding: utf-8 -*-
"""
Mock-deck Generator — PPT 작성/QA 엔진 (4단계 + 5단계)

템플릿(3종: 칸 1/2/3개) 박스 구성 (칸 n = 1·2·3):
- head_msg      : Head Message
- subtitle      : 중제목
- sub{n}        : 소제목 (그룹 'sub{n} group' 내부 텍스트박스; 슬라이드1은 'sub1 line')
- KQ{n}         : 핵심 질문(Key Question)
- Hypothesis{n} : 가설
- Structure{n}  : 구조화 방향 제안(도식 유형 + 메모)

원칙: 박스의 위치·글씨크기·폰트는 절대 불변(run 텍스트만 교체). 사용자가 칸을 N개
만들면 n_sub(1/2/3) 템플릿을 골라 복제·채운다. 복제는 think-cell OLE만 제외하면
외부 relationship 참조가 없어 rId 재매핑 없이 안전.
"""

import os
import copy
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

_HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(_HERE, "template", "template.pptx")

TEMPLATE_SLIDE_BY_NSUB = {1: 0, 2: 1, 3: 2}
_THINKCELL = "think-cell"
GROUP_TYPE = 6
OLE_TYPE = 7


# --------------------------------------------------------------------------
# 도형 순회 / 식별
# --------------------------------------------------------------------------
def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == GROUP_TYPE:
            yield from iter_shapes(sh.shapes)


def _is_ole(sh):
    name = (sh.name or "").lower()
    return sh.shape_type == OLE_TYPE or _THINKCELL in name


# --------------------------------------------------------------------------
# 역할 -> 도형 찾기
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
    """소제목 i. 'sub{i} text' 우선, 없으면 'sub{i} group' 내부 텍스트박스(슬라이드1 'sub1 line' 흡수)."""
    direct = _find_named_text(slide, "sub%d text" % i)
    if direct is not None:
        return direct
    for s in slide.shapes:
        if s.shape_type == GROUP_TYPE and s.name == "sub%d group" % i:
            cands = [c for c in s.shapes if c.has_text_frame]
            if cands:
                return cands[0]
    return None


def find_kq(slide, i):
    return _find_named_text(slide, "KQ%d" % i)


def find_hyp(slide, i):
    return _find_named_text(slide, "Hypothesis%d" % i)


def find_struct(slide, i):
    return _find_named_text(slide, "Structure%d" % i)


# --------------------------------------------------------------------------
# 텍스트 교체 (서식 보존, 멀티라인)
# --------------------------------------------------------------------------
def _set_para_text(p_elem, text):
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
        r = p_elem.makeelement(qn("a:r"), {})
        end_rpr = p_elem.find(qn("a:endParaRPr"))
        if end_rpr is not None:
            rpr = copy.deepcopy(end_rpr)
            rpr.tag = qn("a:rPr")
            r.append(rpr)
        t = p_elem.makeelement(qn("a:t"), {})
        t.text = text
        r.append(t)
        ppr = p_elem.find(qn("a:pPr"))
        if ppr is not None:
            ppr.addnext(r)
        else:
            p_elem.insert(0, r)


def set_text(shape, text):
    """도형 텍스트 교체. '\n'은 문단 분리. 위치/폰트 서식은 건드리지 않음."""
    tf = shape.text_frame
    txbody = tf._txBody
    paras = txbody.findall(qn("a:p"))
    if not paras:
        tf.text = text
        return
    p0 = paras[0]
    lines = text.split("\n") if text != "" else [""]
    _set_para_text(p0, lines[0])
    for p in paras[1:]:
        txbody.remove(p)
    anchor = p0
    for line in lines[1:]:
        newp = copy.deepcopy(p0)
        _set_para_text(newp, line)
        anchor.addnext(newp)
        anchor = newp


# --------------------------------------------------------------------------
# 슬라이드 복제 (think-cell OLE 제외)
# --------------------------------------------------------------------------
def clone_slide(prs, src_slide):
    layout = src_slide.slide_layout
    dst = prs.slides.add_slide(layout)
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    sptree = dst.shapes._spTree
    for sh in src_slide.shapes:
        if _is_ole(sh):
            continue
        sptree.append(copy.deepcopy(sh._element))
    return dst


def _remove_leading_slides(prs, count):
    sld_id_lst = prs.slides._sldIdLst
    for sid in list(sld_id_lst)[:count]:
        sld_id_lst.remove(sid)


# --------------------------------------------------------------------------
# 슬라이드 채우기
# spec = {head, subtitle, sections:[{sub, kq, hyp, struct}, ...]}  (1~3개)
# --------------------------------------------------------------------------
def fill_slide(slide, spec):
    n = len(spec["sections"])
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
        sec = spec["sections"][i - 1]
        _do(find_sub(slide, i), "sub%d" % i, sec.get("sub", ""))
        _do(find_kq(slide, i), "KQ%d" % i, sec.get("kq", ""))
        _do(find_hyp(slide, i), "Hypothesis%d" % i, sec.get("hyp", ""))
        _do(find_struct(slide, i), "Structure%d" % i, sec.get("struct", ""))
    return filled, missing


# --------------------------------------------------------------------------
# 덱 빌드 (4단계)
# --------------------------------------------------------------------------
def build_deck(slides_spec, out_path, template_path=TEMPLATE_PATH):
    prs = Presentation(template_path)
    template_slides = list(prs.slides)  # [1칸, 2칸, 3칸]
    report = []
    for idx, spec in enumerate(slides_spec, 1):
        n = len(spec["sections"])
        if n not in TEMPLATE_SLIDE_BY_NSUB:
            raise ValueError("슬라이드 %d: 칸 개수는 1~3개여야 합니다 (현재 %d)" % (idx, n))
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
def _role_map(slide, n):
    m = {"head": find_head(slide), "subtitle": find_subtitle(slide)}
    for i in range(1, n + 1):
        m["sub%d" % i] = find_sub(slide, i)
        m["KQ%d" % i] = find_kq(slide, i)
        m["Hypothesis%d" % i] = find_hyp(slide, i)
        m["Structure%d" % i] = find_struct(slide, i)
    return m


def _expected(spec):
    e = {"head": spec.get("head", ""), "subtitle": spec.get("subtitle", "")}
    for i, sec in enumerate(spec["sections"], 1):
        e["sub%d" % i] = sec.get("sub", "")
        e["KQ%d" % i] = sec.get("kq", "")
        e["Hypothesis%d" % i] = sec.get("hyp", "")
        e["Structure%d" % i] = sec.get("struct", "")
    return e


def _sig(shape):
    geo = (shape.left, shape.top, shape.width, shape.height)
    font = None
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            f = r.font
            font = (f.name, (f.size.pt if f.size else None), f.bold)
            break
        break
    return geo, font


def qa_report(output_path, slides_spec, template_path=TEMPLATE_PATH):
    """(1) 컨펌 내용 완전성 (2) 위치·폰트·크기 무결성 (3) 오버플로우 경고."""
    tmpl = Presentation(template_path)
    tmpl_slides = list(tmpl.slides)
    out = Presentation(output_path)
    out_slides = list(out.slides)
    issues, warnings = [], []

    if len(out_slides) != len(slides_spec):
        issues.append("출력 슬라이드 수(%d)가 컨펌 수(%d)와 다릅니다." % (len(out_slides), len(slides_spec)))

    for idx, (spec, oslide) in enumerate(zip(slides_spec, out_slides), 1):
        n = len(spec["sections"])
        tslide = tmpl_slides[TEMPLATE_SLIDE_BY_NSUB[n]]
        tmap, omap = _role_map(tslide, n), _role_map(oslide, n)
        exp = _expected(spec)

        # (2) 포맷 무결성
        for role, osh in omap.items():
            tsh = tmap.get(role)
            if tsh is None or osh is None:
                issues.append("슬라이드 %d: 역할 '%s' 박스를 찾지 못함." % (idx, role))
                continue
            tg, tf = _sig(tsh)
            og, of = _sig(osh)
            if tg != og:
                issues.append("슬라이드 %d: '%s' 위치/크기 변경됨." % (idx, role))
            if tf != of:
                issues.append("슬라이드 %d: '%s' 폰트 변경됨." % (idx, role))

        # (1) 내용 완전성
        for role, extext in exp.items():
            if not extext.strip():
                continue
            osh = omap.get(role)
            if osh is None:
                issues.append("슬라이드 %d: '%s' 박스 없음." % (idx, role))
                continue
            got = "\n".join(p.text for p in osh.text_frame.paragraphs).replace("•", " ")
            exp_lines = [l.replace("•", "").strip() for l in extext.split("\n") if l.strip()]
            if not all(el in got for el in exp_lines):
                issues.append("슬라이드 %d: '%s' 내용 누락/불일치." % (idx, role))

        # (3) 오버플로우 (KQ/Hypothesis/Structure)
        for i in range(1, n + 1):
            sec = spec["sections"][i - 1]
            for role, key in (("KQ%d" % i, "kq"), ("Hypothesis%d" % i, "hyp"), ("Structure%d" % i, "struct")):
                sh = omap.get(role)
                if sh is None:
                    continue
                w = _overflow_warn(sh, sec.get(key, ""))
                if w:
                    warnings.append("슬라이드 %d: %s %s" % (idx, role, w))

    return {"ok": len(issues) == 0, "issues": issues, "warnings": warnings}


def _overflow_warn(shape, text):
    try:
        size_pt = None
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size:
                    size_pt = r.font.size.pt
                break
            if size_pt:
                break
        size_pt = size_pt or 14
        w_cm, h_cm = Emu(shape.width).cm, Emu(shape.height).cm
        char_w, line_h = size_pt * 0.032, size_pt * 0.045
        chars_per_line = max(1, int(w_cm / char_w))
        max_lines = max(1, int(h_cm / line_h))
        used = sum(max(1, (len(l) + chars_per_line - 1) // chars_per_line)
                   for l in text.split("\n"))
        if used > max_lines:
            return "오버플로우 가능 (예상 %d줄 > 박스 %d줄)" % (used, max_lines)
    except Exception:
        return None
    return None
