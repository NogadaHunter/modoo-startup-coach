#!/usr/bin/env python3
"""모두의 창업 코칭 리포트(DOCX) 생성기.

사용법:
  python3 build_report.py report.json -o 결과.docx
  python3 build_report.py --verify 결과.docx
  python3 build_report.py --example > example.json

JSON 형식은 references/report-format.md 참고.
"""
import argparse
import json
import re
import sys
from datetime import date

try:
    import docx
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    sys.stderr.write("python-docx가 없습니다. 먼저 실행: pip install python-docx\n")
    sys.exit(1)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

FONT = "맑은 고딕"
PAGE_W = 17.0  # A4 21cm - 여백 2cm x 2

C = {
    "primary": "1E4E8C",
    "primary_light": "EAF1FA",
    "accent": "F08A24",
    "text": "222222",
    "muted": "6B7280",
    "line": "D5DCE6",
    "good": "2E7D32",
    "good_bg": "E8F5E9",
    "warn": "B26A00",
    "warn_bg": "FFF6E0",
    "bad": "C62828",
    "bad_bg": "FDECEA",
    "info_bg": "EEF3FA",
    "gray_bg": "F5F6F8",
}

GRADE_COLOR = {"A": "2E7D32", "B": "1E5BB8", "C": "E07B00", "D": "C62828"}

DOC_TYPES = {
    "review": ("도전신청서 피드백 리포트", "1E4E8C"),
    "compare": ("고친 원고 다시 보기 리포트", "1E4E8C"),
    "plan_review": ("사업계획서 피드백 리포트", "22577A"),
    "ir_review": ("오디션 발표자료 피드백 리포트", "5B3F8C"),
    "blueprint": ("도전신청서 작성 설계도", "2E6B4F"),
    "interview": ("아이디어 코칭 인터뷰 질문지", "2E6B4F"),
    "eligibility": ("신청 자격·트랙 진단표", "4A5568"),
}

STATUS_FILL = [
    (("제외 위험", "미반영", "불일치", "고칠 것", "누락", "위험", "보완 필요"), C["bad_bg"], C["bad"]),
    (("확인 필요", "일부 반영", "진행", "부분", "주의", "예정"), C["warn_bg"], C["warn"]),
    (("통과", "반영", "일치", "완료", "확인", "적정", "있음"), C["good_bg"], C["good"]),
]

DISCLAIMERS = [
    "이 리포트의 점수와 등급은 공고문과 멘토기관 공개 평가지표를 바탕으로 한 **코칭용 추정**이며, 공식 평가 결과나 선정 가능성을 뜻하지 않습니다.",
    "공고문(중소벤처기업부 공고 제2026-511호)에 따라 사업계획서 등을 타인이 대신 작성해 제출하면 작성자와 도전자 모두 처벌될 수 있습니다. 리포트의 고쳐 쓰기 예시는 참고용입니다. 최종 제출문은 신청자 본인의 경험과 표현으로 직접 작성해 주세요.",
    "신청서에 쓰는 실적·수치·협력 관계는 사실만 적고, 확인되지 않은 내용은 '목표'나 '계획'으로 표현해 주세요.",
    "문항·분량·자격·일정은 차수와 연도에 따라 바뀔 수 있습니다. 최종 확인은 모두의 창업 누리집(www.modoo.or.kr)의 공고문과 신청 화면, 대표번호 1357(5번)로 해 주세요.",
]


# ---------- 낮은 수준 도우미 ----------

def rgb(hexstr):
    return RGBColor.from_string(hexstr.upper())


def set_cell_fill(cell, hexstr):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexstr)
    tcPr.append(shd)


def set_cell_margins(cell, top=80, bottom=80, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("bottom", bottom), ("start", left), ("end", right), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def set_table_borders(table, color=None, size=4, inside=True):
    color = color or C["line"]
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    for edge in edges:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblPr.append(borders)


def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblPr.append(borders)


def set_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")) + tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tw.set(qn("w:type"), "dxa")
    tblPr.append(tw)
    grid = table._tbl.tblGrid
    cols = grid.findall(qn("w:gridCol"))
    for i, gc in enumerate(cols):
        if i < len(widths_cm):
            gc.set(qn("w:w"), str(int(widths_cm[i] * 567)))
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def keep_row_together(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def font_run(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(k), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color:
        run.font.color.rgb = rgb(color)
    return run


def add_rich(paragraph, text, size=None, color=None, bold_color=None):
    """**굵게** 표시를 굵은 글씨로 넣는다."""
    text = "" if text is None else str(text)
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            font_run(paragraph.add_run(part[2:-2]), size=size, bold=True, color=bold_color or color)
        else:
            lines = part.split("\n")
            for i, line in enumerate(lines):
                if i:
                    paragraph.add_run().add_break()
                if line:
                    font_run(paragraph.add_run(line), size=size, color=color)
    return paragraph


def para_format(p, before=0, after=4, line=1.25, align=None, keep_next=False):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if align:
        p.alignment = align
    if keep_next:
        pf.keep_with_next = True
    return p


def cell_text(cell, text, size=9.5, bold=False, color=None, align=None, fill=None, keep_next=False):
    cell.text = ""
    p = cell.paragraphs[0]
    para_format(p, after=0, line=1.2, align=align, keep_next=keep_next)
    if bold:
        font_run(p.add_run(str(text)), size=size, bold=True, color=color)
    else:
        add_rich(p, text, size=size, color=color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if fill:
        set_cell_fill(cell, fill)
    set_cell_margins(cell)
    return cell


def cell_bullets(cell, items, size=9.5, color=None, marker="•"):
    cell.text = ""
    first = True
    for it in items:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        para_format(p, after=1, line=1.2)
        p.paragraph_format.left_indent = Cm(0.35)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        if marker:
            font_run(p.add_run(f"{marker} "), size=size, color=color or C["muted"])
        else:
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.first_line_indent = Cm(0)
        add_rich(p, it, size=size)
    set_cell_margins(cell)


def status_style(value):
    if not isinstance(value, str):
        return None, None
    v = value
    if len(v.strip()) > 12:  # 긴 글은 판정 칸이 아니다
        return None, None
    for words, fill, color in STATUS_FILL:
        for w in words:
            if v.strip().startswith(w) or v.strip() == w:
                return fill, color
    return None, None


# ---------- 문서 구성 ----------

class Report:
    def __init__(self, data):
        self.data = data
        self.meta = data.get("meta", {})
        self.doc_type = data.get("doc_type", "review")
        default_title, band = DOC_TYPES.get(self.doc_type, DOC_TYPES["review"])
        self.title = self.meta.get("title") or default_title
        self.band = band
        self.doc = Document()
        self._setup()

    def _setup(self):
        sec = self.doc.sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.orientation = WD_ORIENT.PORTRAIT
        for side in ("left_margin", "right_margin"):
            setattr(sec, side, Cm(2.0))
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(1.8)
        st = self.doc.styles["Normal"]
        st.font.name = FONT
        st.font.size = Pt(10.5)
        rPr = st.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rFonts.set(qn(k), FONT)
        st.font.color.rgb = rgb(C["text"])
        # 표지는 머리글·바닥글 없이
        sec.different_first_page_header_footer = True
        self._footer(sec)

    def _footer(self, sec):
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        idea = self.meta.get("idea") or ""
        label = f"{idea} · {self.title}" if idea else self.title
        font_run(p.add_run(f"{label}   |   "), size=8, color=C["muted"])
        run = p.add_run()
        font_run(run, size=8, color=C["muted"])
        for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
            if kind:
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), kind)
                run._r.append(fc)
            else:
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = txt
                run._r.append(it)
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        font_run(hp.add_run("모두의 창업 코칭 리포트"), size=8, color=C["muted"])

    # ---- 기본 요소
    def p(self, text, size=10.5, color=None, after=4, align=None, keep_next=False):
        para = self.doc.add_paragraph()
        para_format(para, after=after, align=align, keep_next=keep_next)
        add_rich(para, text, size=size, color=color)
        return para

    def spacer(self, pt=4, keep_next=False):
        para = self.doc.add_paragraph()
        para_format(para, after=0, line=1.0, keep_next=keep_next)
        font_run(para.add_run(" "), size=pt)

    def h1(self, text):
        t = self.doc.add_table(rows=1, cols=1)
        no_borders(t)
        c = t.cell(0, 0)
        cell_text(c, text, size=14, bold=True, color="FFFFFF", fill=self.band, keep_next=True)
        set_cell_margins(c, top=120, bottom=120, left=200, right=200)
        set_widths(t, [PAGE_W])
        self.spacer(6, keep_next=True)

    def h2(self, text):
        para = self.doc.add_paragraph()
        para_format(para, before=8, after=4, keep_next=True)
        font_run(para.add_run("■ "), size=11, color=self.band)
        font_run(para.add_run(text), size=11.5, bold=True, color=C["primary"])
        return para

    def bullets(self, items, numbered=False, size=10.5):
        for i, it in enumerate(items, 1):
            para = self.doc.add_paragraph()
            para_format(para, after=2)
            para.paragraph_format.left_indent = Cm(0.6)
            para.paragraph_format.first_line_indent = Cm(-0.45)
            mark = f"{i}. " if numbered else "• "
            font_run(para.add_run(mark), size=size, bold=numbered, color=self.band if numbered else C["muted"])
            add_rich(para, it, size=size)

    def table(self, headers, rows, widths=None, size=9.5, header_fill=None):
        ncol = len(headers) if headers else (len(rows[0]) if rows else 1)
        t = self.doc.add_table(rows=0, cols=ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t)
        if headers:
            r = t.add_row()
            repeat_header(r)
            for i, h in enumerate(headers):
                cell_text(r.cells[i], h, size=size, bold=True, color="FFFFFF", fill=header_fill or self.band,
                          align=WD_ALIGN_PARAGRAPH.CENTER, keep_next=True)
        for ri, row in enumerate(rows):
            r = t.add_row()
            keep_row_together(r)
            for i in range(ncol):
                val = row[i] if i < len(row) else ""
                fill, color = status_style(val)
                if fill:
                    cell_text(r.cells[i], val, size=size, bold=True, color=color, fill=fill,
                              align=WD_ALIGN_PARAGRAPH.CENTER)
                elif isinstance(val, list):
                    cell_bullets(r.cells[i], val, size=size)
                else:
                    cell_text(r.cells[i], val, size=size, fill=C["gray_bg"] if (ri % 2 == 1) else None)
        if not widths:
            widths = [PAGE_W / ncol] * ncol
        else:
            s = sum(widths)
            widths = [w * PAGE_W / s for w in widths]
        set_widths(t, widths)
        self.spacer(4)
        return t

    def callout(self, tone, title, text=None, items=None):
        fills = {"info": (C["info_bg"], C["primary"]), "good": (C["good_bg"], C["good"]),
                 "warn": (C["warn_bg"], C["warn"]), "bad": (C["bad_bg"], C["bad"])}
        fill, color = fills.get(tone, fills["info"])
        t = self.doc.add_table(rows=1, cols=1)
        no_borders(t)
        c = t.cell(0, 0)
        set_cell_fill(c, fill)
        set_cell_margins(c, top=120, bottom=120, left=200, right=200)
        c.text = ""
        p0 = c.paragraphs[0]
        para_format(p0, after=2)
        if title:
            font_run(p0.add_run(title), size=10.5, bold=True, color=color)
        first_used = bool(title)
        if text:
            para = c.add_paragraph() if first_used else p0
            first_used = True
            para_format(para, after=2, line=1.3)
            add_rich(para, text, size=10)
        for it in items or []:
            para = c.add_paragraph() if first_used else p0
            first_used = True
            para_format(para, after=1, line=1.25)
            para.paragraph_format.left_indent = Cm(0.35)
            para.paragraph_format.first_line_indent = Cm(-0.35)
            font_run(para.add_run("• "), size=10, color=color)
            add_rich(para, it, size=10)
        # 왼쪽 굵은 선
        tcPr = c._tc.get_or_add_tcPr()
        b = OxmlElement("w:tcBorders")
        el = OxmlElement("w:left")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "24")
        el.set(qn("w:color"), color)
        b.append(el)
        tcPr.append(b)
        set_widths(t, [PAGE_W])
        self.spacer(5)

    def bar(self, paragraph, score, mx, cells=10, size=11):
        try:
            ratio = max(0.0, min(1.0, float(score) / float(mx)))
        except (TypeError, ValueError, ZeroDivisionError):
            ratio = 0
        filled = int(round(ratio * cells))
        color = C["good"] if ratio >= 0.8 else (self.band if ratio >= 0.6 else (C["accent"] if ratio >= 0.4 else C["bad"]))
        font_run(paragraph.add_run("■" * filled), size=size, color=color)
        font_run(paragraph.add_run("■" * (cells - filled)), size=size, color="DADFE6")

    def scores(self, items, title_cols=("관점", "점수", "", "근거")):
        t = self.doc.add_table(rows=0, cols=4)
        set_table_borders(t)
        r = t.add_row()
        repeat_header(r)
        for i, h in enumerate(title_cols):
            cell_text(r.cells[i], h, size=9.5, bold=True, color="FFFFFF", fill=self.band, align=WD_ALIGN_PARAGRAPH.CENTER)
        for it in items:
            r = t.add_row()
            keep_row_together(r)
            name = it.get("name", "")
            if it.get("weight"):
                name = f"{name}  ({it['weight']}%)"
            cell_text(r.cells[0], name, size=9.5, bold=True)
            c = r.cells[1]
            c.text = ""
            p = c.paragraphs[0]
            para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
            self.bar(p, it.get("score", 0), it.get("max", 5), cells=it.get("max", 5) if it.get("max", 5) <= 10 else 10, size=12)
            set_cell_margins(c)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_text(r.cells[2], f"{it.get('score', '-')}/{it.get('max', 5)}", size=10, bold=True,
                      color=C["primary"], align=WD_ALIGN_PARAGRAPH.CENTER)
            cell_text(r.cells[3], it.get("note", ""), size=9)
        set_widths(t, [4.4, 2.6, 1.4, 8.6])
        self.spacer(4)

    # ---- 표지
    def cover(self):
        m = self.meta
        t = self.doc.add_table(rows=1, cols=1)
        no_borders(t)
        c = t.cell(0, 0)
        set_cell_fill(c, self.band)
        set_cell_margins(c, top=380, bottom=380, left=400, right=400)
        c.text = ""
        p = c.paragraphs[0]
        para_format(p, after=6)
        font_run(p.add_run("모두의 창업 프로젝트 코칭"), size=11, color="DCE6F5")
        p2 = c.add_paragraph()
        para_format(p2, after=4, line=1.1)
        font_run(p2.add_run(self.title), size=24, bold=True, color="FFFFFF")
        if m.get("idea"):
            p3 = c.add_paragraph()
            para_format(p3, before=10, after=0, line=1.2)
            font_run(p3.add_run(m["idea"]), size=16, bold=True, color="FFE2C2")
        set_widths(t, [PAGE_W])
        self.spacer(12)

        rows = [
            ("아이디어", m.get("idea")),
            ("신청자", m.get("applicant")),
            ("트랙", m.get("track")),
            ("단계", m.get("stage")),
            ("멘토기관·운영기관", m.get("mentor_org")),
            ("컨설턴트", " · ".join(x for x in (m.get("consultant"), m.get("consultant_org")) if x)),
            ("작성일", m.get("date") or date.today().isoformat()),
            ("버전", m.get("version") or "v1"),
        ]
        it = self.doc.add_table(rows=0, cols=2)
        set_table_borders(it, inside=True)
        for k, v in rows:
            r = it.add_row()
            cell_text(r.cells[0], k, size=10, bold=True, color=C["primary"], fill=C["primary_light"])
            cell_text(r.cells[1], v if v else "-", size=10)
        set_widths(it, [4.2, PAGE_W - 4.2])
        self.spacer(12)
        self.toc(self.data.get("sections", []), bool(self.data.get("summary")))

        note = {
            "interview": "이 질문지는 신청자가 자기 경험을 정리하도록 돕는 자료입니다. 답은 신청자 본인이 자기 말로 적어 주세요.",
            "blueprint": "이 설계도는 문항별로 무엇을 어떤 순서로 쓸지 정리한 뼈대입니다. 본문은 신청자 본인이 직접 써 주세요.",
        }.get(self.doc_type, "이 리포트는 신청자가 스스로 신청서를 고쳐 쓰도록 돕는 코칭 자료입니다. 점수는 코칭용 추정이며 공식 평가 결과가 아닙니다.")
        self.callout("info", "이 문서에 대하여", note)
        para = self.doc.add_paragraph()
        para.add_run().add_break(WD_BREAK.PAGE)

    def toc(self, sections, has_summary):
        self.h2("목차")
        items = []
        n = 0
        if has_summary:
            items.append("한눈에 보기")
        for s in sections:
            n += 1
            items.append(f"{n}. {s.get('title', '')}")
        items.append("유의사항")
        half = (len(items) + 1) // 2
        cols = [items[:half], items[half:]]
        t = self.doc.add_table(rows=1, cols=2)
        no_borders(t)
        for ci, col in enumerate(cols):
            c = t.cell(0, ci)
            c.text = ""
            for k, it in enumerate(col):
                para = c.paragraphs[0] if k == 0 else c.add_paragraph()
                para_format(para, after=1, line=1.15)
                font_run(para.add_run(it), size=10, color=C["text"])
            set_cell_margins(c, top=40, bottom=40, left=220, right=110)
        set_widths(t, [PAGE_W / 2, PAGE_W / 2])
        self.spacer(8)

    # ---- 한눈에 보기
    def summary(self, s):
        self.h1("한눈에 보기")
        grade = str(s.get("grade", "")).strip().upper()[:1]
        total = s.get("total")
        t = self.doc.add_table(rows=1, cols=2)
        no_borders(t)
        g = t.cell(0, 0)
        set_cell_fill(g, GRADE_COLOR.get(grade, self.band))
        g.text = ""
        gp = g.paragraphs[0]
        para_format(gp, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        font_run(gp.add_run(grade or "-"), size=40, bold=True, color="FFFFFF")
        gp2 = g.add_paragraph()
        para_format(gp2, after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        font_run(gp2.add_run("코칭 등급"), size=9, color="FFFFFF")
        g.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(g, top=160, bottom=160)
        v = t.cell(0, 1)
        set_cell_fill(v, C["primary_light"])
        v.text = ""
        vp = v.paragraphs[0]
        para_format(vp, after=2)
        if total is not None:
            font_run(vp.add_run(f"{total}"), size=26, bold=True, color=C["primary"])
            font_run(vp.add_run(" / 100점"), size=12, color=C["muted"])
            font_run(vp.add_run("   (코칭용 추정 점수)"), size=8.5, color=C["muted"])
        if s.get("verdict"):
            vp2 = v.add_paragraph()
            para_format(vp2, after=0, line=1.3)
            add_rich(vp2, s["verdict"], size=10)
        set_cell_margins(v, top=160, bottom=160, left=240, right=200)
        v.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_widths(t, [3.6, PAGE_W - 3.6])
        self.spacer(8)

        if s.get("scores"):
            self.h2("관점별 점수")
            self.scores(s["scores"])
        if s.get("org_scores"):
            self.h2("멘토기관 평가지표로 다시 보면")
            self.scores(s["org_scores"], title_cols=("기관 지표", "점수", "", "근거"))

        if s.get("strengths") or s.get("top_fixes"):
            t2 = self.doc.add_table(rows=2, cols=2)
            set_table_borders(t2)
            cell_text(t2.cell(0, 0), "잘된 점", size=10.5, bold=True, color=C["good"], fill=C["good_bg"])
            cell_text(t2.cell(0, 1), "가장 먼저 고칠 점", size=10.5, bold=True, color=C["bad"], fill=C["bad_bg"])
            cell_bullets(t2.cell(1, 0), s.get("strengths", []), size=9.5, color=C["good"], marker="✓")
            cell_bullets(t2.cell(1, 1), [f"{x}" for x in s.get("top_fixes", [])], size=9.5, color=C["bad"], marker="▶")
            set_widths(t2, [PAGE_W / 2, PAGE_W / 2])
            self.spacer(6)
        if s.get("red_flags"):
            self.callout("bad", "감점 신호 (먼저 정리할 것)", items=s["red_flags"])
        self.spacer(6)

    # ---- 블록
    def block(self, b):
        tp = b.get("type")
        if tp == "p":
            self.p(b.get("text", ""))
        elif tp == "h":
            self.h2(b.get("text", ""))
        elif tp in ("bullets", "numbers"):
            self.bullets(b.get("items", []), numbered=(tp == "numbers"))
        elif tp == "table":
            self.table(b.get("headers", []), b.get("rows", []), b.get("widths"))
        elif tp == "callout":
            self.callout(b.get("tone", "info"), b.get("title"), b.get("text"), b.get("items"))
        elif tp == "scores":
            self.scores(b.get("items", []))
        elif tp == "question":
            self.question(b)
        elif tp == "rewrite":
            self.rewrite(b)
        elif tp == "blueprint":
            self.blueprint(b)
        elif tp == "interview":
            self.interview(b)
        elif tp == "qa":
            self.qa(b.get("items", []))
        elif tp == "checklist":
            self.checklist(b.get("items", []))
        elif tp == "compare":
            self.compare(b.get("items", []))
        elif tp == "slide":
            self.slide(b)
        elif tp == "pagebreak":
            para = self.doc.add_paragraph()
            para.add_run().add_break(WD_BREAK.PAGE)
        else:
            self.p(f"[알 수 없는 블록: {tp}]", color=C["bad"])

    def card_head(self, left, right_parts=None, score=None, mx=5):
        t = self.doc.add_table(rows=1, cols=2)
        no_borders(t)
        a, b = t.cell(0, 0), t.cell(0, 1)
        for c in (a, b):
            set_cell_fill(c, C["primary_light"])
        cell_text(a, left, size=11.5, bold=True, color=C["primary"], fill=C["primary_light"], keep_next=True)
        b.text = ""
        p = b.paragraphs[0]
        para_format(p, after=0, align=WD_ALIGN_PARAGRAPH.RIGHT, keep_next=True)
        for part in right_parts or []:
            font_run(p.add_run(part + "   "), size=9, color=C["muted"])
        if score is not None:
            self.bar(p, score, mx, cells=mx if mx <= 10 else 10, size=11)
            font_run(p.add_run(f"  {score}/{mx}"), size=10, bold=True, color=C["primary"])
        b.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(b)
        left_w = 9.0 if (right_parts or score is not None) else PAGE_W - 1.0
        set_widths(t, [left_w, PAGE_W - left_w])
        # 위쪽 굵은 선
        for c in (a, b):
            tcPr = c._tc.get_or_add_tcPr()
            bd = OxmlElement("w:tcBorders")
            el = OxmlElement("w:top")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "18")
            el.set(qn("w:color"), self.band)
            bd.append(el)
            tcPr.append(bd)
        self.spacer(3, keep_next=True)

    def labeled(self, label, color, items, numbered=False):
        if not items:
            return
        para = self.doc.add_paragraph()
        para_format(para, before=2, after=2, keep_next=True)
        font_run(para.add_run(label), size=10, bold=True, color=color)
        self.bullets(items, numbered=numbered, size=10)

    def question(self, b):
        parts = []
        chars, limit = b.get("chars"), b.get("limit")
        if chars is not None:
            if limit:
                over = " 초과!" if chars > limit else ""
                parts.append(f"{chars:,} / {limit:,}자{over}")
            else:
                parts.append(f"{chars:,}자")
        title = " ".join(x for x in (b.get("no"), b.get("name")) if x)
        self.card_head(title, parts, b.get("score"), b.get("max", 5))
        if b.get("looking_for"):
            self.p(f"평가자가 찾는 것 : {b['looking_for']}", size=9.5, color=C["muted"], after=3, keep_next=True)
        if b.get("quote"):
            t = self.doc.add_table(rows=1, cols=1)
            no_borders(t)
            c = t.cell(0, 0)
            cell_text(c, f"“{b['quote']}”", size=9.5, color="444444", fill=C["gray_bg"], keep_next=True)
            set_widths(t, [PAGE_W])
            self.spacer(3, keep_next=True)
        self.labeled("잘된 점", C["good"], b.get("good"))
        self.labeled("고칠 점 (중요한 순서)", C["bad"], b.get("fix"), numbered=True)
        if b.get("rewrite"):
            self.rewrite(b["rewrite"])
        self.labeled("신청자에게 물어볼 것", C["primary"], b.get("ask"))
        self.spacer(8)

    def rewrite(self, rw):
        t = self.doc.add_table(rows=2, cols=2)
        set_table_borders(t)
        cell_text(t.cell(0, 0), "지금 문장", size=9.5, bold=True, color=C["muted"], fill=C["gray_bg"],
                  align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(t.cell(0, 1), "이렇게 고쳐 보면 (참고용 · 본인 말로 다시 쓰기)", size=9.5, bold=True,
                  color=C["good"], fill=C["good_bg"], align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(t.cell(1, 0), rw.get("before", ""), size=9.5, color="555555")
        cell_text(t.cell(1, 1), rw.get("after", ""), size=9.5)
        set_widths(t, [PAGE_W / 2, PAGE_W / 2])
        if rw.get("why"):
            self.p(f"왜 : {rw['why']}", size=9, color=C["muted"], after=4)
        else:
            self.spacer(3)

    def blueprint(self, b):
        title = " ".join(x for x in (b.get("no"), b.get("name")) if x)
        self.card_head(title, [b.get("length", "")] if b.get("length") else None)
        rows = []
        if b.get("looking_for"):
            rows.append(("평가자가 찾는 것", b["looking_for"]))
        if b.get("message"):
            rows.append(("핵심 메시지", f"**{b['message']}**"))
        if b.get("facts"):
            rows.append(("들어갈 사실", b["facts"]))
        if b.get("outline"):
            lines = []
            for o in b["outline"]:
                if isinstance(o, dict):
                    s = f"{o.get('step', '')} {o.get('what', '')}"
                    if o.get("chars"):
                        s += f"  ({o['chars']})"
                    lines.append(s.strip())
                else:
                    lines.append(str(o))
            rows.append(("문단 순서", lines))
        if b.get("blanks"):
            rows.append(("신청자가 채울 칸", b["blanks"]))
        if b.get("evidence"):
            rows.append(("확인할 근거", b["evidence"]))
        if b.get("avoid"):
            rows.append(("피할 표현", b["avoid"]))
        t = self.doc.add_table(rows=0, cols=2)
        set_table_borders(t)
        for k, v in rows:
            r = t.add_row()
            keep_row_together(r)
            cell_text(r.cells[0], k, size=9.5, bold=True, color=C["primary"], fill=C["primary_light"])
            if isinstance(v, list):
                numbered = k == "문단 순서" and all(re.match(r"^[①-⑳\d]", x) for x in v)
                cell_bullets(r.cells[1], v, size=9.5, marker="" if numbered else "•")
            else:
                cell_text(r.cells[1], v, size=9.5)
        set_widths(t, [3.4, PAGE_W - 3.4])
        self.spacer(10)

    def interview(self, b):
        lines = int(b.get("lines", 4))
        n = 0
        for g in b.get("groups", []):
            self.h2(g.get("title", ""))
            for q in g.get("questions", []):
                n += 1
                para = self.doc.add_paragraph()
                para_format(para, before=4, after=2, keep_next=True)
                font_run(para.add_run(f"Q{n}. "), size=10.5, bold=True, color=self.band)
                add_rich(para, q, size=10.5)
                t = self.doc.add_table(rows=1, cols=1)
                set_table_borders(t, color="C9D2DE")
                c = t.cell(0, 0)
                c.text = ""
                for i in range(lines):
                    pp = c.paragraphs[0] if i == 0 else c.add_paragraph()
                    para_format(pp, after=0, line=1.5)
                    font_run(pp.add_run(" "), size=10)
                set_widths(t, [PAGE_W])
                self.spacer(2)

    def qa(self, items):
        for i, it in enumerate(items, 1):
            t = self.doc.add_table(rows=3, cols=2)
            set_table_borders(t)
            a = t.cell(0, 0).merge(t.cell(0, 1))
            cell_text(a, f"예상 질문 {i}. {it.get('q', '')}", size=10, bold=True, color=C["primary"],
                      fill=C["primary_light"], keep_next=True)
            cell_text(t.cell(1, 0), "답변 방향", size=9, bold=True, color=C["muted"], fill=C["gray_bg"], keep_next=True)
            cell_text(t.cell(1, 1), it.get("direction", ""), size=9.5, keep_next=True)
            cell_text(t.cell(2, 0), "준비할 근거", size=9, bold=True, color=C["muted"], fill=C["gray_bg"])
            cell_text(t.cell(2, 1), it.get("prepare", ""), size=9.5)
            for r in t.rows:
                keep_row_together(r)
            set_widths(t, [2.6, PAGE_W - 2.6])
            self.spacer(4)

    def checklist(self, items):
        rows = []
        has_status = any(isinstance(x, dict) and x.get("status") for x in items)
        for x in items:
            if isinstance(x, dict):
                rows.append(["☐" if not x.get("status") else "", x.get("text", ""), x.get("status", "")])
            else:
                rows.append(["☐", str(x), ""])
        if has_status:
            self.table(["", "확인할 것", "결과"], [[r[0] or "", r[1], r[2]] for r in rows], widths=[1, 12.5, 3.5])
        else:
            self.table(["", "확인할 것"], [[r[0], r[1]] for r in rows], widths=[1, 16])

    def compare(self, items):
        rows = []
        for it in items:
            bf, af = it.get("before"), it.get("after")
            try:
                d = float(af) - float(bf)
                delta = f"+{d:g}" if d > 0 else (f"{d:g}" if d < 0 else "0")
            except (TypeError, ValueError):
                delta = ""
            rows.append([it.get("name", ""), str(bf), str(af), delta, it.get("reason", "")])
        t = self.table(["항목", "이전", "지금", "변화", "이유"], rows, widths=[4, 1.6, 1.6, 1.6, 8.2])
        for r in t.rows[1:]:
            c = r.cells[3]
            txt = c.text.strip()
            if txt.startswith("+"):
                cell_text(c, txt, size=9.5, bold=True, color=C["good"], fill=C["good_bg"], align=WD_ALIGN_PARAGRAPH.CENTER)
            elif txt.startswith("-"):
                cell_text(c, txt, size=9.5, bold=True, color=C["bad"], fill=C["bad_bg"], align=WD_ALIGN_PARAGRAPH.CENTER)

    def slide(self, b):
        title = f"{b.get('no', '')}장  {b.get('title', '')}".strip()
        self.card_head(title)
        if b.get("message"):
            self.p(f"이 장의 한 줄 메시지 : {b['message']}", size=9.5, color=C["muted"], after=3, keep_next=True)
        self.labeled("잘된 점", C["good"], b.get("good"))
        self.labeled("고칠 점", C["bad"], b.get("fix"), numbered=True)
        if b.get("retitle"):
            self.p(f"제목 바꿔 보기 (참고용) : **{b['retitle']}**", size=10, after=4)
        self.spacer(6)

    def closing(self):
        self.h1("유의사항")
        items = list(DISCLAIMERS)
        if self.doc_type in ("interview", "blueprint"):
            items = [DISCLAIMERS[1].replace("리포트의 고쳐 쓰기 예시는 참고용입니다. ", ""), DISCLAIMERS[2], DISCLAIMERS[3]]
        items += self.data.get("notes", [])
        self.bullets(items, size=9.5)

    def build(self, out):
        self.cover()
        sections = self.data.get("sections", [])
        summary = self.data.get("summary")
        if summary:
            self.summary(summary)
        for i, s in enumerate(sections, 1):
            self.h1(f"{i}. {s.get('title', '')}")
            if s.get("intro"):
                self.p(s["intro"])
            for b in s.get("blocks", []):
                self.block(b)
            self.spacer(6)
        self.closing()
        self.doc.core_properties.title = f"{self.meta.get('idea', '')} {self.title}".strip()
        self.doc.core_properties.author = self.meta.get("consultant") or "모두의 창업 코칭"
        self.doc.save(out)
        return out


# ---------- 점검 ----------

def verify(path):
    d = Document(path)
    texts = [p.text for p in d.paragraphs]
    empty_cells = 0
    cells = 0
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                cells += 1
                texts.append(c.text)
                if not c.text.strip():
                    empty_cells += 1
    full = "\n".join(texts)
    problems = []
    if "�" in full:
        problems.append("깨진 글자(�)가 있습니다")
    if "**" in full:
        problems.append("굵게 표시(**)가 그대로 남아 있습니다")
    for bad in ("None", "null", "undefined", "{", "}"):
        if bad in full:
            problems.append(f"'{bad}' 문자열이 들어 있습니다 (빈 값 확인)")
    if "[알 수 없는 블록" in full:
        problems.append("알 수 없는 블록 종류가 있습니다")
    hangul = len(re.findall(r"[가-힣]", full))
    print(f"파일: {path}")
    print(f"문단 {len(d.paragraphs)}개 · 표 {len(d.tables)}개 · 칸 {cells}개(빈 칸 {empty_cells}개) · 한글 {hangul:,}자")
    confirm = full.count("확인 필요")
    if confirm:
        print(f"'확인 필요' 표시 {confirm}곳 (신청자에게 확인받을 것)")
    if problems:
        print("문제:")
        for p in problems:
            print(" -", p)
        return 1
    print("문제 없음")
    return 0


EXAMPLE = {
    "doc_type": "review",
    "meta": {"title": "도전신청서 피드백 리포트", "idea": "예시 아이디어", "applicant": "홍길동", "track": "일반/기술",
             "stage": "1R 도전신청서 (2026년 2차)", "mentor_org": "○○대학교 창업지원단", "consultant": "",
             "consultant_org": "", "date": "2026-10-30", "version": "v1"},
    "summary": {
        "grade": "B", "total": 72,
        "verdict": "문제는 분명하지만 **차별점의 작동 원리**와 **1R 완료 기준**이 약해요.",
        "scores": [
            {"name": "A 문제 정의·진정성", "score": 4, "max": 5, "weight": 20, "note": "Q2 현장 경험 장면이 구체적"},
            {"name": "B 차별성·혁신성", "score": 3, "max": 5, "weight": 25, "note": "대체재 비교 없음"},
            {"name": "C 효과성·가치", "score": 4, "max": 5, "weight": 20, "note": ""},
            {"name": "D 실현 가능성·구체성", "score": 3, "max": 5, "weight": 20, "note": "완료 기준 숫자 없음"},
            {"name": "E 사업성·성장성", "score": 3, "max": 5, "weight": 15, "note": ""}],
        "strengths": ["Q2 첫 문단의 현장 장면"],
        "top_fixes": ["Q3-1에 대체재 2~3개와 비교 넣기"],
        "red_flags": ["Q3-1 '정확도 95%'를 확정값으로 씀 → 목표로"]},
    "sections": [
        {"title": "자격·요건 점검", "blocks": [
            {"type": "table", "headers": ["확인 항목", "신청자 상황", "판정", "할 일"],
             "rows": [["예비/기창업", "사업자 없음", "통과", "-"], ["동시 수행", "예비창업패키지 신청 여부 모름", "확인 필요", "신청자에게 확인"]],
             "widths": [3, 6, 2.5, 5.5]}]},
        {"title": "문항별 피드백", "blocks": [
            {"type": "question", "no": "Q2", "name": "아이디어 배경", "chars": 1520, "limit": 2000, "score": 4, "max": 5,
             "looking_for": "실제 문제, 직접 겪거나 본 장면, 지금 풀어야 할 이유",
             "quote": "작년 겨울 시장에서 상인 세 분이 같은 불편을 말씀하셨다.",
             "good": ["직접 본 장면으로 시작해요"], "fix": ["통계에 출처·연도를 붙이기"],
             "rewrite": {"before": "시장 규모가 크다.", "after": "[신청자 작성: 자료 이름]에 따르면 [수치]입니다.", "why": "출처 있는 숫자"},
             "ask": ["상인 몇 분에게 더 물어볼 수 있나요?"]}]},
        {"title": "다음 코칭 숙제", "blocks": [
            {"type": "numbers", "items": ["상인 10명 인터뷰"]},
            {"type": "checklist", "items": ["자리표시를 모두 바꿨다", {"text": "출처 표기", "status": "고칠 것"}]}]}],
    "notes": []}


def main():
    ap = argparse.ArgumentParser(description="모두의 창업 코칭 리포트 DOCX 생성")
    ap.add_argument("json", nargs="?", help="report.json")
    ap.add_argument("-o", "--out", help="출력 DOCX 경로")
    ap.add_argument("--verify", metavar="DOCX", help="만든 DOCX 점검")
    ap.add_argument("--example", action="store_true", help="예시 JSON 출력")
    a = ap.parse_args()
    if a.example:
        print(json.dumps(EXAMPLE, ensure_ascii=False, indent=2))
        return 0
    if a.verify:
        return verify(a.verify)
    if not a.json:
        ap.print_help()
        return 2
    with open(a.json, encoding="utf-8-sig") as f:
        data = json.load(f)
    out = a.out or (re.sub(r"[\\/:*?\"<>|\s]+", "", data.get("meta", {}).get("idea", "리포트")) + "_코칭리포트.docx")
    Report(data).build(out)
    print(f"만들었어요: {out}")
    return verify(out)


if __name__ == "__main__":
    sys.exit(main())
