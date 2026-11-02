import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = "/home/user/test/SSC_Krishi_2027_3_Model_Tests.md"
OUT = "/home/user/test/SSC_Krishi_2027_3_Model_Tests.docx"
TEMPLATE = "/home/user/test/SSC_BGS_2027_3_Model_Tests.docx"

FONT = "Noto Sans Bengali"
C_TITLE = RGBColor(0x16, 0x3C, 0x45)
C_H1 = RGBColor(0x00, 0x7B, 0x7D)
C_GREY = RGBColor(0x52, 0x63, 0x6A)
C_DARK = RGBColor(0x1B, 0x2A, 0x2F)

IMG_RE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")


def set_cell_shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def para_shade(p, hexcolor):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    pPr.append(shd)


def add_field(paragraph, code):
    r = paragraph.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), code)
    t = OxmlElement("w:t")
    t.text = "1"
    r._r.addnext(fld)
    fld.append(t)


def style_run(run, size, bold=False, color=None, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(a), FONT)
    if not bold:
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(size * 2)))
        rPr.append(szCs)


INLINE = re.compile(r"(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`)")


def add_rich(paragraph, text, size, color=None, bold=False):
    text = text.replace("  \n", "\n")
    for chunk in INLINE.split(text):
        if not chunk:
            continue
        if chunk.startswith("**") and chunk.endswith("**") and len(chunk) > 4:
            r = paragraph.add_run(chunk[2:-2])
            style_run(r, size, bold=True, color=C_DARK)
        elif chunk.startswith("`") and chunk.endswith("`") and len(chunk) > 2:
            r = paragraph.add_run(chunk[1:-1])
            style_run(r, size - 0.5, bold=False, color=RGBColor(0x00, 0x6B, 0x6D))
        elif chunk.startswith("*") and chunk.endswith("*") and len(chunk) > 2:
            r = paragraph.add_run(chunk[1:-1])
            style_run(r, size - 0.5, italic=True, color=C_GREY)
        else:
            r = paragraph.add_run(chunk)
            style_run(r, size, bold=bold, color=color)


def main():
    doc = Document(TEMPLATE)
    body = doc.element.body
    for child in list(body):
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(1.9), Cm(1.8)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.paragraph_format.line_spacing = 1.28
    normal.paragraph_format.space_after = Pt(4)

    # header / footer
    for s in doc.sections:
        hp = s.header.paragraphs[0]
        hp.text = ""
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = hp.add_run("এসএসসি ২০২৭  •  কৃষিশিক্ষা")
        style_run(r, 9, color=C_GREY)
        fp = s.footer.paragraphs[0]
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_field(fp, "PAGE")
        for r in fp.runs:
            style_run(r, 9, color=C_GREY)

    lines = open(SRC, encoding="utf-8").read().split("\n")
    i = 0
    first_h1 = True
    while i < len(lines):
        ln = lines[i].rstrip()
        stripped = ln.strip()
        if not stripped:
            i += 1
            continue

        if stripped == "---":
            i += 1
            continue

        m = IMG_RE.match(stripped)
        if m:
            path = m.group(2)
            if not path.startswith("/"):
                path = "/home/user/test/" + path
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(path, width=Cm(15.6))
            i += 1
            continue

        if stripped.startswith("# "):
            p = doc.add_paragraph()
            if not first_h1:
                p.paragraph_format.page_break_before = True
            first_h1 = False
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(8)
            r = p.add_run(stripped[2:])
            style_run(r, 17, bold=True, color=C_H1)
            i += 1
            continue

        if stripped.startswith("## "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(stripped[3:])
            style_run(r, 13, bold=True, color=C_TITLE)
            i += 1
            continue

        if stripped.startswith("### "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(stripped[4:])
            style_run(r, 11.5, bold=True, color=C_TITLE)
            i += 1
            continue

        if stripped == "TITLE":
            i += 1
            continue

        if stripped.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = lines[i].strip()
                if not re.match(r"^\|[\s\-:|]+\|$", row):
                    cells = [c.strip() for c in row.strip("|").split("|")]
                    rows.append(cells)
                i += 1
            if not rows:
                continue
            ncol = max(len(r) for r in rows)
            t = doc.add_table(rows=0, cols=ncol)
            t.style = "Table Grid"
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            for ri, row in enumerate(rows):
                cells = t.add_row().cells
                for ci in range(ncol):
                    txt = row[ci] if ci < len(row) else ""
                    cell = cells[ci]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.1
                    add_rich(p, txt, 9.5, bold=(ri == 0))
                    if ri == 0:
                        set_cell_shade(cell, "E7F3F1")
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            continue

        if stripped.startswith("> "):
            buf = []
            while i < len(lines) and lines[i].strip().startswith("> "):
                buf.append(lines[i].strip()[2:])
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.4)
            p.paragraph_format.right_indent = Cm(0.3)
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(7)
            add_rich(p, " ".join(buf), 10, color=RGBColor(0x24, 0x3C, 0x42))
            para_shade(p, "F2F7F6")
            continue

        m = re.match(r"^- (.*)$", stripped)
        if m:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.35)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run("•  ")
            style_run(r, 10.5, color=C_H1)
            add_rich(p, m.group(1), 10.5)
            i += 1
            continue

        # paragraph (may absorb following indented lines)
        buf = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i]
            if not nxt.strip():
                break
            if nxt.strip().startswith(("#", "|", ">", "- ", "![", "---")):
                break
            buf.append(nxt.strip())
            i += 1
        text = "\n".join(buf)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if text.startswith("*") and text.endswith("*") and text.count("*") == 2:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(9)
            r = p.add_run(text.strip("*"))
            style_run(r, 9, italic=True, color=C_GREY)
        elif text.startswith("**") and re.match(r"^\*\*[০-৯]+।\*\*", text):
            add_rich(p, text, 10.5, bold=True, color=C_DARK)
        elif text.startswith("**উত্তর:"):
            add_rich(p, text, 10.5)
        else:
            add_rich(p, text, 11)

    doc.save(OUT)
    d = Document(OUT)
    print("saved", OUT, "paras", len(d.paragraphs),
          "tables", len(d.tables),
          "images", len(d.inline_shapes))
    bad = [p.text for p in d.paragraphs if "**" in p.text or "`" in p.text]
    print("stray markup:", len(bad))


if __name__ == "__main__":
    main()
