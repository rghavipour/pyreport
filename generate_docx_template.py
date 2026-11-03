"""Build the editable Word version of the weekly report (A4, 3 pages, RTL).

The layout mirrors the PDF: paper texture + botanical corners (anchored behind
text), circular section numbers, rounded badges, solid white writing boxes,
member cards with a forest accent edge, plan table and signature card.

Compatibility notes (Word, LibreOffice and Google Docs):
* Tables are written in *visual* column order (no `bidiVisual`), so every
  application places columns identically - Google Docs ignores `bidiVisual`.
* RTL text paragraphs use `w:bidi` + justified alignment, which every
  application renders as right-aligned for Persian (no left/right ambiguity).
* Decorations are anchored in the body (not the header) so Google Docs keeps
  them; the page colour is set as a fallback for the paper texture.
* Vazirmatn (Regular + Bold) is embedded inside the .docx for Word/LibreOffice;
  Google Docs uses its own copy of Vazirmatn from Google Fonts.
"""
import io
import os
import re
import shutil
import tempfile
import uuid
import zipfile

import docx
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm
from PIL import Image

import report_content as rc

os.chdir(rc.BASE_DIR)
OUT_DOCX = "weekly_report_psychology.docx"

FONT = rc.FONT_FAMILY
TEXT_W = 174          # usable width in mm (210 - 2*18)

NS = nsdecls("w", "r", "wp", "a", "pic")


def tw(mm):
    """mm -> twips (dxa)."""
    return int(round(mm * 56.6929))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ------------------------------------------------------------------ runs
def run(text, size=9, bold=False, color=rc.DARK_TEXT):
    b = "<w:b/><w:bCs/>" if bold else '<w:b w:val="0"/><w:bCs w:val="0"/>'
    hs = int(round(size * 2))
    return parse_xml(
        f'<w:r {NS}><w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
        f'{b}<w:color w:val="{color}"/><w:sz w:val="{hs}"/><w:szCs w:val="{hs}"/>'
        f'<w:rtl/><w:lang w:val="fa-IR" w:bidi="fa-IR"/>'
        f'</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    )


class Pics:
    """Creates inline picture runs (optionally converted to page anchors)."""

    def _inline(self, part, stream, w_mm, h_mm):
        stream.seek(0)
        return part.new_pic_inline(stream, Mm(w_mm), Mm(h_mm) if h_mm else None)

    def inline_run(self, part, stream, w_mm, h_mm=None, lower_pt=0):
        """lower_pt: shift the picture below the baseline (to centre it on text)."""
        inline = self._inline(part, stream, w_mm, h_mm)
        pos = f"<w:position w:val='{-int(round(lower_pt * 2))}'/>" if lower_pt else ""
        r = parse_xml(f"<w:r {NS}><w:rPr><w:noProof/>{pos}<w:rtl w:val='0'/></w:rPr><w:drawing/></w:r>")
        r.find(qn("w:drawing")).append(inline)
        return r

    def anchor_run(self, part, stream, x_mm, y_mm, w_mm, h_mm, z):
        inline = self._inline(part, stream, w_mm, h_mm)
        extent = inline.find(qn("wp:extent"))
        graphic = inline.find(qn("a:graphic"))
        anchor = parse_xml(
            f'<wp:anchor {NS} distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{z}" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(x_mm * 36000)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(y_mm * 36000)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{extent.get("cx")}" cy="{extent.get("cy")}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="1" name="Decor"/>'
            f'<wp:cNvGraphicFramePr/></wp:anchor>'
        )
        anchor.append(graphic)
        r = parse_xml(f"<w:r {NS}><w:rPr><w:noProof/><w:rtl w:val='0'/></w:rPr><w:drawing/></w:r>")
        r.find(qn("w:drawing")).append(anchor)
        return r


PICS = Pics()


# ------------------------------------------------------------------ paragraphs
def para(children=(), align="start", before=0, after=0, line=None, exact=False,
         keep_next=False, bottom_border=None, tiny=False, ltr=False):
    """RTL paragraph.  align: 'start' (right, via justify) or 'center'.
    ltr=True -> left-to-right paragraph for image-only content; align is then
    the physical 'left' / 'right' / 'center'."""
    if ltr:
        jc = align if align in ("left", "right", "center") else "right"
    else:
        jc = "center" if align == "center" else "both"
    bdr = ""
    if bottom_border:
        v, sz, c, sp = bottom_border
        bdr = f'<w:pBdr><w:bottom w:val="{v}" w:sz="{sz}" w:space="{sp}" w:color="{c}"/></w:pBdr>'
    if tiny:
        line, exact = 1.5, True
    line_xml = ""
    if line is not None:
        if exact:
            line_xml = f' w:line="{int(line * 20)}" w:lineRule="exact"'
        else:
            line_xml = f' w:line="{int(line * 240)}" w:lineRule="auto"'
    keep = "<w:keepNext/>" if keep_next else ""
    mark = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>' if tiny else ""
    bidi = '<w:bidi w:val="0"/>' if ltr else "<w:bidi/>"
    p = parse_xml(
        f'<w:p {NS}><w:pPr>{keep}{bdr}{bidi}'
        f'<w:spacing w:before="{tw(before)}" w:after="{tw(after)}"{line_xml}/>'
        f'<w:jc w:val="{jc}"/>{mark}</w:pPr></w:p>'
    )
    for c in children:
        p.append(c)
    return p


# ------------------------------------------------------------------ tables
def border_xml(tag, spec):
    if not spec:
        return f'<w:{tag} w:val="nil"/>'
    val, sz, color = spec
    return f'<w:{tag} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'


def cell(content, width_mm, fill=None, borders=None, valign="top", margins=None, span=1):
    """borders: dict with keys top/bottom/start/end (start = visual RIGHT edge).
    margins: (top, start, bottom, end) in mm."""
    return dict(content=list(content), w=width_mm, fill=fill, borders=borders or {},
                valign=valign, margins=margins, span=span)


def _mar(tag, m):
    t, s, b, e = m
    return (f'<w:{tag}><w:top w:w="{tw(t)}" w:type="dxa"/><w:left w:w="{tw(e)}" w:type="dxa"/>'
            f'<w:bottom w:w="{tw(b)}" w:type="dxa"/><w:right w:w="{tw(s)}" w:type="dxa"/></w:{tag}>')


def table(col_widths, rows, tbl_borders=None, fill=None, margins=(0, 0, 0, 0)):
    """Columns/cells are given in reading order (right -> left) and written in
    visual order (left -> right), so no application has to mirror the table.
    rows: list of (height_mm or None, rule, [cells])."""
    total = sum(col_widths)
    tb = tbl_borders or {}
    tblb = ("<w:tblBorders>" + border_xml("top", tb.get("top")) +
            border_xml("left", tb.get("end")) + border_xml("bottom", tb.get("bottom")) +
            border_xml("right", tb.get("start")) + border_xml("insideH", tb.get("insideH")) +
            border_xml("insideV", tb.get("insideV")) + "</w:tblBorders>")
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    grid = "".join(f'<w:gridCol w:w="{tw(w)}"/>' for w in reversed(col_widths))
    tbl = parse_xml(
        f'<w:tbl {NS}><w:tblPr><w:tblW w:w="{tw(total)}" w:type="dxa"/>'
        f'<w:jc w:val="center"/><w:tblInd w:w="0" w:type="dxa"/>{tblb}{shd}'
        f'<w:tblLayout w:type="fixed"/>{_mar("tblCellMar", margins)}'
        f'<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
        f'w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid></w:tbl>'
    )
    for height, rule, cells in rows:
        trh = f'<w:trHeight w:val="{tw(height)}" w:hRule="{rule}"/>' if height else ""
        tr = parse_xml(f'<w:tr {NS}><w:trPr><w:cantSplit/>{trh}</w:trPr></w:tr>')
        for c in reversed(cells):
            bd = c["borders"]
            tcb = ""
            if bd:
                tcb = ("<w:tcBorders>" + (border_xml("top", bd["top"]) if "top" in bd else "") +
                       (border_xml("left", bd["end"]) if "end" in bd else "") +
                       (border_xml("bottom", bd["bottom"]) if "bottom" in bd else "") +
                       (border_xml("right", bd["start"]) if "start" in bd else "") + "</w:tcBorders>")
            cshd = f'<w:shd w:val="clear" w:color="auto" w:fill="{c["fill"]}"/>' if c["fill"] else ""
            span = f'<w:gridSpan w:val="{c["span"]}"/>' if c["span"] > 1 else ""
            mar = _mar("tcMar", c["margins"]) if c["margins"] else ""
            tc = parse_xml(
                f'<w:tc {NS}><w:tcPr><w:tcW w:w="{tw(c["w"])}" w:type="dxa"/>{span}{tcb}{cshd}{mar}'
                f'<w:vAlign w:val="{c["valign"]}"/></w:tcPr></w:tc>'
            )
            content = c["content"] or [para(tiny=True)]
            for el in content:
                tc.append(el)
            if content[-1].tag != qn("w:p"):
                tc.append(para(tiny=True))
            tr.append(tc)
        tbl.append(tr)
    return tbl


# ------------------------------------------------------------------ building blocks
SOFT = ("single", 4, rc.BORDER_SOFT)
LINE = ("single", 4, rc.BORDER)
DOTTED = ("dotted", 6, rc.LINE_DOT)
DASHED = ("dashed", 4, rc.BORDER)
ALL = lambda s: {"top": s, "bottom": s, "start": s, "end": s}  # noqa: E731

LEAF = rc.leaf_png()


def pill_run(part, text, size=8, fg="5A7463", bg=rc.BEIGE, height=None, pad=3.0):
    stream, w, h = rc.pill_png(text, size, fg, bg, height_mm=height, pad_mm=pad)
    return PICS.inline_run(part, stream, w, h), w


def leaf_label(part, text, size=9, before=0, after=1.2, color="2D4C3E"):
    return para([PICS.inline_run(part, LEAF, 2.1, 2.1), run("  " + text, size, True, color)],
                before=before, after=after, keep_next=True)


def writing_box(width, height):
    m = (1.5, 3, 1.5, 3)
    return table([width], [(height, "atLeast", [cell([para(line=1.5)], width, fill=rc.FIELD,
                                                       margins=m)])],
                 tbl_borders=ALL(SOFT), margins=m)


def section_header(part, num_title, before=0, after=2):
    """[circle number] title ———— as a borderless 3-cell row; vAlign=center keeps
    badge, title and line on one axis.  Image cells keep ~4 mm spare width
    because some renderers inset in-cell pictures by ~3 mm."""
    num, title = num_title
    d = 6.2
    num_w = d + 4
    tw_mm = rc.text_width_mm(title, 12, 700) * 1.12 + 5
    line_w = TEXT_W - num_w - tw_mm
    badge = rc.number_badge_png(num)
    line_img = rc.fade_line_png(w=1600, h=10, canvas_h=int(1600 * d / (line_w - 4)))
    cells = [
        cell([para([PICS.inline_run(part, badge, d, d)], align="center")], num_w, valign="center"),
        cell([para([run(title, 12, True, rc.FOREST)])], tw_mm, valign="center",
             margins=(0, 2.2, 0, 0)),
        cell([para([PICS.inline_run(part, line_img, line_w - 4, d)], align="center")], line_w,
             valign="center"),
    ]
    t = table([num_w, tw_mm, line_w], [(None, "auto", cells)])
    return [para(tiny=True, before=before), t, para(tiny=True, after=after)]


def compact_header(part, subtitle, tag):
    tag_run, _ = pill_run(part, tag, 7.5, rc.FOREST_2, height=4.2)
    return [
        para([run(rc.DOC_TITLE, 13.5, True, rc.FOREST), run("   |   ", 13, False, rc.SAGE),
              run(subtitle, 13.5, True, rc.FOREST)], align="center", after=0.5, line=1.15),
        para([run(rc.INSTITUTION, 8.5, False, "728779"), run("   •   ", 8.5, False, rc.SAGE),
              tag_run], align="center", after=4, bottom_border=("single", 4, rc.BORDER, 7)),
    ]


def meta_card():
    lab_w, val_w = 34, 53
    rows = []
    for i, (a, b) in enumerate(rc.SEC1_FIELDS):
        pad_t = 3 if i == 0 else 1.5
        pad_b = 1.5 if i == 0 else 3
        cells = []
        for j, lbl in enumerate((a, b)):
            cells.append(cell([para([run(lbl, 9.5, True, rc.FOREST_2)])], lab_w,
                              valign="bottom", margins=(pad_t, 4 if j == 0 else 2, pad_b, 1)))
            cells.append(cell([para([run(" ", 9.5)], bottom_border=("dotted", 6, rc.LINE_DOT, 1))],
                              val_w, valign="bottom",
                              margins=(pad_t, 1, pad_b, 5 if j == 0 else 4)))
        rows.append((None, "auto", cells))
    return table([lab_w, val_w, lab_w, val_w], rows, tbl_borders=ALL(SOFT), fill=rc.FIELD)


def member_card(part, num, heights):
    inner = TEXT_W - 8.5
    badge_run, bw = pill_run(part, f"عضو شماره {num}", 8, "5A7463", height=4.8)
    badge_w = bw + 5
    hdr = table(
        [17, 72, inner - 17 - 72 - badge_w, badge_w],
        [(None, "auto", [
            cell([para([run(rc.MEMBER_NAME, 9.5, True, rc.FOREST)])], 17, valign="bottom"),
            cell([para([run(" ", 9.5)])], 72, valign="bottom", borders={"bottom": DOTTED}),
            cell([], inner - 17 - 72 - badge_w),
            cell([para([badge_run], align="center")], badge_w, valign="center"),
        ])])
    content = [hdr, para(tiny=True, after=2, bottom_border=("dashed", 4, rc.BORDER_SOFT, 4))]
    for i, (lbl, h) in enumerate(zip(rc.MEMBER_FIELDS, heights)):
        content.append(leaf_label(part, lbl, before=0 if i == 0 else 2))
        content.append(writing_box(inner, h))
    m = (3, 4.5, 3.5, 4)
    return table([TEXT_W], [(None, "auto", [cell(content, TEXT_W, fill=rc.CARD, margins=m)])],
                 tbl_borders={"top": LINE, "bottom": LINE, "end": LINE,
                              "start": ("single", 28, rc.FOREST)}, margins=m)


def plan_table():
    widths = [16, 82, 42, 34]
    aligns = ["center", "start", "start", "center"]
    head = [cell([para([run(h, 8.5, True, rc.FOREST)], align=a)], w, fill=rc.BEIGE,
                 valign="center", borders={"bottom": ("single", 10, "CBD5BC")})
            for h, w, a in zip(rc.PLAN_HEADERS, widths, aligns)]
    rows = [(8.5, "atLeast", head)]
    for i, n in enumerate(rc.PLAN_ROWS):
        fill = rc.FIELD if i % 2 == 0 else rc.CARD
        cells = [cell([para([run(n, 8.5, True, rc.FOREST)], align="center")], widths[0],
                      fill=fill, valign="center")]
        cells += [cell([para([run(" ", 8.5)], align=a)], w, fill=fill, valign="center")
                  for w, a in zip(widths[1:], aligns[1:])]
        rows.append((9, "atLeast", cells))
    return table(widths, rows, margins=(0, 2.5, 0, 2.5),
                 tbl_borders={**ALL(SOFT), "insideH": ("single", 4, rc.BEIGE),
                              "insideV": ("single", 4, rc.BEIGE)})


def sign_card(part):
    col = TEXT_W / 2
    inner = col - 8
    cells = []
    for s in rc.SIGN_BOXES:
        nm = (1.5, 2.5, 1, 2.5)
        notes = table([inner], [(16, "atLeast", [cell([para([run(s["notes"], 7.5, False, "8F9E94")])],
                                                       inner, fill=rc.FIELD, margins=nm)])],
                      tbl_borders=ALL(DASHED), margins=nm)
        sign_w = rc.text_width_mm(s["sign"], 7.8, 400) + 3
        top = {"top": ("dashed", 4, rc.LINE_DOT)}
        names = table([inner - sign_w, sign_w], [(9, "atLeast", [
            cell([para([run(s["name"], 7.8, False, "758679")])], inner - sign_w,
                 borders=top, margins=(1.8, 0, 0, 0)),
            cell([para([run(s["sign"], 7.8, False, "758679")])], sign_w,
                 borders=top, margins=(1.8, 0, 0, 0)),
        ])])
        content = [
            leaf_label(part, s["title"], size=9, color=rc.FOREST, after=0.8),
            para([run(s["desc"], 7.8, False, "697F71")], after=2),
            notes,
            para(tiny=True, after=2.5),
            names,
        ]
        cells.append(cell(content, col, fill=rc.CARD, margins=(3.5, 4, 2, 4)))
    return table([col, col], [(None, "auto", cells)], tbl_borders=ALL(LINE),
                 margins=(3.5, 4, 2, 4))


def footer_table(part, idx):
    right, center, badge = rc.FOOTERS[idx]
    badge_run, bw = pill_run(part, badge, 7.5, "426354", height=4.2, pad=2.8)
    top = {"top": ("single", 4, rc.BORDER_SOFT)}
    badge_w = bw + 5
    right_w = max(67, rc.text_width_mm(right, 7.5, 400) * 1.12 + 4)
    center_w = max(40, rc.text_width_mm(center, 7.5, 400) * 1.12 + 4)
    if right_w > 67:            # long text: keep the centre cell as close to centre as fits
        center_w = rc.text_width_mm(center, 7.5, 400) * 1.12 + 4
    widths = [right_w, center_w, TEXT_W - right_w - center_w - badge_w, badge_w]
    m = (2, 0, 0, 0)
    return table(widths, [(None, "auto", [
        cell([para([run(right, 7.5, False, "809687")])], widths[0], borders=top, margins=m),
        cell([para([run(center, 7.5, False, "809687")], align="center")], widths[1], borders=top,
             margins=m),
        cell([], widths[2], borders=top, margins=m),
        cell([para([badge_run], align="center")], widths[3], borders=top, margins=m),
    ])], margins=m)


# ------------------------------------------------------------------ page decor
def faded(path, opacity, max_w=700):
    img = Image.open(path).convert("RGBA")
    if img.width > max_w:
        img = img.resize((max_w, int(img.height * max_w / img.width)), Image.LANCZOS)
    a = img.getchannel("A").point(lambda v: int(v * opacity))
    img.putalpha(a)
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    buf.seek(0)
    return buf


CORNER_RATIO = 1456 / 1288


def decor_runs(part, corners):
    """Paper texture + botanical corners, anchored to the page behind text."""
    paper = rc.downscaled(rc.asset("paper.jpg"), 1240, fmt="JPEG", quality=80)
    runs = [PICS.anchor_run(part, paper, 0, 0, 210, 297, z=1)]
    for z, (name, w, op) in enumerate(corners, start=2):
        h = w * CORNER_RATIO
        x = 0 if name.endswith("l") else 210 - w
        y = 0 if name.startswith("t") else 297 - h
        runs.append(PICS.anchor_run(part, faded(rc.asset(f"corner_{name}.png"), op), x, y, w, h, z=z))
    return runs


def set_footer(section, idx):
    ftr = section.footer
    ftr.is_linked_to_previous = False
    el = ftr._element
    for child in list(el):
        el.remove(child)
    el.append(footer_table(ftr.part, idx))
    el.append(para(tiny=True))


# ------------------------------------------------------------------ styles & fonts
def setup_styles(doc):
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    if dd is not None:
        styles.remove(dd)
    styles.insert(0, parse_xml(
        f'<w:docDefaults {NS}><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
        f'<w:color w:val="{rc.DARK_TEXT}"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        f'<w:rtl/><w:lang w:val="fa-IR" w:eastAsia="en-US" w:bidi="fa-IR"/></w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr><w:bidi/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'))
    doc.styles["Normal"].font.name = FONT


def obfuscate(font_bytes, guid):
    key = bytes.fromhex(guid.strip("{}").replace("-", ""))[::-1]
    data = bytearray(font_bytes)
    for i in range(32):
        data[i] ^= key[i % 16]
    return bytes(data)


def finalize_package(path):
    """Embed Vazirmatn (ECMA-376 17.8.1 obfuscated fonts) and enable the page colour."""
    fonts = {"Regular": rc.FONT_FILES[400], "Bold": rc.FONT_FILES[700]}
    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        embeds, rels, payload = "", "", {}
        for i, (style, fpath) in enumerate(fonts.items(), start=1):
            guid = "{" + str(uuid.uuid4()).upper() + "}"
            with open(fpath, "rb") as f:
                payload[f"word/fonts/font{i}.odttf"] = obfuscate(f.read(), guid)
            embeds += f'<w:embed{style} r:id="rIdFont{i}" w:fontKey="{guid}"/>'
            rels += (f'<Relationship Id="rIdFont{i}" Type="http://schemas.openxmlformats.org/'
                     f'officeDocument/2006/relationships/font" Target="fonts/font{i}.odttf"/>')
        font_entry = (f'<w:font w:name="{FONT}"><w:charset w:val="00"/><w:family w:val="auto"/>'
                      f'<w:pitch w:val="variable"/>{embeds}</w:font>')
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/fontTable.xml":
                s = data.decode("utf-8")
                if "xmlns:r=" not in s.split(">", 2)[1]:
                    s = s.replace("<w:fonts ", '<w:fonts xmlns:r="http://schemas.openxmlformats.org/'
                                  'officeDocument/2006/relationships" ', 1)
                s = re.sub(rf'<w:font w:name="{FONT}".*?</w:font>', "", s, flags=re.S)
                s = s.replace("</w:fonts>", font_entry + "</w:fonts>")
                data = s.encode("utf-8")
            elif item.filename == "word/settings.xml":
                s = data.decode("utf-8")
                extra = "<w:displayBackgroundShape/><w:embedTrueTypeFonts/>"
                m = re.search(r"<w:zoom[^>]*/>", s)
                if m:
                    s = s.replace(m.group(0), m.group(0) + extra, 1)
                else:
                    s = re.sub(r"(<w:settings[^>]*>)", r"\1" + extra, s, count=1)
                data = s.encode("utf-8")
            elif item.filename == "word/document.xml":
                s = data.decode("utf-8")
                s = s.replace("<w:body>", f'<w:background w:color="{rc.IVORY}"/><w:body>', 1)
                data = s.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                if 'Extension="odttf"' not in s:
                    s = s.replace("<Default ", '<Default Extension="odttf" ContentType="application/'
                                  'vnd.openxmlformats-officedocument.obfuscatedFont"/><Default ', 1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("word/_rels/fontTable.xml.rels",
                      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      + rels + "</Relationships>")
        for name, data in payload.items():
            zout.writestr(name, data)
    shutil.move(tmp, path)


# ------------------------------------------------------------------ document
def build_docx():
    doc = docx.Document()
    setup_styles(doc)
    doc.core_properties.title = rc.DOC_TITLE
    doc.core_properties.subject = rc.DOC_SUBTITLE
    doc.core_properties.author = rc.INSTITUTION

    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_section(WD_SECTION.NEW_PAGE)
    for s in doc.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.top_margin, s.bottom_margin = Mm(13), Mm(19)
        s.left_margin = s.right_margin = Mm(18)
        s.header_distance, s.footer_distance = Mm(6), Mm(8)
        # no section-level <w:bidi/>: some applications would then mirror the
        # (already visually ordered) tables a second time

    body = doc.element.body
    # section-break paragraphs created by add_section (carry sectPr of pages 1 & 2)
    breaks = [p for p in body.iterchildren(qn("w:p"))
              if p.find(qn("w:pPr")) is not None and p.find(qn("w:pPr")).find(qn("w:sectPr")) is not None]
    for bp in breaks:  # make them occupy (almost) no height
        ppr = bp.find(qn("w:pPr"))
        ppr.insert(0, parse_xml(f'<w:spacing {NS} w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'))
        ppr.insert(0, parse_xml(f"<w:bidi {NS}/>"))
        ppr.append(parse_xml(f'<w:rPr {NS}><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'))
    anchors = [breaks[0], breaks[1], body.find(qn("w:sectPr"))]
    part = doc.part

    def add(page, *els):
        for el in els:
            if isinstance(el, list):
                add(page, *el)
            else:
                anchors[page].addprevious(el)

    # ---------------- PAGE 1
    logo = para([PICS.inline_run(part, rc.downscaled(rc.asset("logo_clean_transparent.png"), 1100),
                                 69, 25)], align="center", after=1)
    for r in decor_runs(part, [("tl", 46, 0.85), ("tr", 44, 0.85), ("br", 46, 0.85)]):
        logo.append(r)
    add(0, logo)
    add(0, para([run(rc.INSTITUTION, 10.5, False, "5D7A68")], align="center"))
    add(0, para([run(rc.DOC_TITLE, 20, True, rc.FOREST)], align="center", line=1.1))
    add(0, para([run(rc.DOC_SUBTITLE, 8.5, False, "7A8F82")], align="center", before=0.5, after=1))
    add(0, para([PICS.inline_run(part, faded(rc.asset("botanical_divider_transparent.png"), 0.75, 1600),
                                 58, 4.5)], align="center", after=3))
    add(0, section_header(part, rc.SEC1))
    add(0, meta_card())
    add(0, section_header(part, rc.SEC2, before=3))
    for i, (lbl, h) in enumerate(zip(rc.SEC2_FIELDS, (49, 42, 39))):
        add(0, leaf_label(part, lbl, before=0 if i == 0 else 2), writing_box(TEXT_W, h))

    # ---------------- PAGE 2
    hdr2 = compact_header(part, rc.PAGE2_HEADER, rc.PAGE2_TAG)
    for r in decor_runs(part, [("tl", 38, 0.7), ("br", 38, 0.7)]):
        hdr2[0].append(r)
    add(1, hdr2)
    add(1, section_header(part, rc.SEC3))
    for i, n in enumerate(rc.MEMBERS):
        if i:
            add(1, para(tiny=True, after=3.5))
        add(1, member_card(part, n, (26, 20, 20)))

    # ---------------- PAGE 3
    hdr3 = compact_header(part, rc.PAGE3_HEADER, rc.PAGE3_TAG)
    for r in decor_runs(part, [("tl", 38, 0.7), ("br", 44, 0.85)]):
        hdr3[0].append(r)
    add(2, hdr3)
    add(2, section_header(part, rc.SEC4))
    for i, lbl in enumerate(rc.SEC4_FIELDS):
        add(2, leaf_label(part, lbl, before=0 if i == 0 else 2), writing_box(TEXT_W, 26))
    add(2, section_header(part, rc.SEC5, before=3))
    add(2, plan_table())
    add(2, para(tiny=True, after=3.5))
    add(2, sign_card(part))

    for i, s in enumerate(doc.sections):
        set_footer(s, i)

    # unique drawing ids across all parts
    n = 1
    for p in [doc.part] + [s.footer.part for s in doc.sections]:
        for el in p.element.iter(qn("wp:docPr")):
            el.set("id", str(n))
            el.set("name", f"Picture {n}")
            n += 1

    doc.save(OUT_DOCX)
    finalize_package(OUT_DOCX)
    print(f"Saved {OUT_DOCX} successfully!")


if __name__ == "__main__":
    build_docx()
