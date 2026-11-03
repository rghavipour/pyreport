"""Build the weekly report HTML and print it to PDF (A4, 3 pages).

Font: Vazirmatn is embedded into the HTML (base64 @font-face), so the PDF is
identical on every machine - no internet / system font required.

Chrome lookup order: $CHROME_BIN, Playwright's bundled Chromium, macOS Chrome,
then chromium / google-chrome on PATH.
"""
import base64
import glob
import os
import shutil
import subprocess
import sys

import pypdfium2 as pdfium
from PIL import Image

import report_content as rc

os.chdir(rc.BASE_DIR)

OUT_HTML = "weekly_report_psychology.html"
OUT_PDF = "weekly_report_psychology.pdf"


def b64_file(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def b64_stream(stream):
    return base64.b64encode(stream.getvalue()).decode("ascii")


tl_b64 = b64_stream(rc.downscaled(rc.asset("corner_tl.png"), 700))
tr_b64 = b64_stream(rc.downscaled(rc.asset("corner_tr.png"), 700))
br_b64 = b64_stream(rc.downscaled(rc.asset("corner_br.png"), 700))
logo_b64 = b64_stream(rc.downscaled(rc.asset("logo_clean_transparent.png"), 1100))
div_b64 = b64_file(rc.asset("botanical_divider_transparent.png"))
paper_b64 = b64_stream(rc.downscaled(rc.asset("paper.jpg"), 1240, fmt="JPEG", quality=82))

font_faces = "\n".join(
    f"""  @font-face {{
    font-family: '{rc.FONT_FAMILY}';
    src: url('data:font/ttf;base64,{b64_file(path)}') format('truetype');
    font-weight: {weight};
    font-style: normal;
  }}"""
    for weight, path in rc.FONT_FILES.items()
)


# ---------------------------------------------------------------- fragments
def section_header(num_title):
    num, title = num_title
    return f"""
      <div class="section-header">
        <span class="section-number">{num}</span>
        <h2 class="section-title">{title}</h2>
        <div class="section-decor-line"></div>
      </div>"""


def field(label, flex=1.0, min_h=14):
    return f"""
        <div class="field-group" style="flex: {flex};">
          <div class="field-label-bar"><span class="bullet-leaf"></span><span>{label}</span></div>
          <div class="writing-area" style="min-height: {min_h}mm;"></div>
        </div>"""


def compact_header(subtitle, tag):
    return f"""
    <div class="header-compact">
      <div class="header-compact-title">{rc.DOC_TITLE}<span class="hc-sep">|</span>{subtitle}</div>
      <div class="header-compact-meta">
        <span>{rc.INSTITUTION}</span>
        <span class="hc-dot"></span>
        <span class="header-compact-tag">{tag}</span>
      </div>
    </div>"""


def footer(idx):
    right, center, badge = rc.FOOTERS[idx]
    return f"""
    <div class="page-footer">
      <div class="footer-side">{right}</div>
      <div class="footer-center">{center}</div>
      <div class="footer-side footer-end"><span class="footer-badge">{badge}</span></div>
    </div>"""


def member_card(num):
    fields = "".join(
        field(lbl, flex=f, min_h=h)
        for lbl, f, h in zip(rc.MEMBER_FIELDS, (1.15, 1, 1), (18, 15, 15))
    )
    return f"""
      <div class="member-card">
        <div class="member-card-header">
          <div class="member-name-row">
            <span class="member-name-label">{rc.MEMBER_NAME}</span>
            <div class="member-name-fill"></div>
          </div>
          <span class="member-badge">عضو شماره {num}</span>
        </div>
        <div class="member-fields">{fields}
        </div>
      </div>"""


meta_rows = "".join(
    f"""
        <div class="meta-field"><span class="meta-label">{a}</span><div class="meta-line"></div></div>
        <div class="meta-field"><span class="meta-label">{b}</span><div class="meta-line"></div></div>"""
    for a, b in rc.SEC1_FIELDS
)

plan_head = "".join(
    f'<th class="{c}">{h}</th>'
    for h, c in zip(rc.PLAN_HEADERS, ("col-num", "col-act", "col-own", "col-time"))
)
plan_body = "".join(
    f'<tr><td class="col-num">{n}</td><td></td><td></td><td></td></tr>' for n in rc.PLAN_ROWS
)

sign_boxes = "".join(
    f"""
      <div class="sign-box">
        <div>
          <div class="sign-title"><span class="bullet-leaf"></span><span>{s['title']}</span></div>
          <div class="sign-desc">{s['desc']}</div>
        </div>
        <div class="sign-notes-box">{s['notes']}</div>
        <div class="sign-area-line"><span>{s['name']}</span><span>{s['sign']}</span></div>
      </div>"""
    for s in rc.SIGN_BOXES
)

C = {k: "#" + getattr(rc, k) for k in
     ("FOREST", "FOREST_2", "SAGE", "BEIGE", "BORDER", "BORDER_SOFT", "IVORY",
      "CARD", "FIELD", "MUTED", "DARK_TEXT", "LINE_DOT")}

html_content = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<title>{rc.DOC_TITLE} - {rc.INSTITUTION}</title>
<style>
{font_faces}

  @page {{ size: A4 portrait; margin: 0; }}

  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }}

  body {{
    margin: 0;
    padding: 0;
    font-family: '{rc.FONT_FAMILY}', Tahoma, sans-serif;
    color: {C['FOREST']};
    background-color: {C['IVORY']};
    direction: rtl;
    text-align: right;
  }}

  .page {{
    width: 210mm;
    height: 297mm;
    position: relative;
    padding: 14mm 18mm 12mm 18mm;
    page-break-after: always;
    break-after: page;
    overflow: hidden;
    background: {C['IVORY']} url('data:image/jpeg;base64,{paper_b64}') no-repeat center / cover;
  }}
  .page:last-child {{ page-break-after: auto; break-after: auto; }}

  .corner {{ position: absolute; pointer-events: none; z-index: 1; }}
  .corner-tl {{ top: 0; left: 0; }}
  .corner-tr {{ top: 0; right: 0; }}
  .corner-br {{ bottom: 0; right: 0; }}

  .content-container {{
    position: relative;
    z-index: 2;
    height: 100%;
    display: flex;
    flex-direction: column;
  }}

  /* ---------- Page 1 header ---------- */
  .header-main {{ text-align: center; margin-bottom: 1mm; }}
  .logo-img {{ height: 25mm; max-width: 90%; object-fit: contain; display: block; margin: 0 auto 1mm; }}
  .institution-name {{ font-size: 10.5pt; font-weight: 500; color: #5D7A68; margin: 0 0 0.5mm; }}
  .document-title {{ font-size: 21pt; font-weight: 800; color: {C['FOREST']}; margin: 0; line-height: 1.35; }}
  .document-subtitle {{ font-size: 8.5pt; font-weight: 400; color: #7A8F82; margin: 0.5mm 0 1.5mm; }}
  .botanical-divider {{ text-align: center; margin: 0 0 3mm; line-height: 0; }}
  .botanical-divider img {{ height: 4.5mm; width: 65%; object-fit: contain; opacity: 0.75; }}

  /* ---------- Compact header (pages 2 & 3): centered, never wraps the meta ---------- */
  .header-compact {{
    text-align: center;
    border-bottom: 1px solid {C['BORDER']};
    padding-bottom: 2.5mm;
    margin-bottom: 4mm;
  }}
  .header-compact-title {{
    font-size: 13.5pt;
    font-weight: 800;
    color: {C['FOREST']};
    line-height: 1.5;
  }}
  .hc-sep {{ color: {C['SAGE']}; font-weight: 400; margin: 0 2.5mm; }}
  .header-compact-meta {{
    margin-top: 1mm;
    font-size: 8.5pt;
    color: #728779;
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 2.5mm;
  }}
  .hc-dot {{ width: 1.2mm; height: 1.2mm; border-radius: 50%; background: {C['SAGE']}; display: inline-block; }}
  .header-compact-tag {{
    background: {C['BEIGE']};
    padding: 0.5mm 3mm;
    border-radius: 3mm;
    font-size: 7.5pt;
    font-weight: 600;
    color: {C['FOREST_2']};
  }}

  /* ---------- Sections ---------- */
  .section-block {{ margin-bottom: 3mm; display: flex; flex-direction: column; }}
  .section-block.grow {{ flex: 1; min-height: 0; }}

  .section-header {{ display: flex; align-items: center; gap: 2.5mm; margin-bottom: 2mm; }}
  /* one numbering style everywhere: a perfect circle */
  .section-number {{
    flex: 0 0 auto;
    width: 6.4mm;
    height: 6.4mm;
    border-radius: 50%;
    background: linear-gradient(135deg, {C['FOREST']}, {C['FOREST_2']});
    color: {C['IVORY']};
    font-size: 9pt;
    font-weight: 700;
    line-height: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    padding-top: 0.3mm;
  }}
  .section-title {{ font-size: 12pt; font-weight: 700; color: {C['FOREST']}; margin: 0; white-space: nowrap; }}
  .section-decor-line {{
    flex: 1;
    height: 1px;
    background: linear-gradient(to left, {C['FOREST']}, {C['BORDER']}, transparent);
    opacity: 0.55;
  }}

  /* ---------- Meta card ---------- */
  .meta-card {{
    background: {C['FIELD']};
    border: 1px solid {C['BORDER_SOFT']};
    border-radius: 3.5mm;
    padding: 3mm 4.5mm;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 3mm 8mm;
  }}
  .meta-field {{ display: flex; align-items: flex-end; gap: 2mm; }}
  .meta-label {{ font-size: 9.5pt; font-weight: 600; color: {C['FOREST_2']}; white-space: nowrap; min-width: 27mm; }}
  .meta-line {{ flex: 1; border-bottom: 1px dotted {C['LINE_DOT']}; height: 5mm; }}

  /* ---------- Writing boxes: one solid fill everywhere ---------- */
  .field-group {{ margin-bottom: 2mm; display: flex; flex-direction: column; min-height: 0; }}
  .field-group:last-child {{ margin-bottom: 0; }}
  .field-label-bar {{
    font-size: 9pt;
    font-weight: 600;
    color: #2D4C3E;
    margin-bottom: 1.2mm;
    display: flex;
    align-items: center;
    gap: 1.5mm;
  }}
  .bullet-leaf {{
    width: 2mm;
    height: 2mm;
    background-color: {C['SAGE']};
    border-radius: 50% 0 50% 0;
    display: inline-block;
    flex: 0 0 auto;
  }}
  .writing-area {{
    flex: 1;
    background: {C['FIELD']};
    border: 1px solid {C['BORDER_SOFT']};
    border-radius: 3mm;
  }}

  /* ---------- Member card ---------- */
  .member-card {{
    flex: 1;
    display: flex;
    flex-direction: column;
    background: {C['CARD']};
    border: 1px solid {C['BORDER']};
    border-right: 3.5px solid {C['FOREST']};
    border-radius: 3.5mm;
    padding: 3mm 4mm 3.5mm;
    margin-bottom: 3.5mm;
  }}
  .member-card:last-child {{ margin-bottom: 0; }}
  .member-card-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2mm;
    padding-bottom: 1.8mm;
    border-bottom: 1px dashed {C['BORDER_SOFT']};
  }}
  .member-name-row {{ display: flex; align-items: flex-end; gap: 2mm; flex: 1; }}
  .member-name-label {{ font-size: 9.5pt; font-weight: 700; color: {C['FOREST']}; }}
  .member-name-fill {{ flex: 1; border-bottom: 1px dotted {C['LINE_DOT']}; height: 4.5mm; max-width: 70mm; }}
  .member-badge {{
    font-size: 8pt;
    font-weight: 600;
    color: #5A7463;
    background: {C['BEIGE']};
    padding: 0.6mm 3mm;
    border-radius: 2.5mm;
  }}
  .member-fields {{ flex: 1; display: flex; flex-direction: column; }}

  /* ---------- Plan table ---------- */
  .plan-wrap {{
    border: 1px solid {C['BORDER_SOFT']};
    border-radius: 3mm;
    overflow: hidden;
    background: {C['FIELD']};
  }}
  .plan-table {{ width: 100%; border-collapse: collapse; font-size: 8.5pt; }}
  .plan-table th {{
    background: {C['BEIGE']};
    color: {C['FOREST']};
    font-weight: 700;
    padding: 2.2mm 2.5mm;
    border-bottom: 1.5px solid #CBD5BC;
    border-left: 1px solid {C['BORDER_SOFT']};
    text-align: right;
  }}
  .plan-table td {{
    height: 9mm;
    padding: 0 2.5mm;
    border-bottom: 1px solid {C['BEIGE']};
    border-left: 1px solid {C['BEIGE']};
    color: #33443C;
  }}
  .plan-table th:last-child, .plan-table td:last-child {{ border-left: none; }}
  .plan-table tr:last-child td {{ border-bottom: none; }}
  .plan-table tbody tr:nth-child(even) td {{ background: {C['CARD']}; }}
  .plan-table .col-num {{ width: 9%; text-align: center; font-weight: 700; color: {C['FOREST']}; }}
  .plan-table .col-act {{ width: 47%; }}
  .plan-table .col-own {{ width: 24%; }}
  .plan-table .col-time {{ width: 20%; text-align: center; }}

  /* ---------- Signatures ---------- */
  .sign-card {{
    background: {C['CARD']};
    border: 1px solid {C['BORDER']};
    border-radius: 3.5mm;
    padding: 3.5mm 4.5mm;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 7mm;
  }}
  .sign-box {{ display: flex; flex-direction: column; justify-content: space-between; }}
  .sign-title {{ font-size: 9pt; font-weight: 700; color: {C['FOREST']}; margin-bottom: 1mm; display: flex; align-items: center; gap: 1.5mm; }}
  .sign-desc {{ font-size: 7.8pt; color: #697F71; margin-bottom: 2mm; line-height: 1.6; }}
  .sign-notes-box {{
    border: 1px dashed {C['BORDER']};
    background: {C['FIELD']};
    border-radius: 2mm;
    height: 16mm;
    padding: 1.8mm 2.5mm;
    font-size: 7.5pt;
    color: #8F9E94;
    margin-bottom: 2.5mm;
  }}
  .sign-area-line {{
    border-top: 1px dashed {C['LINE_DOT']};
    padding-top: 1.8mm;
    padding-bottom: 6mm;
    display: flex;
    justify-content: space-between;
    font-size: 7.8pt;
    color: #758679;
  }}

  /* ---------- Footer ---------- */
  .page-footer {{
    margin-top: auto;
    padding-top: 2.2mm;
    border-top: 0.8px solid {C['BORDER_SOFT']};
    display: flex;
    align-items: center;
    font-size: 7.5pt;
    color: #809687;
  }}
  .footer-side {{ flex: 1; }}
  .footer-end {{ text-align: left; }}
  .footer-center {{ flex: 0 0 auto; font-weight: 500; text-align: center; padding: 0 4mm; }}
  .footer-badge {{
    background: {C['BEIGE']};
    padding: 0.6mm 2.8mm;
    border-radius: 2mm;
    font-weight: 600;
    color: #426354;
  }}
  .spacer {{ height: 3.5mm; flex: 0 0 auto; }}
</style>
</head>
<body>

<!-- ==================== PAGE 1 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner corner-tl" style="width: 46mm; opacity: .85;" alt="">
  <img src="data:image/png;base64,{tr_b64}" class="corner corner-tr" style="width: 44mm; opacity: .85;" alt="">
  <img src="data:image/png;base64,{br_b64}" class="corner corner-br" style="width: 46mm; opacity: .85;" alt="">

  <div class="content-container">
    <div class="header-main">
      <img src="data:image/png;base64,{logo_b64}" class="logo-img" alt="لوگو">
      <div class="institution-name">{rc.INSTITUTION}</div>
      <h1 class="document-title">{rc.DOC_TITLE}</h1>
      <div class="document-subtitle">{rc.DOC_SUBTITLE}</div>
      <div class="botanical-divider"><img src="data:image/png;base64,{div_b64}" alt=""></div>
    </div>

    <div class="section-block">{section_header(rc.SEC1)}
      <div class="meta-card">{meta_rows}
      </div>
    </div>

    <div class="section-block grow">{section_header(rc.SEC2)}
      {field(rc.SEC2_FIELDS[0], 1.15, 22)}
      {field(rc.SEC2_FIELDS[1], 1.0, 18)}
      {field(rc.SEC2_FIELDS[2], 0.95, 16)}
    </div>
{footer(0)}
  </div>
</div>

<!-- ==================== PAGE 2 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner corner-tl" style="width: 38mm; opacity: .7;" alt="">
  <img src="data:image/png;base64,{br_b64}" class="corner corner-br" style="width: 38mm; opacity: .7;" alt="">

  <div class="content-container">{compact_header(rc.PAGE2_HEADER, rc.PAGE2_TAG)}

    <div class="section-block grow">{section_header(rc.SEC3)}
      {"".join(member_card(n) for n in rc.MEMBERS)}
    </div>
{footer(1)}
  </div>
</div>

<!-- ==================== PAGE 3 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner corner-tl" style="width: 38mm; opacity: .7;" alt="">
  <img src="data:image/png;base64,{br_b64}" class="corner corner-br" style="width: 44mm; opacity: .85;" alt="">

  <div class="content-container">{compact_header(rc.PAGE3_HEADER, rc.PAGE3_TAG)}

    <div class="section-block grow">{section_header(rc.SEC4)}
      {"".join(field(l, 1, 20) for l in rc.SEC4_FIELDS)}
    </div>

    <div class="section-block">{section_header(rc.SEC5)}
      <div class="plan-wrap">
        <table class="plan-table">
          <thead><tr>{plan_head}</tr></thead>
          <tbody>{plan_body}</tbody>
        </table>
      </div>
    </div>

    <div class="sign-card">{sign_boxes}
    </div>
    <div class="spacer"></div>
{footer(2)}
  </div>
</div>

</body>
</html>
"""


def find_chrome():
    candidates = [os.environ.get("CHROME_BIN")]
    candidates += sorted(glob.glob("/opt/pw-browsers/chromium-*/chrome-linux*/chrome"), reverse=True)
    candidates += sorted(glob.glob(os.path.expanduser(
        "~/.cache/ms-playwright/chromium-*/chrome-linux*/chrome")), reverse=True)
    candidates += [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ]
    candidates += [shutil.which(n) for n in
                   ("chromium", "chromium-browser", "google-chrome", "google-chrome-stable", "chrome")]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    sys.exit("Chrome/Chromium not found - set CHROME_BIN")


def main():
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_content)

    subprocess.run([
        find_chrome(),
        "--headless=new",
        "--no-sandbox",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=10000",
        f"--print-to-pdf={os.path.abspath(OUT_PDF)}",
        "file://" + os.path.abspath(OUT_HTML),
    ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    pdf = pdfium.PdfDocument(OUT_PDF)
    os.makedirs("preview", exist_ok=True)
    for i, page in enumerate(pdf):
        img = page.render(scale=2).to_pil()
        img.save(f"preview/page_{i + 1}.png")
        w = 600
        thumb = img.resize((w, int(img.height * w / img.width)), Image.LANCZOS)
        thumb.save(f"preview/page_{i + 1}_thumb.png")
    print(f"Saved {OUT_PDF} ({len(pdf)} pages) and previews.")


if __name__ == "__main__":
    main()
