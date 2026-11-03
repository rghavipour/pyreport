import base64
import os
import subprocess
import pypdfium2 as pdfium
from PIL import Image

def get_b64(path):
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')

tl_b64 = get_b64('assets/corner_tl.png')
tr_b64 = get_b64('assets/corner_tr.png')
br_b64 = get_b64('assets/corner_br.png')
logo_b64 = get_b64('assets/logo_clean_transparent.png')
div_b64 = get_b64('assets/botanical_divider_transparent.png')
paper_b64 = get_b64('assets/paper.jpg')

html_content = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<title>گزارش هفتگی گروه‌ها - مرکز سلامت روان</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;600;700;800&display=swap');

  @page {{
    size: A4 portrait;
    margin: 0;
  }}

  * {{
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }}

  body {{
    margin: 0;
    padding: 0;
    font-family: 'Vazirmatn', Tahoma, sans-serif;
    color: #294B40;
    background-color: #FAF9F4;
    direction: rtl;
    text-align: right;
    font-feature-settings: "ss01";
  }}

  .page {{
    width: 210mm;
    height: 297mm;
    position: relative;
    padding: 15mm 18mm 14mm 18mm;
    page-break-after: always;
    page-break-inside: avoid;
    overflow: hidden;
    background: #FAF9F4 url('data:image/jpeg;base64,{paper_b64}') repeat;
    background-size: cover;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }}

  .page:last-child {{
    page-break-after: auto;
  }}

  /* Botanical Corner Accents */
  .corner-tl {{
    position: absolute;
    top: 0;
    left: 0;
    width: 46mm;
    pointer-events: none;
    z-index: 1;
    opacity: 0.85;
  }}

  .corner-tr {{
    position: absolute;
    top: 0;
    right: 0;
    width: 44mm;
    pointer-events: none;
    z-index: 1;
    opacity: 0.85;
  }}

  .corner-br {{
    position: absolute;
    bottom: 0;
    right: 0;
    width: 46mm;
    pointer-events: none;
    z-index: 1;
    opacity: 0.85;
  }}

  /* Content Wrapper */
  .content-container {{
    position: relative;
    z-index: 2;
    height: 100%;
    display: flex;
    flex-direction: column;
  }}

  /* Header Section */
  .header-main {{
    text-align: center;
    margin-bottom: 2mm;
  }}

  .logo-wrapper {{
    text-align: center;
    margin-bottom: 1mm;
  }}

  .logo-img {{
    height: 26mm;
    max-width: 90%;
    object-fit: contain;
  }}

  .institution-name {{
    font-size: 11pt;
    font-weight: 500;
    color: #5D7A68;
    letter-spacing: 0.5px;
    margin: 0 0 1mm 0;
  }}

  .document-title {{
    font-size: 21pt;
    font-weight: 800;
    color: #294B40;
    margin: 0;
    letter-spacing: -0.5px;
  }}

  .document-subtitle {{
    font-size: 8.5pt;
    font-weight: 400;
    color: #7A8F82;
    margin: 1mm 0 2mm 0;
  }}

  .botanical-divider {{
    text-align: center;
    margin: 1mm 0 3mm 0;
    line-height: 0;
  }}

  .botanical-divider img {{
    height: 4.5mm;
    width: 65%;
    object-fit: contain;
    opacity: 0.75;
  }}

  /* Compact Header for Pages 2 & 3 */
  .header-compact {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid #D5DEC9;
    padding-bottom: 2.2mm;
    margin-bottom: 3.5mm;
  }}

  .header-compact-title {{
    font-size: 12pt;
    font-weight: 700;
    color: #294B40;
  }}

  .header-compact-meta {{
    font-size: 8.5pt;
    color: #728779;
    display: flex;
    align-items: center;
    gap: 3mm;
  }}

  .header-compact-tag {{
    background: #EAE5D8;
    padding: 0.8mm 3mm;
    border-radius: 3mm;
    font-size: 7.5pt;
    font-weight: 600;
    color: #385E4F;
  }}

  /* Section Styling */
  .section-block {{
    margin-bottom: 3mm;
  }}

  .section-header {{
    display: flex;
    align-items: center;
    gap: 2.5mm;
    margin-bottom: 1.8mm;
  }}

  .section-number-pill {{
    background: linear-gradient(135deg, #294B40, #385E4F);
    color: #FAF9F4;
    font-size: 8.5pt;
    font-weight: 700;
    padding: 0.8mm 3.2mm;
    border-radius: 4mm;
    box-shadow: 0 1px 3px rgba(41, 75, 64, 0.15);
  }}

  .section-title {{
    font-size: 12pt;
    font-weight: 700;
    color: #294B40;
    margin: 0;
  }}

  .section-decor-line {{
    flex: 1;
    height: 1px;
    background: linear-gradient(to left, #294B40, #D5DEC9, transparent);
    opacity: 0.5;
  }}

  /* Metadata Card (Section 1) */
  .meta-card {{
    background: rgba(255, 255, 255, 0.80);
    border: 1px solid #DFE3D5;
    border-radius: 3.5mm;
    padding: 3mm 4mm;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 2.5mm 6mm;
    box-shadow: 0 1px 4px rgba(41, 75, 64, 0.04);
  }}

  .meta-field {{
    display: flex;
    align-items: baseline;
    gap: 2mm;
  }}

  .meta-label {{
    font-size: 9.5pt;
    font-weight: 600;
    color: #385E4F;
    white-space: nowrap;
    min-width: 25mm;
  }}

  .meta-line {{
    flex: 1;
    border-bottom: 1px dotted #A8B39A;
    height: 5mm;
    color: #555;
    font-size: 9.5pt;
    padding-right: 2mm;
  }}

  /* Field Writing Boxes */
  .field-group {{
    margin-bottom: 2mm;
  }}

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
    background-color: #89957B;
    border-radius: 50% 0 50% 0;
    display: inline-block;
  }}

  .writing-area {{
    background: rgba(255, 255, 255, 0.70);
    border: 1px solid #DFE3D5;
    border-radius: 3mm;
    padding: 2.5mm 3.5mm;
    font-size: 9pt;
    color: #2A3831;
    line-height: 1.8;
    box-shadow: inset 0 1px 2px rgba(41, 75, 64, 0.02);
  }}

  /* Member Card Component */
  .member-card {{
    background: rgba(255, 255, 255, 0.85);
    border: 1px solid #D5DEC9;
    border-right: 3.5px solid #294B40;
    border-radius: 3.5mm;
    padding: 3mm 4mm;
    margin-bottom: 3.5mm;
    box-shadow: 0 2px 5px rgba(41, 75, 64, 0.03);
  }}

  .member-card-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 2mm;
    padding-bottom: 1.5mm;
    border-bottom: 1px dashed #DFE3D5;
  }}

  .member-name-row {{
    display: flex;
    align-items: baseline;
    gap: 2mm;
    flex: 1;
  }}

  .member-name-label {{
    font-size: 9.5pt;
    font-weight: 700;
    color: #294B40;
  }}

  .member-name-fill {{
    flex: 1;
    border-bottom: 1px dotted #A8B39A;
    height: 4.5mm;
    max-width: 55mm;
  }}

  .member-badge {{
    font-size: 8pt;
    font-weight: 600;
    color: #5A7463;
    background: #EAE6DB;
    padding: 0.8mm 3mm;
    border-radius: 2.5mm;
  }}

  .member-fields-grid {{
    display: flex;
    flex-direction: column;
    gap: 2mm;
  }}

  .hint-banner {{
    font-size: 8pt;
    color: #5B7263;
    background: rgba(233, 227, 210, 0.5);
    border-right: 2.5px solid #89957B;
    padding: 1.8mm 3.5mm;
    border-radius: 2mm;
    margin-bottom: 2.5mm;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }}

  /* Planning Table (Section 5) */
  .plan-table {{
    width: 100%;
    border-collapse: collapse;
    background: rgba(255, 255, 255, 0.85);
    border-radius: 3mm;
    overflow: hidden;
    border: 1px solid #DFE3D5;
    font-size: 8.5pt;
    box-shadow: 0 1px 4px rgba(41, 75, 64, 0.03);
  }}

  .plan-table th {{
    background: #EAE6DB;
    color: #294B40;
    font-weight: 700;
    padding: 2.2mm 2.5mm;
    border-bottom: 1.5px solid #CBD5BC;
    border-left: 1px solid #DFE3D5;
    text-align: right;
  }}

  .plan-table th:last-child {{
    border-left: none;
  }}

  .plan-table td {{
    padding: 2.2mm 2.5mm;
    border-bottom: 1px solid #EAE6DB;
    border-left: 1px solid #EAE6DB;
    color: #33443C;
    vertical-align: middle;
  }}

  .plan-table td:last-child {{
    border-left: none;
  }}

  .plan-table tr:last-child td {{
    border-bottom: none;
  }}

  .plan-table tr:nth-child(even) {{
    background: rgba(250, 249, 244, 0.7);
  }}

  /* Signatures Card */
  .sign-card {{
    background: rgba(255, 255, 255, 0.85);
    border: 1px solid #D5DEC9;
    border-radius: 3.5mm;
    padding: 3.5mm 4.5mm;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6mm;
    margin-top: 3.5mm;
    box-shadow: 0 1px 4px rgba(41, 75, 64, 0.03);
  }}

  .sign-box {{
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    min-height: 24mm;
  }}

  .sign-title {{
    font-size: 9pt;
    font-weight: 700;
    color: #294B40;
    margin-bottom: 1mm;
    display: flex;
    align-items: center;
    gap: 1.5mm;
  }}

  .sign-desc {{
    font-size: 7.8pt;
    color: #697F71;
    margin-bottom: 2mm;
    line-height: 1.6;
  }}

  .sign-notes-box {{
    border: 1px dashed #D5DEC9;
    background: rgba(250, 249, 244, 0.6);
    border-radius: 2mm;
    height: 9mm;
    margin-bottom: 2mm;
  }}

  .sign-area-line {{
    border-top: 1px dashed #A8B39A;
    padding-top: 1.8mm;
    display: flex;
    justify-content: space-between;
    font-size: 7.8pt;
    color: #758679;
  }}

  /* Footer Styling */
  .page-footer {{
    margin-top: auto;
    padding-top: 2.2mm;
    border-top: 0.8px solid #DFE3D5;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 7.5pt;
    color: #809687;
    position: relative;
    z-index: 3;
  }}

  .footer-center {{
    font-weight: 500;
    letter-spacing: 0.3px;
  }}

  .footer-badge {{
    background: #EAE6DB;
    padding: 0.6mm 2.5mm;
    border-radius: 2mm;
    font-weight: 600;
    color: #426354;
  }}
</style>
</head>
<body>

<!-- ==================== PAGE 1 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner-tl" alt="botanical-tl">
  <img src="data:image/png;base64,{tr_b64}" class="corner-tr" alt="botanical-tr">
  <img src="data:image/png;base64,{br_b64}" class="corner-br" alt="botanical-br">

  <div class="content-container">
    <!-- Header -->
    <div class="header-main">
      <div class="logo-wrapper">
        <img src="data:image/png;base64,{logo_b64}" class="logo-img" alt="لوگوی مرکز سلامت روان">
      </div>
      <div class="institution-name">مرکز مشاوره، روانشناسی و سلامت روان</div>
      <h1 class="document-title">گزارش هفتگی گروه‌ها</h1>
      <div class="document-subtitle">فرم اختصاصی ارزیابی هفتگی فعالیت‌ها، پایش سلامت تیمی و برنامه‌ریزی یکپارچه</div>
      <div class="botanical-divider">
        <img src="data:image/png;base64,{div_b64}" alt="divider">
      </div>
    </div>

    <!-- Section 1: مشخصات گروه -->
    <div class="section-block">
      <div class="section-header">
        <span class="section-number-pill">۱</span>
        <h2 class="section-title">مشخصات گروه</h2>
        <div class="section-decor-line"></div>
      </div>
      <div class="meta-card">
        <div class="meta-field">
          <span class="meta-label">نام گروه:</span>
          <div class="meta-line"></div>
        </div>
        <div class="meta-field">
          <span class="meta-label">نام سرگروه:</span>
          <div class="meta-line"></div>
        </div>
        <div class="meta-field">
          <span class="meta-label">بازه زمانی گزارش:</span>
          <div class="meta-line"></div>
        </div>
        <div class="meta-field">
          <span class="meta-label">تاریخ ارائه:</span>
          <div class="meta-line"></div>
        </div>
      </div>
    </div>

    <!-- Section 2: گزارش عملکرد سرگروه -->
    <div class="section-block" style="flex: 1; display: flex; flex-direction: column;">
      <div class="section-header">
        <span class="section-number-pill">۲</span>
        <h2 class="section-title">گزارش عملکرد سرگروه</h2>
        <div class="section-decor-line"></div>
      </div>

      <div class="field-group" style="flex: 1.15; display: flex; flex-direction: column;">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>اقدامات و فعالیت‌های انجام‌شده:</span>
        </div>
        <div class="writing-area" style="flex: 1; min-height: 22mm;"></div>
      </div>

      <div class="field-group" style="flex: 1; display: flex; flex-direction: column;">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>برنامه‌ها و ایده‌های پیشنهادی:</span>
        </div>
        <div class="writing-area" style="flex: 1; min-height: 18mm;"></div>
      </div>

      <div class="field-group" style="flex: 0.95; display: flex; flex-direction: column;">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>دستاوردها و نتایج:</span>
        </div>
        <div class="writing-area" style="flex: 1; min-height: 16mm;"></div>
      </div>
    </div>

    <!-- Footer Page 1 -->
    <div class="page-footer">
      <div>محرمانه — ویژه ارزیابی داخلی مرکز سلامت روان</div>
      <div class="footer-center">گزارش هفتگی گروه‌ها</div>
      <div class="footer-badge">صفحه ۱ از ۳</div>
    </div>
  </div>
</div>

<!-- ==================== PAGE 2 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner-tl" alt="botanical-tl" style="width: 38mm; opacity: 0.7;">
  <img src="data:image/png;base64,{br_b64}" class="corner-br" alt="botanical-br" style="width: 38mm; opacity: 0.7;">

  <div class="content-container">
    <!-- Compact Header -->
    <div class="header-compact">
      <div class="header-compact-title">گزارش هفتگی گروه‌ها | بخش اعضای گروه</div>
      <div class="header-compact-meta">
        <span>مرکز روانشناسی و سلامت روان</span>
        <span class="header-compact-tag">پایش تیمی هفتگی</span>
      </div>
    </div>

    <!-- Section 3: گزارش اعضای گروه -->
    <div class="section-block" style="flex: 1; display: flex; flex-direction: column;">
      <div class="section-header">
        <span class="section-number-pill">۳</span>
        <h2 class="section-title">گزارش اعضای گروه</h2>
        <div class="section-decor-line"></div>
      </div>

      <div class="hint-banner">
        <span>این بخش به ازای هر یک از اعضای گروه تکمیل می‌گردد. در فایل ورد (DOCX) می‌توانید این قالب را برای اعضای بعدی کپی و تکثیر فرمایید.</span>
        <span style="font-weight: 600; color: #385E4F;">قالب تکرارپذیر اعضا</span>
      </div>

      <!-- Member 1 -->
      <div class="member-card" style="flex: 1; display: flex; flex-direction: column;">
        <div class="member-card-header">
          <div class="member-name-row">
            <span class="member-name-label">نام عضو:</span>
            <div class="member-name-fill"></div>
          </div>
          <span class="member-badge">عضو شماره ۱</span>
        </div>
        <div class="member-fields-grid" style="flex: 1; display: flex; flex-direction: column; justify-content: space-between;">
          <div class="field-group" style="margin-bottom: 1.5mm; flex: 1.1; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>شرح فعالیت‌ها و اقدامات انجام‌شده:</span>
            </div>
            <div class="writing-area" style="min-height: 18mm; flex: 1;"></div>
          </div>
          <div class="field-group" style="margin-bottom: 1.5mm; flex: 0.95; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>نیازها و درخواست‌های حمایتی:</span>
            </div>
            <div class="writing-area" style="min-height: 14mm; flex: 1;"></div>
          </div>
          <div class="field-group" style="margin-bottom: 0; flex: 0.95; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>پیشنهادها و برنامه‌های پیش‌رو:</span>
            </div>
            <div class="writing-area" style="min-height: 14mm; flex: 1;"></div>
          </div>
        </div>
      </div>

      <!-- Member 2 -->
      <div class="member-card" style="flex: 1; display: flex; flex-direction: column;">
        <div class="member-card-header">
          <div class="member-name-row">
            <span class="member-name-label">نام عضو:</span>
            <div class="member-name-fill"></div>
          </div>
          <span class="member-badge">عضو شماره ۲</span>
        </div>
        <div class="member-fields-grid" style="flex: 1; display: flex; flex-direction: column; justify-content: space-between;">
          <div class="field-group" style="margin-bottom: 1.5mm; flex: 1.1; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>شرح فعالیت‌ها و اقدامات انجام‌شده:</span>
            </div>
            <div class="writing-area" style="min-height: 18mm; flex: 1;"></div>
          </div>
          <div class="field-group" style="margin-bottom: 1.5mm; flex: 0.95; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>نیازها و درخواست‌های حمایتی:</span>
            </div>
            <div class="writing-area" style="min-height: 14mm; flex: 1;"></div>
          </div>
          <div class="field-group" style="margin-bottom: 0; flex: 0.95; display: flex; flex-direction: column;">
            <div class="field-label-bar">
              <span class="bullet-leaf"></span>
              <span>پیشنهادها و برنامه‌های پیش‌رو:</span>
            </div>
            <div class="writing-area" style="min-height: 14mm; flex: 1;"></div>
          </div>
        </div>
      </div>
    </div>

    <!-- Footer Page 2 -->
    <div class="page-footer">
      <div>محرمانه — ویژه ارزیابی داخلی مرکز سلامت روان</div>
      <div class="footer-center">گزارش هفتگی اعضای گروه</div>
      <div class="footer-badge">صفحه ۲ از ۳</div>
    </div>
  </div>
</div>

<!-- ==================== PAGE 3 ==================== -->
<div class="page">
  <img src="data:image/png;base64,{tl_b64}" class="corner-tl" alt="botanical-tl" style="width: 38mm; opacity: 0.7;">
  <img src="data:image/png;base64,{br_b64}" class="corner-br" alt="botanical-br" style="width: 44mm; opacity: 0.85;">

  <div class="content-container">
    <!-- Compact Header -->
    <div class="header-compact">
      <div class="header-compact-title">گزارش هفتگی گروه‌ها | چالش‌ها و برنامه‌ریزی آتی</div>
      <div class="header-compact-meta">
        <span>مرکز روانشناسی و سلامت روان</span>
        <span class="header-compact-tag">برنامه‌ریزی و ارزیابی</span>
      </div>
    </div>

    <!-- Section 4: چالش‌ها و موانع -->
    <div class="section-block">
      <div class="section-header">
        <span class="section-number-pill">۴</span>
        <h2 class="section-title">چالش‌ها و موانع</h2>
        <div class="section-decor-line"></div>
      </div>

      <div class="field-group">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>مشکلات و موانع موجود:</span>
        </div>
        <div class="writing-area" style="min-height: 22mm;"></div>
      </div>

      <div class="field-group">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>موارد نیازمند پیگیری:</span>
        </div>
        <div class="writing-area" style="min-height: 22mm;"></div>
      </div>

      <div class="field-group">
        <div class="field-label-bar">
          <span class="bullet-leaf"></span>
          <span>راهکارهای پیشنهادی:</span>
        </div>
        <div class="writing-area" style="min-height: 22mm;"></div>
      </div>
    </div>

    <!-- Section 5: برنامه هفته آینده -->
    <div class="section-block" style="margin-top: 2mm;">
      <div class="section-header">
        <span class="section-number-pill">۵</span>
        <h2 class="section-title">برنامه هفته آینده</h2>
        <div class="section-decor-line"></div>
      </div>

      <table class="plan-table">
        <thead>
          <tr>
            <th style="width: 8%; text-align: center;">ردیف</th>
            <th style="width: 48%;">اقدامات پیشنهادی</th>
            <th style="width: 24%;">مسئول انجام</th>
            <th style="width: 20%; text-align: center;">زمان‌بندی</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: center; font-weight: 600;">۱</td>
            <td style="height: 9mm;"></td>
            <td></td>
            <td style="text-align: center;"></td>
          </tr>
          <tr>
            <td style="text-align: center; font-weight: 600;">۲</td>
            <td style="height: 9mm;"></td>
            <td></td>
            <td style="text-align: center;"></td>
          </tr>
          <tr>
            <td style="text-align: center; font-weight: 600;">۳</td>
            <td style="height: 9mm;"></td>
            <td></td>
            <td style="text-align: center;"></td>
          </tr>
          <tr>
            <td style="text-align: center; font-weight: 600;">۴</td>
            <td style="height: 9mm;"></td>
            <td></td>
            <td style="text-align: center;"></td>
          </tr>
          <tr>
            <td style="text-align: center; font-weight: 600;">۵</td>
            <td style="height: 9mm;"></td>
            <td></td>
            <td style="text-align: center;"></td>
          </tr>
        </tbody>
      </table>
    </div>

    <!-- Signatures Card -->
    <div class="sign-card">
      <div class="sign-box">
        <div>
          <div class="sign-title">
            <span class="bullet-leaf"></span>
            <span>امضا و تأیید سرگروه</span>
          </div>
          <div class="sign-desc">صحت مطالب و گزارش‌های ثبت‌شده فوق مورد تأیید است.</div>
        </div>
        <div class="sign-notes-box" style="height: 15mm; padding: 2mm; font-size: 7.5pt; color: #8F9E94;">
          ملاحظات تکمیلی سرگروه:
        </div>
        <div class="sign-area-line">
          <span>نام و نام خانوادگی:</span>
          <span>امضا و تاریخ:</span>
        </div>
      </div>

      <div class="sign-box">
        <div>
          <div class="sign-title">
            <span class="bullet-leaf"></span>
            <span>تأییدیه و بازخورد سوپروایزر / مدیریت مرکز</span>
          </div>
          <div class="sign-desc">ملاحظات نظارتی، ارزیابی کیفی و تأمین منابع.</div>
        </div>
        <div class="sign-notes-box" style="height: 15mm; padding: 2mm; font-size: 7.5pt; color: #8F9E94;">
          دستور اقدامات نظارتی و بازخورد مدیریت:
        </div>
        <div class="sign-area-line">
          <span>نام مسئول / سوپروایزر:</span>
          <span>امضا و مهر مرکز:</span>
        </div>
      </div>
    </div>

    <!-- Footer Page 3 -->
    <div class="page-footer">
      <div>مرکز مشاوره، روانشناسی و سلامت روان — پرونده محرمانه سازمانی</div>
      <div class="footer-center">پایان گزارش هفتگی</div>
      <div class="footer-badge">صفحه ۳ از ۳</div>
    </div>
  </div>
</div>

</body>
</html>
"""

with open('weekly_report_psychology.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

chrome_cmd = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "--headless=new",
    "--disable-gpu",
    "--run-all-compositor-stages-before-draw",
    "--print-to-pdf=weekly_report_psychology.pdf",
    "weekly_report_psychology.html"
]

subprocess.run(chrome_cmd, check=True)

# Re-render previews
pdf = pdfium.PdfDocument("weekly_report_psychology.pdf")
for i, page in enumerate(pdf):
    img = page.render(scale=2).to_pil()
    img.save(f"preview/page_{i+1}.png")
    w = 600
    h = int(img.height * w / img.width)
    thumb = img.resize((w, h), Image.LANCZOS)
    thumb.save(f"preview/page_{i+1}_thumb.png")

print("Updated PDF and previews successfully!")
