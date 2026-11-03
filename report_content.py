"""Shared content, colors and graphics for the weekly report templates.

Both generators (PDF and DOCX) import this module so the two outputs always
carry identical text, numbering and color palette.
"""
import io
import os

from PIL import Image, ImageDraw, ImageFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(BASE_DIR, "assets")
FONT_DIR = os.path.join(ASSETS, "fonts")

FONT_FAMILY = "Vazirmatn"
FONT_FILES = {
    400: os.path.join(FONT_DIR, "Vazirmatn-Regular.ttf"),
    500: os.path.join(FONT_DIR, "Vazirmatn-Medium.ttf"),
    600: os.path.join(FONT_DIR, "Vazirmatn-SemiBold.ttf"),
    700: os.path.join(FONT_DIR, "Vazirmatn-Bold.ttf"),
    800: os.path.join(FONT_DIR, "Vazirmatn-ExtraBold.ttf"),
}

# ---------------------------------------------------------------- palette
FOREST = "294B40"
FOREST_2 = "385E4F"
SAGE = "89957B"
BEIGE = "EAE6DB"
BORDER = "D5DEC9"
BORDER_SOFT = "DFE3D5"
IVORY = "FAF9F4"
CARD = "FBFAF6"       # solid card fill (member card / sign card)
FIELD = "FFFFFF"      # solid fill used for EVERY writing box
MUTED = "6E8575"
DARK_TEXT = "2A3831"
LINE_DOT = "A8B39A"

# ---------------------------------------------------------------- texts
INSTITUTION = "مرکز مشاوره، روانشناسی و سلامت روان"
DOC_TITLE = "گزارش هفتگی گروه‌ها"
DOC_SUBTITLE = "فرم اختصاصی ارزیابی هفتگی فعالیت‌ها، پایش سلامت تیمی و برنامه‌ریزی یکپارچه"
CONFIDENTIAL = "محرمانه — ویژه ارزیابی داخلی مرکز سلامت روان"

PAGE2_HEADER = "بخش اعضای گروه"
PAGE2_TAG = "پایش تیمی هفتگی"
PAGE3_HEADER = "چالش‌ها و برنامه‌ریزی آتی"
PAGE3_TAG = "برنامه‌ریزی و ارزیابی"

FOOTERS = [
    (CONFIDENTIAL, "گزارش هفتگی گروه‌ها", "صفحه ۱ از ۳"),
    (CONFIDENTIAL, "گزارش هفتگی اعضای گروه", "صفحه ۲ از ۳"),
    (INSTITUTION + " — پرونده محرمانه سازمانی", "پایان گزارش هفتگی", "صفحه ۳ از ۳"),
]

SEC1 = ("۱", "مشخصات گروه")
SEC1_FIELDS = [("نام گروه:", "نام سرگروه:"), ("بازه زمانی گزارش:", "تاریخ ارائه:")]

SEC2 = ("۲", "گزارش عملکرد سرگروه")
SEC2_FIELDS = [
    "اقدامات و فعالیت‌های انجام‌شده:",
    "برنامه‌ها و ایده‌های پیشنهادی:",
    "دستاوردها و نتایج:",
]

SEC3 = ("۳", "گزارش اعضای گروه")
MEMBER_NAME = "نام عضو:"
MEMBER_FIELDS = [
    "شرح فعالیت‌ها و اقدامات انجام‌شده:",
    "نیازها و درخواست‌های حمایتی:",
    "پیشنهادها و برنامه‌های پیش‌رو:",
]
MEMBERS = ["۱", "۲"]

SEC4 = ("۴", "چالش‌ها و موانع")
SEC4_FIELDS = [
    "مشکلات و موانع موجود:",
    "موارد نیازمند پیگیری:",
    "راهکارهای پیشنهادی:",
]

SEC5 = ("۵", "برنامه هفته آینده")
PLAN_HEADERS = ["ردیف", "اقدامات پیشنهادی", "مسئول انجام", "زمان‌بندی"]
PLAN_ROWS = ["۱", "۲", "۳", "۴", "۵"]

SIGN_BOXES = [
    {
        "title": "امضا و تأیید سرگروه",
        "desc": "صحت مطالب و گزارش‌های ثبت‌شده فوق مورد تأیید است.",
        "notes": "ملاحظات تکمیلی سرگروه:",
        "name": "نام و نام خانوادگی:",
        "sign": "امضا و تاریخ:",
    },
    {
        "title": "تأییدیه و بازخورد سوپروایزر / مدیریت مرکز",
        "desc": "ملاحظات نظارتی، ارزیابی کیفی و تأمین منابع.",
        "notes": "دستور اقدامات نظارتی و بازخورد مدیریت:",
        "name": "نام مسئول / سوپروایزر:",
        "sign": "امضا و مهر مرکز:",
    },
]


def asset(name):
    return os.path.join(ASSETS, name)


def _rgb(hex_str, alpha=255):
    return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4)) + (alpha,)


# ---------------------------------------------------------------- graphics
def number_badge_png(digit, px=240):
    """Perfect circle with a Persian digit — the single numbering style used
    for every section number in both outputs."""
    scale = 4
    size = px * scale
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    # diagonal gradient forest -> forest_2, clipped to a circle
    grad = Image.new("RGBA", (size, size))
    c1, c2 = _rgb(FOREST), _rgb(FOREST_2)
    gd = ImageDraw.Draw(grad)
    for i in range(size * 2):
        t = i / (size * 2)
        col = tuple(int(c1[k] + (c2[k] - c1[k]) * t) for k in range(3)) + (255,)
        gd.line([(i, 0), (0, i)], fill=col)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size - 1, size - 1], fill=255)
    img.paste(grad, (0, 0), mask)

    font = ImageFont.truetype(FONT_FILES[700], int(size * 0.56))
    d = ImageDraw.Draw(img)
    bbox = d.textbbox((0, 0), digit, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((size - w) / 2 - bbox[0], (size - h) / 2 - bbox[1]), digit,
           font=font, fill=_rgb(IVORY))
    img = img.resize((px, px), Image.LANCZOS)
    return _to_stream(img)


def leaf_png(px=80):
    scale = 4
    s = px * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # leaf: rounded top-left and bottom-right corners (like CSS 50% 0 50% 0)
    col = _rgb(SAGE)
    d.rectangle([0, 0, s, s], fill=col)
    cut = Image.new("L", (s, s), 255)
    cd = ImageDraw.Draw(cut)
    cd.rectangle([0, 0, s // 2, s // 2], fill=0)
    cd.rectangle([s // 2, s // 2, s, s], fill=0)
    cd.pieslice([0, 0, s - 1, s - 1], 180, 270, fill=255)
    cd.pieslice([0, 0, s - 1, s - 1], 0, 90, fill=255)
    img.putalpha(cut)
    img = img.resize((px, px), Image.LANCZOS)
    return _to_stream(img)


def fade_line_png(w=1600, h=12, canvas_h=None):
    """Section decor line: strong next to the title (right) fading to the left.
    canvas_h: optional taller transparent canvas with the line in its middle
    (lets the line share the vertical alignment of the number badge)."""
    canvas_h = canvas_h or h
    img = Image.new("RGBA", (w, canvas_h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    c1, c2 = _rgb(FOREST), _rgb(BORDER)
    y0, y1 = canvas_h // 2 - 1, canvas_h // 2 + 1
    for x in range(w):
        t = 1 - x / (w - 1)          # 1 at right edge, 0 at left edge
        if t > 0.5:
            k = (t - 0.5) * 2
            col = [int(c2[i] + (c1[i] - c2[i]) * k) for i in range(3)]
            a = 150
        else:
            col = list(c2[:3])
            a = int(150 * t * 2)
        d.line([(x, y0), (x, y1)], fill=tuple(col) + (a,))
    return _to_stream(img)


def downscaled(path, max_w, fmt=None, quality=85):
    """Return a smaller copy of an asset (keeps DOCX/HTML lightweight)."""
    img = Image.open(path)
    if img.width > max_w:
        h = int(img.height * max_w / img.width)
        img = img.resize((max_w, h), Image.LANCZOS)
    buf = io.BytesIO()
    fmt = fmt or ("JPEG" if img.mode == "RGB" else "PNG")
    if fmt == "JPEG":
        img.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True)
    else:
        img.save(buf, "PNG", optimize=True)
    buf.seek(0)
    return buf


def text_width_mm(text, size_pt, weight=700):
    """Approximate rendered width of a Persian string (Vazirmatn) in mm."""
    font = ImageFont.truetype(FONT_FILES[weight], 400)
    try:
        bbox = ImageDraw.Draw(Image.new("L", (1, 1))).textbbox(
            (0, 0), text, font=font, direction="rtl", features=["-kern"])
    except Exception:
        bbox = ImageDraw.Draw(Image.new("L", (1, 1))).textbbox((0, 0), text, font=font)
    width_units = (bbox[2] - bbox[0]) / 400.0        # in em
    return width_units * size_pt * 0.3528


def _to_stream(img):
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    buf.seek(0)
    return buf
