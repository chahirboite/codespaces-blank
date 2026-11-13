import base64
import datetime
import io
import os
import re
import urllib.request
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    from reportlab.pdfbase import pdfmetrics

    HAS_PDF_LIBS = True
except ImportError:
    HAS_PDF_LIBS = False


REMOTE_ARABIC_FONTS = {
    "amiri": {
        "filename": "Amiri-Regular.ttf",
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/amiri/Amiri-Regular.ttf",
    },
    "noto": {
        "filename": "NotoNaskhArabic-Regular.ttf",
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notonaskharabic/NotoNaskhArabic%5Bwght%5D.ttf",
    },
}


def get_auto_school_year():
    """حساب الموسم الدراسي تلقائياً"""
    today = datetime.date.today()
    current_year = today.year
    if today.month < 9:
        return f"{current_year - 1}-{current_year}"
    else:
        return f"{current_year}-{current_year + 1}"


def format_room_label(room_val, is_preview=False):
    """تنسيق رقم القاعة تلقائياً عند إدخال الرقم مباشرة"""
    r = room_val.strip()
    if not r:
        return "[القاعة]" if is_preview else "القاعة ........"
    prefixes = ("القاعة", "قاعة", "مختبر", "المختبر", "ورشة", "الورشة", "ق.")
    if r.startswith(prefixes):
        return r
    return f"القاعة {r}"


def get_static_logo_data_uri():
    """قراءة الشعار المحلي خلف العنوان"""
    for filename in ["logo.jpg", "logo.png", "logo.jpeg", "LOGO.JPG", "LOGO.PNG"]:
        if os.path.exists(filename):
            try:
                ext = filename.split(".")[-1].lower()
                mime_type = "image/jpeg" if ext in ["jpg", "jpeg"] else "image/png"
                with open(filename, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode()
                return f"data:{mime_type};base64,{encoded}"
            except Exception:
                pass
    return None


def get_logo_bytes(uploaded_logo):
    """تحديد مصدر صورة الشعار"""
    if uploaded_logo is not None:
        return io.BytesIO(uploaded_logo.getvalue())
    for filename in ["logo.jpg", "logo.png", "logo.jpeg", "LOGO.JPG", "LOGO.PNG"]:
        if os.path.exists(filename):
            return filename
    return None


def get_logo_base64_src(uploaded_logo):
    """إرجاع مصدر الصورة لـ HTML Preview"""
    if uploaded_logo is not None:
        try:
            bytes_data = uploaded_logo.getvalue()
            mime = uploaded_logo.type
            encoded = base64.b64encode(bytes_data).decode()
            return f"data:{mime};base64,{encoded}"
        except Exception:
            pass
    return get_static_logo_data_uri()


def get_arabic_pdf_font(preferred_key="amiri"):
    """البحث عن خط عربي احترافي أو تحميله مؤقتاً لتوليد PDF"""
    font_info = REMOTE_ARABIC_FONTS.get(preferred_key, REMOTE_ARABIC_FONTS["amiri"])
    cache_dir = "/tmp" if os.path.exists("/tmp") else "."
    cached_path = os.path.join(cache_dir, font_info["filename"])

    if os.path.exists(cached_path) and os.path.getsize(cached_path) > 15000:
        return cached_path

    try:
        req = urllib.request.Request(font_info["url"], headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = resp.read()
            if len(data) > 15000:
                with open(cached_path, "wb") as f:
                    f.write(data)
                return cached_path
    except Exception:
        if os.path.exists(cached_path):
            try:
                os.remove(cached_path)
            except Exception:
                pass

    candidates = [
        "/usr/share/fonts/truetype/hosny-amiri/Amiri-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path

    if os.path.exists("/usr/share/fonts"):
        for root, _, files in os.walk("/usr/share/fonts"):
            for f in files:
                if f.lower() in [
                    "amiri-regular.ttf",
                    "notonaskharabic-regular.ttf",
                    "notosansarabic-regular.ttf",
                    "dejavusans.ttf",
                    "freesans.ttf",
                ]:
                    return os.path.join(root, f)
    return None


def shape_ar(text):
    """تهيئة النص العربي للعرض الصحيح في PDF مع حماية السنوات والتواريخ من الانعكاس"""
    if not HAS_PDF_LIBS:
        return text
    protected = re.sub(r"(\d{4}[-/]\d{2,4}(?:/\d{2})?)", "\u202A\\1\u202C", str(text))
    reshaped = arabic_reshaper.reshape(protected)
    visual = get_display(reshaped, base_dir="R")
    return re.sub(r"[\u202A-\u202E]", "", visual)


def format_ar_paragraph(text, font_name, font_size, max_width):
    """تقسيم الفقرة العربية إلى أسطر منطقية قبل التشكيل البصري لمنع انكسار الكلمات المفردة"""
    if not HAS_PDF_LIBS:
        return text
    safe_width = max_width - 4
    raw_lines = str(text).split("\n")
    visual_lines = []
    for raw_line in raw_lines:
        stripped = raw_line.strip()
        if not stripped:
            visual_lines.append("")
            continue
        words = stripped.split(" ")
        current_line = ""
        for w in words:
            candidate = f"{current_line} {w}" if current_line else w
            shaped_cand = shape_ar(candidate)
            if pdfmetrics.stringWidth(shaped_cand, font_name, font_size) <= safe_width:
                current_line = candidate
            else:
                if current_line:
                    visual_lines.append(shape_ar(current_line))
                    current_line = w
                else:
                    visual_lines.append(shape_ar(w))
                    current_line = ""
        if current_line:
            visual_lines.append(shape_ar(current_line))
    return "<br/>".join(visual_lines)


def clear_and_prepare_headers(section):
    """فك ارتباط الترويسات وتفريغ محتواها القديم بأمان"""
    section.header.is_linked_to_previous = False
    section.first_page_header.is_linked_to_previous = False
    for h in [section.header, section.first_page_header]:
        for t in list(h.tables):
            t._tbl.getparent().remove(t._tbl)
        for p in list(h.paragraphs):
            p._p.getparent().remove(p._p)


def clear_all_footers_completely(section):
    """حذف وتفريغ جميع التذييلات القديمة نهائياً"""
    section.footer.is_linked_to_previous = False
    section.first_page_footer.is_linked_to_previous = False
    for f in [section.footer, section.first_page_footer]:
        for t in list(f.tables):
            t._tbl.getparent().remove(t._tbl)
        for p in list(f.paragraphs):
            p._p.getparent().remove(p._p)

        p_clean = f.add_paragraph()
        p_clean.paragraph_format.space_before = Pt(0)
        p_clean.paragraph_format.space_after = Pt(0)

        pPr = p_clean._p.get_or_add_pPr()
        pBdr = pPr.find(qn("w:pBdr"))
        if pBdr is not None:
            pPr.remove(pBdr)


def ensure_valid_header_xml(section):
    """ضمان وجود فقرة في الترويسة غير المستعملة لمنع تلف XML"""
    for h in [section.header, section.first_page_header]:
        if len(h.paragraphs) == 0 and len(h.tables) == 0:
            h.add_paragraph()


def set_cell_rtl(cell):
    """إجبار الخلية على التوجيه من اليمين إلى اليسار"""
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:bidi")) is None:
        tcPr.append(OxmlElement("w:bidi"))


def set_cell_no_wrap(cell):
    """منع انكسار النص داخل الخلايا على مستوى XML"""
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:noWrap")) is None:
        tcPr.append(OxmlElement("w:noWrap"))


def add_styled_run(paragraph, text, is_arabic=True, font_size=14.0, bold=False, font_name=None):
    """إضافة نص منسق يدعم العربية والفرنسية مع ضبط خط النصوص المركبة"""
    if is_arabic:
        chosen_font = font_name if font_name else "Simplified Arabic"
        paragraph.paragraph_format.bidi = True
        pPr = paragraph._p.get_or_add_pPr()
        if pPr.find(qn("w:bidi")) is None:
            pPr.append(OxmlElement("w:bidi"))
        run = paragraph.add_run(text)
        run.font.name = chosen_font
        run.font.size = Pt(font_size)
        run.bold = bold
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), chosen_font)
        rFonts.set(qn("w:hAnsi"), chosen_font)
        rFonts.set(qn("w:cs"), chosen_font)
        if rPr.find(qn("w:rtl")) is None:
            rPr.append(OxmlElement("w:rtl"))
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), str(int(font_size * 2)))
    else:
        chosen_font = font_name if font_name else "Arial"
        run = paragraph.add_run(text)
        run.font.name = chosen_font
        run.font.size = Pt(font_size)
        run.bold = bold
    return run


def add_header_line_to_cell(cell, text, is_arabic, align, is_first=False, font_name=None, base_size=11.0):
    """إضافة سطر مستقل داخل الخلية مع تعديل حجم الخط تلقائياً لمنع الرجوع للسطر"""
    p = cell.paragraphs[0] if is_first else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05

    length = len(text)
    if length > 38:
        font_size = max(8.5, base_size - 2.5)
    elif length > 30:
        font_size = max(9.0, base_size - 1.8)
    elif length > 22:
        font_size = max(9.5, base_size - 1.0)
    else:
        font_size = base_size

    add_styled_run(p, text, is_arabic=is_arabic, font_size=font_size, bold=True, font_name=font_name)
