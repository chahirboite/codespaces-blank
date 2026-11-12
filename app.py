import base64
import datetime
import io
import os
import re
import zipfile
import docx
import streamlit as st
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

# مكتبات توليد PDF باللغة العربية
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Image as RLImage
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    HAS_PDF_LIBS = True
except ImportError:
    HAS_PDF_LIBS = False


# ==========================================
# 1. الدوال المساعدة المشتركة (Utility Functions)
# ==========================================

def get_auto_school_year():
    """حساب الموسم الدراسي تلقائياً"""
    today = datetime.date.today()
    current_year = today.year
    if today.month < 9:
        return f"{current_year - 1}-{current_year}"
    else:
        return f"{current_year}-{current_year + 1}"


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


def find_arabic_font():
    """البحث عن خط عربي مثبت في النظام لتوليد PDF"""
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
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
                if f.lower() in ["dejavusans.ttf", "notosansarabic-regular.ttf", "freesans.ttf"]:
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
    """تقسيم الفقرة العربية إلى أسطر منطقية قبل التشكيل البصري لمنع انعكاس ترتيب الأسطر في PDF"""
    if not HAS_PDF_LIBS:
        return text
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
            if pdfmetrics.stringWidth(shaped_cand, font_name, font_size) <= max_width:
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


# ==========================================
# 2. معالجة هياكل Word (docx XML Helpers)
# ==========================================

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


def add_styled_run(paragraph, text, is_arabic=True, font_size=10.0, bold=False):
    """إضافة نص منسق يدعم العربية والفرنسية"""
    if is_arabic:
        paragraph.paragraph_format.bidi = True
        pPr = paragraph._p.get_or_add_pPr()
        if pPr.find(qn("w:bidi")) is None:
            pPr.append(OxmlElement("w:bidi"))
        run = paragraph.add_run(text)
        run.font.name = "Traditional Arabic"
        run.font.size = Pt(font_size)
        run.bold = bold
        rPr = run._r.get_or_add_rPr()
        if rPr.find(qn("w:rtl")) is None:
            rPr.append(OxmlElement("w:rtl"))
    else:
        run = paragraph.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(font_size)
        run.bold = bold
    return run


def add_header_line_to_cell(cell, text, is_arabic, align, is_first=False):
    """إضافة سطر مستقل داخل الخلية مع تعديل حجم الخط تلقائياً لمنع الرجوع للسطر"""
    p = cell.paragraphs[0] if is_first else cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0

    length = len(text)
    if length > 38:
        font_size = 7.5
    elif length > 30:
        font_size = 8.2
    elif length > 22:
        font_size = 8.8
    else:
        font_size = 10.0

    add_styled_run(p, text, is_arabic=is_arabic, font_size=font_size, bold=True)


# ==========================================
# 3. قوائم وثوابت الخيارات (بصياغة إدارية وتربوية رصينة)
# ==========================================

LEVELS_AR = [
    "الجذع المشترك العلمي والتكنولوجي",
    "الجذع المشترك الأدبي والإنساني",
    "الجذع المشترك الأصيل",
    "الأولى بكالوريا علوم تجريبية",
    "الأولى بكالوريا علوم رياضية",
    "الأولى بكالوريا آداب وعلوم إنسانية",
    "الأولى بكالوريا علوم اقتصادية وتدبير",
    "الثانية بكالوريا علوم فيزيائية",
    "الثانية بكالوريا علوم الحياة والأرض",
    "الثانية بكالوريا علوم رياضية (أ / ب)",
    "الثانية بكالوريا آداب وعلوم إنسانية",
    "الثانية بكالوريا علوم اقتصادية وتدبير",
    "تخصيص مستوى آخر...",
]

LEVELS_FR = [
    "Tronc Commun Scientifique et Technologique",
    "Tronc Commun Lettres et Sciences Humaines",
    "Tronc Commun Enseignement Originel",
    "1ère Bac Sciences Expérimentales",
    "1ère Bac Sciences Mathématiques",
    "1ère Bac Lettres et Sciences Humaines",
    "1ère Bac Sciences Économiques et Gestion",
    "2ème Bac Sciences Physiques",
    "2ème Bac Sciences de la Vie et de la Terre (SVT)",
    "2ème Bac Sciences Mathématiques (A / B)",
    "2ème Bac Lettres et Sciences Humaines",
    "2ème Bac Sciences Économiques et Gestion",
    "Autre niveau...",
]

REPORT_TYPES = {
    "1️⃣ تقرير حول الإخلال بالنظام وعرقلة الدرس": {
        "title": "تقرير تربوي حول الإخلال بالنظام العام وعرقلة السير العادي للحصة الدراسية",
        "situations": [
            "إحداث الفوضى والتشويش المستمر أثناء تقديم الدرس، ومقاطعة الأستاذ بشكل متعمد دون استئذان.",
            "عدم الالتزام بالانضباط الصفي والتنقل غير المبرر داخل القاعة، مما أربك السير العادي للحصة وحرم باقي المتعلمين من حقهم في التحصيل.",
            "التمادي في مخالفة التوجيهات التربوية وعدم الامتثال لتنبيهات الأستاذ المتكررة.",
            "الاستعمال غير المشروع للهاتف النقال / الأجهزة الإلكترونية داخل الفصل الدراسي في خرق سافر لمقتضيات النظام الداخلي للمؤسسة.",
        ],
        "default_action": "استدعاء ولي أمر التلميذ(ة) وإحالة الملف على أنظار مجلس الانضباط لاتخاذ المتعين",
    },
    "2️⃣ تقرير حول السلوك غير اللائق والامتناع عن العمل": {
        "title": "تقرير إداري وتربوي حول سلوك غير لائق والامتناع عن أداء الواجبات الصفية",
        "situations": [
            "الامتناع المتعمد وغير المبرر عن تتبع مجريات الدرس، ورفض تدوين الملخصات وإنجاز الأنشطة الصفية والفروض.",
            "التلفظ بعبارات نابية وغير لائقة، وإبداء سلوك استفزازي يمس بالاحترام الواجب للأطر التربوية وزملائه في القسم.",
            "إلحاق أضرار مادية بتجهيزات القاعة الدراسية والعبث بممتلكات المؤسسة.",
            "مغادرة القاعة الدراسية بصفة أحادية ودون الحصول على إذن مسبق من الأستاذ.",
        ],
        "default_action": "عرض حالة التلميذ(ة) على مجلس الانضباط وتطبيق مقتضيات النظام الداخلي للمؤسسة",
    },
    "3️⃣ تقرير حول التمادي في التأخر والغياب غير المبرر": {
        "title": "تقرير حول التمادي في التأخر والتغيب غير المبرر عن الحصص الدراسية",
        "situations": [
            "التأخر المتكرر وغير المبرر عن موعد انطلاق الحصة الدراسية، مما يربك انطلاقة الدرس.",
            "التغيب المستمر بدون عذر مقبول عن حصص المادة، مع التهاون في تتبع الدروس.",
            "الامتناع عن الإدلاء بورقة السماح بالدخول المسلمة من لدن الإدارة التربوية عقب الغياب.",
        ],
        "default_action": "إشعار ولي الأمر رسمياً واتخاذ الإجراءات الإدارية الجاري بها العمل في شأن المواظبة",
    },
}

# ==========================================
# 4. إعدادات Streamlit وذاكرة الجلسة
# ==========================================

st.set_page_config(
    page_title="منصة الأدوات التربوية 📝", page_icon="📝", layout="centered"
)

if "processed_data" not in st.session_state:
    st.session_state.processed_data = None
if "processed_filename" not in st.session_state:
    st.session_state.processed_filename = ""
if "processed_mime" not in st.session_state:
    st.session_state.processed_mime = ""
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""
if "report_data_docx" not in st.session_state:
    st.session_state.report_data_docx = None
if "report_filename_docx" not in st.session_state:
    st.session_state.report_filename_docx = ""
if "report_data_pdf" not in st.session_state:
    st.session_state.report_data_pdf = None
if "report_filename_pdf" not in st.session_state:
    st.session_state.report_filename_pdf = ""

# ==========================================
# 5. القائمة الجانبية (Sidebar Navigation)
# ==========================================

with st.sidebar:
    st.title("🛠️ المنصة التربوية")

    selected_tool = st.radio(
        "اختر الأداة المطلوبة / Choisir l'outil:",
        ["📝 الترويسة المدرسية الذكية", "📋 محرر التقارير السريعة"],
        index=0,
    )

    st.markdown("---")
    st.subheader("📌 نبذة عن المنصة")
    st.write(
        "منصة تربوية متكاملة طورها **الأستاذ عثمان شهير** لتبسيط وإعادة هيكلة وتوحيد الوثائق التربوية والإدارية لرجال ونساء التعليم بنقرات بسيطة."
    )

    st.markdown("---")
    st.subheader("✨ مميزات المنصة")
    st.markdown(
        """
        * 🚀 **أتمتة كاملة:** توفير الوقت والجهد في الهيكلة والتنسيق.
        * 👁️ **معاينة حية:** مشاهدة تفاعلية للمستند أثناء الكتابة.
        * 📄 **تصدير مزدوج:** تنزيل المستندات بصيغتي Word (`.docx`) و PDF (`.pdf`).
        """
    )

    st.markdown("---")
    st.subheader("💬 التواصل والملاحظات")
    st.caption("رأيك يهمنا لتطوير الأداة! لا تتردد في التواصل المباشر مع المطور.")
    developer_email = "chahirboite@gmail.com"
    st.markdown("**المطور:** OTHMAN CHAHIR")
    st.link_button(
        label="📧 إرسال بريد إلكتروني",
        url=f"mailto:{developer_email}?subject=Plateforme Pedagogique Feedback",
        use_container_width=True,
    )

st.markdown(
    """
    <style>
    .stApp { background-color: #f8fafc; }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        height: 48px;
        background-color: #2563eb;
        color: white;
        border: none;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
    }
    .stButton>button:hover { background-color: #1d4ed8; }
    .paper-preview {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 20px 14px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
        direction: rtl;
        text-align: right;
        font-family: 'Segoe UI', Tahoma, Geneva, sans-serif;
        color: #1e293b;
        margin-top: 15px;
        margin-bottom: 25px;
        box-sizing: border-box;
        width: 100%;
        overflow: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 6. منطق الأداة 1: الترويسة المدرسية الذكية
# ==========================================

def process_docx_header(
    uploaded_file_obj, direction_val, school_val, teacher_name_val, subject_val,
    class_level_val, school_year_val, show_all_pages=False, add_footer=False,
    is_arabic=True, uploaded_logo=None, header_space=18, border_color_hex="#444444",
):
    file_bytes = uploaded_file_obj.getvalue()
    doc = docx.Document(io.BytesIO(file_bytes))
    logo_file = get_logo_bytes(uploaded_logo)

    d_name = direction_val.strip() if direction_val.strip() else ("تاونات" if is_arabic else "Taounate")
    s_name = school_val.strip() if school_val.strip() else ("ثانوية أولاد داود التأهيلية" if is_arabic else "Lycée Qualifiant Oulad Daoud")
    t_name = teacher_name_val.strip() if teacher_name_val.strip() else ("عثمان شهير" if is_arabic else "Othmane Chahir")
    subj = subject_val.strip() if subject_val.strip() else ("الرياضيات" if is_arabic else "Mathématiques")
    c_level = class_level_val.strip()
    s_year = school_year_val

    clean_hex = border_color_hex.lstrip("#")

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.2)
        section.header_distance = Inches(0.4)
        section.footer_distance = Inches(0.4)

        clear_and_prepare_headers(section)

        if not show_all_pages:
            section.different_first_page_header_footer = True
            header = section.first_page_header
        else:
            section.different_first_page_header_footer = False
            header = section.header

        tbl = header.add_table(rows=1, cols=3, width=Inches(7.2))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.tblPr

        if is_arabic:
            tblPr.insert(0, OxmlElement("w:bidiVisual"))

        tblBorders = OxmlElement("w:tblBorders")
        bottom_border = OxmlElement("w:bottom")
        bottom_border.set(qn("w:val"), "single")
        bottom_border.set(qn("w:sz"), "8")
        bottom_border.set(qn("w:space"), "0")
        bottom_border.set(qn("w:color"), clean_hex)
        tblBorders.append(bottom_border)
        tblPr.append(tblBorders)

        row = tbl.rows[0]
        col_widths = [Inches(3.1), Inches(1.0), Inches(3.1)]
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_no_wrap(cell)

        if is_arabic:
            cell_right = row.cells[0]
            set_cell_rtl(cell_right)
            add_header_line_to_cell(cell_right, f"مادة {subj}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True)
            add_header_line_to_cell(cell_right, c_level, is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False)
            add_header_line_to_cell(cell_right, f"الموسم الدراسي: {s_year}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False)

            cell_left = row.cells[2]
            set_cell_rtl(cell_left)
            add_header_line_to_cell(cell_left, f"المديرية الإقليمية: {d_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True)
            add_header_line_to_cell(cell_left, s_name, is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False)
            add_header_line_to_cell(cell_left, f"الأستاذ: {t_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False)
        else:
            cell_left = row.cells[0]
            add_header_line_to_cell(cell_left, f"Direction Provinciale: {d_name}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True)
            add_header_line_to_cell(cell_left, s_name, is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False)
            add_header_line_to_cell(cell_left, f"Professeur: {t_name}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False)

            cell_right = row.cells[2]
            add_header_line_to_cell(cell_right, f"Matière: {subj}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True)
            add_header_line_to_cell(cell_right, c_level, is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False)
            add_header_line_to_cell(cell_right, f"Année scolaire: {s_year}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False)

        cell_1 = row.cells[1]
        if is_arabic:
            set_cell_rtl(cell_1)
        p_1 = cell_1.paragraphs[0]
        p_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if logo_file:
            try:
                if isinstance(logo_file, io.BytesIO):
                    logo_file.seek(0)
                p_1.add_run().add_picture(logo_file, width=Inches(1.2))
            except Exception:
                pass

        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(header_space)

        ensure_valid_header_xml(section)

        if add_footer:
            footers_to_process = [section.footer]
            if section.different_first_page_header_footer:
                footers_to_process.append(section.first_page_footer)

            page_w = section.page_width if section.page_width else Inches(8.27)
            margin_l = section.left_margin if section.left_margin else Inches(0.75)
            margin_r = section.right_margin if section.right_margin else Inches(0.75)
            printable_w = page_w - margin_l - margin_r

            center_pos = int(printable_w / 2)
            right_pos = int(printable_w)

            for footer_target in footers_to_process:
                for p in list(footer_target.paragraphs):
                    p._p.getparent().remove(p._p)
                for t in list(footer_target.tables):
                    t._tbl.getparent().remove(t._tbl)

                p_foot = footer_target.add_paragraph()
                p_foot.paragraph_format.space_before = Pt(0)
                p_foot.paragraph_format.space_after = Pt(0)
                p_foot.paragraph_format.tab_stops.add_tab_stop(center_pos, WD_TAB_ALIGNMENT.CENTER)

                if is_arabic:
                    p_foot.paragraph_format.tab_stops.add_tab_stop(right_pos, WD_TAB_ALIGNMENT.LEFT)
                    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    add_styled_run(p_foot, t_name, is_arabic=True, font_size=10, bold=False)
                    p_foot.add_run("\t\t")
                    add_styled_run(p_foot, s_year, is_arabic=False, font_size=10, bold=False)
                else:
                    p_foot.paragraph_format.tab_stops.add_tab_stop(right_pos, WD_TAB_ALIGNMENT.RIGHT)
                    p_foot.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    add_styled_run(p_foot, s_year, is_arabic=False, font_size=10, bold=False)
                    p_foot.add_run("\t\t")
                    add_styled_run(p_foot, t_name, is_arabic=False, font_size=10, bold=False)
        else:
            clear_all_footers_completely(section)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def render_header_tool():
    lang = st.radio("Langue de l'en-tête / لغة الترويسة:", ["Français", "العربية"], horizontal=True)
    is_arabic = lang == "العربية"
    text_dir = "rtl" if is_arabic else "ltr"

    if is_arabic:
        page_title_html = "📝 الترويسة المدرسية<br><span style='font-size: 0.9em;'>الذكية</span>"
        page_subtitle_text = "حل سريع وأنيق لتخصيص وتوحيد الترويسات المدرسية وشعارات المؤسسات على جميع ملفات Word (.docx) التعليمية."
    else:
        page_title_html = "📝 Générateur d'En-tête<br><span style='font-size: 0.9em;'>Scolaire</span>"
        page_subtitle_text = "Solution rapide et élégante pour personnaliser et uniformiser les en-têtes et pieds de page de vos fichiers Word (.docx)."

    static_logo_uri = get_static_logo_data_uri()
    bg_style = ""
    if static_logo_uri:
        bg_style = f"background-image: linear-gradient(rgba(248, 250, 252, 0.90), rgba(248, 250, 252, 0.90)), url('{static_logo_uri}'); background-size: contain; background-repeat: no-repeat; background-position: center;"

    st.markdown(
        f"""
        <div dir="{text_dir}" style="position: relative; padding: 20px; border-radius: 14px; text-align: center; margin-bottom: 25px; border: 1px solid #e2e8f0; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.04); {bg_style}">
            <h1 style="color: #1e293b; margin: 0; font-size: 1.6rem; font-weight: 800;">{page_title_html}</h1>
            <p style="color: #1e293b; margin-top: 10px; margin-bottom: 0; font-size: 0.95rem; font-weight: 600; line-height: 1.6;">{page_subtitle_text}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_file1, col_file2 = st.columns(2)
    with col_file1:
        uploaded_files = st.file_uploader(
            "📂 " + ("اختر ملفات Word (.docx):" if is_arabic else "Sélectionner les fichiers Word (.docx):"),
            type=["docx"], accept_multiple_files=True
        )
    with col_file2:
        uploaded_logo = st.file_uploader(
            "🖼️ " + ("رفع شعار المؤسسة (اختياري):" if is_arabic else "Logo de l'établissement (Optionnel):"),
            type=["jpg", "png", "jpeg"]
        )

    auto_school_year = get_auto_school_year()

    col_sch1, col_sch2 = st.columns(2)
    with col_sch1:
        direction_name = st.text_input("المديرية الإقليمية / Direction Provinciale:", value="", placeholder="تاونات" if is_arabic else "Taounate")
    with col_sch2:
        school_name = st.text_input("المؤسسة التعليمية / Établissement:", value="", placeholder="ثانوية أولاد داود التأهيلية" if is_arabic else "Lycée Qualifiant Oulad Daoud")

    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("اسم الأستاذ / Professeur:", value="", placeholder="مثال: عثمان شهير" if is_arabic else "Ex: Othmane Chahir")
        subject = st.text_input("المادة / Matière:", value="", placeholder="مثال: الرياضيات" if is_arabic else "Ex: Mathématiques")
    with col2:
        levels_options = LEVELS_AR if is_arabic else LEVELS_FR
        selected_level = st.selectbox("المستوى / القسم - Niveau:", options=levels_options, index=0)
        if selected_level in ["تخصيص مستوى آخر...", "Autre niveau..."]:
            class_level = st.text_input("أدخل اسم المستوى / Autre niveau:", value="")
        else:
            class_level = selected_level

    st.subheader("⚙️ " + ("خيارات التنسيق والهوامش" if is_arabic else "Options de mise en page"))
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        show_all_pages = st.checkbox("إظهار الترويسة في جميع الصفحات" if is_arabic else "Afficher l'en-tête sur toutes les pages", value=False)
    with col_opt2:
        add_footer = st.checkbox("تذييل الصفحات / Pied de page", value=False)

    col_sp1, col_sp2 = st.columns(2)
    with col_sp1:
        header_space = st.slider("المسافة أسفل الترويسة (نقطة):" if is_arabic else "Espace sous l'en-tête (pt):", min_value=6, max_value=36, value=18, step=3)
    with col_sp2:
        border_color = st.color_picker("لون الخط الفاصل أسفل الترويسة:" if is_arabic else "Couleur de la ligne de séparation:", value="#444444")

    btn_label = "🚀 توليد الملفات المحدثة" if is_arabic else "🚀 Générer les fichiers mis à jour"
    if uploaded_files and st.button(btn_label):
        num_files = len(uploaded_files)
        if num_files == 1:
            single_file = uploaded_files[0]
            processed_file = process_docx_header(
                uploaded_file_obj=single_file, direction_val=direction_name, school_val=school_name,
                teacher_name_val=teacher_name, subject_val=subject, class_level_val=class_level,
                school_year_val=auto_school_year, show_all_pages=show_all_pages, add_footer=add_footer,
                is_arabic=is_arabic, uploaded_logo=uploaded_logo, header_space=header_space, border_color_hex=border_color
            )
            st.session_state.processed_data = processed_file
            st.session_state.processed_filename = f"محدث_{single_file.name}"
            st.session_state.processed_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            st.session_state.success_msg = "تم تحديث الملف بنجاح!" if is_arabic else "Le fichier a été mis à jour avec succès!"
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file_obj in uploaded_files:
                    processed_data = process_docx_header(
                        uploaded_file_obj=file_obj, direction_val=direction_name, school_val=school_name,
                        teacher_name_val=teacher_name, subject_val=subject, class_level_val=class_level,
                        school_year_val=auto_school_year, show_all_pages=show_all_pages, add_footer=add_footer,
                        is_arabic=is_arabic, uploaded_logo=uploaded_logo, header_space=header_space, border_color_hex=border_color
                    )
                    zip_file.writestr(f"محدث_{file_obj.name}", processed_data.getvalue())
            zip_buffer.seek(0)
            st.session_state.processed_data = zip_buffer
            st.session_state.processed_filename = "الملفات_المحدثة.zip"
            st.session_state.processed_mime = "application/zip"
            st.session_state.success_msg = f"تم تحديث {num_files} ملفات بنجاح!" if is_arabic else f"{num_files} fichiers mis à jour avec succès!"

    if st.session_state.processed_data is not None and uploaded_files:
        st.success(st.session_state.success_msg)
        st.download_button(
            label="📥 " + ("تحميل الملف الجاهز" if is_arabic else "Télécharger le fichier") if len(uploaded_files) == 1 else "📥 " + ("تحميل جميع الملفات المحدثة (.zip)" if is_arabic else "Télécharger tous les fichiers (.zip)"),
            data=st.session_state.processed_data,
            file_name=st.session_state.processed_filename,
            mime=st.session_state.processed_mime,
        )


# ==========================================
# 7. منطق الأداة 2: محرر التقارير السريعة
# ==========================================

def build_report_texts(student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val):
    """بناء نصوص التقرير بصياغة إدارية وتربوية رسمية موحدة"""
    student_str = student_val.strip() if student_val.strip() else "............................................"
    class_str = class_val.strip() if class_val.strip() else "......................"
    time_str = time_val.strip() if time_val.strip() else "......................"
    room_str = room_val.strip() if room_val.strip() else "......................"
    date_str = date_val.strftime("%Y/%m/%d")

    salutation = "سلام تام بوجود مولانا الإمام المؤيد بالله، وبعد؛"

    intro_text = (
        f"يشرفني، السيد المدير المحترم، أن أرفع إلى كريم علمكم هذا التقرير التربوي قصد الإخبار واتخاذ ما ترونه مناسباً، "
        f"في شأن التصرفات غير المنضبطة الصادرة عن التلميذ(ة): « {student_str} »، المتمدرس(ة) بالقسم: « {class_str} »، "
        f"وذلك خلال حصة مادة « {subject_val} » المنعقدة بتاريخ {date_str} (التوقيت: {time_str} / {room_str})."
    )

    section_heading = "حيثيات الواقعة والأفعال المنسوبة للتلميذ(ة):"

    custom_text = f"معطيات وملاحظات تكميلية: {custom_notes.strip()}" if custom_notes.strip() else ""

    closing_body = (
        f"وحرصاً منا على صون حرمة الفضاء التربوي، وتكريس مبادئ الانضباط والمسؤولية المنصوص عليها في النظام الداخلي للمؤسسة، "
        f"ألتمس من سيادتكم التدخل لاتخاذ الإجراءات التربوية والإدارية اللازمة في حق المعني(ة) بالأمر، ولا سيما: {action_val}."
    )

    closing_salutation = "وتفضلوا، السيد المدير المحترم، بقبول أسمى عبارات التقدير والاحترام.\nوالسلام."

    return {
        "salutation": salutation,
        "intro_text": intro_text,
        "section_heading": section_heading,
        "custom_text": custom_text,
        "closing_body": closing_body,
        "closing_salutation": closing_salutation,
        "date_str": date_str,
    }


def generate_report_docx(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj
):
    doc = docx.Document()
    logo_file = get_logo_bytes(uploaded_logo_obj)

    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)

    clear_and_prepare_headers(section)
    header = section.header

    # جدول الترويسة الهيكلي
    tbl = header.add_table(rows=1, cols=3, width=Inches(6.7))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.tblPr
    tblPr.insert(0, OxmlElement("w:bidiVisual"))

    tblBorders = OxmlElement("w:tblBorders")
    bottom_border = OxmlElement("w:bottom")
    bottom_border.set(qn("w:val"), "single")
    bottom_border.set(qn("w:sz"), "8")
    bottom_border.set(qn("w:space"), "0")
    bottom_border.set(qn("w:color"), "444444")
    tblBorders.append(bottom_border)
    tblPr.append(tblBorders)

    row = tbl.rows[0]
    col_widths = [Inches(2.7), Inches(1.3), Inches(2.7)]
    for i, cell in enumerate(row.cells):
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_no_wrap(cell)

    d_name = direction_val.strip() if direction_val.strip() else "تاونات"
    s_name = school_val.strip() if school_val.strip() else "ثانوية أولاد داود التأهيلية"
    t_name = teacher_val.strip() if teacher_val.strip() else "عثمان شهير"

    cell_r = row.cells[0]
    set_cell_rtl(cell_r)
    add_header_line_to_cell(cell_r, f"المديرية الإقليمية: {d_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True)
    add_header_line_to_cell(cell_r, s_name, is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False)

    cell_c = row.cells[1]
    set_cell_rtl(cell_c)
    p_c = cell_c.paragraphs[0]
    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_file:
        try:
            if isinstance(logo_file, io.BytesIO):
                logo_file.seek(0)
            p_c.add_run().add_picture(logo_file, width=Inches(1.1))
        except Exception:
            pass

    cell_l = row.cells[2]
    set_cell_rtl(cell_l)
    add_header_line_to_cell(cell_l, f"الموسم الدراسي: {year_val}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True)
    add_header_line_to_cell(cell_l, f"الأستاذ(ة): {t_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False)

    p_space = header.add_paragraph()
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(6)

    texts = build_report_texts(
        student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val
    )

    # عنوان التقرير
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(14)
    r_title = add_styled_run(p_title, f"الموضوع: {report_title}", is_arabic=True, font_size=14, bold=True)
    r_title.font.underline = True

    # الافتتاحية
    p_sal = doc.add_paragraph()
    p_sal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sal.paragraph_format.space_before = Pt(2)
    p_sal.paragraph_format.space_after = Pt(10)
    add_styled_run(p_sal, texts["salutation"], is_arabic=True, font_size=12.5, bold=True)

    # ديباجة التقرير
    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.paragraph_format.space_before = Pt(4)
    p_body1.paragraph_format.space_after = Pt(8)
    p_body1.paragraph_format.line_spacing = 1.3
    add_styled_run(p_body1, texts["intro_text"], is_arabic=True, font_size=12, bold=False)

    # تفاصيل الوقائع
    p_situations_head = doc.add_paragraph()
    p_situations_head.paragraph_format.space_before = Pt(6)
    p_situations_head.paragraph_format.space_after = Pt(4)
    add_styled_run(p_situations_head, texts["section_heading"], is_arabic=True, font_size=12, bold=True)

    if situations_list:
        for sit in situations_list:
            p_bullet = doc.add_paragraph(style='List Bullet')
            p_bullet.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_bullet.paragraph_format.space_before = Pt(2)
            p_bullet.paragraph_format.space_after = Pt(3)
            p_bullet.paragraph_format.line_spacing = 1.2
            add_styled_run(p_bullet, sit, is_arabic=True, font_size=11.5, bold=False)

    if texts["custom_text"]:
        p_custom = doc.add_paragraph()
        p_custom.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_custom.paragraph_format.space_before = Pt(4)
        p_custom.paragraph_format.space_after = Pt(4)
        add_styled_run(p_custom, texts["custom_text"], is_arabic=True, font_size=11.5, bold=False)

    # الإجراء والخاتمة
    p_action = doc.add_paragraph()
    p_action.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_action.paragraph_format.space_before = Pt(10)
    p_action.paragraph_format.space_after = Pt(8)
    p_action.paragraph_format.line_spacing = 1.3
    add_styled_run(p_action, texts["closing_body"], is_arabic=True, font_size=12, bold=False)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(4)
    p_end.paragraph_format.space_after = Pt(16)
    p_end.paragraph_format.line_spacing = 1.25
    add_styled_run(p_end, texts["closing_salutation"], is_arabic=True, font_size=12, bold=False)

    # 1. التواقيع داخل جسم الصفحة مع مساحة كافية وأسطر منقطة لرأي الإدارة
    sig_tbl = doc.add_table(rows=1, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.autofit = False
    sig_tblPr = sig_tbl._tbl.tblPr
    sig_tblPr.insert(0, OxmlElement("w:bidiVisual"))

    sigBorders = OxmlElement("w:tblBorders")
    for b_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "none")
        sigBorders.append(b)
    sig_tblPr.append(sigBorders)

    cell_sig_r = sig_tbl.rows[0].cells[0]
    cell_sig_r.width = Inches(3.2)
    cell_sig_r.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_rtl(cell_sig_r)
    p_sig_r = cell_sig_r.paragraphs[0]
    p_sig_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_r.paragraph_format.space_after = Pt(75)
    add_styled_run(p_sig_r, "توقيع الأستاذ(ة):", is_arabic=True, font_size=11.5, bold=True)

    cell_sig_l = sig_tbl.rows[0].cells[1]
    cell_sig_l.width = Inches(3.5)
    cell_sig_l.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_rtl(cell_sig_l)
    p_sig_l = cell_sig_l.paragraphs[0]
    p_sig_l.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_l.paragraph_format.space_after = Pt(8)
    add_styled_run(p_sig_l, "رأي وتوقيع الإدارة التربوية:", is_arabic=True, font_size=11.5, bold=True)

    for idx in range(3):
        p_dots = cell_sig_l.add_paragraph()
        p_dots.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_dots.paragraph_format.space_before = Pt(4)
        p_dots.paragraph_format.space_after = Pt(28 if idx == 2 else 6)
        add_styled_run(p_dots, "...........................................................", is_arabic=True, font_size=11, bold=False)

    # 2. تذييل الصفحات الرسمي (Pied de page): اسم الأستاذ والتاريخ فقط
    clear_all_footers_completely(section)
    footer = section.footer
    p_foot = footer.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(0)
    p_foot.paragraph_format.space_after = Pt(0)

    center_pos = int(Inches(3.35))
    right_pos = int(Inches(6.7))

    p_foot.paragraph_format.tab_stops.add_tab_stop(center_pos, WD_TAB_ALIGNMENT.CENTER)
    p_foot.paragraph_format.tab_stops.add_tab_stop(right_pos, WD_TAB_ALIGNMENT.LEFT)
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    add_styled_run(p_foot, f"الأستاذ(ة): {t_name}", is_arabic=True, font_size=9.5, bold=False)
    p_foot.add_run("\t\t")
    add_styled_run(p_foot, f"حرر في: {texts['date_str']}", is_arabic=True, font_size=9.5, bold=False)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def generate_report_pdf(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj
):
    """توليد نسخة PDF رسمية باللغة العربية جاهزة للطباعة بترتيب أسطر سليم"""
    if not HAS_PDF_LIBS:
        return None
    font_path = find_arabic_font()
    if not font_path:
        return None

    try:
        pdfmetrics.registerFont(TTFont("ArabicFont", font_path))
    except Exception:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=45,
        leftMargin=45,
        topMargin=36,
        bottomMargin=36,
    )

    usable_width = A4[0] - 90

    style_right = ParagraphStyle("ArRight", fontName="ArabicFont", fontSize=11, leading=18, alignment=TA_RIGHT)
    style_center = ParagraphStyle("ArCenter", fontName="ArabicFont", fontSize=11.5, leading=18, alignment=TA_CENTER)
    style_title = ParagraphStyle("ArTitle", fontName="ArabicFont", fontSize=12.5, leading=19, alignment=TA_CENTER)
    style_hdr_r = ParagraphStyle("ArHdrR", fontName="ArabicFont", fontSize=9.5, leading=14, alignment=TA_RIGHT)
    style_hdr_l = ParagraphStyle("ArHdrL", fontName="ArabicFont", fontSize=9.5, leading=14, alignment=TA_LEFT)
    style_foot_r = ParagraphStyle("ArFootR", fontName="ArabicFont", fontSize=8.5, leading=12, alignment=TA_RIGHT, textColor=colors.grey)
    style_foot_l = ParagraphStyle("ArFootL", fontName="ArabicFont", fontSize=8.5, leading=12, alignment=TA_LEFT, textColor=colors.grey)

    d_name = direction_val.strip() if direction_val.strip() else "تاونات"
    s_name = school_val.strip() if school_val.strip() else "ثانوية أولاد داود التأهيلية"
    t_name = teacher_val.strip() if teacher_val.strip() else "عثمان شهير"

    texts = build_report_texts(
        student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val
    )

    story = []

    # 1. الترويسة
    logo_file = get_logo_bytes(uploaded_logo_obj)
    logo_cell = ""
    if logo_file:
        try:
            if isinstance(logo_file, io.BytesIO):
                logo_file.seek(0)
            logo_cell = RLImage(logo_file, width=65, height=48, kind="proportional")
        except Exception:
            logo_cell = ""

    hdr_data = [[
        Paragraph(f"{shape_ar(f'الموسم الدراسي: {year_val}')}<br/>{shape_ar(f'الأستاذ(ة): {t_name}')}", style_hdr_l),
        logo_cell,
        Paragraph(f"{shape_ar(f'المديرية الإقليمية: {d_name}')}<br/>{shape_ar(s_name)}", style_hdr_r),
    ]]
    hdr_table = Table(hdr_data, colWidths=[200, 105, 200])
    hdr_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor("#334155")),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
    ]))
    story.append(hdr_table)
    story.append(Spacer(1, 14))

    # 2. الموضوع والعنوان
    wrapped_title = format_ar_paragraph(f"الموضوع: {report_title}", "ArabicFont", 12.5, usable_width)
    story.append(Paragraph(f"<u>{wrapped_title}</u>", style_title))
    story.append(Spacer(1, 12))

    # 3. الافتتاحية والديباجة
    story.append(Paragraph(shape_ar(texts["salutation"]), style_center))
    story.append(Spacer(1, 10))

    wrapped_intro = format_ar_paragraph(texts["intro_text"], "ArabicFont", 11, usable_width)
    story.append(Paragraph(wrapped_intro, style_right))
    story.append(Spacer(1, 10))

    # 4. الوقائع
    story.append(Paragraph(shape_ar(texts["section_heading"]), style_right))
    story.append(Spacer(1, 5))
    for sit in situations_list:
        wrapped_sit = format_ar_paragraph(f"• {sit}", "ArabicFont", 11, usable_width - 10)
        story.append(Paragraph(wrapped_sit, style_right))
        story.append(Spacer(1, 4))

    if texts["custom_text"]:
        story.append(Spacer(1, 4))
        wrapped_custom = format_ar_paragraph(texts["custom_text"], "ArabicFont", 11, usable_width)
        story.append(Paragraph(wrapped_custom, style_right))

    story.append(Spacer(1, 10))

    # 5. الإجراء والخاتمة
    wrapped_closing = format_ar_paragraph(texts["closing_body"], "ArabicFont", 11, usable_width)
    story.append(Paragraph(wrapped_closing, style_right))
    story.append(Spacer(1, 8))

    wrapped_end = format_ar_paragraph(texts["closing_salutation"], "ArabicFont", 11.5, usable_width)
    story.append(Paragraph(wrapped_end, style_center))
    story.append(Spacer(1, 18))

    # 6. التواقيع ومساحة رأي الإدارة بالنقط
    admin_col = (
        f"{shape_ar('رأي وتوقيع الإدارة التربوية:')}<br/><br/>"
        f"..................................................<br/><br/>"
        f"..................................................<br/><br/>"
        f".................................................."
    )
    teacher_col = f"{shape_ar('توقيع الأستاذ(ة):')}"

    sig_data = [[
        Paragraph(admin_col, style_right),
        Paragraph(teacher_col, style_right),
    ]]
    sig_table = Table(sig_data, colWidths=[260, 245])
    sig_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 40),
    ]))
    story.append(sig_table)
    story.append(Spacer(1, 16))

    # 7. التذييل السفلي
    foot_data = [[
        Paragraph(shape_ar(f"حرر في: {texts['date_str']}"), style_foot_l),
        Paragraph(shape_ar(f"الأستاذ(ة): {t_name}"), style_foot_r),
    ]]
    foot_table = Table(foot_data, colWidths=[252, 253])
    foot_table.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 0.6, colors.lightgrey),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
    ]))
    story.append(foot_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


def render_live_preview_html(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj
):
    """توليد المعاينة الحية التفاعلية بالصياغة الإدارية الرسمية"""
    logo_src = get_logo_base64_src(uploaded_logo_obj)
    logo_html = f'<img src="{logo_src}" style="max-height: 48px; max-width: 90px; height: auto; width: auto; object-fit: contain;" />' if logo_src else ''

    d_name = direction_val.strip() if direction_val.strip() else "تاونات"
    s_name = school_val.strip() if school_val.strip() else "ثانوية أولاد داود التأهيلية"
    t_name = teacher_val.strip() if teacher_val.strip() else "عثمان شهير"

    student_disp = student_val.strip() if student_val.strip() else "<span style='color:#94a3b8;'>[اسم التلميذ(ة)]</span>"
    class_disp = class_val.strip() if class_val.strip() else "<span style='color:#94a3b8;'>[المستوى]</span>"
    time_disp = time_val.strip() if time_val.strip() else "[التوقيت]"
    room_disp = room_val.strip() if room_val.strip() else "[القاعة]"
    date_str = date_val.strftime("%Y/%m/%d")

    bullets_html = "".join([f"<li style='margin-bottom: 5px; text-align: justify;'>{s}</li>" for s in situations_list])
    custom_html = f"<p style='margin-top: 6px; font-weight: 500; text-align: justify;'><b>معطيات وملاحظات تكميلية:</b> {custom_notes.strip()}</p>" if custom_notes.strip() else ""

    html = f"""<div class="paper-preview" style="box-sizing: border-box; width: 100%; max-width: 100%; overflow: hidden; padding: 18px 14px;">
<div style="display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 4px; border-bottom: 2px solid #334155; padding-bottom: 8px; margin-bottom: 15px;">
<div style="text-align: right; font-size: clamp(0.60rem, 2.0vw, 0.82rem); line-height: 1.4; font-weight: bold; overflow: hidden;">
<div style="white-space: nowrap;">المديرية الإقليمية: {d_name}</div>
<div style="white-space: nowrap;">{s_name}</div>
</div>
<div style="text-align: center; display: flex; justify-content: center; align-items: center; padding: 0 2px;">
{logo_html}
</div>
<div style="text-align: left; font-size: clamp(0.60rem, 2.0vw, 0.82rem); line-height: 1.4; font-weight: bold; overflow: hidden;">
<div style="white-space: nowrap;">الموسم الدراسي: <span style="display: inline-block; direction: ltr; unicode-bidi: embed; white-space: nowrap;">{year_val}</span></div>
<div style="white-space: nowrap;">الأستاذ(ة): {t_name}</div>
</div>
</div>
<div style="text-align: center; margin: 14px 0;">
<h3 style="margin: 0; font-size: clamp(0.90rem, 3.0vw, 1.10rem); text-decoration: underline; color: #0f172a; line-height: 1.5;">الموضوع: {report_title}</h3>
</div>
<div style="text-align: center; font-weight: bold; font-size: clamp(0.85rem, 2.8vw, 0.98rem); margin-bottom: 12px;">
سلام تام بوجود مولانا الإمام المؤيد بالله، وبعد؛
</div>
<p style="line-height: 1.75; text-align: justify; font-size: clamp(0.83rem, 2.7vw, 0.94rem); margin-bottom: 12px; word-break: break-word;">
يشرفني، السيد المدير المحترم، أن أرفع إلى كريم علمكم هذا التقرير التربوي قصد الإخبار واتخاذ ما ترونه مناسباً، 
في شأن التصرفات غير المنضبطة الصادرة عن التلميذ(ة): <b>« {student_disp} »</b>، المتمدرس(ة) بالقسم: <b>« {class_disp} »</b>، 
وذلك خلال حصة مادة <b>« {subject_val} »</b> المنعقدة بتاريخ <b>{date_str}</b> (التوقيت: {time_disp} / {room_disp}).
</p>
<div style="font-weight: bold; font-size: clamp(0.83rem, 2.7vw, 0.94rem); margin-top: 10px; margin-bottom: 6px;">حيثيات الواقعة والأفعال المنسوبة للتلميذ(ة):</div>
<ul style="padding-right: 20px; font-size: clamp(0.80rem, 2.6vw, 0.90rem); line-height: 1.65; margin-top: 4px; margin-bottom: 10px;">
{bullets_html}
</ul>
{custom_html}
<p style="line-height: 1.75; text-align: justify; font-size: clamp(0.83rem, 2.7vw, 0.94rem); margin-top: 14px; margin-bottom: 10px; word-break: break-word;">
وحرصاً منا على صون حرمة الفضاء التربوي، وتكريس مبادئ الانضباط والمسؤولية المنصوص عليها في النظام الداخلي للمؤسسة، 
ألتمس من سيادتكم التدخل لاتخاذ الإجراءات التربوية والإدارية اللازمة في حق المعني(ة) بالأمر، ولا سيما: <b>{action_val}</b>.
</p>
<div style="text-align: center; font-size: clamp(0.83rem, 2.7vw, 0.94rem); line-height: 1.6; margin-top: 8px; margin-bottom: 20px;">
وتفضلوا، السيد المدير المحترم، بقبول أسمى عبارات التقدير والاحترام.<br>
<b>والسلام.</b>
</div>
<div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-top: 25px; margin-bottom: 70px; padding-top: 10px; font-size: clamp(0.75rem, 2.5vw, 0.9rem);">
<div style="width: 45%; text-align: right; font-weight: bold;">
<div>توقيع الأستاذ(ة):</div>
</div>
<div style="width: 50%; text-align: right;">
<div style="font-weight: bold; margin-bottom: 8px;">رأي وتوقيع الإدارة التربوية:</div>
<div style="color: #64748b; line-height: 1.9; font-size: clamp(0.70rem, 2.2vw, 0.85rem); overflow: hidden; white-space: nowrap;">
....................................................<br>
....................................................<br>
....................................................
</div>
</div>
</div>
<div style="display: flex; justify-content: space-between; align-items: center; margin-top: 20px; border-top: 1px dashed #cbd5e1; padding-top: 8px; font-size: clamp(0.65rem, 2.2vw, 0.78rem); color: #64748b;">
<div>الأستاذ(ة): {t_name}</div>
<div>حرر في: {date_str}</div>
</div>
</div>"""
    return html


def render_report_tool():
    st.markdown(
        """
        <div style="text-align: center; padding: 15px; border-radius: 10px; background-color: #ffffff; border: 1px solid #e2e8f0; margin-bottom: 20px;">
            <h1 style="color: #1e293b; margin: 0; font-size: 1.6rem;">📝 محرر التقارير السريعة</h1>
            <p style="color: #64748b; margin-top: 5px; font-size: 0.95rem;">توليد تقارير انضباطية رسمية ومسبوكة بصيغتي Word (.docx) و PDF جاهزة للطباعة والتوقيع</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("1️⃣ البيانات العامة والترويسة")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        direction_name = st.text_input("المديرية الإقليمية:", value="تاونات")
        school_name = st.text_input("المؤسسة التعليمية:", value="ثانوية أولاد داود التأهيلية")
    with col_g2:
        teacher_name = st.text_input("اسم الأستاذ(ة):", value="عثمان شهير")
        school_year = st.text_input("الموسم الدراسي:", value=get_auto_school_year())

    uploaded_logo = st.file_uploader(
        "🖼️ رفع شعار المؤسسة الخاص (اختياري - يحل محل الشعار الافتراضي):",
        type=["jpg", "png", "jpeg"],
    )

    st.subheader("2️⃣ بيانات التلميذ والواقعة")
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        student_name = st.text_input("اسم التلميذ(ة):", placeholder="مثال: محمد العلمي")
    with col_s2:
        class_level = st.text_input("القسم / المستوى:", placeholder="مثال: الثانية بكالوريا علوم فيزيائية 1")
    with col_s3:
        subject_name = st.text_input("المادة الدراسية:", value="الرياضيات")

    incident_date = datetime.date.today()

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        incident_time = st.text_input("توقيت الحصة:", placeholder="مثال: من 10:00 إلى 11:00")
    with col_t2:
        room_number = st.text_input("رقم القاعة / المختبر:", value="القاعة 4")

    st.subheader("3️⃣ تحديد نوع التقرير والمخالفات")
    selected_type_key = st.radio(
        "اختر نوع التقرير المناسب (تتغير الصيغة والمخالفات تلقائياً):",
        list(REPORT_TYPES.keys()),
        horizontal=False,
    )

    current_config = REPORT_TYPES[selected_type_key]
    selected_situations = st.multiselect(
        "اختر السلوكات والمخالفات الملاحظة (يمكن اختيار أكثر من موقف):",
        options=current_config["situations"],
        default=[current_config["situations"][0]],
    )

    custom_details = st.text_area(
        "إضافة أو تحرير تفاصيل الواقعة (اختياري):",
        placeholder="أدخل أي معلومات تفصيلية إضافية توضح ملابسات الواقعة...",
    )

    proposed_action = st.text_input(
        "الإجراء الإداري / التأديبي المقترح:",
        value=current_config["default_action"],
    )

    # --- قسم المعاينة الحية التفاعلية ---
    st.markdown("---")
    st.subheader("👁️ المعاينة الحية للتصميم والوثيقة الرسمية")
    st.caption("يتحدث نموذج الوثيقة أوتوماتيكياً أسفله مع كل تغيير تقوم به في الاستمارة أعلاه:")

    preview_html_code = render_live_preview_html(
        direction_val=direction_name, school_val=school_name, teacher_val=teacher_name,
        year_val=school_year, student_val=student_name, class_val=class_level,
        subject_val=subject_name, date_val=incident_date, time_val=incident_time,
        room_val=room_number, report_title=current_config["title"],
        situations_list=selected_situations, custom_notes=custom_details,
        action_val=proposed_action, uploaded_logo_obj=uploaded_logo,
    )
    st.markdown(preview_html_code, unsafe_allow_html=True)

    # --- أزرار التوليد والتنزيل (Word & PDF) ---
    st.markdown("---")
    if st.button("🚀 توليد وتجهيز التقرير (Word & PDF)", type="primary"):
        if not student_name.strip():
            st.warning("⚠️ يرجى إدخال اسم التلميذ(ة) لتوليد التقرير.")
        else:
            safe_name = student_name.strip().replace(" ", "_")
            docx_bytes = generate_report_docx(
                direction_val=direction_name, school_val=school_name, teacher_val=teacher_name,
                year_val=school_year, student_val=student_name, class_val=class_level,
                subject_val=subject_name, date_val=incident_date, time_val=incident_time,
                room_val=room_number, report_title=current_config["title"],
                situations_list=selected_situations, custom_notes=custom_details,
                action_val=proposed_action, uploaded_logo_obj=uploaded_logo,
            )
            st.session_state.report_data_docx = docx_bytes
            st.session_state.report_filename_docx = f"تقرير_{safe_name}.docx"

            pdf_bytes = generate_report_pdf(
                direction_val=direction_name, school_val=school_name, teacher_val=teacher_name,
                year_val=school_year, student_val=student_name, class_val=class_level,
                subject_val=subject_name, date_val=incident_date, time_val=incident_time,
                room_val=room_number, report_title=current_config["title"],
                situations_list=selected_situations, custom_notes=custom_details,
                action_val=proposed_action, uploaded_logo_obj=uploaded_logo,
            )
            st.session_state.report_data_pdf = pdf_bytes
            st.session_state.report_filename_pdf = f"تقرير_{safe_name}.pdf"

    if st.session_state.report_data_docx is not None:
        st.success("✅ التقرير جاهز للتحميل!")
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📥 تحميل بصيغة Word (.docx)",
                data=st.session_state.report_data_docx,
                file_name=st.session_state.report_filename_docx,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with col_dl2:
            if st.session_state.report_data_pdf is not None:
                st.download_button(
                    label="📕 تحميل بصيغة PDF (.pdf)",
                    data=st.session_state.report_data_pdf,
                    file_name=st.session_state.report_filename_pdf,
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.info("ℹ️ لتفعيل تنزيل PDF المباشر، تأكد من وجود `reportlab` و `arabic-reshaper` و `python-bidi` في ملف `requirements.txt`.")


# ==========================================
# 8. التوجيه الرئيسي وتوصيل الأجزاء (Router)
# ==========================================

if selected_tool == "📝 الترويسة المدرسية الذكية":
    render_header_tool()
else:
    render_report_tool()

# تذييل الصفحة الثابت وتوقيع المطور
st.markdown(
    """
    <hr style="margin-top: 50px; margin-bottom: 15px; border: 0; border-top: 1px solid #e0e0e0;">
    <div style="text-align: center; color: #888888; font-size: 13px; font-weight: 500; letter-spacing: 0.5px;">
        Développé par CHAHIR OTHMAN
    </div>
    """,
    unsafe_allow_html=True,
)
