import io
import os
import docx
import streamlit as st
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def clear_existing_headers(section):
    """حذف أي ترويسات سابقة لتفادي التداخل"""
    section.header.is_linked_to_previous = False
    for h in [section.header, section.first_page_header]:
        for p in list(h.paragraphs):
            p._p.getparent().remove(p._p)
        for t in list(h.tables):
            t._tbl.getparent().remove(t._tbl)


def get_logo_path():
    """البحث عن ملف الشعار تلقائياً"""
    for filename in ["logo.jpg", "logo.png", "logo.jpeg", "LOGO.JPG", "LOGO.PNG"]:
        if os.path.exists(filename):
            return filename
    return None


def set_cell_rtl(cell):
    """إجبار الخلية على التوجيه من اليمين إلى اليسار (خاص باللغة العربية)"""
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:bidi")) is None:
        tcPr.append(OxmlElement("w:bidi"))


def add_styled_run(
    paragraph, text, is_arabic=True, font_size=11, bold=False
):
    """إضافة نص مع مراعاة اتجاه اللغة (RTL للعربية و LTR للفرنسية)"""
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
        paragraph.paragraph_format.bidi = False
        run = paragraph.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(font_size)
        run.bold = bold

    return run


# --- إعدادات صفحة Streamlit ---
st.set_page_config(
    page_title="الترويسة المدرسية الذكية 📝", page_icon="📝", layout="centered"
)

st.title("📝 الترويسة المدرسية الذكية")
st.write("دمج وتحديث ترويسة وتذييل الملفات التعليمية تلقائياً وبشكل أنيق.")

uploaded_file = st.file_uploader("اختر ملف Word (.docx)", type=["docx"])

# اختيار لغة الترويسة
lang = st.radio(
    "لغة الترويسة / Langue de l'en-tête:",
    ["العربية", "Français"],
    horizontal=True,
)
is_arabic = lang == "العربية"

# تعيين القيم الافتراضية بناءً على اللغة المختارة
default_teacher = "ذ. عثمان شهير" if is_arabic else "Pr. Othmane Chahir"
default_subject = "الرياضيات" if is_arabic else "Mathématiques"
default_class = (
    "الجذع المشترك العلمي والتكنولوجي"
    if is_arabic
    else "Tronc Commun Scientifique"
)
default_year = "2026-2027"

col1, col2 = st.columns(2)
with col1:
    teacher_name = st.text_input(
        "اسم الأستاذ / Professeur:", value=default_teacher
    )
    subject = st.text_input("المادة / Matière:", value=default_subject)
with col2:
    class_level = st.text_input(
        "المستوى / القسم - Niveau:", value=default_class
    )
    school_year = st.text_input(
        "الموسم الدراسي / Année scolaire:", value=default_year
    )

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    show_all_pages = st.checkbox(
        "هل تريد أن تظهر الترويسة في جميع الصفحات؟",
        value=False,
        help="عند عدم التفعيل (الافتراضي)، تظهر الترويسة في الصفحة الأولى فقط.",
    )
with col_opt2:
    add_footer = st.checkbox(
        "تذييل الصفحات / Pied de page",
        value=False,
        help="تفعيل إظهار تذييل الصفحات لجميع صفحات الملف (يحتوي على اسم الأستاذ والموسم الدراسي).",
    )


def process_docx(
    file_bytes, show_all_pages=False, add_footer=False, is_arabic=True
):
    doc = docx.Document(file_bytes)
    logo_path = get_logo_path()

    for section in doc.sections:
        # --- 1. معالجة الترويسة (En-tête) ---
        clear_existing_headers(section)

        if not show_all_pages:
            section.different_first_page_header_footer = True
            header = section.first_page_header
        else:
            section.different_first_page_header_footer = False
            header = section.header

        tbl = header.add_table(rows=1, cols=3, width=Inches(7.2))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        # تطبيق اتجاه الجدول الخاص بالعربية
        if is_arabic:
            tbl._tbl.tblPr.append(OxmlElement("w:bidiVisual"))

        row = tbl.rows[0]

        # صياغة النصوص حسب اللغة
        if is_arabic:
            right_text = (
                f"مادة {subject}\n{class_level}\nالموسم الدراسي: {school_year}"
            )
            left_text = f"المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: {teacher_name}"
        else:
            left_text = f"Direction Provinciale: Taounate\nLycée Qualifiant Oulad Daoud\nProfesseur: {teacher_name}"
            right_text = f"Matière: {subject}\nNiveau: {class_level}\nAnnée scolaire: {school_year}"

        # العمود الأيمن (في العربية: مادة/مستوى | في الفرنسية: مادة/مستوى)
        cell_0 = row.cells[0]
        if is_arabic:
            set_cell_rtl(cell_0)
        p_0 = cell_0.paragraphs[0]
        p_0.alignment = (
            WD_ALIGN_PARAGRAPH.RIGHT if is_arabic else WD_ALIGN_PARAGRAPH.LEFT
        )
        add_styled_run(
            p_0,
            right_text if is_arabic else left_text,
            is_arabic,
            font_size=11,
            bold=True,
        )

        # العمود الأوسط (صورة الشعار)
        cell_1 = row.cells[1]
        if is_arabic:
            set_cell_rtl(cell_1)
        p_1 = cell_1.paragraphs[0]
        p_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if logo_path:
            try:
                p_1.add_run().add_picture(logo_path, width=Inches(1.8))
            except Exception:
                pass

        # العمود الأيسر (في العربية: مديرية/أستاذ | في الفرنسية: مديرية/أستاذ)
        cell_2 = row.cells[2]
        if is_arabic:
            set_cell_rtl(cell_2)
        p_2 = cell_2.paragraphs[0]
        p_2.alignment = (
            WD_ALIGN_PARAGRAPH.LEFT if is_arabic else WD_ALIGN_PARAGRAPH.RIGHT
        )
        add_styled_run(
            p_2,
            left_text if is_arabic else right_text,
            is_arabic,
            font_size=11,
            bold=True,
        )

        # مسافة فاصلة أسفل الترويسة لتفادي التداخل
        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(24)

        # --- 2. معالجة تذييل الصفحات (Pied de page) ---
        if add_footer:
            footers_to_process = [section.footer]
            if section.different_first_page_header_footer:
                footers_to_process.append(section.first_page_footer)

            for footer_target in footers_to_process:
                for p in list(footer_target.paragraphs):
                    p._p.getparent().remove(p._p)
                for t in list(footer_target.tables):
                    t._tbl.getparent().remove(t._tbl)

                tbl_footer = footer_target.add_table(
                    rows=1, cols=2, width=Inches(7.2)
                )
                tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
                if is_arabic:
                    tbl_footer._tbl.tblPr.append(OxmlElement("w:bidiVisual"))

                row_f = tbl_footer.rows[0]

                # الجهة 1: اسم الأستاذ
                cell_f_0 = row_f.cells[0]
                if is_arabic:
                    set_cell_rtl(cell_f_0)
                p_f_0 = cell_f_0.paragraphs[0]
                p_f_0.alignment = (
                    WD_ALIGN_PARAGRAPH.RIGHT
                    if is_arabic
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                f_text_0 = (
                    f"الأستاذ: {teacher_name}"
                    if is_arabic
                    else f"Professeur: {teacher_name}"
                )
                add_styled_run(
                    p_f_0, f_text_0, is_arabic, font_size=10, bold=False
                )

                # الجهة 2: الموسم الدراسي
                cell_f_1 = row_f.cells[1]
                if is_arabic:
                    set_cell_rtl(cell_f_1)
                p_f_1 = cell_f_1.paragraphs[0]
                p_f_1.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT
                    if is_arabic
                    else WD_ALIGN_PARAGRAPH.RIGHT
                )
                f_text_1 = (
                    f"الموسم الدراسي: {school_year}"
                    if is_arabic
                    else f"Année scolaire: {school_year}"
                )
                add_styled_run(
                    p_f_1, f_text_1, is_arabic, font_size=10, bold=False
                )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


if uploaded_file and st.button("توليد الملف المحدث"):
    processed_file = process_docx(
        uploaded_file, show_all_pages, add_footer, is_arabic
    )
    st.success("تم تحديث الملف بنجاح!")
    st.download_button(
        label="📥 تحميل الملف الجاهز",
        data=processed_file,
        file_name=f"محدث_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
