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
    """إجبار الخلية على التوجيه من اليمين إلى اليسار"""
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:bidi")) is None:
        tcPr.append(OxmlElement("w:bidi"))


def add_rtl_run(
    paragraph, text, font_name="Traditional Arabic", font_size=11, bold=False
):
    """تطبيق التوجيه الصريح RTL على الفقرة والنص لمنع تعكيس الكلمات"""
    paragraph.paragraph_format.bidi = True

    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        pPr.append(OxmlElement("w:bidi"))

    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold

    rPr = run._r.get_or_add_rPr()
    if rPr.find(qn("w:rtl")) is None:
        rPr.append(OxmlElement("w:rtl"))

    return run


# --- إعدادات صفحة Streamlit ---
st.set_page_config(
    page_title="الترويسة المدرسية الذكية 📝", page_icon="📝", layout="centered"
)

st.title("📝 الترويسة المدرسية الذكية")
st.write("دمج وتحديث ترويسة وتذييل الملفات التعليمية تلقائياً وبشكل أنيق.")

uploaded_file = st.file_uploader("اختر ملف Word (.docx)", type=["docx"])

col1, col2 = st.columns(2)
with col1:
    teacher_name = st.text_input("اسم الأستاذ:", value="ذ. عثمان شهير")
    subject = st.text_input("المادة:", value="الرياضيات")
with col2:
    class_level = st.text_input(
        "المستوى / القسم:", value="الجذع المشترك العلمي والتكنولوجي"
    )
    school_year = st.text_input("الموسم الدراسي:", value="2026-2027")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    show_all_pages = st.checkbox(
        "هل تريد أن تظهر الترويسة في جميع الصفحات؟",
        value=False,
        help="عند عدم التفعيل (الافتراضي)، تظهر الترويسة في الصفحة الأولى فقط.",
    )
with col_opt2:
    add_footer = st.checkbox(
        "تذييل الصفحات",
        value=False,
        help="تفعيل إظهار تذييل الصفحات لجميع صفحات الملف (يحتوي على اسم الأستاذ والموسم الدراسي).",
    )


def process_docx(file_bytes, show_all_pages=False, add_footer=False):
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
        tbl._tbl.tblPr.append(OxmlElement("w:bidiVisual"))

        row = tbl.rows[0]

        # العمود الأيمن: مادة + مستوى + موسم
        cell_r = row.cells[0]
        set_cell_rtl(cell_r)
        p_r = cell_r.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        right_text = (
            f"مادة {subject}\n{class_level}\nالموسم الدراسي: {school_year}"
        )
        add_rtl_run(p_r, right_text, font_size=11, bold=True)

        # العمود الأوسط: صورة الشعار فقط
        cell_c = row.cells[1]
        set_cell_rtl(cell_c)
        p_c = cell_c.paragraphs[0]
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if logo_path:
            try:
                p_c.add_run().add_picture(logo_path, width=Inches(1.8))
            except Exception:
                pass

        # العمود الأيسر: مديرية + ثانوية + أستاذ
        cell_l = row.cells[2]
        set_cell_rtl(cell_l)
        p_l = cell_l.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        left_text = f"المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: {teacher_name}"
        add_rtl_run(p_l, left_text, font_size=11, bold=True)

        # مسافة فاصلة أسفل الترويسة
        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(24)

        # --- 2. معالجة تذييل الصفحات (Pied de page) ---
        if add_footer:
            footers_to_process = [section.footer]
            if section.different_first_page_header_footer:
                footers_to_process.append(section.first_page_footer)

            for footer_target in footers_to_process:
                # مسح التذييل القديم
                for p in list(footer_target.paragraphs):
                    p._p.getparent().remove(p._p)
                for t in list(footer_target.tables):
                    t._tbl.getparent().remove(t._tbl)

                # إضافة جدول التذييل (عمودان لتوزيع الأستاذ والموسم الدراسي مع وجود مسافة بينهما)
                tbl_footer = footer_target.add_table(
                    rows=1, cols=2, width=Inches(7.2)
                )
                tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl_footer._tbl.tblPr.append(OxmlElement("w:bidiVisual"))

                row_f = tbl_footer.rows[0]

                # الجهة اليمنى: اسم الأستاذ
                cell_f_r = row_f.cells[0]
                set_cell_rtl(cell_f_r)
                p_f_r = cell_f_r.paragraphs[0]
                p_f_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                add_rtl_run(
                    p_f_r,
                    f"الأستاذ: {teacher_name}",
                    font_size=10,
                    bold=False,
                )

                # الجهة اليسرى: الموسم الدراسي
                cell_f_l = row_f.cells[1]
                set_cell_rtl(cell_f_l)
                p_f_l = cell_f_l.paragraphs[0]
                p_f_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
                add_rtl_run(
                    p_f_l,
                    f"الموسم الدراسي: {school_year}",
                    font_size=10,
                    bold=False,
                )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


if uploaded_file and st.button("توليد الملف المحدث"):
    processed_file = process_docx(uploaded_file, show_all_pages, add_footer)
    st.success("تم تحديث الملف بنجاح!")
    st.download_button(
        label="📥 تحميل الملف الجاهز",
        data=processed_file,
        file_name=f"محدث_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
