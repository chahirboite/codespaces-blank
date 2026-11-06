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
    section.header.is_linked_to_previous = False
    for h in [section.header, section.first_page_header]:
        for p in list(h.paragraphs):
            p._p.getparent().remove(p._p)
        for t in list(h.tables):
            t._tbl.getparent().remove(t._tbl)


def get_logo_path():
    """البحث عن ملف الصورة سواء كان logo.jpg أو logo.png"""
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
    """تطبيق التوجيه الصريح RTL على الفقرة والنص (Run) لمنع تعكيس الكلمات في Word"""
    paragraph.paragraph_format.bidi = True

    # ضبط وسم الفقرة
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        pPr.append(OxmlElement("w:bidi"))

    # إضافة النص وضبط الخط
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold

    # إجبار Word على تطبيق RTL على مستوى النص نفسه
    rPr = run._r.get_or_add_rPr()
    if rPr.find(qn("w:rtl")) is None:
        rPr.append(OxmlElement("w:rtl"))

    return run


st.set_page_config(
    page_title="مدمج الترويسة المدرسية", page_icon="📝", layout="centered"
)

st.title("📝 تطبيق دمج الترويسة المدرسية الآلي")
st.write(
    "قم برفع ملف Word، وستتم إضافة الترويسة الرسمية وإلغاء أي ترويسة قديمة تلقائياً."
)

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

show_all_pages = st.checkbox(
    "هل تريد أن تظهر الترويسة في جميع الصفحات؟",
    value=False,
    help="عند عدم التفعيل (الافتراضي)، تظهر الترويسة في الصفحة الأولى فقط.",
)


def process_docx(file_bytes, show_all_pages=False):
    doc = docx.Document(file_bytes)
    logo_path = get_logo_path()

    for section in doc.sections:
        clear_existing_headers(section)

        # تحديد إعدادات التكرار (الصفحة الأولى فقط أو جميع الصفحات)
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

        # 1. العمود الأيمن: مادة + مستوى + موسم
        cell_r = row.cells[0]
        set_cell_rtl(cell_r)
        p_r = cell_r.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        right_text = (
            f"مادة {subject}\n{class_level}\nالموسم الدراسي: {school_year}"
        )
        add_rtl_run(p_r, right_text, font_size=11, bold=True)

        # 2. العمود الأوسط: صورة الشعار فقط (logo.jpg / logo.png)
        cell_c = row.cells[1]
        set_cell_rtl(cell_c)
        p_c = cell_c.paragraphs[0]
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if logo_path:
            try:
                p_c.add_run().add_picture(logo_path, width=Inches(1.8))
            except Exception:
                pass

        # 3. العمود الأيسر: مديرية + ثانوية + أستاذ
        cell_l = row.cells[2]
        set_cell_rtl(cell_l)
        p_l = cell_l.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        left_text = f"المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: {teacher_name}"
        add_rtl_run(p_l, left_text, font_size=11, bold=True)

        # مسافة فاصلة تعادل سطرين أسفل الترويسة لتفادي التداخل مع متن الملف
        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(24)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


if uploaded_file and st.button("توليد الملف المحدث"):
    processed_file = process_docx(uploaded_file, show_all_pages)
    st.success("تم تحديث الملف بنجاح!")
    st.download_button(
        label="📥 تحميل الملف الجاهز",
        data=processed_file,
        file_name=f"محدث_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
