import io
import docx
import streamlit as st
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches


def clear_existing_headers(section):
    section.header.is_linked_to_previous = False
    header = section.header
    for p in list(header.paragraphs):
        p._p.getparent().remove(p._p)
    for t in list(header.tables):
        t._tbl.getparent().remove(t._tbl)


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


def process_docx(file_bytes):
    doc = docx.Document(file_bytes)
    logo_path = "logo.jpg"

    for section in doc.sections:
        clear_existing_headers(section)

        header = section.header
        tbl = header.add_table(rows=1, cols=3, width=Inches(7.2))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl._tbl.tblPr.append(OxmlElement("w:bidiVisual"))

        row = tbl.rows[0]

        p_r = row.cells[0].paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_r.add_run(
            f"المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: {teacher_name}"
        )

        try:
            p_c = row.cells[1].paragraphs[0]
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_c.add_run().add_picture(logo_path, width=Inches(2.1))
        except Exception:
            pass

        p_l = row.cells[2].paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_l.add_run(
            f"مادة {subject}\n{class_level}\nالموسم الدراسي: {school_year}"
        )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


if uploaded_file and st.button("توليد الملف المحدث"):
    processed_file = process_docx(uploaded_file)
    st.success("تم تحديث الملف بنجاح!")
    st.download_button(
        label="📥 تحميل الملف الجاهز",
        data=processed_file,
        file_name=f"محدث_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )