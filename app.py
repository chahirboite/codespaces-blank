import io
import os
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor
import streamlit as st

# --- إعدادات صفحة Streamlit ---
st.set_page_config(
    page_title="مُنشئ ترويسة الوورد مع الشعار",
    page_icon="📄",
    layout="centered"
)

# --- دوال ضبط اتجاه الكتابة من اليمين إلى اليسار (RTL) ---
def set_cell_rtl(cell):
    """ضبط الخلية لتدعم اتجاه RTL"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(OxmlElement('w:bidi'))

def set_table_rtl(table):
    """ضبط اتجاه أعمدة الجدول مرئياً من اليمين إلى اليسار"""
    tblPr = table._tbl.tblPr
    tblPr.append(OxmlElement('w:bidiVisual'))

# --- دالة بناء الترويسة وإدراج الشعار والمسافات ---
def create_header(doc, right_text, center_text, left_text, logo_file=None, show_on_all_pages=False):
    section = doc.sections[0]
    
    # 1. إعدادات تكرار الترويسة (الافتراضي: الصفحة الأولى فقط)
    if not show_on_all_pages:
        section.different_first_page_header_footer = True
        header_target = section.first_page_header
    else:
        section.different_first_page_header_footer = False
        header_target = section.header

    # 2. إنشاء جدول الترويسة (3 أعمدة) وتحديد اتجاه RTL
    table = header_target.add_table(rows=1, cols=3, width=Inches(6.5))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_rtl(table)

    # --- العمود الأيمن (معلومات المادة والجذع) ---
    cell_right = table.cell(0, 0)
    set_cell_rtl(cell_right)
    p_r = cell_right.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.bidi = True
    run_r = p_r.add_run(right_text)
    run_r.font.name = 'Traditional Arabic'
    run_r.font.size = Pt(11)
    run_r.font.bold = True

    # --- العمود الأوسط (المملكة والشعار) ---
    cell_center = table.cell(0, 1)
    set_cell_rtl(cell_center)
    p_c = cell_center.paragraphs[0]
    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c.paragraph_format.bidi = True

    # كتابة النص العلوي
    run_c = p_c.add_run(center_text + "\n")
    run_c.font.name = 'Traditional Arabic'
    run_c.font.size = Pt(10)

    # إدراج الشعار (استخدام الملف المرفوع إن وجد، أو الاعتماد على logo.jpg المحلي)
    img_target = logo_file
    if img_target is None:
        if os.path.exists("logo.jpg"):
            img_target = "logo.jpg"
        elif os.path.exists("logo.png"):
            img_target = "logo.png"

    if img_target:
        run_img = p_c.add_run()
        run_img.add_picture(img_target, width=Inches(0.85))

    # --- العمود الأيسر (المديرية والمؤسسة والأستاذ) ---
    cell_left = table.cell(0, 2)
    set_cell_rtl(cell_left)
    p_l = cell_left.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l.paragraph_format.bidi = True
    run_l = p_l.add_run(left_text)
    run_l.font.name = 'Traditional Arabic'
    run_l.font.size = Pt(11)
    run_l.font.bold = True

    # 3. إبعاد الترويسة بمسافة تعادل سطرين أسفل الجدول لمنع التداخل مع النص
    p_spacer = header_target.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(24)

# --- دالة توليد المستند ---
def generate_docx(right_text, center_text, left_text, title_text, content_text, logo_file, show_all_pages):
    doc = docx.Document()

    # بناء الترويسة بالخيارات والشعار
    create_header(doc, right_text, center_text, left_text, logo_file=logo_file, show_on_all_pages=show_all_pages)

    # عنوان الدرس
    if title_text:
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.bidi = True
        run_title = p_title.add_run(title_text)
        run_title.font.name = 'Traditional Arabic'
        run_title.font.size = Pt(20)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(200, 0, 0)

    # محتوى النص
    if content_text:
        p_content = doc.add_paragraph()
        p_content.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_content.paragraph_format.bidi = True
        run_content = p_content.add_run(content_text)
        run_content.font.name = 'Traditional Arabic'
        run_content.font.size = Pt(14)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- واجهة المستخدم (Streamlit UI) ---
st.title("📝 مُنشئ ترويسة المستندات مع الشعار")

st.subheader("1. بيانات الترويسة والشعار")
col1, col2, col3 = st.columns(3)

with col1:
    right_text = st.text_area(
        "العمود الأيمن",
        value="مادة الرياضيات\nالجذع المشترك العلمي والتكنولوجي\nالموسم الدراسي: 2026-2027",
        height=110
    )

with col2:
    center_text = st.text_area(
        "العمود الأوسط",
        value="المملكة المغربية\nوزارة التربية الوطنية والتعليم الأولي والرياضة",
        height=110
    )

with col3:
    left_text = st.text_area(
        "العمود الأيسر",
        value="المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: ذ. هشام الشومب",
        height=110
    )

# رفع الشعار (اختياري)
uploaded_logo = st.file_uploader("رفع شعار آخر (اختياري - افتراضياً يستخدم logo.jpg المرفق)", type=["jpg", "png", "jpeg"])

st.subheader("2. خيارات الصفحات والتنسيق")
# خيار تحديد ظهور الترويسة (غير مفعل افتراضياً = الصفحة الأولى فقط)
show_all_pages = st.checkbox(
    "هل تريد أن تظهر الترويسة في جميع الصفحات؟",
    value=False,
    help="عند ترك هذا الخيار غير مفعل (الافتراضي)، ستظهر الترويسة في الصفحة الأولى فقط دون غيرها."
)

st.subheader("3. محتوى الملف")
title_text = st.text_input("عنوان الدرس", value="درس: الإشتقاق")
content_text = st.text_area("محتوى البداية (اختياري)", value="لتكن f دالة عددية معرفة على مجال I مفتوح مركزه x0...", height=100)

st.markdown("---")

if st.button("🚀 إنشاء وتنسيق المستند", type="primary"):
    file_buffer = generate_docx(right_text, center_text, left_text, title_text, content_text, uploaded_logo, show_all_pages)
    
    # تحديد اسم الملف المحمل ديناميكياً من عنوان الدرس بدلاً من الاسم المباشر
    clean_title = "".join(c for c in title_text if c.isalnum() or c in (" ", "_", "-")).strip()
    download_filename = f"{clean_title}.docx" if clean_title else "document.docx"
    
    st.success("تم إنشاء الملف وتطبيق جميع التنسيقات بنجاح!")
    st.download_button(
        label="📥 تحميل ملف الوورد (.docx)",
        data=file_buffer,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
