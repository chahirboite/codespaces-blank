import io
import os
import docx
import streamlit as st
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def clear_and_prepare_headers(section):
    """فك ارتباط الترويسات وتفريغ محتواها القديم بأمان"""
    section.header.is_linked_to_previous = False
    section.first_page_header.is_linked_to_previous = False

    for h in [section.header, section.first_page_header]:
        for t in list(h.tables):
            t._tbl.getparent().remove(t._tbl)
        for p in list(h.paragraphs):
            p._p.getparent().remove(p._p)


def ensure_valid_header_xml(section):
    """ضمان وجود فقرة واحدة على الأقل في الترويسات غير المستعملة لمنع تلف XML"""
    for h in [section.header, section.first_page_header]:
        if len(h.paragraphs) == 0 and len(h.tables) == 0:
            h.add_paragraph()


def get_logo_bytes(uploaded_logo):
    """تحديد الشعار وإرجاعه كـ BytesIO أو مسار ملف محلي"""
    if uploaded_logo is not None:
        return io.BytesIO(uploaded_logo.getvalue())

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
    paragraph, text, is_arabic=True, font_size=10.5, bold=False
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


# --- قائمة مستويات الثانوي التأهيلي المغربي ---
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


# --- إعدادات صفحة Streamlit ---
st.set_page_config(
    page_title="الترويسة المدرسية الذكية 📝", page_icon="📝", layout="centered"
)

st.title("📝 الترويسة المدرسية الذكية")
st.write("دمج وتحديث ترويسة وتذييل الملفات التعليمية تلقائياً وبشكل أنيق.")

# اختيار لغة الترويسة
lang = st.radio(
    "لغة الترويسة / Langue de l'en-tête:",
    ["العربية", "Français"],
    horizontal=True,
)
is_arabic = lang == "العربية"

col_file1, col_file2 = st.columns(2)
with col_file1:
    uploaded_file = st.file_uploader(
        "اختر ملف Word (.docx):", type=["docx"]
    )
with col_file2:
    uploaded_logo = st.file_uploader(
        "رفع شعار المؤسسة (اختياري) / Logo:",
        type=["jpg", "png", "jpeg"],
        help="في حال عدم رفع شعار، سيتم اعتماد logo.jpg تلقائياً إن وجد.",
    )

default_teacher = "ذ. عثمان شهير" if is_arabic else "Pr. Othmane Chahir"
default_subject = "الرياضيات" if is_arabic else "Mathématiques"
default_year = "2026-2027"

col1, col2 = st.columns(2)
with col1:
    teacher_name = st.text_input(
        "اسم الأستاذ / Professeur:", value=default_teacher
    )
    subject = st.text_input("المادة / Matière:", value=default_subject)

with col2:
    levels_options = LEVELS_AR if is_arabic else LEVELS_FR
    selected_level = st.selectbox(
        "المستوى / القسم - Niveau:",
        options=levels_options,
        index=0,
    )

    if selected_level in ["تخصيص مستوى آخر...", "Autre niveau..."]:
        class_level = st.text_input(
            "أدخل اسم المستوى / Autre niveau:", value=""
        )
    else:
        class_level = selected_level

    school_year = st.text_input(
        "الموسم الدراسي / Année scolaire:", value=default_year
    )

st.subheader("⚙️ خيارات التنسيق والهوامش")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    show_all_pages = st.checkbox(
        "إظهار الترويسة في جميع الصفحات",
        value=False,
        help="عند عدم التفعيل (الافتراضي)، تظهر الترويسة في الصفحة الأولى فقط.",
    )
with col_opt2:
    add_footer = st.checkbox(
        "تذييل الصفحات / Pied de page",
        value=False,
        help="تفعيل إظهار تذييل الصفحات لجميع صفحات الملف.",
    )

header_space = st.slider(
    "المسافة الفاصلة أسفل الترويسة (نقطة) / Espace sous l'en-tête (pt):",
    min_value=6,
    max_value=36,
    value=18,
    step=3,
    help="اختر قيمة أصغر لتوفير المساحة عند الطباعة.",
)


def process_docx(
    uploaded_file_obj,
    show_all_pages=False,
    add_footer=False,
    is_arabic=True,
    uploaded_logo=None,
    header_space=18,
):
    file_bytes = uploaded_file_obj.getvalue()
    doc = docx.Document(io.BytesIO(file_bytes))
    logo_file = get_logo_bytes(uploaded_logo)

    for section in doc.sections:
        section.top_margin = Inches(1.2)
        section.header_distance = Inches(0.4)

        clear_and_prepare_headers(section)

        if not show_all_pages:
            section.different_first_page_header_footer = True
            header = section.first_page_header
        else:
            section.different_first_page_header_footer = False
            header = section.header

        # بناء جدول الترويسة
        tbl = header.add_table(rows=1, cols=3, width=Inches(7.2))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.tblPr

        if is_arabic:
            tblPr.insert(0, OxmlElement("w:bidiVisual"))

        # إضافة خط سفلي أنيق وفاصل للجدول
        tblBorders = OxmlElement("w:tblBorders")
        bottom_border = OxmlElement("w:bottom")
        bottom_border.set(qn("w:val"), "single")
        bottom_border.set(qn("w:sz"), "8")  # سمك الخط 1pt
        bottom_border.set(qn("w:space"), "0")
        bottom_border.set(qn("w:color"), "444444")
        tblBorders.append(bottom_border)
        tblPr.append(tblBorders)

        row = tbl.rows[0]

        # ضبط التناظر والمحاذاة العمودية للأعمدة
        col_widths = [Inches(2.8), Inches(1.6), Inches(2.8)]
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        if is_arabic:
            right_text = (
                f"مادة {subject}\n{class_level}\nالموسم الدراسي: {school_year}"
            )
            left_text = f"المديرية الإقليمية تاونات\nثانوية أولاد داود التأهيلية\nالأستاذ: {teacher_name}"
        else:
            left_text = f"Direction Provinciale: Taounate\nLycée Qualifiant Oulad Daoud\nProfesseur: {teacher_name}"
            right_text = f"Matière: {subject}\nNiveau: {class_level}\nAnnée scolaire: {school_year}"

        # العمود الأول
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
            font_size=10.5,
            bold=True,
        )

        # العمود الأوسط (الشعار)
        cell_1 = row.cells[1]
        if is_arabic:
            set_cell_rtl(cell_1)
        p_1 = cell_1.paragraphs[0]
        p_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if logo_file:
            try:
                if isinstance(logo_file, io.BytesIO):
                    logo_file.seek(0)
                p_1.add_run().add_picture(logo_file, width=Inches(1.5))
            except Exception:
                pass

        # العمود الثالث
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
            font_size=10.5,
            bold=True,
        )

        # مسافة فاصلة أسفل الجدول
        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(header_space)

        ensure_valid_header_xml(section)

        # --- معالجة تذييل الصفحات (Pied de page) ---
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
                    tbl_footer._tbl.tblPr.insert(0, OxmlElement("w:bidiVisual"))

                row_f = tbl_footer.rows[0]
                row_f.cells[0].width = Inches(3.6)
                row_f.cells[1].width = Inches(3.6)

                if is_arabic:
                    # الخلية اليمنى: اسم الأستاذ فقط
                    cell_f_r = row_f.cells[0]
                    set_cell_rtl(cell_f_r)
                    p_f_r = cell_f_r.paragraphs[0]
                    p_f_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    add_styled_run(
                        p_f_r, teacher_name, is_arabic, font_size=10, bold=False
                    )

                    # الخلية اليسرى: الموسم الدراسي فقط
                    cell_f_l = row_f.cells[1]
                    set_cell_rtl(cell_f_l)
                    p_f_l = cell_f_l.paragraphs[0]
                    p_f_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    add_styled_run(
                        p_f_l, school_year, is_arabic, font_size=10, bold=False
                    )
                else:
                    # الخلية اليسرى: الموسم الدراسي فقط
                    cell_f_l = row_f.cells[0]
                    p_f_l = cell_f_l.paragraphs[0]
                    p_f_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    add_styled_run(
                        p_f_l, school_year, is_arabic, font_size=10, bold=False
                    )

                    # الخلية اليمنى: اسم الأستاذ فقط
                    cell_f_r = row_f.cells[1]
                    p_f_r = cell_f_r.paragraphs[0]
                    p_f_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    add_styled_run(
                        p_f_r, teacher_name, is_arabic, font_size=10, bold=False
                    )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


if uploaded_file and st.button("توليد الملف المحدث", type="primary"):
    processed_file = process_docx(
        uploaded_file_obj=uploaded_file,
        show_all_pages=show_all_pages,
        add_footer=add_footer,
        is_arabic=is_arabic,
        uploaded_logo=uploaded_logo,
        header_space=header_space,
    )
    st.success("تم تحديث الملف بنجاح!")
    st.download_button(
        label="📥 تحميل الملف الجاهز",
        data=processed_file,
        file_name=f"محدث_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# --- تذييل الصفحة وتوقيع المطور (حقوق التأليف) ---
st.markdown(
    """
    <hr style="margin-top: 50px; margin-bottom: 15px; border: 0; border-top: 1px solid #e0e0e0;">
    <div style="text-align: center; color: #888888; font-size: 13px; font-weight: 500; letter-spacing: 0.5px;">
        Développé par CHAHIR OTHMAN
    </div>
    """,
    unsafe_allow_html=True,
)
