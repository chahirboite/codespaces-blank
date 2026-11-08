import base64
import datetime
import io
import os
import zipfile
import docx
import streamlit as st
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def get_auto_school_year():
    """حساب الموسم الدراسي تلقائياً خلف الكواليس"""
    today = datetime.date.today()
    current_year = today.year
    if today.month < 9:
        return f"{current_year - 1}-{current_year}"
    else:
        return f"{current_year}-{current_year + 1}"


def get_static_logo_data_uri():
    """قراءة صورة الشعار المحلي الثابت لاستخدامها كخلفية دائمية خلف العنوان"""
    for filename in ["logo.jpg", "logo.png", "logo.jpeg", "LOGO.JPG", "LOGO.PNG"]:
        if os.path.exists(filename):
            try:
                ext = filename.split(".")[-1].lower()
                mime_type = (
                    "image/jpeg" if ext in ["jpg", "jpeg"] else "image/png"
                )
                with open(filename, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode()
                return f"data:{mime_type};base64,{encoded}"
            except Exception:
                pass
    return None


def get_logo_bytes(uploaded_logo):
    """تحديد الشعار الخاص بملف Word"""
    if uploaded_logo is not None:
        return io.BytesIO(uploaded_logo.getvalue())

    for filename in ["logo.jpg", "logo.png", "logo.jpeg", "LOGO.JPG", "LOGO.PNG"]:
        if os.path.exists(filename):
            return filename
    return None


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
    """ضمان وجود فقرة واحدة على الأقل في الترويسات غير المستعملة لمنع تلف XML"""
    for h in [section.header, section.first_page_header]:
        if len(h.paragraphs) == 0 and len(h.tables) == 0:
            h.add_paragraph()


def set_cell_rtl(cell):
    """إجبار الخلية على التوجيه من اليمين إلى اليسار"""
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:bidi")) is None:
        tcPr.append(OxmlElement("w:bidi"))


def add_styled_run(
    paragraph, text, is_arabic=True, font_size=10.5, bold=False
):
    """إضافة نص مع مراعاة اتجاه اللغة"""
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
    page_title="Générateur d'En-tête Scolaire 📝", page_icon="📝", layout="centered"
)

# 1. اختيار لغة الواجهة والترويسة
lang = st.radio(
    "Langue de l'en-tête / لغة الترويسة:",
    ["Français", "العربية"],
    horizontal=True,
)
is_arabic = lang == "العربية"
text_dir = "rtl" if is_arabic else "ltr"

# 2. تحديد العنوان والوصف التفصيلي
if is_arabic:
    page_title_html = "📝 الترويسة المدرسية<br><span style='font-size: 0.9em;'>الذكية</span>"
    page_subtitle_text = "حل سريع وأنيق لتخصيص وتوحيد الترويسات المدرسية وشعارات المؤسسات على جميع ملفات Word (.docx) التعليمية."
else:
    page_title_html = "📝 Générateur d'En-tête<br><span style='font-size: 0.9em;'>Scolaire</span>"
    page_subtitle_text = "Solution rapide et élégante pour personnaliser, automatiser et uniformiser les en-têtes et pieds de page de vos fichiers Word (.docx) pédagogiques."

words = page_subtitle_text.split()
animated_subtitle_html = "".join(
    [
        f'<span class="animated-word" style="animation-delay: {i * 0.05:.2f}s;">{word}</span> '
        for i, word in enumerate(words)
    ]
)

# 3. بناء خلفية العنوان الثابتة
header_container = st.container()
static_logo_uri = get_static_logo_data_uri()
bg_style = ""
if static_logo_uri:
    bg_style = f"background-image: linear-gradient(rgba(248, 250, 252, 0.90), rgba(248, 250, 252, 0.90)), url('{static_logo_uri}'); background-size: contain; background-repeat: no-repeat; background-position: center;"

with header_container:
    st.markdown(
        f"""
        <style>
        @keyframes wordPop {{
            0% {{
                opacity: 0;
                transform: translateY(8px) scale(0.95);
            }}
            100% {{
                opacity: 1;
                transform: translateY(0) scale(1);
            }}
        }}
        .animated-word {{
            display: inline-block;
            opacity: 0;
            animation: wordPop 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
            margin: 0 2px;
        }}
        .responsive-title {{
            color: #1e293b;
            margin: 0;
            font-size: clamp(1.2rem, 4vw, 1.6rem);
            font-weight: 800;
            font-family: 'Segoe UI', Tahoma, Geneva, sans-serif;
            line-height: 1.3;
            text-align: center;
        }}
        </style>
        <div dir="{text_dir}" style="
            position: relative;
            padding: 24px 16px;
            border-radius: 14px;
            text-align: center;
            margin-bottom: 25px;
            border: 1px solid #e2e8f0;
            background-color: #ffffff;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
            direction: {text_dir};
            {bg_style}
        ">
            <h1 class="responsive-title" style="direction: {text_dir};">
                {page_title_html}
            </h1>
            <p style="color: #1e293b; margin-top: 12px; margin-bottom: 0; font-size: clamp(0.85rem, 2.9vw, 0.98rem); font-weight: 600; line-height: 1.6; max-width: 650px; margin-left: auto; margin-right: auto; direction: {text_dir}; text-align: center;">
                {animated_subtitle_html}
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# 4. أزرار وحقول رفع الملفات
col_file1, col_file2 = st.columns(2)
with col_file1:
    uploaded_files = st.file_uploader(
        "📂 " + ("اختر ملفات Word (.docx):" if is_arabic else "Sélectionner les fichiers Word (.docx):"),
        type=["docx"],
        accept_multiple_files=True,
        help="يمكنك اختيار ملف واحد أو عدة ملفات دفعة واحدة." if is_arabic else "Vous pouvez sélectionner un ou plusieurs fichiers à la fois.",
    )
with col_file2:
    uploaded_logo = st.file_uploader(
        "🖼️ " + ("رفع شعار المؤسسة (اختياري):" if is_arabic else "Logo de l'établissement (Optionnel):"),
        type=["jpg", "png", "jpeg"],
        help="سيتم دمج الشعار المرفوع في ترويسة ملفات Word المنتَجة." if is_arabic else "Le logo sera intégré dans les documents Word.",
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
    </style>
    """,
    unsafe_allow_html=True,
)

auto_school_year = get_auto_school_year()

# 5. خانات المديرية والمؤسسة
col_sch1, col_sch2 = st.columns(2)
with col_sch1:
    direction_name = st.text_input(
        "المديرية الإقليمية / Direction Provinciale:",
        value="",
        placeholder="تاونات" if is_arabic else "Taounate",
    )
with col_sch2:
    school_name = st.text_input(
        "المؤسسة التعليمية / Établissement:",
        value="",
        placeholder="ثانوية أولاد داود التأهيلية"
        if is_arabic
        else "Lycée Qualifiant Oulad Daoud",
    )

# 6. باقي خانات الإدخال
col1, col2 = st.columns(2)
with col1:
    teacher_name = st.text_input(
        "اسم الأستاذ / Professeur:",
        value="",
        placeholder="مثال: عثمان شهير" if is_arabic else "Ex: Othmane Chahir",
    )
    subject = st.text_input(
        "المادة / Matière:",
        value="",
        placeholder="مثال: الرياضيات" if is_arabic else "Ex: Mathématiques",
    )

with col2:
    levels_options = LEVELS_AR if is_arabic else LEVELS_FR
    selected_level = st.selectbox(
        "المستوى / القسم - Niveau:",
        options=levels_options,
        index=0,
    )

    if selected_level in ["تخصيص مستوى آخر...", "Autre niveau..."]:
        class_level = st.text_input(
            "أدخل اسم المستوى / Autre niveau:",
            value="",
            placeholder="أدخل اسم المستوى..." if is_arabic else "Saisir le niveau...",
        )
    else:
        class_level = selected_level

st.subheader("⚙️ " + ("خيارات التنسيق والهوامش" if is_arabic else "Options de mise en page"))

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    show_all_pages = st.checkbox(
        "إظهار الترويسة في جميع الصفحات" if is_arabic else "Afficher l'en-tête sur toutes les pages",
        value=False,
        help="عند عدم التفعيل (الافتراضي)، تظهر الترويسة في الصفحة الأولى فقط." if is_arabic else "Par défaut, l'en-tête apparaît uniquement sur la 1ère page.",
    )
with col_opt2:
    add_footer = st.checkbox(
        "تذييل الصفحات / Pied de page",
        value=False,
        help="تفعيل إظهار تذييل الصفحات لجميع صفحات الملف." if is_arabic else "Activer le pied de page sur toutes les pages.",
    )

col_sp1, col_sp2 = st.columns(2)
with col_sp1:
    header_space = st.slider(
        "المسافة أسفل الترويسة (نقطة):" if is_arabic else "Espace sous l'en-tête (pt):",
        min_value=6,
        max_value=36,
        value=18,
        step=3,
        help="اختر قيمة أصغر لتوفير المساحة عند الطباعة." if is_arabic else "Diminuer la valeur pour économiser de l'espace à l'impression.",
    )
with col_sp2:
    border_color = st.color_picker(
        "لون الخط الفاصل أسفل الترويسة:" if is_arabic else "Couleur de la ligne de séparation:",
        value="#444444",
        help="اختر لون الخط الفاصل بين الترويسة ومحتوى المستند." if is_arabic else "Couleur du trait sous l'en-tête.",
    )


def process_docx(
    uploaded_file_obj,
    direction_val,
    school_val,
    teacher_name_val,
    subject_val,
    class_level_val,
    school_year_val,
    show_all_pages=False,
    add_footer=False,
    is_arabic=True,
    uploaded_logo=None,
    header_space=18,
    border_color_hex="#444444",
):
    file_bytes = uploaded_file_obj.getvalue()
    doc = docx.Document(io.BytesIO(file_bytes))
    logo_file = get_logo_bytes(uploaded_logo)

    d_name = (
        direction_val.strip()
        if direction_val.strip()
        else ("تاونات" if is_arabic else "Taounate")
    )
    s_name = (
        school_val.strip()
        if school_val.strip()
        else ("ثانوية أولاد داود التأهيلية" if is_arabic else "Lycée Qualifiant Oulad Daoud")
    )
    t_name = (
        teacher_name_val.strip()
        if teacher_name_val.strip()
        else ("عثمان شهير" if is_arabic else "Othmane Chahir")
    )
    subj = (
        subject_val.strip()
        if subject_val.strip()
        else ("الرياضيات" if is_arabic else "Mathématiques")
    )
    c_level = class_level_val.strip() if class_level_val.strip() else selected_level
    s_year = school_year_val

    clean_hex = border_color_hex.lstrip("#")

    for section in doc.sections:
        # إبعاد محتوى الصفحة عن التذييل وضبط مسافة التذييل أسفل الصفحة
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

        col_widths = [Inches(2.8), Inches(1.6), Inches(2.8)]
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        if is_arabic:
            right_text = (
                f"مادة {subj}\n{c_level}\nالموسم الدراسي: {s_year}"
            )
            left_text = f"المديرية الإقليمية: {d_name}\n{s_name}\nالأستاذ: {t_name}"
        else:
            left_text = f"Direction Provinciale: {d_name}\n{s_name}\nProfesseur: {t_name}"
            right_text = f"Matière: {subj}\nNiveau: {c_level}\nAnnée scolaire: {s_year}"

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

        p_space = header.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(header_space)

        ensure_valid_header_xml(section)

        # --- معالجة التذييل المتماثل ديناميكياً مع مركز الصفحة ---
        if add_footer:
            footers_to_process = [section.footer]
            if section.different_first_page_header_footer:
                footers_to_process.append(section.first_page_footer)

            # حساب عرض الطباعة ومركز الصفحة الهندسي بدقة
            page_w = section.page_width if section.page_width else Inches(8.27)
            margin_l = section.left_margin if section.left_margin else Inches(0.75)
            margin_r = section.right_margin if section.right_margin else Inches(0.75)
            printable_w = page_w - margin_l - margin_r

            center_pos = printable_w / 2
            right_pos = printable_w

            for footer_target in footers_to_process:
                for p in list(footer_target.paragraphs):
                    p._p.getparent().remove(p._p)
                for t in list(footer_target.tables):
                    t._tbl.getparent().remove(t._tbl)

                p_foot = footer_target.add_paragraph()
                p_foot.paragraph_format.space_before = Pt(0)
                p_foot.paragraph_format.space_after = Pt(0)
                p_foot.paragraph_format.left_indent = Inches(0)
                p_foot.paragraph_format.right_indent = Inches(0)

                # إضافة نقطة تبويب ممتدة للمركز بالتموضع المتماثل صراحة
                p_foot.paragraph_format.tab_stops.add_tab_stop(
                    center_pos, WD_TAB_ALIGNMENT.CENTER
                )

                if is_arabic:
                    p_foot.paragraph_format.tab_stops.add_tab_stop(
                        right_pos, WD_TAB_ALIGNMENT.LEFT
                    )
                    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                    # 1. العنصر الأيمن: اسم الأستاذ
                    add_styled_run(
                        p_foot, t_name, is_arabic=True, font_size=10, bold=False
                    )
                    # 2. التبويب الأول (الانتقال لمركز الصفحة - محجوز لرقم الصفحة)
                    p_foot.add_run("\t")
                    # 3. التبويب الثاني (الانتقال لأقصى اليسار المتماثل)
                    p_foot.add_run("\t")
                    # 4. العنصر الأيسر: الموسم الدراسي
                    add_styled_run(
                        p_foot, s_year, is_arabic=False, font_size=10, bold=False
                    )
                else:
                    p_foot.paragraph_format.tab_stops.add_tab_stop(
                        right_pos, WD_TAB_ALIGNMENT.RIGHT
                    )
                    p_foot.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    # 1. العنصر الأيسر: الموسم الدراسي
                    add_styled_run(
                        p_foot, s_year, is_arabic=False, font_size=10, bold=False
                    )
                    # 2. التبويب الأول (الانتقال لمركز الصفحة - محجوز لرقم الصفحة)
                    p_foot.add_run("\t")
                    # 3. التبويب الثاني (الانتقال لأقصى اليمين المتماثل)
                    p_foot.add_run("\t")
                    # 4. العنصر الأيمن: اسم الأستاذ
                    add_styled_run(
                        p_foot, t_name, is_arabic=False, font_size=10, bold=False
                    )
        else:
            clear_all_footers_completely(section)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# زر المعالجة والتوليد
btn_label = "🚀 توليد الملفات المحدثة" if is_arabic else "🚀 Générer les fichiers mis à jour"
if uploaded_files and st.button(btn_label):
    num_files = len(uploaded_files)

    if num_files == 1:
        single_file = uploaded_files[0]
        processed_file = process_docx(
            uploaded_file_obj=single_file,
            direction_val=direction_name,
            school_val=school_name,
            teacher_name_val=teacher_name,
            subject_val=subject,
            class_level_val=class_level,
            school_year_val=auto_school_year,
            show_all_pages=show_all_pages,
            add_footer=add_footer,
            is_arabic=is_arabic,
            uploaded_logo=uploaded_logo,
            header_space=header_space,
            border_color_hex=border_color,
        )
        st.success("تم تحديث الملف بنجاح!" if is_arabic else "Le fichier a été mis à jour avec succès!")
        st.download_button(
            label="📥 " + ("تحميل الملف الجاهز" if is_arabic else "Télécharger le fichier"),
            data=processed_file,
            file_name=f"محدث_{single_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_obj in uploaded_files:
                processed_data = process_docx(
                    uploaded_file_obj=file_obj,
                    direction_val=direction_name,
                    school_val=school_name,
                    teacher_name_val=teacher_name,
                    subject_val=subject,
                    class_level_val=class_level,
                    school_year_val=auto_school_year,
                    show_all_pages=show_all_pages,
                    add_footer=add_footer,
                    is_arabic=is_arabic,
                    uploaded_logo=uploaded_logo,
                    header_space=header_space,
                    border_color_hex=border_color,
                )
                zip_file.writestr(
                    f"محدث_{file_obj.name}", processed_data.getvalue()
                )

        zip_buffer.seek(0)
        st.success(f"تم تحديث {num_files} ملفات بنجاح!" if is_arabic else f"{num_files} fichiers mis à jour avec succès!")
        st.download_button(
            label="📥 " + ("تحميل جميع الملفات المحدثة (.zip)" if is_arabic else "Télécharger tous les fichiers (.zip)"),
            data=zip_buffer,
            file_name="الملفات_المحدثة.zip",
            mime="application/zip",
        )

# --- تذييل الصفحة وتوقيع المطور ---
st.markdown(
    """
    <hr style="margin-top: 50px; margin-bottom: 15px; border: 0; border-top: 1px solid #e0e0e0;">
    <div style="text-align: center; color: #888888; font-size: 13px; font-weight: 500; letter-spacing: 0.5px;">
        Développé par CHAHIR OTHMAN
    </div>
    """,
    unsafe_allow_html=True,
)
