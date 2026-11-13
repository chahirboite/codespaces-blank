import io
import zipfile
import docx
import streamlit as st
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from utils import (
    add_header_line_to_cell,
    add_styled_run,
    clear_all_footers_completely,
    clear_and_prepare_headers,
    ensure_valid_header_xml,
    get_auto_school_year,
    get_logo_bytes,
    get_static_logo_data_uri,
    set_cell_no_wrap,
    set_cell_rtl,
)


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
    ar_header_font = "Simplified Arabic"
    ar_header_size = 11.5

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
            add_header_line_to_cell(cell_right, f"مادة {subj}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True, font_name=ar_header_font, base_size=ar_header_size)
            add_header_line_to_cell(cell_right, c_level, is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False, font_name=ar_header_font, base_size=ar_header_size)
            add_header_line_to_cell(cell_right, f"الموسم الدراسي: {s_year}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False, font_name=ar_header_font, base_size=ar_header_size)

            cell_left = row.cells[2]
            set_cell_rtl(cell_left)
            add_header_line_to_cell(cell_left, f"المديرية الإقليمية: {d_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True, font_name=ar_header_font, base_size=ar_header_size)
            add_header_line_to_cell(cell_left, s_name, is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False, font_name=ar_header_font, base_size=ar_header_size)
            add_header_line_to_cell(cell_left, f"الأستاذ: {t_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False, font_name=ar_header_font, base_size=ar_header_size)
        else:
            cell_left = row.cells[0]
            add_header_line_to_cell(cell_left, f"Direction Provinciale: {d_name}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True, base_size=10.0)
            add_header_line_to_cell(cell_left, s_name, is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False, base_size=10.0)
            add_header_line_to_cell(cell_left, f"Professeur: {t_name}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False, base_size=10.0)

            cell_right = row.cells[2]
            add_header_line_to_cell(cell_right, f"Matière: {subj}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True, base_size=10.0)
            add_header_line_to_cell(cell_right, c_level, is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False, base_size=10.0)
            add_header_line_to_cell(cell_right, f"Année scolaire: {s_year}", is_arabic=False, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False, base_size=10.0)

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
                    add_styled_run(p_foot, t_name, is_arabic=True, font_size=11, bold=False, font_name="Simplified Arabic")
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
    if "processed_data" not in st.session_state:
        st.session_state.processed_data = None
    if "processed_filename" not in st.session_state:
        st.session_state.processed_filename = ""
    if "processed_mime" not in st.session_state:
        st.session_state.processed_mime = ""
    if "success_msg" not in st.session_state:
        st.session_state.success_msg = ""

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
        direction_name = st.text_input("المديرية الإقليمية / Direction Provinciale:", value="", placeholder="مثال: تاونات" if is_arabic else "Ex: Taounate")
    with col_sch2:
        school_name = st.text_input("المؤسسة التعليمية / Établissement:", value="", placeholder="مثال: ثانوية أولاد داود التأهيلية" if is_arabic else "Ex: Lycée Qualifiant Oulad Daoud")

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
