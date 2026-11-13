import datetime
import io
import docx
import streamlit as st
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Image as RLImage
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
except ImportError:
    pass

from utils import (
    HAS_PDF_LIBS,
    add_header_line_to_cell,
    add_styled_run,
    clear_all_footers_completely,
    clear_and_prepare_headers,
    format_ar_paragraph,
    format_room_label,
    get_arabic_pdf_font,
    get_auto_school_year,
    get_logo_base64_src,
    get_logo_bytes,
    set_cell_no_wrap,
    set_cell_rtl,
    shape_ar,
)


FONT_PROFILES = {
    "نمط 1": {
        "word_font": "Simplified Arabic",
        "word_body_size": 14.0,
        "word_bullet_size": 13.5,
        "word_title_size": 15.5,
        "word_sal_size": 14.0,
        "word_header_base": 11.5,
        "word_footer_size": 10.5,
        "word_line_spacing": 1.25,
        "css_font": "'Amiri', 'Noto Naskh Arabic', serif",
        "css_body_size": "clamp(0.95rem, 3.1vw, 1.08rem)",
        "css_title_size": "clamp(1.05rem, 3.5vw, 1.22rem)",
        "css_line_height": "1.85",
        "pdf_font_key": "amiri",
        "pdf_body_size": 12.5,
        "pdf_title_size": 14.5,
        "pdf_leading": 19,
    },
    "نمط 2": {
        "word_font": "Arial",
        "word_body_size": 12.0,
        "word_bullet_size": 11.5,
        "word_title_size": 14.0,
        "word_sal_size": 12.5,
        "word_header_base": 10.5,
        "word_footer_size": 10.0,
        "word_line_spacing": 1.3,
        "css_font": "'Noto Naskh Arabic', 'Cairo', sans-serif",
        "css_body_size": "clamp(0.90rem, 2.9vw, 1.02rem)",
        "css_title_size": "clamp(1.00rem, 3.3vw, 1.18rem)",
        "css_line_height": "1.8",
        "pdf_font_key": "noto",
        "pdf_body_size": 11.0,
        "pdf_title_size": 13.0,
        "pdf_leading": 18,
    },
}

# 3 صيغ تعبيرية مختلفة لنص التقرير يتم التبديل بينها عبر زر 🔃
PHRASING_STYLES = [
    {
        "salutation": "سلام تام بوجود مولانا الإمام المؤيد بالله، وبعد؛",
        "intro": (
            "يشرفني، السيد المدير المحترم، أن أرفع إلى علمكم هذا التقرير التربوي قصد الإخبار واتخاذ المتعين، "
            "بشأن السلوكات المخلة بالنظام الصفي الصادرة عن التلميذ(ة): « {student} »، بالقسم: « {class_name} »، "
            "وذلك خلال حصة مادة « {subject} » بتاريخ {date} (التوقيت: {time} / {room})."
        ),
        "heading": "تفاصيل الواقعة والأفعال المسجلة:",
        "closing": (
            "وبناءً على ما سبق، وحفاظاً على السير العادي للدراسة وهيبة الفضاء التربوي، "
            "ألتمس من سيادتكم التدخل لاتخاذ الإجراءات الإدارية والتربوية المناسبة في حق المعني(ة) بالأمر، والمتمثلة في: {action}."
        ),
        "ending": "وتفضلوا، السيد المدير المحترم، بقبول فائق التقدير والاحترام.\nوالسلام.",
    },
    {
        "salutation": "سلام تام بوجود مولانا الإمام المؤيد بالله، وبعد؛",
        "intro": (
            "يشرفني، السيد المدير المحترم، أن أوافي سيادتكم بهذا التقرير التربوي قصد الاطلاع واتخاذ ما ترونه مناسباً، "
            "بخصوص التصرفات غير المنضبطة التي بدرت من التلميذ(ة): « {student} »، المتمدرس(ة) بالقسم: « {class_name} »، "
            "أثناء حصة مادة « {subject} » المنعقدة بتاريخ {date} (التوقيت: {time} / {room})."
        ),
        "heading": "حيثيات الواقعة والأفعال المنسوبة للتلميذ(ة):",
        "closing": (
            "وحرصاً منا على صون حرمة المؤسسة التعليمية، وتكريس مبادئ الانضباط والمسؤولية المنصوص عليها في النظام الداخلي، "
            "ألتمس من سيادتكم التفضل باتخاذ التدابير التربوية والتأديبية اللازمة في حق المعني(ة) بالأمر، ولا سيما: {action}."
        ),
        "ending": "وتفضلوا، السيد المدير المحترم، بقبول أسمى عبارات التقدير والاحترام.\nوالسلام.",
    },
    {
        "salutation": "سلام تام بوجود مولانا الإمام المؤيد بالله، وبعد؛",
        "intro": (
            "تبعاً لمقتضيات النظام الداخلي للمؤسسة، أحيطكم علماً، السيد المدير المحترم، "
            "بالوقائع المسجلة في حق التلميذ(ة): « {student} »، قسم: « {class_name} »، "
            "خلال حصة مادة « {subject} » ليوم {date} (التوقيت: {time} / {room})."
        ),
        "heading": "ملخص المخالفات المرصودة:",
        "closing": (
            "لذا، أضع هذا التقرير رهن إشارتكم لاتخاذ ما يلزم من إجراءات تربوية وتأديبية زاجرة "
            "في حق المعني(ة) بالأمر، وعلى رأسها: {action}."
        ),
        "ending": "وتفضلوا بقبول خالص التحية والتقدير.\nوالسلام.",
    },
]

SITUATION_VARIANTS = {
    "إحداث الفوضى والتشويش المستمر أثناء تقديم الدرس، ومقاطعة الأستاذ بشكل متعمد دون استئذان.": [
        "إحداث الفوضى والتشويش المستمر أثناء تقديم الدرس، ومقاطعة الأستاذ بشكل متعمد دون استئذان.",
        "تعمد إحداث الضجيج وعرقلة سير الشرح داخل الفصل، مع مقاطعة الأستاذ بشكل متكرر دون إذن مسبق.",
        "التشويش المستمر أثناء الحصة ومقاطعة الأستاذ عمداً مما حال دون متابعة الدرس.",
    ],
    "عدم الالتزام بالانضباط الصفي والتنقل غير المبرر داخل القاعة، مما أربك السير العادي للحصة وحرم باقي المتعلمين من حقهم في التحصيل.": [
        "عدم الالتزام بالانضباط الصفي والتنقل غير المبرر داخل القاعة، مما أربك السير العادي للحصة وحرم باقي المتعلمين من حقهم في التحصيل.",
        "الإخلال بقواعد الانضباط الصفي والتنقل العشوائي داخل القاعة، الأمر الذي أثر سلباً على تركيز المتعلمين والسير الطبيعي للحصة.",
        "عدم الانضباط والتنقل غير المبرر داخل القاعة، مما تسبب في إرباك الحصة الدراسية.",
    ],
    "التمادي في مخالفة التوجيهات التربوية وعدم الامتثال لتنبيهات الأستاذ المتكررة.": [
        "التمادي في مخالفة التوجيهات التربوية وعدم الامتثال لتنبيهات الأستاذ المتكررة.",
        "الإصرار على عدم الامتثال للتوجيهات التربوية وتجاهل الملاحظات والتنبيهات المتكررة للأستاذ.",
        "عدم الاستجابة لتنبيهات الأستاذ المتكررة والتمادي في مخالفة ضوابط القسم.",
    ],
    "الاستعمال غير المشروع للهاتف النقال / الأجهزة الإلكترونية داخل الفصل الدراسي في خرق سافر لمقتضيات النظام الداخلي للمؤسسة.": [
        "الاستعمال غير المشروع للهاتف النقال / الأجهزة الإلكترونية داخل الفصل الدراسي في خرق سافر لمقتضيات النظام الداخلي للمؤسسة.",
        "حيازة واستعمال الهاتف النقال داخل قاعة الدرس رغم المنع الصريح المنصوص عليه في النظام الداخلي للمؤسسة.",
        "استعمال الهاتف النقال داخل الفصل الدراسي بالمخالفة لمقتضيات النظام الداخلي.",
    ],
    "الامتناع المتعمد وغير المبرر عن تتبع مجريات الدرس، ورفض تدوين الملخصات وإنجاز الأنشطة الصفية والفروض.": [
        "الامتناع المتعمد وغير المبرر عن تتبع مجريات الدرس، ورفض تدوين الملخصات وإنجاز الأنشطة الصفية والفروض.",
        "الرفض التام للانخراط في الأنشطة التعليمية الصفية، والامتناع دون مبرر عن كتابة الدروس وإنجاز الفروض والتمارين.",
        "الامتناع عن تتبع الدرس ورفض تدوين الملخصات وإنجاز الأنشطة الصفية المطلوبة.",
    ],
    "التلفظ بعبارات نابية وغير لائقة، وإبداء سلوك استفزازي يمس بالاحترام الواجب للأطر التربوية وزملائه في القسم.": [
        "التلفظ بعبارات نابية وغير لائقة، وإبداء سلوك استفزازي يمس بالاحترام الواجب للأطر التربوية وزملائه في القسم.",
        "صدر عن المعني(ة) بالأمر كلام غير لائق وسلوك استفزازي يتنافى مع قواعد اللياقة والاحترام الواجب داخل الفضاء التربوي.",
        "التلفظ بعبارات غير لائقة وإبداء سلوك استفزازي يخل بالاحترام الواجب داخل القسم.",
    ],
    "إلحاق أضرار مادية بتجهيزات القاعة الدراسية والعبث بممتلكات المؤسسة.": [
        "إلحاق أضرار مادية بتجهيزات القاعة الدراسية والعبث بممتلكات المؤسسة.",
        "العبث بممتلكات المؤسسة وإتلاف تجهيزات القاعة الدراسية بشكل متعمد.",
        "إتلاف تجهيزات القاعة الدراسية والعبث بممتلكات المؤسسة.",
    ],
    "مغادرة القاعة الدراسية بصفة أحادية ودون الحصول على إذن مسبق من الأستاذ.": [
        "مغادرة القاعة الدراسية بصفة أحادية ودون الحصول على إذن مسبق من الأستاذ.",
        "الانصراف من قاعة الدرس بشكل مفاجئ ودون استئذان الأستاذ المشرف على الحصة.",
        "مغادرة القاعة الدراسية دون إذن مسبق من الأستاذ.",
    ],
    "التأخر المتكرر وغير المبرر عن موعد انطلاق الحصة الدراسية، مما يربك انطلاقة الدرس.": [
        "التأخر المتكرر وغير المبرر عن موعد انطلاق الحصة الدراسية، مما يربك انطلاقة الدرس.",
        "التمادي في التأخر عن الالتحاق بالحصة الدراسية في وقتها المحدد دون عذر مقبول، مما يعرقل الانطلاقة السليمة للدرس.",
        "التأخر المتكرر بدون مبرر عن موعد انطلاق الحصة الدراسية.",
    ],
    "التغيب المستمر بدون عذر مقبول عن حصص المادة، مع التهاون في تتبع الدروس.": [
        "التغيب المستمر بدون عذر مقبول عن حصص المادة، مع التهاون في تتبع الدروس.",
        "كثرة التغيبات غير المبررة عن حصص المادة الدراسية، وما يترتب عنها من انقطاع وتهاون في التحصيل الدراسي.",
        "التغيب المتكرر بدون عذر مقبول عن حصص المادة والتهاون في متابعة الدروس.",
    ],
    "الامتناع عن الإدلاء بورقة السماح بالدخول المسلمة من لدن الإدارة التربوية عقب الغياب.": [
        "الامتناع عن الإدلاء بورقة السماح بالدخول المسلمة من لدن الإدارة التربوية عقب الغياب.",
        "عدم الالتزام بالإجراءات الإدارية المعمول بها، ورفض الإدلاء بورقة السماح بالدخول بعد التغيب عن الحصص السابقة.",
        "عدم الإدلاء بورقة السماح بالدخول المسلمة من الإدارة التربوية إثر الغياب.",
    ],
}

GRADED_ACTIONS_MAP = {
    "1️⃣ تنبيه تربوي: توجيه تنبيه شفوي والتذكير بالانضباط": "توجيه تنبيه تربوي شفوي وتذكير المعني(ة) بالأمر بضرورة الالتزام بالانضباط الصفي",
    "2️⃣ إنذار كتابي: توجيه إنذار مع توقيع التزام خطي": "توجيه إنذار كتابي مع إلزام التلميذ(ة) بتوقيع التزام خطي بعدم تكرار المخالفة",
    "3️⃣ توبيخ إداري: توجيه توبيخ وإشعار ولي الأمر": "توجيه توبيخ إداري وإشعار ولي الأمر رسمياً بالسلوك المسجل",
    "4️⃣ استدعاء ولي الأمر: الحضور للإدارة وتوقيع تعهد": "استدعاء ولي أمر التلميذ(ة) للحضور الفوري إلى الإدارة التربوية قصد التدارس وتوقيع تعهد",
    "5️⃣ مواكبة تربوية: الإحالة على خلية الإنصات / التوجيه": "إحالة التلميذ(ة) على خلية الإنصات والوساطة التربوية أو مستشار التوجيه للمواكبة",
    "6️⃣ مجلس الانضباط: عرض الملف لاتخاذ المتعين": "عرض ملف التلميذ(ة) على أنظار مجلس الانضباط لاتخاذ العقوبات التأديبية المناسبة",
    "7️⃣ توقيف مؤقت: تطبيق مقتضيات النظام الداخلي": "تطبيق عقوبة التوقيف المؤقت عن الدراسة وفق ما ينص عليه النظام الداخلي للمؤسسة",
    "✍️ تحرير إجراء مخصص يدوياً...": "",
}

REPORT_TYPES = {
    "1️⃣ تقرير حول الإخلال بالنظام وعرقلة الدرس": {
        "title": "تقرير تربوي حول الإخلال بالنظام العام وعرقلة السير العادي للحصة الدراسية",
        "situations": [
            "إحداث الفوضى والتشويش المستمر أثناء تقديم الدرس، ومقاطعة الأستاذ بشكل متعمد دون استئذان.",
            "عدم الالتزام بالانضباط الصفي والتنقل غير المبرر داخل القاعة، مما أربك السير العادي للحصة وحرم باقي المتعلمين من حقهم في التحصيل.",
            "التمادي في مخالفة التوجيهات التربوية وعدم الامتثال لتنبيهات الأستاذ المتكررة.",
            "الاستعمال غير المشروع للهاتف النقال / الأجهزة الإلكترونية داخل الفصل الدراسي في خرق سافر لمقتضيات النظام الداخلي للمؤسسة.",
        ],
        "default_action_key": "4️⃣ استدعاء ولي الأمر: الحضور للإدارة وتوقيع تعهد",
    },
    "2️⃣ تقرير حول السلوك غير اللائق والامتناع عن العمل": {
        "title": "تقرير إداري وتربوي حول سلوك غير لائق والامتناع عن أداء الواجبات الصفية",
        "situations": [
            "الامتناع المتعمد وغير المبرر عن تتبع مجريات الدرس، ورفض تدوين الملخصات وإنجاز الأنشطة الصفية والفروض.",
            "التلفظ بعبارات نابية وغير لائقة، وإبداء سلوك استفزازي يمس بالاحترام الواجب للأطر التربوية وزملائه في القسم.",
            "إلحاق أضرار مادية بتجهيزات القاعة الدراسية والعبث بممتلكات المؤسسة.",
            "مغادرة القاعة الدراسية بصفة أحادية ودون الحصول على إذن مسبق من الأستاذ.",
        ],
        "default_action_key": "6️⃣ مجلس الانضباط: عرض الملف لاتخاذ المتعين",
    },
    "3️⃣ تقرير حول التمادي في التأخر والغياب غير المبرر": {
        "title": "تقرير حول التمادي في التأخر والتغيب غير المبرر عن الحصص الدراسية",
        "situations": [
            "التأخر المتكرر وغير المبرر عن موعد انطلاق الحصة الدراسية، مما يربك انطلاقة الدرس.",
            "التغيب المستمر بدون عذر مقبول عن حصص المادة، مع التهاون في تتبع الدروس.",
            "الامتناع عن الإدلاء بورقة السماح بالدخول المسلمة من لدن الإدارة التربوية عقب الغياب.",
        ],
        "default_action_key": "3️⃣ توبيخ إداري: توجيه توبيخ وإشعار ولي الأمر",
    },
}


def combine_selected_actions(selected_keys):
    """دمج الإجراءات المختارة (حتى 3 إجراءات) في صياغة متناسقة"""
    phrases = [
        GRADED_ACTIONS_MAP[k]
        for k in selected_keys
        if GRADED_ACTIONS_MAP.get(k, "").strip()
    ]
    if not phrases:
        return ""
    return "، و".join(phrases)


def get_styled_situations(situations_list, style_idx):
    """ترجمة قائمة المخالفات المختارة إلى الصيغة التعبيرية الحالية"""
    idx = style_idx % len(PHRASING_STYLES)
    return [SITUATION_VARIANTS.get(s, [s, s, s])[idx] for s in situations_list]


def build_report_texts(student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val, style_idx=0):
    """بناء نصوص التقرير وفق الصيغة التعبيرية المختارة"""
    template = PHRASING_STYLES[style_idx % len(PHRASING_STYLES)]

    student_str = student_val.strip() if student_val.strip() else "............................................"
    class_str = class_val.strip() if class_val.strip() else "......................"
    subj_str = subject_val.strip() if subject_val.strip() else "......................"
    time_str = time_val.strip() if time_val.strip() else "......................"
    room_str = format_room_label(room_val, is_preview=False)
    date_str = date_val.strftime("%Y/%m/%d")
    action_str = action_val.strip() if action_val.strip() else "اتخاذ المتعين وفق مقتضيات النظام الداخلي للمؤسسة"

    intro_text = template["intro"].format(
        student=student_str,
        class_name=class_str,
        subject=subj_str,
        date=date_str,
        time=time_str,
        room=room_str,
    )

    custom_text = f"معطيات وملاحظات تكميلية: {custom_notes.strip()}" if custom_notes.strip() else ""
    closing_body = template["closing"].format(action=action_str)

    return {
        "salutation": template["salutation"],
        "intro_text": intro_text,
        "section_heading": template["heading"],
        "custom_text": custom_text,
        "closing_body": closing_body,
        "closing_salutation": template["ending"],
        "date_str": date_str,
        "action_str": action_str,
    }


def generate_report_docx(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj, font_profile, style_idx=0
):
    doc = docx.Document()
    logo_file = get_logo_bytes(uploaded_logo_obj)

    is_dense = len(situations_list) >= 3 or len(custom_notes.strip()) > 100 or len(action_val.strip()) > 130
    is_very_dense = (len(situations_list) + (1 if custom_notes.strip() else 0)) >= 4 and len(action_val.strip()) > 150
    size_adj = -1.0 if is_very_dense else (-0.5 if is_dense else 0.0)

    w_font = font_profile["word_font"]
    w_body = font_profile["word_body_size"] + size_adj
    w_bullet = font_profile["word_bullet_size"] + size_adj
    w_title = font_profile["word_title_size"] + size_adj
    w_sal = font_profile["word_sal_size"] + size_adj
    w_hdr = font_profile["word_header_base"]
    w_foot = font_profile["word_footer_size"]
    w_spacing = 1.08 if is_very_dense else (1.15 if is_dense else font_profile["word_line_spacing"])

    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.header_distance = Inches(0.25)
    section.footer_distance = Inches(0.25)

    clear_and_prepare_headers(section)
    header = section.header

    tbl = header.add_table(rows=1, cols=3, width=Inches(6.8))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.tblPr
    tblPr.insert(0, OxmlElement("w:bidiVisual"))

    tblBorders = OxmlElement("w:tblBorders")
    bottom_border = OxmlElement("w:bottom")
    bottom_border.set(qn("w:val"), "single")
    bottom_border.set(qn("w:sz"), "8")
    bottom_border.set(qn("w:space"), "0")
    bottom_border.set(qn("w:color"), "334155")
    tblBorders.append(bottom_border)
    tblPr.append(tblBorders)

    row = tbl.rows[0]
    col_widths = [Inches(2.75), Inches(1.3), Inches(2.75)]
    for i, cell in enumerate(row.cells):
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_no_wrap(cell)

    d_name = direction_val.strip() if direction_val.strip() else "........................"
    s_name = school_val.strip() if school_val.strip() else "...................................."
    t_name = teacher_val.strip() if teacher_val.strip() else "........................"

    cell_r = row.cells[0]
    set_cell_rtl(cell_r)
    add_header_line_to_cell(cell_r, f"المديرية الإقليمية: {d_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=True, font_name=w_font, base_size=w_hdr)
    add_header_line_to_cell(cell_r, s_name, is_arabic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, is_first=False, font_name=w_font, base_size=w_hdr)

    cell_c = row.cells[1]
    set_cell_rtl(cell_c)
    p_c = cell_c.paragraphs[0]
    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_file:
        try:
            if isinstance(logo_file, io.BytesIO):
                logo_file.seek(0)
            p_c.add_run().add_picture(logo_file, width=Inches(1.05))
        except Exception:
            pass

    cell_l = row.cells[2]
    set_cell_rtl(cell_l)
    add_header_line_to_cell(cell_l, f"الموسم الدراسي: {year_val}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=True, font_name=w_font, base_size=w_hdr)
    add_header_line_to_cell(cell_l, f"الأستاذ(ة): {t_name}", is_arabic=True, align=WD_ALIGN_PARAGRAPH.LEFT, is_first=False, font_name=w_font, base_size=w_hdr)

    p_space = header.add_paragraph()
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(4)

    texts = build_report_texts(
        student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val, style_idx=style_idx
    )
    styled_situations = get_styled_situations(situations_list, style_idx)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4 if is_very_dense else 6)
    p_title.paragraph_format.space_after = Pt(6 if is_very_dense else (8 if is_dense else 12))
    r_title = add_styled_run(p_title, f"الموضوع: {report_title}", is_arabic=True, font_size=w_title, bold=True, font_name=w_font)
    r_title.font.underline = True

    p_sal = doc.add_paragraph()
    p_sal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sal.paragraph_format.space_before = Pt(2)
    p_sal.paragraph_format.space_after = Pt(4 if is_very_dense else (6 if is_dense else 8))
    add_styled_run(p_sal, texts["salutation"], is_arabic=True, font_size=w_sal, bold=True, font_name=w_font)

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.paragraph_format.space_before = Pt(2)
    p_body1.paragraph_format.space_after = Pt(5 if is_very_dense else 6)
    p_body1.paragraph_format.line_spacing = w_spacing
    add_styled_run(p_body1, texts["intro_text"], is_arabic=True, font_size=w_body, bold=False, font_name=w_font)

    p_situations_head = doc.add_paragraph()
    p_situations_head.paragraph_format.space_before = Pt(3)
    p_situations_head.paragraph_format.space_after = Pt(2)
    add_styled_run(p_situations_head, texts["section_heading"], is_arabic=True, font_size=w_body, bold=True, font_name=w_font)

    if styled_situations:
        for sit in styled_situations:
            p_bullet = doc.add_paragraph(style='List Bullet')
            p_bullet.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_bullet.paragraph_format.space_before = Pt(1)
            p_bullet.paragraph_format.space_after = Pt(2)
            p_bullet.paragraph_format.line_spacing = w_spacing
            add_styled_run(p_bullet, sit, is_arabic=True, font_size=w_bullet, bold=False, font_name=w_font)

    if texts["custom_text"]:
        p_custom = doc.add_paragraph()
        p_custom.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_custom.paragraph_format.space_before = Pt(3)
        p_custom.paragraph_format.space_after = Pt(3)
        p_custom.paragraph_format.line_spacing = w_spacing
        add_styled_run(p_custom, texts["custom_text"], is_arabic=True, font_size=w_bullet, bold=False, font_name=w_font)

    p_action = doc.add_paragraph()
    p_action.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_action.paragraph_format.space_before = Pt(5 if is_very_dense else 6)
    p_action.paragraph_format.space_after = Pt(5 if is_very_dense else 6)
    p_action.paragraph_format.line_spacing = w_spacing
    add_styled_run(p_action, texts["closing_body"], is_arabic=True, font_size=w_body, bold=False, font_name=w_font)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(2)
    p_end.paragraph_format.space_after = Pt(8 if is_very_dense else (10 if is_dense else 14))
    p_end.paragraph_format.line_spacing = 1.15
    add_styled_run(p_end, texts["closing_salutation"], is_arabic=True, font_size=w_body, bold=False, font_name=w_font)

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
    p_sig_r.paragraph_format.space_after = Pt(35 if is_very_dense else (45 if is_dense else 65))
    add_styled_run(p_sig_r, "توقيع الأستاذ(ة):", is_arabic=True, font_size=w_bullet, bold=True, font_name=w_font)

    cell_sig_l = sig_tbl.rows[0].cells[1]
    cell_sig_l.width = Inches(3.6)
    cell_sig_l.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_rtl(cell_sig_l)
    p_sig_l = cell_sig_l.paragraphs[0]
    p_sig_l.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_l.paragraph_format.space_after = Pt(4)
    add_styled_run(p_sig_l, "رأي وتوقيع الإدارة التربوية:", is_arabic=True, font_size=w_bullet, bold=True, font_name=w_font)

    for idx in range(3):
        p_dots = cell_sig_l.add_paragraph()
        p_dots.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_dots.paragraph_format.space_before = Pt(2)
        p_dots.paragraph_format.space_after = Pt(8 if idx == 2 else 3)
        add_styled_run(p_dots, "...........................................................", is_arabic=True, font_size=11, bold=False, font_name=w_font)

    clear_all_footers_completely(section)
    footer = section.footer
    p_foot = footer.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(0)
    p_foot.paragraph_format.space_after = Pt(0)

    center_pos = int(Inches(3.4))
    right_pos = int(Inches(6.8))

    p_foot.paragraph_format.tab_stops.add_tab_stop(center_pos, WD_TAB_ALIGNMENT.CENTER)
    p_foot.paragraph_format.tab_stops.add_tab_stop(right_pos, WD_TAB_ALIGNMENT.LEFT)
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    add_styled_run(p_foot, f"الأستاذ(ة): {t_name}", is_arabic=True, font_size=w_foot, bold=False, font_name=w_font)
    p_foot.add_run("\t\t")
    add_styled_run(p_foot, f"حرر في: {texts['date_str']}", is_arabic=True, font_size=w_foot, bold=False, font_name=w_font)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def generate_report_pdf(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj, font_profile, style_idx=0
):
    """توليد نسخة PDF رسمية في صفحة واحدة دائماً مع تذييل مثبت أسفل الصفحة"""
    if not HAS_PDF_LIBS:
        return None
    font_path = get_arabic_pdf_font(font_profile["pdf_font_key"])
    if not font_path:
        return None

    registered_font = f"ArabicFont_{font_profile['pdf_font_key']}"
    try:
        pdfmetrics.registerFont(TTFont(registered_font, font_path))
    except Exception:
        return None

    left_m = 42
    right_m = 42
    top_m = 28
    bottom_m = 44
    usable_width = A4[0] - left_m - right_m
    max_body_height = A4[1] - top_m - bottom_m

    d_name = direction_val.strip() if direction_val.strip() else "........................"
    s_name = school_val.strip() if school_val.strip() else "...................................."
    t_name = teacher_val.strip() if teacher_val.strip() else "........................"

    texts = build_report_texts(
        student_val, class_val, subject_val, date_val, time_val, room_val, custom_notes, action_val, style_idx=style_idx
    )
    styled_situations = get_styled_situations(situations_list, style_idx)

    def build_story_elements(scale=1.0):
        p_body = font_profile["pdf_body_size"] * scale
        p_title = font_profile["pdf_title_size"] * scale
        p_lead = font_profile["pdf_leading"] * scale

        style_right = ParagraphStyle(
            "ArRight", fontName=registered_font, fontSize=p_body, leading=p_lead,
            alignment=TA_RIGHT, leftPadding=0, rightPadding=0
        )
        style_center = ParagraphStyle(
            "ArCenter", fontName=registered_font, fontSize=p_body + 0.5, leading=p_lead,
            alignment=TA_CENTER, leftPadding=0, rightPadding=0
        )
        style_title = ParagraphStyle(
            "ArTitle", fontName=registered_font, fontSize=p_title, leading=p_lead + 2,
            alignment=TA_CENTER, leftPadding=0, rightPadding=0
        )
        style_hdr_r = ParagraphStyle(
            "ArHdrR", fontName=registered_font, fontSize=10.5 * scale, leading=14.5 * scale,
            alignment=TA_RIGHT, leftPadding=0, rightPadding=0
        )
        style_hdr_l = ParagraphStyle(
            "ArHdrL", fontName=registered_font, fontSize=10.5 * scale, leading=14.5 * scale,
            alignment=TA_LEFT, leftPadding=0, rightPadding=0
        )

        items = []

        logo_file = get_logo_bytes(uploaded_logo_obj)
        logo_cell = ""
        if logo_file:
            try:
                if isinstance(logo_file, io.BytesIO):
                    logo_file.seek(0)
                logo_cell = RLImage(logo_file, width=62 * scale, height=45 * scale, kind="proportional")
            except Exception:
                logo_cell = ""

        hdr_data = [[
            Paragraph(f"{shape_ar(f'الموسم الدراسي: {year_val}')}<br/>{shape_ar(f'الأستاذ(ة): {t_name}')}", style_hdr_l),
            logo_cell,
            Paragraph(f"{shape_ar(f'المديرية الإقليمية: {d_name}')}<br/>{shape_ar(s_name)}", style_hdr_r),
        ]]
        hdr_table = Table(hdr_data, colWidths=[203, 105, 203])
        hdr_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor("#334155")),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6 * scale),
        ]))
        items.append(hdr_table)
        items.append(Spacer(1, 10 * scale))

        wrapped_title = format_ar_paragraph(f"الموضوع: {report_title}", registered_font, p_title, usable_width)
        items.append(Paragraph(f"<u>{wrapped_title}</u>", style_title))
        items.append(Spacer(1, 8 * scale))

        items.append(Paragraph(shape_ar(texts["salutation"]), style_center))
        items.append(Spacer(1, 7 * scale))

        wrapped_intro = format_ar_paragraph(texts["intro_text"], registered_font, p_body, usable_width)
        items.append(Paragraph(wrapped_intro, style_right))
        items.append(Spacer(1, 7 * scale))

        items.append(Paragraph(shape_ar(texts["section_heading"]), style_right))
        items.append(Spacer(1, 4 * scale))
        for sit in styled_situations:
            wrapped_sit = format_ar_paragraph(f"• {sit}", registered_font, p_body, usable_width - 8)
            items.append(Paragraph(wrapped_sit, style_right))
            items.append(Spacer(1, 3 * scale))

        if texts["custom_text"]:
            items.append(Spacer(1, 3 * scale))
            wrapped_custom = format_ar_paragraph(texts["custom_text"], registered_font, p_body, usable_width)
            items.append(Paragraph(wrapped_custom, style_right))

        items.append(Spacer(1, 7 * scale))

        wrapped_closing = format_ar_paragraph(texts["closing_body"], registered_font, p_body, usable_width)
        items.append(Paragraph(wrapped_closing, style_right))
        items.append(Spacer(1, 6 * scale))

        wrapped_end = format_ar_paragraph(texts["closing_salutation"], registered_font, p_body + 0.5, usable_width)
        items.append(Paragraph(wrapped_end, style_center))
        items.append(Spacer(1, 12 * scale))

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
        sig_table = Table(sig_data, colWidths=[265, 246])
        sig_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 16 * scale),
        ]))
        items.append(sig_table)
        return items

    story = build_story_elements(scale=1.0)
    total_height = sum(el.wrap(usable_width, max_body_height)[1] for el in story)
    if total_height > max_body_height:
        fit_scale = max(0.68, (max_body_height / total_height) * 0.95)
        story = build_story_elements(scale=fit_scale)

    def draw_pdf_footer(canvas, doc_obj):
        canvas.saveState()
        canvas.setStrokeColor(colors.lightgrey)
        canvas.setLineWidth(0.6)
        canvas.line(left_m, 34, A4[0] - right_m, 34)

        canvas.setFont(registered_font, 9.5)
        canvas.setFillColor(colors.grey)
        canvas.drawString(left_m, 20, shape_ar(f"حرر في: {texts['date_str']}"))
        canvas.drawRightString(A4[0] - right_m, 20, shape_ar(f"الأستاذ(ة): {t_name}"))
        canvas.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=right_m,
        leftMargin=left_m,
        topMargin=top_m,
        bottomMargin=bottom_m,
    )
    doc.build(story, onFirstPage=draw_pdf_footer, onLaterPages=draw_pdf_footer)
    buffer.seek(0)
    return buffer


def render_live_preview_html(
    direction_val, school_val, teacher_val, year_val, student_val, class_val,
    subject_val, date_val, time_val, room_val, report_title, situations_list,
    custom_notes, action_val, uploaded_logo_obj, font_profile, style_idx=0
):
    """توليد المعاينة الحية التفاعلية وفق الخط والصيغة التعبيرية المختارة"""
    template = PHRASING_STYLES[style_idx % len(PHRASING_STYLES)]
    styled_situations = get_styled_situations(situations_list, style_idx)

    css_font = font_profile["css_font"]
    css_body_sz = font_profile["css_body_size"]
    css_title_sz = font_profile["css_title_size"]
    css_lh = font_profile["css_line_height"]

    logo_src = get_logo_base64_src(uploaded_logo_obj)
    logo_html = f'<img src="{logo_src}" style="max-height: 48px; max-width: 90px; height: auto; width: auto; object-fit: contain;" />' if logo_src else ''

    d_name = direction_val.strip() if direction_val.strip() else "<span style='color:#94a3b8;'>[المديرية الإقليمية]</span>"
    s_name = school_val.strip() if school_val.strip() else "<span style='color:#94a3b8;'>[المؤسسة التعليمية]</span>"
    t_name = teacher_val.strip() if teacher_val.strip() else "<span style='color:#94a3b8;'>[اسم الأستاذ(ة)]</span>"

    student_disp = student_val.strip() if student_val.strip() else "<span style='color:#94a3b8;'>[اسم التلميذ(ة)]</span>"
    class_disp = class_val.strip() if class_val.strip() else "<span style='color:#94a3b8;'>[المستوى]</span>"
    subj_disp = subject_val.strip() if subject_val.strip() else "<span style='color:#94a3b8;'>[المادة]</span>"
    time_disp = time_val.strip() if time_val.strip() else "<span style='color:#94a3b8;'>[التوقيت]</span>"
    room_disp = format_room_label(room_val, is_preview=True)
    if not room_val.strip():
        room_disp = f"<span style='color:#94a3b8;'>{room_disp}</span>"

    date_str = date_val.strftime("%Y/%m/%d")
    action_disp = action_val.strip() if action_val.strip() else "اتخاذ المتعين وفق مقتضيات النظام الداخلي للمؤسسة"

    intro_html = template["intro"].format(
        student=f"<b>« {student_disp} »</b>",
        class_name=f"<b>« {class_disp} »</b>",
        subject=f"<b>« {subj_disp} »</b>",
        date=f"<b>{date_str}</b>",
        time=time_disp,
        room=room_disp,
    )
    closing_html = template["closing"].format(action=f"<b>{action_disp}</b>")
    ending_html = template["ending"].replace("\n", "<br><b>") + "</b>"

    bullets_html = "".join([f"<li style='margin-bottom: 6px; text-align: justify;'>{s}</li>" for s in styled_situations])
    custom_html = f"<p style='margin-top: 6px; font-weight: 500; text-align: justify;'><b>معطيات وملاحظات تكميلية:</b> {custom_notes.strip()}</p>" if custom_notes.strip() else ""

    html = f"""<div class="paper-preview" style="font-family: {css_font}; box-sizing: border-box; width: 100%; max-width: 100%; overflow: hidden; padding: 20px 15px;">
<div style="display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 4px; border-bottom: 2px solid #334155; padding-bottom: 8px; margin-bottom: 16px;">
<div style="text-align: right; font-size: clamp(0.68rem, 2.2vw, 0.88rem); line-height: 1.45; font-weight: bold; overflow: hidden;">
<div style="white-space: nowrap;">المديرية الإقليمية: {d_name}</div>
<div style="white-space: nowrap;">{s_name}</div>
</div>
<div style="text-align: center; display: flex; justify-content: center; align-items: center; padding: 0 2px;">
{logo_html}
</div>
<div style="text-align: left; font-size: clamp(0.68rem, 2.2vw, 0.88rem); line-height: 1.45; font-weight: bold; overflow: hidden;">
<div style="white-space: nowrap;">الموسم الدراسي: <span style="display: inline-block; direction: ltr; unicode-bidi: embed; white-space: nowrap;">{year_val}</span></div>
<div style="white-space: nowrap;">الأستاذ(ة): {t_name}</div>
</div>
</div>
<div style="text-align: center; margin: 14px 0;">
<h3 style="font-family: {css_font}; margin: 0; font-size: {css_title_sz}; font-weight: 700; text-decoration: underline; color: #0f172a; line-height: 1.6;">الموضوع: {report_title}</h3>
</div>
<div style="text-align: center; font-weight: bold; font-size: {css_body_sz}; margin-bottom: 12px;">
{template["salutation"]}
</div>
<p style="line-height: {css_lh}; text-align: justify; font-size: {css_body_sz}; margin-bottom: 12px; word-break: break-word;">
{intro_html}
</p>
<div style="font-weight: bold; font-size: {css_body_sz}; margin-top: 10px; margin-bottom: 6px;">{template["heading"]}</div>
<ul style="padding-right: 22px; font-size: {css_body_sz}; line-height: {css_lh}; margin-top: 4px; margin-bottom: 10px;">
{bullets_html}
</ul>
{custom_html}
<p style="line-height: {css_lh}; text-align: justify; font-size: {css_body_sz}; margin-top: 14px; margin-bottom: 10px; word-break: break-word;">
{closing_html}
</p>
<div style="text-align: center; font-size: {css_body_sz}; line-height: 1.7; margin-top: 8px; margin-bottom: 20px;">
{ending_html}
</div>
<div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-top: 25px; margin-bottom: 70px; padding-top: 10px; font-size: clamp(0.82rem, 2.7vw, 0.98rem);">
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
<div style="display: flex; justify-content: space-between; align-items: center; margin-top: 20px; border-top: 1px dashed #cbd5e1; padding-top: 8px; font-size: clamp(0.72rem, 2.3vw, 0.84rem); color: #64748b;">
<div>الأستاذ(ة): {t_name}</div>
<div>حرر في: {date_str}</div>
</div>
</div>"""
    return html


def render_report_tool():
    if "report_data_docx" not in st.session_state:
        st.session_state.report_data_docx = None
    if "report_filename_docx" not in st.session_state:
        st.session_state.report_filename_docx = ""
    if "report_data_pdf" not in st.session_state:
        st.session_state.report_data_pdf = None
    if "report_filename_pdf" not in st.session_state:
        st.session_state.report_filename_pdf = ""
    if "phrasing_style_idx" not in st.session_state:
        st.session_state.phrasing_style_idx = 0
    if "font_style_choice" not in st.session_state:
        st.session_state.font_style_choice = "نمط 1"

    active_font_profile = FONT_PROFILES.get(st.session_state.font_style_choice, FONT_PROFILES["نمط 1"])

    st.markdown(
        """
        <div style="text-align: center; padding: 15px; border-radius: 10px; background-color: #ffffff; border: 1px solid #e2e8f0; margin-bottom: 20px;">
            <h1 style="color: #1e293b; margin: 0; font-size: 1.6rem; font-weight: 800;">📝 محرر التقارير السريعة</h1>
            <p style="color: #64748b; margin-top: 5px; font-size: 0.95rem;">توليد تقارير انضباطية رسمية ومسبوكة بصيغتي Word (.docx) و PDF جاهزة للطباعة والتوقيع</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("1️⃣ البيانات العامة والترويسة")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        direction_name = st.text_input("المديرية الإقليمية:", value="", placeholder="مثال: تاونات")
    with col_g2:
        school_name = st.text_input("المؤسسة التعليمية:", value="", placeholder="مثال: ثانوية أولاد داود التأهيلية")

    teacher_name = st.text_input("اسم الأستاذ(ة):", value="", placeholder="مثال: عثمان شهير")
    school_year = get_auto_school_year()

    uploaded_logo = st.file_uploader(
        "🖼️ رفع شعار المؤسسة الخاص (اختياري - يحل محل الشعار الافتراضي):",
        type=["jpg", "png", "jpeg"],
    )

    st.subheader("2️⃣ بيانات التلميذ والواقعة")
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        student_name = st.text_input("اسم التلميذ(ة):", value="", placeholder="مثال: محمد العلمي")
    with col_s2:
        class_level = st.text_input("القسم / المستوى:", value="", placeholder="مثال: الثانية بكالوريا علوم فيزيائية 1")
    with col_s3:
        subject_name = st.text_input("المادة الدراسية:", value="", placeholder="مثال: الرياضيات")

    incident_date = datetime.date.today()

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        incident_time = st.text_input("توقيت الحصة:", value="", placeholder="مثال: من 10:00 إلى 11:00")
    with col_t2:
        room_number = st.text_input("رقم القاعة:", value="", placeholder="مثال: 4 (يكفي كتابة الرقم مباشرة)")

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

    if "last_report_type" not in st.session_state:
        st.session_state.last_report_type = selected_type_key
    if "selected_action_options" not in st.session_state:
        st.session_state.selected_action_options = [current_config["default_action_key"]]
    if "proposed_action_text" not in st.session_state:
        st.session_state.proposed_action_text = combine_selected_actions(st.session_state.selected_action_options)

    if st.session_state.last_report_type != selected_type_key:
        st.session_state.last_report_type = selected_type_key
        default_key = current_config["default_action_key"]
        st.session_state.selected_action_options = [default_key]
        st.session_state.proposed_action_text = combine_selected_actions([default_key])

    def on_action_select_change():
        chosen_list = st.session_state.selected_action_options
        st.session_state.proposed_action_text = combine_selected_actions(chosen_list)

    st.multiselect(
        "الإجراء الإداري / التأديبي المقترح (يمكن الجمع بين 3 إجراءات على الأكثر):",
        options=list(GRADED_ACTIONS_MAP.keys()),
        max_selections=3,
        key="selected_action_options",
        on_change=on_action_select_change,
    )

    proposed_action = st.text_area(
        "صياغة الإجراء المقترح (قابلة للتعديل أو التحرير اليدوي):",
        key="proposed_action_text",
        height=90,
        placeholder="اكتب أو عدّل صياغة الإجراء الإداري أو التأديبي المقترح هنا...",
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
        font_profile=active_font_profile,
        style_idx=st.session_state.phrasing_style_idx,
    )
    st.markdown(preview_html_code, unsafe_allow_html=True)

    # --- شريط التحكم السفلي: اختيار الخط (نمط 1 / نمط 2) + زر تبديل الصياغة 🔃 + زر توليد التقرير ---
    def reset_generated_docs():
        st.session_state.report_data_docx = None
        st.session_state.report_data_pdf = None

    st.markdown('<div class="inline-action-marker"></div>', unsafe_allow_html=True)
    col_font, col_switch, col_gen = st.columns([0.24, 0.14, 0.62])
    with col_font:
        st.selectbox(
            "الخط",
            options=["نمط 1", "نمط 2"],
            key="font_style_choice",
            label_visibility="collapsed",
            on_change=reset_generated_docs,
        )
    with col_switch:
        if st.button("🔃", key="cycle_phrasing_btn"):
            st.session_state.phrasing_style_idx = (st.session_state.phrasing_style_idx + 1) % len(PHRASING_STYLES)
            reset_generated_docs()
            st.rerun()
    with col_gen:
        generate_clicked = st.button("🚀 توليد وتجهيز التقرير (Word & PDF)", type="primary")

    if generate_clicked:
        if not student_name.strip():
            st.warning("⚠️ يرجى إدخال اسم التلميذ(ة) لتوليد التقرير.")
        else:
            safe_name = student_name.strip().replace(" ", "_")
            current_style = st.session_state.phrasing_style_idx
            chosen_profile = FONT_PROFILES[st.session_state.font_style_choice]
            docx_bytes = generate_report_docx(
                direction_val=direction_name, school_val=school_name, teacher_val=teacher_name,
                year_val=school_year, student_val=student_name, class_val=class_level,
                subject_val=subject_name, date_val=incident_date, time_val=incident_time,
                room_val=room_number, report_title=current_config["title"],
                situations_list=selected_situations, custom_notes=custom_details,
                action_val=proposed_action, uploaded_logo_obj=uploaded_logo,
                font_profile=chosen_profile, style_idx=current_style,
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
                font_profile=chosen_profile, style_idx=current_style,
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
