import streamlit as st
from header_tool import render_header_tool
from report_tool import render_report_tool

# إعدادات الصفحة الرئيسية
st.set_page_config(
    page_title="منصة الأدوات التربوية 📝", page_icon="📝", layout="centered"
)

# التنسيق العام والخطوط
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Cairo:wght@400;600;700;800&family=Noto+Naskh+Arabic:wght@400;600;700&display=swap');

    .stApp {
        background-color: #f8fafc;
        font-family: 'Cairo', sans-serif;
    }
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: 700;
        font-family: 'Cairo', sans-serif;
        height: 46px;
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
        padding: 22px 16px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.06);
        direction: rtl;
        text-align: right;
        color: #0f172a;
        margin-top: 15px;
        margin-bottom: 16px;
        box-sizing: border-box;
        width: 100%;
        overflow: hidden;
    }
    /* إبقاء أزرار التحكم السفلية (نمط 1 / نمط 2 + 🔃 + التوليد) متناسقة على الهاتف والحاسوب */
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        gap: 8px !important;
    }
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(1) {
        width: 105px !important;
        min-width: 105px !important;
        max-width: 105px !important;
        flex: 0 0 105px !important;
    }
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) {
        width: 50px !important;
        min-width: 50px !important;
        max-width: 50px !important;
        flex: 0 0 50px !important;
    }
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(3) {
        flex: 1 1 auto !important;
        min-width: 0 !important;
    }
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) button {
        background-color: #f1f5f9 !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        padding: 0 !important;
        font-size: 1.2rem !important;
        box-shadow: none !important;
    }
    .element-container:has(.inline-action-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) button:hover {
        background-color: #e2e8f0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# القائمة الجانبية
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

# التوجيه بين الأداتين
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
