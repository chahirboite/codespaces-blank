import streamlit as st
import streamlit.components.v1 as components
from massar_engine import (
    export_updated_massar_xlsx,
    parse_massar_qualifying_workbook,
    update_student_assessment_in_memory,
)


PARTIAL_GRADE_ROWS = [
    [(0.25, "+0,25"), (0.50, "+0,5"), (0.75, "+0,75"), (1.00, "+1")],
    [(1.25, "+1,25"), (1.50, "+1,5"), (1.75, "+1,75"), (2.00, "+2")],
    [(2.50, "+2,5"), (3.00, "+3"), (4.00, "+4"), (5.00, "+5")],
]


def _inject_massar_cards_css():
    st.markdown(
        """
        <style>
        /* ضبط اتجاه الكتابة من اليمين إلى اليسار لجميع عناصر الواجهة والتسميات */
        div[data-testid="stWidgetLabel"],
        div[data-testid="stWidgetLabel"] p,
        div[data-testid="stCaptionContainer"],
        div[data-testid="stCaptionContainer"] p,
        div[data-testid="stExpander"] summary,
        div[data-testid="stExpander"] p,
        div[data-testid="stMarkdownContainer"] p {
            direction: rtl !important;
            text-align: right !important;
            unicode-bidi: plaintext !important;
        }

        div[data-testid="stSelectbox"] > div,
        div[data-testid="stTextInput"] input {
            direction: rtl !important;
            text-align: right !important;
        }

        /* إخفاء عبارة Press Enter to apply لتفعيل البحث اللحظي النظيف */
        div[data-testid="InputInstructions"] {
            display: none !important;
        }

        /* إبقاء حقل النقطة وزر الحاسبة في سطر واحد بمحاذاة بعضهما على الهاتف والحاسوب */
        .element-container:has(.grade-calc-row-marker) + .element-container div[data-testid="stHorizontalBlock"] {
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            align-items: center !important;
            gap: 8px !important;
        }
        .element-container:has(.grade-calc-row-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(1) {
            flex: 1 1 auto !important;
            min-width: 0 !important;
        }
        .element-container:has(.grade-calc-row-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) {
            width: 118px !important;
            min-width: 118px !important;
            max-width: 118px !important;
            flex: 0 0 118px !important;
        }
        .element-container:has(.grade-calc-row-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) button {
            background: linear-gradient(135deg, #0d9488 0%, #0284c7 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px !important;
            height: 44px !important;
            font-weight: 800 !important;
            font-size: 0.95rem !important;
            padding: 0 8px !important;
            box-shadow: 0 3px 8px rgba(13, 148, 136, 0.25) !important;
        }
        .element-container:has(.grade-calc-row-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) button:hover {
            opacity: 0.92 !important;
        }

        /* إبقاء أزرار لوحة الآلة الحاسبة (4 أعمدة) وأزرار التحكم أفقية داخل النافذة المنبثقة على الهاتف */
        .element-container:has(.calc-grid-marker) + .element-container div[data-testid="stHorizontalBlock"] {
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            gap: 6px !important;
        }
        .element-container:has(.calc-grid-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            flex: 1 1 0% !important;
            min-width: 0 !important;
        }
        .element-container:has(.calc-grid-marker) + .element-container div[data-testid="stHorizontalBlock"] button {
            background-color: #f8fafc !important;
            color: #0f172a !important;
            border: 1.5px solid #cbd5e1 !important;
            border-radius: 10px !important;
            height: 44px !important;
            font-weight: 800 !important;
            font-size: 0.95rem !important;
            padding: 0 2px !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05) !important;
            direction: ltr !important;
        }
        .element-container:has(.calc-grid-marker) + .element-container div[data-testid="stHorizontalBlock"] button:hover {
            background-color: #e0f2fe !important;
            border-color: #0284c7 !important;
        }

        .element-container:has(.calc-tools-marker) + .element-container div[data-testid="stHorizontalBlock"] {
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            gap: 6px !important;
        }
        .element-container:has(.calc-tools-marker) + .element-container div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            flex: 1 1 0% !important;
            min-width: 0 !important;
        }
        .element-container:has(.calc-tools-marker) + .element-container div[data-testid="stHorizontalBlock"] button {
            background-color: #f1f5f9 !important;
            color: #334155 !important;
            border: 1px solid #cbd5e1 !important;
            height: 40px !important;
            font-size: 0.84rem !important;
            padding: 0 4px !important;
            box-shadow: none !important;
        }

        .calc-screen-box {
            direction: rtl;
            background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
            color: #f8fafc;
            border-radius: 14px;
            padding: 14px 16px;
            margin-bottom: 12px;
            border: 1.5px solid #334155;
            box-shadow: inset 0 2px 6px rgba(0, 0, 0, 0.35);
        }
        .calc-tape {
            direction: ltr;
            text-align: left;
            font-size: 0.82rem;
            color: #94a3b8;
            min-height: 20px;
            white-space: nowrap;
            overflow-x: auto;
            margin-bottom: 4px;
        }
        .calc-total-row {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
        }
        .calc-total-val {
            direction: ltr;
            font-size: 1.85rem;
            font-weight: 800;
            color: #38bdf8;
            letter-spacing: 0.5px;
        }

        .cycle-card {
            background: #ffffff;
            border: 1.5px solid #e2e8f0;
            border-radius: 16px;
            padding: 18px 12px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
            margin-bottom: 8px;
            min-height: 155px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            direction: rtl;
        }
        .cycle-card-active {
            border: 2px solid #0284c7;
            background: linear-gradient(180deg, #f0f9ff 0%, #ffffff 100%);
        }
        .cycle-icon {
            font-size: 2.3rem;
            margin-bottom: 6px;
        }
        .cycle-title {
            font-size: 1.05rem;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 4px;
        }
        .cycle-badge-ready {
            background: #dcfce7;
            color: #166534;
            font-size: 0.75rem;
            font-weight: 700;
            padding: 2px 10px;
            border-radius: 999px;
        }
        .cycle-badge-soon {
            background: #f1f5f9;
            color: #64748b;
            font-size: 0.75rem;
            font-weight: 600;
            padding: 2px 10px;
            border-radius: 999px;
        }
        .student-mobile-card {
            background: #ffffff;
            border: 2px solid #0284c7;
            border-radius: 14px;
            padding: 14px 16px;
            margin: 10px 0 12px 0;
            box-shadow: 0 4px 12px rgba(2, 132, 199, 0.08);
            direction: rtl;
            text-align: right;
        }
        .rtl-status-bar {
            direction: rtl;
            text-align: right;
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 9px 12px;
            font-size: 0.88rem;
            color: #1e293b;
            line-height: 1.7;
            margin: 6px 0 12px 0;
        }
        .rtl-hint-box {
            direction: rtl;
            text-align: right;
            background: #eff6ff;
            border-right: 4px solid #3b82f6;
            color: #1e40af;
            padding: 10px 12px;
            border-radius: 8px;
            font-size: 0.88rem;
            margin: 8px 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _bind_live_search_js():
    """ربط خانة البحث بتحديث تلقائي فوري أثناء الكتابة دون الحاجة للضغط على Enter"""
    components.html(
        """
        <script>
        (function() {
            const parentDoc = window.parent.document;
            const inputs = parentDoc.querySelectorAll('input[type="text"]');
            inputs.forEach(function(inp) {
                const aria = inp.getAttribute('aria-label') || '';
                if (aria.includes('ابحث بالرقم الترتيبي') && !inp.dataset.liveBound) {
                    inp.dataset.liveBound = '1';
                    let timer = null;
                    inp.addEventListener('input', function() {
                        clearTimeout(timer);
                        timer = setTimeout(function() {
                            const ev = new KeyboardEvent('keydown', {
                                key: 'Enter',
                                code: 'Enter',
                                keyCode: 13,
                                which: 13,
                                bubbles: true,
                                cancelable: true
                            });
                            inp.dispatchEvent(ev);
                        }, 380);
                    });
                }
            });
        })();
        </script>
        """,
        height=0,
    )


def _render_calculator_body(target_num_key, student_name):
    """محتوى الآلة الحاسبة الذكية لجمع النقط الجزئية"""
    hist_key = f"calc_hist_{target_num_key}"
    if hist_key not in st.session_state:
        st.session_state[hist_key] = []

    history = st.session_state[hist_key]
    raw_total = round(sum(history), 2)
    exceeded = raw_total > 20.0

    tape_str = " + ".join(f"{v:g}" for v in history) if history else "اضغط على النقط الجزئية للجمع..."
    val_color = "#f87171" if exceeded else "#38bdf8"
    warn_html = (
        "<div style='color:#fca5a5; font-size:0.78rem; margin-top:4px;'>⚠️ المجموع تجاوز 20.00 (سيُضبط تلقائياً في 20.00 عند التأكيد)</div>"
        if exceeded
        else ""
    )

    st.markdown(
        f"""
        <div class="calc-screen-box">
            <div style="font-size:0.82rem; color:#cbd5e1; margin-bottom:4px;">
                👤 التلميذ(ة): <b><bdi>{student_name}</bdi></b>
            </div>
            <div class="calc-tape">{tape_str}</div>
            <div class="calc-total-row">
                <span style="font-size:0.9rem; color:#94a3b8;">المجموع المحصل:</span>
                <span class="calc-total-val" style="color:{val_color};">{raw_total:.2f} <span style="font-size:0.95rem; color:#94a3b8;">/ 20</span></span>
            </div>
            {warn_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # شبكة النقط الجزئية (3 صفوف × 4 أزرار)
    for r_idx, row_items in enumerate(PARTIAL_GRADE_ROWS):
        st.markdown('<div class="calc-grid-marker"></div>', unsafe_allow_html=True)
        cols = st.columns(4)
        for c_idx, (val, label) in enumerate(row_items):
            with cols[c_idx]:
                if st.button(label, key=f"cbtn_{target_num_key}_{r_idx}_{c_idx}", use_container_width=True):
                    st.session_state[hist_key].append(val)
                    st.rerun()

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    # أزرار التراجع، التصفير، والإلغاء
    st.markdown('<div class="calc-tools-marker"></div>', unsafe_allow_html=True)
    col_undo, col_reset, col_cancel = st.columns(3)
    with col_undo:
        if st.button("↩️ تراجع", key=f"c_undo_{target_num_key}", use_container_width=True):
            if st.session_state[hist_key]:
                st.session_state[hist_key].pop()
                st.rerun()
    with col_reset:
        if st.button("🔄 تصفير", key=f"c_reset_{target_num_key}", use_container_width=True):
            st.session_state[hist_key] = []
            st.rerun()
    with col_cancel:
        if st.button("✖️ إلغاء", key=f"c_cancel_{target_num_key}", use_container_width=True):
            st.session_state.pop(hist_key, None)
            st.session_state["active_calc_modal_key"] = None
            st.rerun()

    # زر التأكيد الرئيسي لإدخال النقطة في الحقل
    if st.button(
        f"✅ تأكيد وإدخال ({min(20.0, raw_total):.2f})",
        type="primary",
        key=f"c_confirm_{target_num_key}",
        use_container_width=True,
    ):
        st.session_state[target_num_key] = min(20.0, raw_total)
        st.session_state.pop(hist_key, None)
        st.session_state["active_calc_modal_key"] = None
        st.rerun()


if hasattr(st, "dialog"):
    @st.dialog("🧮 حاسبة النقط الجزئية")
    def _open_calculator_dialog(target_num_key, student_name):
        _render_calculator_body(target_num_key, student_name)
else:
    def _open_calculator_dialog(target_num_key, student_name):
        with st.container(border=True):
            st.markdown("#### 🧮 حاسبة النقط الجزئية")
            _render_calculator_body(target_num_key, student_name)


def _render_cycle_selector():
    """شاشة بوابة الدخول ذات البطاقات المربعة للأسلاك الثلاثة"""
    _inject_massar_cards_css()

    st.markdown(
        """
        <div dir="rtl" style="text-align: center; padding: 15px; border-radius: 12px; background-color: #ffffff; border: 1px solid #e2e8f0; margin-bottom: 20px;">
            <h1 style="color: #1e293b; margin: 0; font-size: 1.55rem; font-weight: 800;">📊 مساعد التصحيح الذكي واستثمار نقط مسار</h1>
            <p style="color: #64748b; margin-top: 6px; font-size: 0.94rem;">اختر السلك التعليمي للبدء في مسك النقط، حقن الملاحظات، وإنتاج تقارير الاستثمار</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            """
            <div class="cycle-card">
                <div class="cycle-icon">🏫</div>
                <div class="cycle-title">التعليم الابتدائي</div>
                <span class="cycle-badge-soon">قيد الإعداد</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("دخول الابتدائي", key="btn_cycle_primary", use_container_width=True):
            st.session_state.massar_selected_cycle = "primary"
            st.rerun()

    with col2:
        st.markdown(
            """
            <div class="cycle-card">
                <div class="cycle-icon">🏛️</div>
                <div class="cycle-title">الثانوي الإعدادي</div>
                <span class="cycle-badge-soon">قيد الإعداد</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("دخول الإعدادي", key="btn_cycle_middle", use_container_width=True):
            st.session_state.massar_selected_cycle = "middle"
            st.rerun()

    with col3:
        st.markdown(
            """
            <div class="cycle-card cycle-card-active">
                <div class="cycle-icon">🎓</div>
                <div class="cycle-title">الثانوي التأهيلي</div>
                <span class="cycle-badge-ready">✅ جاهز للاستعمال</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("دخول التأهيلي", key="btn_cycle_qualifying", type="primary", use_container_width=True):
            st.session_state.massar_selected_cycle = "qualifying"
            st.rerun()


def _render_qualifying_fast_correction_tab(parsed_doc):
    """التبويب الأول: التصحيح والمسك السريع المصمم للهاتف"""
    assessments = parsed_doc["assessments"]
    students = parsed_doc["students"]
    meta = parsed_doc["meta"]

    assess_options = {a["code"]: f"{a['label']} ({a['code']})" for a in assessments}
    selected_code = st.selectbox(
        "اختر الفرض أو النشاط المراد مسك نقطه:",
        options=list(assess_options.keys()),
        format_func=lambda k: assess_options[k],
        key=f"sel_assess_{meta['filename']}",
    )

    total_students = len(students)
    entered_count = sum(
        1
        for s in students
        if s["grades"].get(selected_code) is not None or bool(s["absences"].get(selected_code))
    )
    remaining_count = total_students - entered_count
    progress_ratio = (entered_count / total_students) if total_students > 0 else 0.0

    st.progress(progress_ratio)
    st.markdown(
        f"""
        <div class="rtl-status-bar">
            📌 القسم: <b><bdi>{meta['class_name']}</bdi></b> &nbsp;|&nbsp; المادة: <b><bdi>{meta['subject']}</bdi></b><br>
            ✅ المُنجز: <b>{entered_count} / {total_students}</b> &nbsp;|&nbsp; ⏳ المتبقي: <b>{remaining_count}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # تفريغ خانة البحث بأمان قبل إنشاء حقل الإدخال
    search_key = f"search_query_{meta['filename']}"
    clear_flag_key = f"clear_flag_{meta['filename']}"

    if st.session_state.get(clear_flag_key, False):
        st.session_state[search_key] = ""
        st.session_state[clear_flag_key] = False
    elif search_key not in st.session_state:
        st.session_state[search_key] = ""

    query = st.text_input(
        "🔍 ابحث بالرقم الترتيبي للورقة (مثلاً: 14) أو باسم التلميذ(ة):",
        key=search_key,
        placeholder="اكتب الرقم الترتيبي (1 - 40) أو جزءاً من الاسم...",
    ).strip()

    # تفعيل التحديث الفوري أثناء الكتابة بدون الحاجة للضغط على Enter
    _bind_live_search_js()

    matched_students = []
    if query:
        if query.isdigit():
            ord_num = int(query)
            matched_students = [s for s in students if s["ordinal"] == ord_num]
        else:
            q_lower = query.lower()
            matched_students = [
                s
                for s in students
                if q_lower in s["name"].lower() or q_lower in s["massar_code"].lower()
            ]
    else:
        ungraded = [
            s
            for s in students
            if s["grades"].get(selected_code) is None and not s["absences"].get(selected_code)
        ]
        if ungraded:
            matched_students = [ungraded[0]]
            st.markdown(
                """
                <div class="rtl-hint-box">
                    💡 يتم عرض أول ورقة متبقية تلقائياً، أو اكتب رقم الورقة في خانة البحث أعلاه.
                </div>
                """,
                unsafe_allow_html=True,
            )

    if query and not matched_students:
        st.warning("⚠️ لم يتم العثور على أي تلميذ مطابق. تأكد من الرقم الترتيبي أو الاسم.")
    elif matched_students:
        if len(matched_students) > 1:
            chosen_ord = st.selectbox(
                "تطابقت عدة أسماء، اختر التلميذ(ة) المعني:",
                options=[s["ordinal"] for s in matched_students],
                format_func=lambda o: next(
                    f"رقم {x['ordinal']} — {x['name']} ({x['massar_code']})"
                    for x in matched_students
                    if x["ordinal"] == o
                ),
            )
            active_student = next(s for s in matched_students if s["ordinal"] == chosen_ord)
        else:
            active_student = matched_students[0]

        cur_grade = active_student["grades"].get(selected_code)
        cur_abs = active_student["absences"].get(selected_code, "")

        if cur_abs:
            status_badge = "<span style='background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:6px;font-weight:700;'>غائب (مبرر)</span>"
        elif cur_grade is not None:
            status_badge = f"<span style='background:#dcfce7;color:#166534;padding:2px 8px;border-radius:6px;font-weight:700;'>النقطة المسجلة: <bdi>{cur_grade:.2f} / 20</bdi></span>"
        else:
            status_badge = "<span style='background:#f1f5f9;color:#475569;padding:2px 8px;border-radius:6px;'>لم تُمسك بعد</span>"

        st.markdown(
            f"""
            <div dir="rtl" class="student-mobile-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                    <span style="background:#0284c7; color:#ffffff; font-weight:800; padding:3px 11px; border-radius:8px; font-size:0.95rem;">
                        الرقم الترتيبي: {active_student['ordinal']}
                    </span>
                    {status_badge}
                </div>
                <div style="font-size:1.18rem; font-weight:800; color:#0f172a; margin:6px 0;">
                    <bdi>{active_student['name']}</bdi>
                </div>
                <div style="font-size:0.84rem; color:#475569;">
                    رمز مسار: <b><bdi>{active_student['massar_code']}</bdi></b> &nbsp;|&nbsp; تاريخ الازدياد: <bdi>{active_student['birth_date']}</bdi>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        num_key = f"num_in_{meta['filename']}_{selected_code}_{active_student['ordinal']}"
        if num_key not in st.session_state:
            st.session_state[num_key] = float(cur_grade) if cur_grade is not None else 10.00

        st.markdown(
            "<div dir='rtl' style='text-align:right; font-weight:700; font-size:0.9rem; color:#1e293b; margin-bottom:4px;'>أدخل النقطة (من 0.00 إلى 20.00):</div>",
            unsafe_allow_html=True,
        )

        # صف أفقي يجمع حقل إدخال النقطة وبمحاذاته زر الآلة الحاسبة الذكية
        st.markdown('<div class="grade-calc-row-marker"></div>', unsafe_allow_html=True)
        col_num, col_calc = st.columns([0.68, 0.32])
        with col_num:
            grade_input = st.number_input(
                "أدخل النقطة",
                min_value=0.0,
                max_value=20.0,
                step=0.25,
                format="%.2f",
                key=num_key,
                label_visibility="collapsed",
            )
        with col_calc:
            if st.button("🧮 حاسبة", key=f"btn_open_calc_{num_key}", use_container_width=True):
                st.session_state[f"calc_hist_{num_key}"] = []
                st.session_state["active_calc_modal_key"] = num_key
                if hasattr(st, "dialog"):
                    _open_calculator_dialog(num_key, active_student["name"])

        # في حال عدم توفر st.dialog في نسخة قديمة، تُعرض الحاسبة كإطار منبثق مباشرة هنا
        if not hasattr(st, "dialog") and st.session_state.get("active_calc_modal_key") == num_key:
            _open_calculator_dialog(num_key, active_student["name"])

        col_save, col_abs = st.columns([0.62, 0.38])
        with col_save:
            if st.button(
                "💾 حفظ النقطة",
                type="primary",
                use_container_width=True,
                key=f"save_btn_{meta['filename']}_{selected_code}_{active_student['ordinal']}",
            ):
                update_student_assessment_in_memory(
                    parsed_doc,
                    student_ordinal=active_student["ordinal"],
                    assessment_code=selected_code,
                    grade_value=grade_input,
                    is_absent=False,
                )
                st.session_state.pop(num_key, None)
                st.session_state["active_calc_modal_key"] = None
                st.session_state[clear_flag_key] = True
                st.toast(f"✅ تم حفظ نقطة {active_student['name']} ({grade_input:.2f})")
                st.rerun()

        with col_abs:
            if st.button(
                "🚫 غياب مبرر",
                use_container_width=True,
                key=f"abs_btn_{meta['filename']}_{selected_code}_{active_student['ordinal']}",
            ):
                update_student_assessment_in_memory(
                    parsed_doc,
                    student_ordinal=active_student["ordinal"],
                    assessment_code=selected_code,
                    grade_value=None,
                    is_absent=True,
                )
                st.session_state.pop(num_key, None)
                st.session_state["active_calc_modal_key"] = None
                st.session_state[clear_flag_key] = True
                st.toast(f"📝 تم تسجيل غياب مبرر لـ {active_student['name']}")
                st.rerun()

    # كشف الأوراق المتبقية
    missing_list = [
        s
        for s in students
        if s["grades"].get(selected_code) is None and not s["absences"].get(selected_code)
    ]
    with st.expander(f"⚠️ لائحة الأوراق المتبقية ({len(missing_list)} تلميذ(ة))", expanded=False):
        if not missing_list:
            st.success("🎉 تم مسك جميع نقط هذا التقويم بالكامل!")
        else:
            for ms in missing_list:
                st.markdown(
                    f"<div dir='rtl' style='text-align:right; padding:3px 0;'>• <b>رقم {ms['ordinal']}</b>: <bdi>{ms['name']}</bdi> (<bdi>{ms['massar_code']}</bdi>)</div>",
                    unsafe_allow_html=True,
                )

    # زر التحميل الفوري لملف مسار المحدث
    st.markdown("---")
    updated_xlsx_io = export_updated_massar_xlsx(parsed_doc)
    base_fn = meta["filename"].rsplit(".", 1)[0]
    st.download_button(
        label=f"📥 تحميل ملف مسار المُحيّن ({meta['class_name']}.xlsx)",
        data=updated_xlsx_io,
        file_name=f"{base_fn}_محيّن.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


def _render_qualifying_workspace():
    """فضاء العمل الخاص بالسلك الثانوي التأهيلي"""
    _inject_massar_cards_css()

    col_back, col_title = st.columns([0.28, 0.72])
    with col_back:
        if st.button("🔙 تغيير السلك", use_container_width=True):
            st.session_state.massar_selected_cycle = None
            st.rerun()
    with col_title:
        st.markdown("<h3 dir='rtl' style='text-align:right; margin:4px 0;'>🎓 السلك الثانوي التأهيلي</h3>", unsafe_allow_html=True)

    if "massar_qualifying_docs" not in st.session_state:
        st.session_state.massar_qualifying_docs = {}

    uploaded_files = st.file_uploader(
        "📂 ارفع ملف أو عدة ملفات Excel المستخرجة من مسار (.xlsx):",
        type=["xlsx"],
        accept_multiple_files=True,
        key="qualifying_uploader",
    )

    if uploaded_files:
        for uf in uploaded_files:
            if uf.name not in st.session_state.massar_qualifying_docs:
                try:
                    parsed = parse_massar_qualifying_workbook(uf.read(), uf.name)
                    st.session_state.massar_qualifying_docs[uf.name] = parsed
                except Exception as exc:
                    st.error(f"❌ تعذرت قراءة الملف '{uf.name}': {exc}")

    docs_dict = st.session_state.massar_qualifying_docs
    if not docs_dict:
        st.info("👆 ارفع ملف مسار واحد أو أكثر للبدء في التصحيح السريع أو التحليل.")
        return

    doc_keys = list(docs_dict.keys())
    selected_doc_key = st.selectbox(
        "🏫 اختر القسم / الملف النشط:",
        options=doc_keys,
        format_func=lambda k: f"{docs_dict[k]['meta']['class_name']} — {docs_dict[k]['meta']['subject']} ({k})",
    )
    active_doc = docs_dict[selected_doc_key]

    tab1, tab2, tab3 = st.tabs(
        [
            "✍️ التصحيح والمسك السريع",
            "💬 حقن الملاحظات التربوية",
            "📈 استثمار النتائج وخطة الدعم",
        ]
    )

    with tab1:
        _render_qualifying_fast_correction_tab(active_doc)

    with tab2:
        st.info(
            "🚧 **المرحلة القادمة:** هنا سيتم حقن الملاحظات التربوية تلقائياً في عمود `#100#` "
            "بناءً على احتساب المعدل (75% فروض + 25% أنشطة مندمجة أو معدل الفروض الصرف للرياضيات)."
        )

    with tab3:
        st.info(
            "🚧 **المرحلة الموالية:** هنا سيتم عرض المبيانات الإحصائية وتوليد تقرير استثمار النتائج "
            "وخطة الدعم بصيغتي Word و PDF."
        )


def render_massar_tool():
    """نقطة الدخول الرئيسية لأداة مساعد التصحيح الذكي واستثمار نقط مسار"""
    if "massar_selected_cycle" not in st.session_state:
        st.session_state.massar_selected_cycle = None

    cycle = st.session_state.massar_selected_cycle

    if cycle is None:
        _render_cycle_selector()
    elif cycle == "qualifying":
        _render_qualifying_workspace()
    else:
        cycle_label = "التعليم الابتدائي" if cycle == "primary" else "التعليم الثانوي الإعدادي"
        if st.button("🔙 الرجوع لاختيار السلك"):
            st.session_state.massar_selected_cycle = None
            st.rerun()
        st.warning(f"⏳ وحدة **{cycle_label}** قيد التجهيز؛ سنفعّلها فور تزويدنا بنماذج ملفات مسار الخاصة بها.")
