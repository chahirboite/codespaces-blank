import streamlit as st
from massar_engine import (
    export_updated_massar_xlsx,
    parse_massar_qualifying_workbook,
    update_student_assessment_in_memory,
)


def _inject_massar_cards_css():
    st.markdown(
        """
        <style>
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
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_cycle_selector():
    """شاشة بوابة الدخول ذات البطاقات المربعة للأسلاك الثلاثة"""
    _inject_massar_cards_css()

    st.markdown(
        """
        <div style="text-align: center; padding: 15px; border-radius: 12px; background-color: #ffffff; border: 1px solid #e2e8f0; margin-bottom: 20px;">
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
    st.caption(
        f"📌 القسم: **{meta['class_name']}** | المادة: **{meta['subject']}** | "
        f"المُنجز: **{entered_count} / {total_students}** | المتبقي: **{remaining_count}**"
    )

    # محرك البحث الخاطف بالرقم الترتيبي (1..N) أو الاسم الكامل أو رمز مسار
    search_key = f"search_query_{meta['filename']}"
    if search_key not in st.session_state:
        st.session_state[search_key] = ""

    query = st.text_input(
        "🔍 ابحث بالرقم الترتيبي للورقة (مثلاً: 14) أو باسم التلميذ(ة):",
        key=search_key,
        placeholder="اكتب الرقم الترتيبي (1 - 40) أو جزءاً من الاسم...",
    ).strip()

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
        # عرض أول تلميذ لم تُمسك نقطته بعد تلقائياً لتسريع العمل المتسلسل
        ungraded = [
            s
            for s in students
            if s["grades"].get(selected_code) is None and not s["absences"].get(selected_code)
        ]
        if ungraded:
            matched_students = [ungraded[0]]
            st.info("💡 يتم عرض أول ورقة متبقية تلقائياً، أو اكتب رقم الورقة في خانة البحث أعلاه.")

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
            status_badge = f"<span style='background:#dcfce7;color:#166534;padding:2px 8px;border-radius:6px;font-weight:700;'>النقطة المسجلة: {cur_grade:.2f} / 20</span>"
        else:
            status_badge = "<span style='background:#f1f5f9;color:#475569;padding:2px 8px;border-radius:6px;'>لم تُمسك بعد</span>"

        st.markdown(
            f"""
            <div class="student-mobile-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                    <span style="background:#0284c7; color:#ffffff; font-weight:800; padding:3px 11px; border-radius:8px; font-size:0.95rem;">
                        الرقم الترتيبي: {active_student['ordinal']}
                    </span>
                    {status_badge}
                </div>
                <div style="font-size:1.18rem; font-weight:800; color:#0f172a; margin:4px 0;">
                    {active_student['name']}
                </div>
                <div style="font-size:0.84rem; color:#64748b;">
                    رمز مسار: <b>{active_student['massar_code']}</b> | تاريخ الازدياد: {active_student['birth_date']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        default_num = float(cur_grade) if cur_grade is not None else 10.00
        grade_input = st.number_input(
            "أدخل النقطة (من 0.00 إلى 20.00):",
            min_value=0.0,
            max_value=20.0,
            value=default_num,
            step=0.25,
            format="%.2f",
            key=f"num_in_{meta['filename']}_{selected_code}_{active_student['ordinal']}",
        )

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
                st.session_state[search_key] = ""
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
                st.session_state[search_key] = ""
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
                st.markdown(f"- **رقم {ms['ordinal']}**: {ms['name']} (`{ms['massar_code']}`)")

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
        st.markdown("### 🎓 السلك الثانوي التأهيلي")

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
