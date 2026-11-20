import io
import openpyxl


def parse_massar_qualifying_workbook(file_bytes, filename=""):
    """
    قراءة ملف Excel المستخرج من مسار (سلك الثانوي التأهيلي)
    واستخراج البيانات الإدارية، الفروض المتاحة ديناميكياً، ولائحة التلاميذ.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    if "NotesCC" not in wb.sheetnames:
        raise ValueError(f"الملف '{filename}' لا يحتوي على ورقة 'NotesCC' المعتمدة في مسار.")

    ws = wb["NotesCC"]

    meta = {
        "filename": filename,
        "academy": str(ws["D7"].value or "").strip(),
        "direction": str(ws["I7"].value or "").strip(),
        "school": str(ws["O7"].value or "").strip(),
        "level": str(ws["D9"].value or "").strip(),
        "class_name": str(ws["I9"].value or "").strip() or filename,
        "teacher": str(ws["O9"].value or "").strip(),
        "semester": str(ws["D11"].value or "").strip(),
        "subject": str(ws["O11"].value or "").strip(),
        "school_year": str(ws["D13"].value or "").strip(),
    }

    # الكشف الديناميكي عن أعمدة الفروض (#1#, #2#, #3#)، الأنشطة المندمجة (#4#)، والملاحظات (#100#)
    assessments = []
    notes_col = None
    for c in range(1, ws.max_column + 1):
        code_val = str(ws.cell(15, c).value or "").strip()
        header_title = str(ws.cell(16, c).value or "").strip()
        if code_val in ("#1#", "#2#", "#3#", "#4#"):
            label = header_title if header_title else f"التقويم {code_val}"
            assessments.append(
                {
                    "code": code_val,
                    "label": label,
                    "grade_col": c,
                    "absence_col": c + 1,
                    "is_integrated": (code_val == "#4#"),
                }
            )
        elif code_val == "#100#":
            notes_col = c

    if not assessments:
        raise ValueError(f"لم يتم العثور على أعمدة فروض صالحة (#1#..#4#) داخل الملف '{filename}'.")

    students = []
    for r in range(18, ws.max_row + 1):
        sid = ws.cell(r, 2).value
        massar_code = ws.cell(r, 3).value
        full_name = ws.cell(r, 4).value
        if sid is None and massar_code is None and full_name is None:
            continue

        ordinal = len(students) + 1
        grades = {}
        absences = {}
        for a in assessments:
            raw_g = ws.cell(r, a["grade_col"]).value
            try:
                grades[a["code"]] = float(raw_g) if raw_g is not None and str(raw_g).strip() != "" else None
            except (ValueError, TypeError):
                grades[a["code"]] = None

            abs_val = ws.cell(r, a["absence_col"]).value
            absences[a["code"]] = str(abs_val or "").strip()

        students.append(
            {
                "ordinal": ordinal,
                "row": r,
                "id": sid,
                "massar_code": str(massar_code or "").strip(),
                "name": str(full_name or "").strip(),
                "birth_date": str(ws.cell(r, 6).value or "").strip(),
                "grades": grades,
                "absences": absences,
                "note": str(ws.cell(r, notes_col).value or "").strip() if notes_col else "",
            }
        )

    return {
        "raw_bytes": file_bytes,
        "meta": meta,
        "assessments": assessments,
        "notes_col": notes_col,
        "students": students,
    }


def update_student_assessment_in_memory(parsed_doc, student_ordinal, assessment_code, grade_value=None, is_absent=False):
    """
    تحديث نقطة تلميذ أو تسجيل تغيبه المبرر في الذاكرة الحية دون المساس ببنية مسار.
    """
    target_assessment = next(
        (a for a in parsed_doc["assessments"] if a["code"] == assessment_code), None
    )
    if not target_assessment:
        return False

    for st_obj in parsed_doc["students"]:
        if st_obj["ordinal"] == student_ordinal:
            if is_absent:
                st_obj["grades"][assessment_code] = None
                st_obj["absences"][assessment_code] = "مبرر"
            else:
                st_obj["grades"][assessment_code] = round(float(grade_value), 2) if grade_value is not None else None
                st_obj["absences"][assessment_code] = ""
            return True
    return False


def compute_qualifying_student_average(student_obj, assessments):
    """
    حساب معدل الدورة للتلميذ في السلك الثانوي التأهيلي:
    - إذا وُجد عمود الأنشطة المندمجة (#4#): 75% لمعدل الفروض + 25% للأنشطة المندمجة.
    - إذا لم يوجد عمود الأنشطة المندمجة (مثل الرياضيات): مجموع نقط الفروض مقسوم على عددها.
    """
    exam_scores = []
    integrated_score = None

    for a in assessments:
        val = student_obj["grades"].get(a["code"])
        if val is None:
            continue
        if a["is_integrated"]:
            integrated_score = float(val)
        else:
            exam_scores.append(float(val))

    if not exam_scores and integrated_score is None:
        return None

    has_integrated_col = any(a["is_integrated"] for a in assessments)
    if has_integrated_col and integrated_score is not None and exam_scores:
        exams_avg = sum(exam_scores) / len(exam_scores)
        return round((exams_avg * 0.75) + (integrated_score * 0.25), 2)
    elif exam_scores:
        return round(sum(exam_scores) / len(exam_scores), 2)
    return round(integrated_score, 2)


def export_updated_massar_xlsx(parsed_doc):
    """
    إعادة كتابة النقط والتغيبات المحيّنة داخل نفس نسخة ملف مسار الأصلية
    مع الحفاظ التام على الخلايا المحمية، الرموز المخفية، وقوائم التحقق.
    """
    wb = openpyxl.load_workbook(io.BytesIO(parsed_doc["raw_bytes"]))
    ws = wb["NotesCC"]

    assess_map = {a["code"]: a for a in parsed_doc["assessments"]}
    notes_col = parsed_doc.get("notes_col")

    for st_obj in parsed_doc["students"]:
        r = st_obj["row"]
        for code, a_info in assess_map.items():
            g_col = a_info["grade_col"]
            abs_col = a_info["absence_col"]

            g_val = st_obj["grades"].get(code)
            abs_val = st_obj["absences"].get(code, "")

            if abs_val:
                ws.cell(r, g_col).value = None
                ws.cell(r, abs_col).value = abs_val
            else:
                ws.cell(r, g_col).value = float(g_val) if g_val is not None else None
                ws.cell(r, abs_col).value = None

        if notes_col and st_obj.get("note"):
            ws.cell(r, notes_col).value = st_obj["note"]

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
