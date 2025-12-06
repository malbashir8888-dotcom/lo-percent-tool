import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="LO from Grades Tool", layout="centered")

st.markdown("## 📊 حساب مخرجات التعلّم من ملف الدرجات")
st.write("""
هذه الأداة تحسب نسب مخرجات التعلّم مباشرة من ملف درجات الطلاب،
بشرط أن يحتوي الملف على:
- صف يوضح رقم مخرج التعلّم (LO)
- صف يوضح الدرجة العظمى لكل سؤال (/5، /10، /3…)
- صفوف الطلاب ودرجاتهم
""")

uploaded_file = st.file_uploader(
    "اختيار ملف درجات (Excel أو ODS)",
    type=["xlsx", "xls", "ods"]
)

def read_grade_sheet(uploaded_file, sheet_name=0):
    buffer = BytesIO(uploaded_file.getvalue())
    ext = uploaded_file.name.lower().split(".")[-1]
    if ext == "ods":
        df = pd.read_excel(buffer, engine="odf", sheet_name=sheet_name, header=None)
    else:
        df = pd.read_excel(buffer, sheet_name=sheet_name, header=None)
    return df

def build_lo_report_from_grades(df,
                                lo_row_index=4,
                                max_row_index=5,
                                student_start_index=6):

    lo_row = df.iloc[lo_row_index]
    max_row = df.iloc[max_row_index]

    lo_stats = {}

    for col in df.columns:
        lo = lo_row[col]
        if pd.isna(lo):
            continue

        max_raw = max_row[col]

        if isinstance(max_raw, str):
            m = re.search(r'(\d+(\.\d+)?)', max_raw)
            if not m:
                continue
            max_each = float(m.group(1))
        else:
            try:
                max_each = float(max_raw)
            except:
                continue

        scores = df.loc[student_start_index:, col]
        scores = pd.to_numeric(scores, errors="coerce").dropna()

        if scores.empty or max_each == 0:
            continue

        total_score = scores.sum()
        n_students = scores.count()
        total_max = max_each * n_students

        lo = str(lo)
        lo_stats.setdefault(lo, {"total_score": 0.0, "total_max": 0.0})
        lo_stats[lo]["total_score"] += total_score
        lo_stats[lo]["total_max"] += total_max

    rows = []

    overall_total = sum(v["total_score"] for v in lo_stats.values())
    overall_max   = sum(v["total_max"]   for v in lo_stats.values())
    overall_percent = (overall_total / overall_max) * 100 if overall_max else None
    rows.append({
        "Learning Objective": "Overall",
        "Total": overall_total,
        "Max": overall_max,
        "Percent": overall_percent
    })

    for lo, v in lo_stats.items():
        ts = v["total_score"]
        tm = v["total_max"]
        p  = (ts / tm) * 100 if tm else None
        rows.append({
            "Learning Objective": lo,
            "Total": ts,
            "Max": tm,
            "Percent": p
        })

    return pd.DataFrame(rows)

st.markdown("### إعداد الصفوف (كما تظهر في Excel)")

col1, col2, col3 = st.columns(3)
with col1:
    lo_row_excel = st.number_input("صف مخرجات التعلم (LO)", min_value=1, value=5)
with col2:
    max_row_excel = st.number_input("صف الدرجة العظمى للسؤال", min_value=1, value=6)
with col3:
    student_start_excel = st.number_input("أول صف للطلاب", min_value=1, value=7)

if st.button("تحليل ملف الدرجات"):
    if not uploaded_file:
        st.error("الرجاء رفع ملف الدرجات أولًا.")
    else:
        df = read_grade_sheet(uploaded_file)

        report = build_lo_report_from_grades(
            df,
            lo_row_index=lo_row_excel - 1,
            max_row_index=max_row_excel - 1,
            student_start_index=student_start_excel - 1
        )

        st.subheader("نتيجة تحليل مخرجات التعلم")
        st.dataframe(report)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, sheet_name="LO_Report", index=False)
        output.seek(0)

        st.download_button(
            "تحميل تقرير المخرجات (Excel)",
            data=output,
            file_name="LO_Report_from_grades.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
