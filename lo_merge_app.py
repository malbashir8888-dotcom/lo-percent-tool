import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merge LO Percent Tool", layout="centered")

st.markdown("## 🧮 أداة دمج نسب مخرجات التعلّم من أكثر من مصدر")
st.write(
    """
ترفع في هذه الأداة جميع الملفات النهائية المتاحة للمقرر الواحد، وتشمل:
- تقارير **Remark** من نوع *Class Learning Objective Report*.
- تقارير الأداة الثانية الناتجة عن تحليل ملف الدرجات (ملف يحتوي أعمدة: *Learning Objective* و *Percent*).

ستقوم الأداة بالتعرّف تلقائيًا على نوع كل ملف، ثم دمج جميع النِّسَب
وحساب متوسط نسبة كل مخرج تعلّم (LO) في تقرير واحد نهائي.
"""
)

# --------------------------------------------------------
# دالة 1: استخراج Percent من تقرير Remark
# --------------------------------------------------------
def extract_from_remark(file_obj, filename):
    try:
        xls = pd.ExcelFile(file_obj)
    except Exception:
        return pd.DataFrame()

    if "Class Learning Objective Report" not in xls.sheet_names:
        return pd.DataFrame()

    df = pd.read_excel(
        xls, sheet_name="Class Learning Objective Report", header=None
    )

    header_row = df.index[df[0] == "Learning Objective"]
    if len(header_row) == 0:
        return pd.DataFrame()

    start = header_row[0] + 1

    rows = []
    for i in range(start, len(df)):
        lo = df.at[i, 0]
        if pd.isna(lo):
            break
        percent = df.at[i, 5]
        rows.append(
            {
                "Learning_Objective": str(lo),
                "Percent": percent,
                "Source_File": filename,
                "Source_Type": "Remark",
            }
        )

    return pd.DataFrame(rows)


# --------------------------------------------------------
# دالة 2: استخراج Percent من تقرير الأداة الثانية
#   (ملف فيه أعمدة: Learning Objective, Percent)
# --------------------------------------------------------
def extract_from_lo_report(file_obj, filename):
    try:
        # نقرأ أول شيت بافتراض أن الهيدر في الصف الأول
        df = pd.read_excel(file_obj, sheet_name=0)
    except Exception:
        return pd.DataFrame()

    # توحيد الأسماء (حساسية صغيرة للاختلافات في الكتابة)
    normalized_cols = {c: str(c).strip().lower() for c in df.columns}

    lo_col = None
    p_col = None
    for orig, norm in normalized_cols.items():
        if norm in ["learning objective", "learning_objective", "lo"]:
            lo_col = orig
        if norm in ["percent", "percentage", "perc"]:
            p_col = orig

    if lo_col is None or p_col is None:
        return pd.DataFrame()

    sub = df[[lo_col, p_col]].copy()
    sub.columns = ["Learning_Objective", "Percent"]
    sub["Source_File"] = filename
    sub["Source_Type"] = "Grades-Report"

    # إزالة الصفوف الفارغة
    sub = sub.dropna(subset=["Learning_Objective"])
    return sub


# --------------------------------------------------------
# واجهة رفع الملفات
# --------------------------------------------------------
uploaded_files = st.file_uploader(
    "اختيار جميع الملفات (Remark + تقارير الأداة الثانية)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if st.button("تنفيذ الدمج"):
    if not uploaded_files:
        st.error("الرجاء رفع ملف واحد على الأقل.")
    else:
        all_rows = []

        for f in uploaded_files:
            # نجرب أولاً: هل هو تقرير Remark؟
            df_r = extract_from_remark(f, f.name)
            if not df_r.empty:
                all_rows.append(df_r)
                continue

            # إن لم يكن Remark نجرب نوع تقرير الأداة الثانية
            f.seek(0)
            df_g = extract_from_lo_report(f, f.name)
            if not df_g.empty:
                all_rows.append(df_g)
                continue

            # إن لم يتعرّف عليه أي نوع:
            st.warning(f"لم يتم التعرّف على نوع الملف: {f.name}")

        if not all_rows:
            st.error("لم يتم استخراج أي بيانات من الملفات المرفوعة.")
        else:
            merged = pd.concat(all_rows, ignore_index=True)

            # تحويل Percent إلى أعداد
            merged["Percent"] = pd.to_numeric(
                merged["Percent"], errors="coerce"
            )
            merged = merged.dropna(subset=["Percent"])

            # جدول ملخّص
            summary = (
                merged.groupby("Learning_Objective", as_index=False)
                .agg(
                    Num_Measurements=("Percent", "count"),
                    Mean_Percent=("Percent", "mean"),
                )
                .sort_values("Learning_Objective")
            )

            st.subheader("النتائج التفصيلية (من جميع الملفات)")
            st.dataframe(merged)

            st.subheader("ملخّص مخرجات التعلّم بعد الدمج")
            st.dataframe(summary)

            # تجهيز ملف Excel للتحميل
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary.to_excel(
                    writer, sheet_name="Summary_Merged_LO", index=False
                )
                merged.to_excel(
                    writer, sheet_name="All_Records_Detail", index=False
                )
            output.seek(0)

            st.download_button(
                "تحميل تقرير الدمج النهائي (Excel)",
                data=output,
                file_name="Merged_LO_Percent_Report.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
