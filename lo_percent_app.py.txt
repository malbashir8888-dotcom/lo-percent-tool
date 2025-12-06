import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="LO Percent Tool", layout="centered")

st.markdown("## أداة تجميع نسب مخرجات التعلّم (Percent)")
st.write("ارفعي تقارير Remark بصيغة Excel، وسيتم حساب متوسط النسبة لكل مخرج تعلّم.")

uploaded_files = st.file_uploader(
    "اختيار ملفات Excel (يمكن اختيار أكثر من ملف)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

run_clicked = st.button("تنفيذ التجميع")

def extract_percent_from_file(file_obj, filename):
    """استخراج عمود Percent لكل Learning Objective من ورقة Class Learning Objective Report"""
    try:
        xls = pd.ExcelFile(file_obj)
    except Exception:
        return pd.DataFrame()

    if "Class Learning Objective Report" not in xls.sheet_names:
        return pd.DataFrame()

    df = pd.read_excel(xls, sheet_name="Class Learning Objective Report", header=None)

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
        rows.append({
            "Learning_Objective": lo,
            "Percent": percent,
            "Source_File": filename
        })

    return pd.DataFrame(rows)

if run_clicked:
    if not uploaded_files:
        st.error("الرجاء رفع ملف واحد على الأقل أولاً.")
    else:
        all_frames = []
        for f in uploaded_files:
            df_lo = extract_percent_from_file(f, f.name)
            if not df_lo.empty:
                all_frames.append(df_lo)

        if not all_frames:
            st.error("لم يتم استخراج أي بيانات Percent من الملفات المرفوعة.")
        else:
            all_lo = pd.concat(all_frames, ignore_index=True)

            summary = (
                all_lo
                .groupby("Learning_Objective", as_index=False)
                .agg(
                    Num_Measurements=("Percent", "count"),
                    Sum_Percent=("Percent", "sum"),
                    Mean_Percent=("Percent", "mean"),
                )
            )

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="All_Objectives", index=False)
                all_lo.to_excel(writer, sheet_name="Raw_Percent_All_Files", index=False)
            output.seek(0)

            st.success("تم تجهيز النتائج. يمكنك تحميل الملف بالزر أدناه.")
            st.download_button(
                label="تحميل النتائج (ملف Excel)",
                data=output,
                file_name="LO_Percent_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
