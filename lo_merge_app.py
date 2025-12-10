import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merge LO Percent Tool", layout="centered")

# --------------------------------------------------------
# UI TEXT (English)
# --------------------------------------------------------
st.markdown("## 🧮 Merged Learning Outcomes Percent Tool")
st.write(
    """
In this tool, you can upload all final files for a single course, including:
- **Remark®** reports of type *Class Learning Objective Report*.
- Reports generated from the second tool (files that contain the columns *Learning Objective* and *Percent*).

The tool will automatically detect the type of each file, extract all LO percentages,
and then compute the average percentage for each Learning Outcome (LO) in one final merged report.
"""
)

# --------------------------------------------------------
# Function 1: Extract Percent from Remark report
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
# Function 2: Extract Percent from second-tool report
#   (file with columns: Learning Objective, Percent)
# --------------------------------------------------------
def extract_from_lo_report(file_obj, filename):
    try:
        # Read the first sheet assuming header is in the first row
        df = pd.read_excel(file_obj, sheet_name=0)
    except Exception:
        return pd.DataFrame()

    # Normalize column names (to handle small variations)
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

    # Remove empty rows
    sub = sub.dropna(subset=["Learning_Objective"])
    return sub


# --------------------------------------------------------
# File upload UI
# --------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload all files (Remark + second-tool reports)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if st.button("Run Merge"):
    if not uploaded_files:
        st.error("Please upload at least one file.")
    else:
        all_rows = []

        for f in uploaded_files:
            # First, try Remark report
            df_r = extract_from_remark(f, f.name)
            if not df_r.empty:
                all_rows.append(df_r)
                continue

            # If not Remark, try second-tool report
            f.seek(0)
            df_g = extract_from_lo_report(f, f.name)
            if not df_g.empty:
                all_rows.append(df_g)
                continue

            # If file type cannot be detected
            st.warning(f"File type could not be detected: {f.name}")

        if not all_rows:
            st.error("No usable data were extracted from the uploaded files.")
        else:
            merged = pd.concat(all_rows, ignore_index=True)

            # Convert Percent to numeric
            merged["Percent"] = pd.to_numeric(
                merged["Percent"], errors="coerce"
            )
            merged = merged.dropna(subset=["Percent"])

            # Summary table
            summary = (
                merged.groupby("Learning_Objective", as_index=False)
                .agg(
                    Num_Measurements=("Percent", "count"),
                    Mean_Percent=("Percent", "mean"),
                )
                .sort_values("Learning_Objective")
            )

            st.subheader("Detailed results (all files)")
            st.dataframe(merged)

            st.subheader("Merged Learning Outcomes Summary")
            st.dataframe(summary)

            # Prepare Excel for download
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
                "Download Final Merged Report (Excel)",
                data=output,
                file_name="Merged_LO_Percent_Report.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
