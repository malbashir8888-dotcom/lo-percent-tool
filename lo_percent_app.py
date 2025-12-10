import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="LO Percent Tool", layout="centered")

# ------------------------------------------------------------
# UI TEXT (English)
# ------------------------------------------------------------
st.markdown("## Learning Outcomes Percent Aggregation Tool (Percent)")
st.write(
    "Upload Excel files generated from Remark® (Class Learning Objective Report). "
    "The tool will extract the percentage for each Learning Objective and compute the combined averages."
)

uploaded_files = st.file_uploader(
    "Select Excel files (you may upload multiple files)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

run_clicked = st.button("Run Aggregation")

# ------------------------------------------------------------
# Extraction Function
# ------------------------------------------------------------
def extract_percent_from_file(file_obj, filename):
    """Extract Percent values for each Learning Objective from 'Class Learning Objective Report'."""
    try:
        xls = pd.ExcelFile(file_obj)
    except Exception:
        return pd.DataFrame()

    # Detect the correct sheet
    if "Class Learning Objective Report" not in xls.sheet_names:
        return pd.DataFrame()

    df = pd.read_excel(xls, sheet_name="Class Learning Objective Report", header=None)

    # Find header row
    header_row = df.index[df[0] == "Learning Objective"]
    if len(header_row) == 0:
        return pd.DataFrame()

    start = header_row[0] + 1

    rows = []
    for i in range(start, len(df)):
        lo = df.at[i, 0]
        if pd.isna(lo):  # stop when empty row appears
            break
        percent = df.at[i, 5]
        rows.append({
            "Learning_Objective": lo,
            "Percent": percent,
            "Source_File": filename
        })

    return pd.DataFrame(rows)

# ------------------------------------------------------------
# Main Logic
# ------------------------------------------------------------
if run_clicked:
    if not uploaded_files:
        st.error("Please upload at least one file.")
    else:
        all_frames = []
        for f in uploaded_files:
            df_lo = extract_percent_from_file(f, f.name)
            if not df_lo.empty:
                all_frames.append(df_lo)

        if not all_frames:
            st.error("No Percent data could be extracted from the uploaded files.")
        else:
            all_lo = pd.concat(all_frames, ignore_index=True)

            # Summary Table
            summary = (
                all_lo
                .groupby("Learning_Objective", as_index=False)
                .agg(
                    Num_Measurements=("Percent", "count"),
                    Sum_Percent=("Percent", "sum"),
                    Mean_Percent=("Percent", "mean"),
                )
            )

            # Excel Output
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary.to_excel(writer, sheet_name="All_Objectives", index=False)
                all_lo.to_excel(writer, sheet_name="Raw_Percent_All_Files", index=False)
            output.seek(0)

            st.success("Results are ready. You can download the Excel file below.")
            st.download_button(
                label="Download Results (Excel File)",
                data=output,
                file_name="LO_Percent_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
