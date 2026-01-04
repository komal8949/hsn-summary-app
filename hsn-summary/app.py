import streamlit as st
import pandas as pd
from io import BytesIO

st.title("HSN/SAC Summary Generator")

def rename_duplicate_columns(df):
    """Rename duplicate columns by adding a suffix"""
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        dup_idx = cols[cols == dup].index.tolist()
        for i, idx in enumerate(dup_idx):
            if i != 0:
                cols[idx] = f"{dup}_{i}"
    df.columns = cols
    return df

uploaded_file = st.file_uploader("Upload HSN Excel file", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ---------------------------
    # Clean the file
    # ---------------------------
    try:
        header = df[df.eq("HSN/SAC").any(axis=1)].index[0]
    except IndexError:
        st.error("No 'HSN/SAC' header found in the uploaded file.")
        st.stop()

    clean = df.drop(range(header))
    clean.columns = df.loc[header]
    clean = clean.reset_index(drop=True)
    clean = clean[2:].reset_index(drop=True)
    clean = clean.drop(index=0).reset_index(drop=True)

    # Drop unnecessary columns if they exist
    columns_to_drop = ["Description", "Type of", "Total Tax"]
    existing_columns_to_drop = [col for col in columns_to_drop if col in clean.columns]
    df_cleaned = clean.drop(columns=existing_columns_to_drop)

    # ---------------------------
    # Rename duplicate columns
    # ---------------------------
    df_cleaned = rename_duplicate_columns(df_cleaned)

    # ---------------------------
    # Summarize
    # ---------------------------
    summary = df_cleaned.groupby(["HSN/SAC", "UQC"], as_index=False).sum()

    st.subheader("Summary")
    st.dataframe(summary)

    # ---------------------------
    # Download as Excel
    # ---------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
# No writer.save() here
    processed_data = output.getvalue()


    st.download_button(
        label="Download Summary Excel",
        data=processed_data,
        file_name="HSN_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
