import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Rate Converter",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# TITLE
# =====================================================

st.title("AVIVA GCL HL → LAP Converter")

st.write("Upload Excel and download transformed output.")

# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# =====================================================
# PROCESS FUNCTION
# =====================================================

@st.cache_data
def process_file(df):

    # Clean columns
    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # First column = Age
    age_column = df.columns[0]

    # Convert wide → long
    df_long = df.melt(
        id_vars=[age_column],
        var_name="Tenure",
        value_name="Premium"
    )

    # Remove blanks
    df_long.dropna(inplace=True)

    # Convert numeric
    df_long["Premium"] = pd.to_numeric(
        df_long["Premium"],
        errors="coerce"
    )

    df_long.dropna(inplace=True)

    # Calculations
    df_long["Wellness"] = (
        df_long["Premium"] * 0.05
    ).round(2)

    df_long["Net"] = (
        df_long["Premium"] +
        df_long["Wellness"]
    ).round(2)

    df_long["Gross"] = (
        df_long["Net"] * 1.18
    ).round(2)

    # Rename age column
    df_long.rename(
        columns={
            age_column: "Age"
        },
        inplace=True
    )

    # Reorder columns
    df_long = df_long[
        [
            "Age",
            "Tenure",
            "Premium",
            "Wellness",
            "Net",
            "Gross"
        ]
    ]

    return df_long

# =====================================================
# EXCEL CREATION
# =====================================================

@st.cache_data
def convert_to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Output"
        )

    output.seek(0)

    return output

# =====================================================
# MAIN
# =====================================================

if uploaded_file is not None:

    try:

        # Read Excel
        df = pd.read_excel(uploaded_file)

        st.success("File uploaded successfully")

        # Process
        final_df = process_file(df)

        # Show only few rows
        st.subheader("Preview")

        st.dataframe(
            final_df.head(50),
            use_container_width=True
        )

        # Row count
        st.info(
            f"Processed Rows: {len(final_df)}"
        )

        # Excel generation
        excel_file = convert_to_excel(final_df)

        # Download button
        st.download_button(
            label="📥 Download Final Excel",
            data=excel_file,
            file_name="AVIVA_GCL_LAP_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(f"Error: {str(e)}")
