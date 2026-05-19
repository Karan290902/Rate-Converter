import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="AVIVA GCL HL → LAP Converter",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# TITLE
# =====================================================

st.title("AVIVA GCL HL → LAP Converter")

st.markdown("""
Upload AVIVA GCL HL Excel file  
Get final transformed LAP output instantly
""")

# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_file = st.file_uploader(
    "Upload AVIVA GCL HL Excel",
    type=["xlsx"]
)

# =====================================================
# PROCESS FUNCTION
# =====================================================

def process_file(df):

    # ---------------------------------------------
    # CLEAN COLUMN NAMES
    # ---------------------------------------------

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # ---------------------------------------------
    # FIND AGE COLUMN
    # ---------------------------------------------

    age_column = df.columns[0]

    # ---------------------------------------------
    # CONVERT TO LONG FORMAT
    # ---------------------------------------------

    df_long = df.melt(
        id_vars=[age_column],
        var_name="Tenure",
        value_name="Premium"
    )

    # ---------------------------------------------
    # REMOVE BLANKS
    # ---------------------------------------------

    df_long.dropna(inplace=True)

    # ---------------------------------------------
    # CONVERT PREMIUM TO NUMERIC
    # ---------------------------------------------

    df_long["Premium"] = pd.to_numeric(
        df_long["Premium"],
        errors="coerce"
    )

    df_long.dropna(inplace=True)

    # ---------------------------------------------
    # CALCULATIONS
    # ---------------------------------------------

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

    # ---------------------------------------------
    # RENAME AGE COLUMN
    # ---------------------------------------------

    df_long.rename(
        columns={
            age_column: "Age"
        },
        inplace=True
    )

    # ---------------------------------------------
    # SORT VALUES
    # ---------------------------------------------

    df_long.sort_values(
        by=["Age", "Tenure"],
        inplace=True
    )

    # ---------------------------------------------
    # RESET INDEX
    # ---------------------------------------------

    df_long.reset_index(
        drop=True,
        inplace=True
    )

    return df_long

# =====================================================
# CREATE EXCEL
# =====================================================

def to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="LAP Output"
        )

        workbook = writer.book
        worksheet = writer.sheets["LAP Output"]

        # -----------------------------------------
        # FORMAT HEADERS
        # -----------------------------------------

        for cell in worksheet[1]:

            cell.font = cell.font.copy(
                bold=True
            )

        # -----------------------------------------
        # AUTO WIDTH
        # -----------------------------------------

        for column_cells in worksheet.columns:

            length = max(
                len(str(cell.value))
                if cell.value else 0
                for cell in column_cells
            )

            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = length + 5

    output.seek(0)

    return output

# =====================================================
# MAIN
# =====================================================

if uploaded_file is not None:

    try:

        # -----------------------------------------
        # READ FILE
        # -----------------------------------------

        df = pd.read_excel(uploaded_file)

        # -----------------------------------------
        # SHOW INPUT
        # -----------------------------------------

        st.subheader("Input File Preview")

        st.dataframe(df.head())

        # -----------------------------------------
        # PROCESS
        # -----------------------------------------

        final_df = process_file(df)

        # -----------------------------------------
        # SHOW OUTPUT
        # -----------------------------------------

        st.subheader("Processed Output")

        st.dataframe(final_df)

        # -----------------------------------------
        # STATS
        # -----------------------------------------

        st.success(
            f"Successfully processed {len(final_df)} rows"
        )

        # -----------------------------------------
        # GENERATE EXCEL
        # -----------------------------------------

        excel_file = to_excel(final_df)

        # -----------------------------------------
        # DOWNLOAD BUTTON
        # -----------------------------------------

        st.download_button(
            label="📥 Download Final Excel",
            data=excel_file,
            file_name="AVIVA_GCL_LAP_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error processing file: {str(e)}"
        )