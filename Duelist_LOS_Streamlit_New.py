import streamlit as st
import pandas as pd
import io

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Duelist for LOS",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# TITLE
# ============================================================

st.title("📊 Duelist for LOS Generator")
st.markdown(
    "Upload the original Duelist file and generate a new "
    "Duelist for LOS."
)

# ============================================================
# REQUIRED COLUMNS
# ============================================================

required_columns = [
    "BBB",
    "AT",
    "AcTypeDesc",
    "CC",
    "OldClientCode",
    "ClientCode",
    "OldAcNum",
    "MainCode",
    "Name",
    "OfficerName",
    "Loan Type",
    "Ageing",
    "Address",
    "Phone",
    "MobileNo",
    "ContactPerson",
    "Mobile1",
    "Nominee",
    "NomBaln",
    "AcOpenDate",
    "FirstInstDueDate",
    "ExpiryDate",
    "IntDrRate",
    "PenalIntRate",
    "Limit",
    "TotDisburse",
    "OutstandingBaln",
    "IntDrAmt",
    "PenalIntAmt",
    "IntOnInt",
    "OvDuePrin",
    "OvDueInt",
    "PastDuedInt",
    "TotCharge",
    "TotOvDue",
    "InsurancePremium",
    "OvDueWithInsurance",
    "AcCodeForChg",
    "OvDueInst",
    "DuesFrom",
    "AgeingDays",
    "Bucket",
    "PrinPaid",
    "IntPaid",
    "PenalPaid",
    "IntOnIntPaid",
    "ManualIntPaid",
    "NoOfPeriods",
    "RemainingPeriods",
    "InstAmt",
    "LastTranDate",
    "LastTranAmt",
    "Remarks",
    "AcOfficer",
    "Securities",
    "BranchName",
    "Dealer Name",
    "BorrowingBank",
    "ModelDetails",
    "LotNumber",
    "DownPayment",
    "VehicleModel",
    "VehicleNo",
    "Region",
    "DistrictCode"
]

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Options")

    uploaded_file = st.file_uploader(
        "Upload Duelist File",
        type=["xlsx", "xls", "xlsm", "xlsb"]
    )

    sheet_name = st.text_input(
        "Sheet Name",
        value="Mainsheet"
    )

    generate = st.button(
        "🚀 Generate Duelist for LOS",
        type="primary",
        use_container_width=True
    )

# ============================================================
# NO FILE
# ============================================================

if uploaded_file is None:

    st.info(
        "Upload original Duelist file and click Generate."
    )

    st.stop()

# ============================================================
# PROCESS
# ============================================================

if generate:

    try:

        with st.spinner("Processing..."):

            # ====================================================
            # READ FILE
            # ====================================================

            if uploaded_file.name.lower().endswith(".xlsb"):

                df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    engine="pyxlsb",
                    dtype={
                        "BBB": str,
                        "ClientCode": str,
                        "MainCode": str,
                        "AcCodeForChg": str,
                        "Nominee": str
                    }
                )

            else:

                df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    dtype={
                        "BBB": str,
                        "ClientCode": str,
                        "MainCode": str,
                        "AcCodeForChg": str,
                        "Nominee": str
                    }
                )

            # ====================================================
            # CLEAN COLUMN NAMES
            # ====================================================

            df.columns = (
                df.columns
                .astype(str)
                .str.strip()
            )

            # ====================================================
            # CHECK MISSING COLUMNS
            # ====================================================

            missing_columns = [
                col
                for col in required_columns
                if col not in df.columns
            ]

            if len(missing_columns) > 0:

                st.error(
                    f"{len(missing_columns)} columns missing."
                )

                st.dataframe(
                    pd.DataFrame({
                        "Missing Columns": missing_columns
                    }),
                    use_container_width=True
                )

                st.stop()

            # ====================================================
            # ORIGINAL SUMMARY
            # ====================================================

            original_rows = len(df)

            original_columns = len(df.columns)

            original_unique_at = (
                df["AT"].nunique(dropna=True)
            )

            original_duplicate_maincode = (
                df["MainCode"]
                .duplicated()
                .sum()
            )

            original_duplicate_clientcode = (
                df["ClientCode"]
                .duplicated()
                .sum()
            )

            original_blank_maincode = (
                df["MainCode"]
                .isna()
                .sum()
            )

            original_blank_clientcode = (
                df["ClientCode"]
                .isna()
                .sum()
            )

            # ====================================================
            # CREATE LOS FILE
            # ====================================================

            result = df[
                required_columns
            ].copy()

            # Add CategoryID

            result["CategoryID"] = 1

            # ====================================================
            # FINAL SUMMARY
            # ====================================================

            final_rows = len(result)

            final_columns = len(result.columns)

            final_unique_at = (
                result["AT"].nunique(dropna=True)
            )

            final_duplicate_maincode = (
                result["MainCode"]
                .duplicated()
                .sum()
            )

            final_duplicate_clientcode = (
                result["ClientCode"]
                .duplicated()
                .sum()
            )

            final_blank_maincode = (
                result["MainCode"]
                .isna()
                .sum()
            )

            final_blank_clientcode = (
                result["ClientCode"]
                .isna()
                .sum()
            )

            # ====================================================
            # SUMMARY TABLE
            # ====================================================

            summary = pd.DataFrame({

                "Metric": [
                    "Total Accounts",
                    "Total Columns",
                    "Unique AT",
                    "Duplicate MainCode",
                    "Duplicate ClientCode",
                    "Blank MainCode",
                    "Blank ClientCode"
                ],

                "Original Duelist": [
                    original_rows,
                    original_columns,
                    original_unique_at,
                    original_duplicate_maincode,
                    original_duplicate_clientcode,
                    original_blank_maincode,
                    original_blank_clientcode
                ],

                "Duelist for LOS": [
                    final_rows,
                    final_columns,
                    final_unique_at,
                    final_duplicate_maincode,
                    final_duplicate_clientcode,
                    final_blank_maincode,
                    final_blank_clientcode
                ]
            })

            # ====================================================
            # CREATE OUTPUT FILE
            # ====================================================

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                result.to_excel(
                    writer,
                    sheet_name="Mainsheet",
                    index=False
                )

            output.seek(0)

            # ====================================================
            # DISPLAY
            # ====================================================

            st.success(
                "✅ Duelist for LOS generated successfully."
            )

            # Summary

            st.subheader(
                "📊 Original Duelist vs Duelist for LOS"
            )

            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True
            )

            # Preview

            st.subheader(
                "👀 Preview"
            )

            st.dataframe(
                result.head(20),
                use_container_width=True,
                height=500
            )

            # Download

            st.download_button(
                label="📥 Download Duelist for LOS New.xlsx",
                data=output.getvalue(),
                file_name="Duelist for LOS New.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    except Exception as e:

        st.error(
            "Unable to process file."
        )

        st.exception(e)