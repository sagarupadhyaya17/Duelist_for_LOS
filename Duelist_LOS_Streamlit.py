import streamlit as st
import pandas as pd
import numpy as np
import io
import time


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Duelist for LOS",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>

.main-title {
    font-size: 34px;
    font-weight: 700;
    color: #1f4e78;
    margin-bottom: 5px;
}

.sub-title {
    font-size: 16px;
    color: #666666;
    margin-bottom: 25px;
}

.section-title {
    font-size: 22px;
    font-weight: 600;
    color: #1f4e78;
    margin-top: 25px;
    margin-bottom: 12px;
}

.sidebar-title {
    font-size: 22px;
    font-weight: 700;
    color: #1f4e78;
}

div.stButton > button {
    width: 100%;
    height: 45px;
    font-size: 16px;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# TITLE
# ============================================================

st.markdown(
    '<div class="main-title">'
    '📊 Duelist for LOS Processing System'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    'Upload Duelist → Process → Analyze → Preview → Download'
    '</div>',
    unsafe_allow_html=True
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
# SESSION STATE
# ============================================================

if "processed" not in st.session_state:
    st.session_state.processed = False

if "result" not in st.session_state:
    st.session_state.result = None

if "output_file" not in st.session_state:
    st.session_state.output_file = None

if "summary" not in st.session_state:
    st.session_state.summary = None

if "loan_type_summary" not in st.session_state:
    st.session_state.loan_type_summary = None

if "multiple_loan_type" not in st.session_state:
    st.session_state.multiple_loan_type = None


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        '<div class="sidebar-title">'
        '⚙️ Processing Options'
        '</div>',
        unsafe_allow_html=True
    )

    st.write("")

    uploaded_file = st.file_uploader(
        "📂 Upload Duelist Excel File",
        type=["xlsx", "xls"]
    )

    sheet_name = st.text_input(
        "📄 Sheet Name",
        value="Mainsheet"
    )

    st.divider()

    process_button = st.button(
        "🚀 Generate New Duelist for LOS",
        type="primary",
        use_container_width=True
    )

    st.divider()

    st.markdown("### 📌 Processing Steps")

    st.write("1. 📂 Read Excel")
    st.write("2. 🔍 Validate Columns")
    st.write("3. 📋 Extract Required Columns")
    st.write("4. 🔄 Find Highest-Count Loan Type")
    st.write("5. 🏷️ Add CategoryID")
    st.write("6. 📊 Create Summaries")
    st.write("7. 📦 Prepare Mainsheet")


# ============================================================
# GENERATE CLICKED WITHOUT FILE
# ============================================================

if process_button and uploaded_file is None:

    st.error(
        "⚠️ Please upload the Duelist Excel file first "
        "before generating the New Duelist for LOS."
    )

    st.stop()


# ============================================================
# NO FILE UPLOADED
# ============================================================

if uploaded_file is None:

    st.info(
        "👈 Please upload the Duelist Excel file "
        "from the Processing Options panel."
    )

    st.markdown("""
    ### How to use

    1. Upload the original Duelist Excel file.
    2. Make sure the sheet name is `Mainsheet`.
    3. Click **Generate New Duelist for LOS**.
    4. Wait for processing to complete.
    5. Review the summaries.
    6. Preview the New Duelist.
    7. Download the final Excel.

    **The downloaded Excel contains only `Mainsheet`.**
    """)

    st.stop()


# ============================================================
# READ EXCEL FILE
# ============================================================

try:

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

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

except Exception as e:

    st.error(
        f"❌ Unable to read Excel file or sheet "
        f"'{sheet_name}'."
    )

    st.exception(e)

    st.stop()


# ============================================================
# FILE INFORMATION
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📁 Uploaded File'
    '</div>',
    unsafe_allow_html=True
)

st.success(
    f"✅ File uploaded successfully: **{uploaded_file.name}**"
)


# ============================================================
# CHECK REQUIRED COLUMNS
# ============================================================

missing_columns = [
    col
    for col in required_columns
    if col not in df.columns
]


if missing_columns:

    st.error(
        f"❌ {len(missing_columns)} required column(s) "
        "are missing."
    )

    st.markdown("### Missing Columns")

    st.dataframe(
        pd.DataFrame({
            "Missing Column": missing_columns
        }),
        use_container_width=True,
        hide_index=True
    )

    st.stop()


st.success(
    "✅ All required columns are available."
)


# ============================================================
# ORIGINAL DUELIST BASIC SUMMARY
# ============================================================

original_rows = len(df)

original_columns = len(df.columns)

original_unique_at = (
    df["AT"].nunique(dropna=True)
)

original_unique_loan_type = (
    df["Loan Type"].nunique(dropna=True)
)


# ============================================================
# ORIGINAL DUELIST METRICS
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📋 Original Duelist Information'
    '</div>',
    unsafe_allow_html=True
)

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Total Accounts",
        original_rows
    )

with col2:
    st.metric(
        "Total Columns",
        original_columns
    )

with col3:
    st.metric(
        "Unique AT",
        original_unique_at
    )

with col4:
    st.metric(
        "Unique Loan Type",
        original_unique_loan_type
    )


# ============================================================
# ORIGINAL AT WITH MORE THAN ONE LOAN TYPE
# ============================================================

st.markdown(
    '<div class="section-title">'
    '🔎 Original Duelist — AT with More Than One Loan Type'
    '</div>',
    unsafe_allow_html=True
)


temp_df = df[
    ["AT", "Loan Type"]
].copy()


# Clean Loan Type
temp_df["Loan Type"] = (
    temp_df["Loan Type"]
    .astype("string")
    .str.strip()
)


# Remove blank Loan Types
temp_df = temp_df[
    temp_df["Loan Type"].notna()
    &
    (temp_df["Loan Type"] != "")
]


# Count unique Loan Types per AT
original_at_loan_type_counts = (
    temp_df
    .groupby(
        "AT",
        dropna=False
    )["Loan Type"]
    .nunique()
    .reset_index(
        name="Number of Loan Types"
    )
)


# Keep only AT having more than one Loan Type
multiple_loan_type = (
    original_at_loan_type_counts[
        original_at_loan_type_counts[
            "Number of Loan Types"
        ] > 1
    ]
    .sort_values(
        "Number of Loan Types",
        ascending=False
    )
    .reset_index(drop=True)
)


if len(multiple_loan_type) > 0:

    st.success(
        f"Found **{len(multiple_loan_type)} AT(s)** "
        "having more than one Loan Type."
    )

    st.dataframe(
        multiple_loan_type,
        use_container_width=True,
        hide_index=True
    )

else:

    st.info(
        "No AT has more than one Loan Type."
    )


# ============================================================
# PROCESSING
# ============================================================

if process_button:

    st.session_state.processed = False
    st.session_state.result = None
    st.session_state.output_file = None
    st.session_state.summary = None
    st.session_state.loan_type_summary = None
    st.session_state.multiple_loan_type = None

    st.markdown(
        '<div class="section-title">'
        '⚙️ Processing Duelist'
        '</div>',
        unsafe_allow_html=True
    )

    progress_bar = st.progress(0)

    status_text = st.empty()

    with st.spinner(
        "⏳ Processing Duelist... Please wait."
    ):

        # ====================================================
        # STEP 1
        # ====================================================

        status_text.info(
            "📂 Step 1/7 — Reading Excel file..."
        )

        progress_bar.progress(10)

        time.sleep(0.3)


        # ====================================================
        # STEP 2
        # ====================================================

        status_text.info(
            "🔍 Step 2/7 — Validating required columns..."
        )

        progress_bar.progress(20)

        time.sleep(0.3)


        # ====================================================
        # STEP 3
        # ====================================================

        status_text.info(
            "📋 Step 3/7 — Extracting required columns..."
        )

        progress_bar.progress(35)

        result = df[
            required_columns
        ].copy()

        time.sleep(0.3)


        # ====================================================
        # STEP 4
        # ====================================================

        status_text.info(
            "🔄 Step 4/7 — Finding highest-count "
            "Loan Type for each AT..."
        )

        progress_bar.progress(50)


        # ====================================================
        # FUNCTION: MOST COMMON LOAN TYPE
        # ====================================================

        def get_most_common_loan_type(series):

            series = (
                series
                .dropna()
                .astype(str)
                .str.strip()
            )

            series = series[
                series != ""
            ]

            if series.empty:
                return np.nan

            counts = series.value_counts()

            return counts.index[0]


        # ====================================================
        # AT -> HIGHEST COUNT LOAN TYPE
        # ====================================================

        at_loan_type = (
            df
            .groupby(
                "AT",
                dropna=False
            )["Loan Type"]
            .agg(
                get_most_common_loan_type
            )
        )


        # ====================================================
        # REPLACE LOAN TYPE
        # ====================================================

        result["Loan Type"] = (
            result["AT"]
            .map(at_loan_type)
        )

        time.sleep(0.3)


        # ====================================================
        # STEP 5
        # ====================================================

        status_text.info(
            "🏷️ Step 5/7 — Adding CategoryID..."
        )

        progress_bar.progress(65)

        result["CategoryID"] = 1

        time.sleep(0.3)


        # ====================================================
        # STEP 6
        # ====================================================

        status_text.info(
            "📊 Step 6/7 — Creating summaries..."
        )

        progress_bar.progress(75)


        # ====================================================
        # ORIGINAL METRICS
        # ====================================================

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
        # ORIGINAL AT + LOAN TYPE COMBINATIONS
        # ====================================================

        original_at_loan_type = (
            df[
                ["AT", "Loan Type"]
            ]
            .dropna(
                subset=["AT", "Loan Type"]
            )
            .drop_duplicates()
        )

        original_at_loan_type_count = (
            len(original_at_loan_type)
        )


        # ====================================================
        # ORIGINAL MULTIPLE LOAN TYPE COUNT
        # ====================================================

        original_multiple_loan_type_count = (
            len(multiple_loan_type)
        )


        # ====================================================
        # FINAL METRICS
        # ====================================================

        final_rows = len(result)

        final_columns = len(result.columns)

        final_unique_at = (
            result["AT"]
            .nunique(dropna=True)
        )

        final_unique_loan_type = (
            result["Loan Type"]
            .nunique(dropna=True)
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
        # FINAL AT + LOAN TYPE COMBINATIONS
        # ====================================================

        final_at_loan_type = (
            result[
                ["AT", "Loan Type"]
            ]
            .dropna(
                subset=["AT", "Loan Type"]
            )
            .drop_duplicates()
        )

        final_at_loan_type_count = (
            len(final_at_loan_type)
        )


        # ====================================================
        # FINAL AT WITH MORE THAN ONE LOAN TYPE
        # ====================================================

        final_at_loan_type_counts = (
            result[
                ["AT", "Loan Type"]
            ]
            .dropna(
                subset=["AT", "Loan Type"]
            )
            .groupby("AT")["Loan Type"]
            .nunique()
        )

        final_multiple_loan_type_count = (
            (
                final_at_loan_type_counts > 1
            ).sum()
        )


        # ====================================================
        # CATEGORY ID
        # ====================================================

        category_id_values = (
            result["CategoryID"]
            .unique()
        )


        # ====================================================
        # ORIGINAL VS FINAL SUMMARY
        # ====================================================

        summary = pd.DataFrame({

            "Summary": [

                "Total Accounts",

                "Total Columns",

                "Unique AT",

                "Unique Loan Type",

                "AT + Loan Type Combinations",

                "AT with More Than 1 Loan Type",

                "Duplicate MainCode",

                "Duplicate ClientCode",

                "Blank MainCode",

                "Blank ClientCode",

                "CategoryID"

            ],

            "Original Duelist": [

                original_rows,

                original_columns,

                original_unique_at,

                original_unique_loan_type,

                original_at_loan_type_count,

                original_multiple_loan_type_count,

                original_duplicate_maincode,

                original_duplicate_clientcode,

                original_blank_maincode,

                original_blank_clientcode,

                "-"

            ],

            "Final Duelist": [

                final_rows,

                final_columns,

                final_unique_at,

                final_unique_loan_type,

                final_at_loan_type_count,

                final_multiple_loan_type_count,

                final_duplicate_maincode,

                final_duplicate_clientcode,

                final_blank_maincode,

                final_blank_clientcode,

                ", ".join(
                    map(
                        str,
                        category_id_values
                    )
                )

            ]

        })


        # ====================================================
        # AT-WISE LOAN TYPE SUMMARY
        # ====================================================

        loan_type_summary = (
            df
            .groupby(
                ["AT", "Loan Type"],
                dropna=False
            )
            .size()
            .reset_index(
                name="Count"
            )
            .sort_values(
                ["AT", "Count"],
                ascending=[True, False]
            )
        )


        # ====================================================
        # FINAL LOAN TYPE
        # ====================================================

        loan_type_summary[
            "Final Loan Type"
        ] = (
            loan_type_summary["AT"]
            .map(at_loan_type)
        )


        time.sleep(0.5)


        # ====================================================
        # STEP 7
        # ====================================================

        status_text.info(
            "📦 Step 7/7 — Creating final Mainsheet..."
        )

        progress_bar.progress(90)


        # ====================================================
        # FINAL VALIDATION
        # ====================================================

        if (
            final_at_loan_type_count
            != final_unique_at
        ):

            st.error(
                "❌ Final Duelist validation failed. "
                "Some ATs have multiple Loan Types."
            )

            st.stop()


        if final_multiple_loan_type_count != 0:

            st.error(
                "❌ Final Duelist contains ATs with "
                "more than one Loan Type."
            )

            st.stop()


        # ====================================================
        # CREATE EXCEL
        #
        # ONLY MAIN SHEET
        # ====================================================

        output_buffer = io.BytesIO()

        with pd.ExcelWriter(
            output_buffer,
            engine="openpyxl"
        ) as writer:

            # IMPORTANT:
            # Write to writer, NOT output_buffer

            result.to_excel(
                writer,
                sheet_name="Mainsheet",
                index=False
            )

        output_buffer.seek(0)


        # ====================================================
        # STORE RESULTS
        # ====================================================

        st.session_state.result = result

        st.session_state.output_file = (
            output_buffer.getvalue()
        )

        st.session_state.summary = summary

        st.session_state.loan_type_summary = (
            loan_type_summary
        )

        st.session_state.multiple_loan_type = (
            multiple_loan_type
        )

        st.session_state.processed = True


        progress_bar.progress(100)

        status_text.success(
            "✅ Processing completed successfully!"
        )


# ============================================================
# DISPLAY RESULTS
# ============================================================

if st.session_state.processed:

    result = st.session_state.result

    summary = st.session_state.summary

    loan_type_summary = (
        st.session_state.loan_type_summary
    )

    multiple_loan_type = (
        st.session_state.multiple_loan_type
    )


    # ========================================================
    # SUCCESS
    # ========================================================

    st.success(
        "🎉 New Duelist for LOS generated successfully!"
    )


    # ========================================================
    # ORIGINAL VS FINAL COMPARISON
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '📊 Original vs Final Duelist Comparison'
        '</div>',
        unsafe_allow_html=True
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # VALIDATION
    # ========================================================

    final_at_count = (
        result["AT"]
        .nunique(dropna=True)
    )

    final_at_loan_count = (
        result[
            ["AT", "Loan Type"]
        ]
        .dropna(
            subset=["AT", "Loan Type"]
        )
        .drop_duplicates()
        .shape[0]
    )


    if (
        final_at_count
        == final_at_loan_count
        and final_at_count > 0
    ):

        st.success(
            "✅ Validation Passed: Every AT in the "
            "Final Duelist has exactly one Loan Type."
        )

    elif final_at_count == 0:

        st.warning(
            "⚠️ No valid AT values were found."
        )

    else:

        st.error(
            "❌ Validation Failed: Some ATs have "
            "multiple Loan Types."
        )


    # ========================================================
    # AT-WISE LOAN TYPE SUMMARY
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '📚 AT-wise Loan Type Summary'
        '</div>',
        unsafe_allow_html=True
    )


    with st.expander(
        "🔎 View AT-wise Loan Type Summary"
    ):

        st.dataframe(
            loan_type_summary,
            use_container_width=True,
            height=500,
            hide_index=True
        )


    # ========================================================
    # FINAL DUELIST METRICS
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '📈 Final Duelist'
        '</div>',
        unsafe_allow_html=True
    )


    col1, col2, col3, col4 = st.columns(4)


    with col1:

        st.metric(
            "Total Accounts",
            len(result)
        )


    with col2:

        st.metric(
            "Total Columns",
            len(result.columns)
        )


    with col3:

        st.metric(
            "Unique AT",
            result["AT"].nunique(
                dropna=True
            )
        )


    with col4:

        st.metric(
            "Unique Loan Type",
            result["Loan Type"].nunique(
                dropna=True
            )
        )


    # ========================================================
    # NEW DUELIST PREVIEW
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '👀 New Duelist Preview'
        '</div>',
        unsafe_allow_html=True
    )


    preview_max = min(
        100,
        len(result)
    )


    if preview_max > 0:

        preview_min = (
            5
            if preview_max >= 5
            else 1
        )

        preview_default = min(
            10,
            preview_max
        )

        preview_rows = st.slider(
            "Number of rows to preview",
            min_value=preview_min,
            max_value=preview_max,
            value=preview_default
        )

        st.dataframe(
            result.head(preview_rows),
            use_container_width=True,
            height=500,
            hide_index=True
        )

    else:

        st.warning(
            "⚠️ Final Duelist contains no rows."
        )


    # ========================================================
    # DOWNLOAD SECTION
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '⬇️ Download New Duelist for LOS'
        '</div>',
        unsafe_allow_html=True
    )


    st.info(
        "📌 The downloaded Excel contains "
        "**only the Mainsheet**. "
        "Summary tables are displayed only "
        "in this Streamlit application."
    )


    st.download_button(
        label="📥 Download Duelist for LOS New.xlsx",
        data=st.session_state.output_file,
        file_name="Duelist for LOS New.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )