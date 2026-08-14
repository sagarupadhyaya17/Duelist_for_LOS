import pandas as pd
import numpy as np
import io

# ============================================================
# FILE PATHS
# ============================================================

input_file = "Input_Files/Duelist 12th Aug, 2026.xlsx"
output_file = "Output_Files/Duelist for LOS New.xlsx"

sheet_name = "Mainsheet"


# ============================================================
# READ EXCEL FILE
# ============================================================

df = pd.read_excel(
    input_file,
    sheet_name=sheet_name,
    dtype={"BBB": str, "ClientCode": str, "MainCode": str, "AcCodeForChg": str, "Nominee": str}
)

# Clean column names
df.columns = df.columns.astype(str).str.strip()

df.head(5)

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
# CHECK WHETHER REQUIRED COLUMNS EXIST
# ============================================================

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:
    print("\n========================================")
    print("MISSING COLUMNS")
    print("========================================")

    for col in missing_columns:
        print(f"- {col}")

    print("========================================\n")

    raise ValueError(
        f"{len(missing_columns)} required column(s) are missing."
    )


# ============================================================
# EXTRACT ONLY REQUIRED COLUMNS
# ============================================================

result = df[required_columns].copy()

result.head(5)

# ============================================================
# FIND MOST COMMON LOAN TYPE FOR EACH AT
# ============================================================

def get_most_common_loan_type(series):

    # Remove blank/null values
    series = (
        series
        .dropna()
        .astype(str)
        .str.strip()
    )

    series = series[series != ""]

    # If no valid Loan Type exists
    if series.empty:
        return np.nan

    # Count each Loan Type
    counts = series.value_counts()

    # Return the most frequent Loan Type
    return counts.index[0]


# Create mapping:
# AT -> Most common Loan Type

at_loan_type = (
    df.groupby("AT", dropna=False)["Loan Type"]
      .agg(get_most_common_loan_type)
)


# ============================================================
# REPLACE LOAN TYPE USING AT
# ============================================================

result["Loan Type"] = result["AT"].map(at_loan_type)


# ============================================================
# ADD CATEGORY ID
# ============================================================

result["CategoryID"] = 1

# ============================================================
# SUMMARY OF ORIGINAL VS FINAL DUELIST
# ============================================================

# ---------- ORIGINAL DUELIST SUMMARY ----------

original_rows = len(df)
original_columns = len(df.columns)

original_unique_at = df["AT"].nunique(dropna=True)
original_unique_loan_type = df["Loan Type"].nunique(dropna=True)

original_duplicate_maincode = df["MainCode"].duplicated().sum()
original_duplicate_clientcode = df["ClientCode"].duplicated().sum()

original_blank_maincode = df["MainCode"].isna().sum()
original_blank_clientcode = df["ClientCode"].isna().sum()


# ---------- FINAL DUELIST SUMMARY ----------

final_rows = len(result)
final_columns = len(result.columns)

final_unique_at = result["AT"].nunique(dropna=True)
final_unique_loan_type = result["Loan Type"].nunique(dropna=True)

final_duplicate_maincode = result["MainCode"].duplicated().sum()
final_duplicate_clientcode = result["ClientCode"].duplicated().sum()

final_blank_maincode = result["MainCode"].isna().sum()
final_blank_clientcode = result["ClientCode"].isna().sum()

category_id_values = result["CategoryID"].unique()


# ============================================================
# CREATE SUMMARY TABLE
# ============================================================

summary = pd.DataFrame({
    "Summary": [
        "Total Accounts",
        "Total Columns",
        "Unique AT",
        "Unique Loan Type",
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
        final_duplicate_maincode,
        final_duplicate_clientcode,
        final_blank_maincode,
        final_blank_clientcode,
        ", ".join(map(str, category_id_values))
    ]
})


# ============================================================
# PRINT SUMMARY
# ============================================================

print("\n")
print("=" * 70)
print("              DUELIST PROCESSING SUMMARY")
print("=" * 70)

print(summary.to_string(index=False))

print("=" * 70)

# ============================================================
# AT-WISE LOAN TYPE SUMMARY
# ============================================================

loan_type_summary = (
    df.groupby(["AT", "Loan Type"], dropna=False)
      .size()
      .reset_index(name="Count")
      .sort_values(
          ["AT", "Count"],
          ascending=[True, False]
      )
)

loan_type_summary["Final Loan Type"] = (
    loan_type_summary["AT"]
    .map(at_loan_type)
)


print("\n")
print("=" * 70)
print("           AT-WISE LOAN TYPE SUMMARY")
print("=" * 70)

print(loan_type_summary.to_string(index=False))

# ============================================================
# EXPORT
# ============================================================

# Final Duelist

result.to_excel(
    output_file,
    sheet_name="Mainsheet",
    index=False
)