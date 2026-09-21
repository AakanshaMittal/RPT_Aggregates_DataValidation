import pandas as pd
from pathlib import Path

# =====================================================
# CONFIGURATION
# =====================================================

RPT_MAPPINGS = {

    "APPRPT1": "RPT6683971780115LC",
    "EXCRPT1": "RPT260820171120D1U",

    "APPRPT2": "RPT3046986096719OH",
    "EXCRPT2": "RPT260820171622X0Y",

    "APPRPT3": "RPT4154543496776MR",
    "EXCRPT3": "RPT260820171356O2P",

    "APPRPT4": "RPT4154543496776MR",
    "EXCRPT4": "RPT260820171356O2P"
    

    # Add more pairs as needed
}
OUTPUT_RPT_APP_FOLDER = Path("data/outputRpt/File1")
OUTPUT_RPT_EXC_FOLDER = Path("data/excel_files")
REPORT_OUTPUT_FOLDER = Path("data/outputRpt/Aggregates")

MAX_T_FILES = 3

SHEET_NAME = "Output RPT"

REPORT_OUTPUT_FOLDER.mkdir(exist_ok=True)

# =====================================================
# LOAD EXCEL FROM OUTPUT RPT SHEET
# =====================================================

def load_excel(path, file_type):

    excel_file = pd.ExcelFile(
        path,
        engine="openpyxl"
    )

    print(f"\nWorkbook: {path.name}")
    print("Available Sheets:")

    for sheet in excel_file.sheet_names:
        print(repr(sheet))

    df = pd.read_excel(
        path,
        sheet_name=SHEET_NAME,
        header=None,
        engine="openpyxl"
    )

    df = df.iloc[5:].reset_index(drop=True)

    result = pd.DataFrame()

    result["Key"] = df.iloc[:, 0]

    # =====================================================
    # Existing C-J logic (UNCHANGED)
    # =====================================================

    for col_idx in range(2, min(10, len(df.columns))):

        col_data = df.iloc[:, col_idx]

        if col_data.notna().any():

            col_name = chr(65 + col_idx)

            result[col_name] = col_data

    # =====================================================
    # K onwards - read actual RPT IDs from row 8
    # =====================================================

    rpt_header_row = 2

    for col_idx in range(10, len(df.columns)):

        rpt_id = df.iloc[rpt_header_row, col_idx]

        if pd.isna(rpt_id):
            continue

        rpt_id = str(rpt_id).strip()

        if rpt_id == "":
            continue

        result[rpt_id] = df.iloc[:, col_idx]

    result = result[result["Key"].notna()]

    result["Key"] = result["Key"].astype(str).str.strip()

    result = result[result["Key"] != ""]
    print(f"\nLoaded Columns from {path.name}:")
    print(list(result.columns))
    return result
# =====================================================
# COMPARE EXCELS
# =====================================================

def compare_excels(file1, file2, report_path):

    # file1 = EXC
    # file2 = APP

    df1 = load_excel(file1, "EXC")
    df2 = load_excel(file2, "APP")

    # =====================================================
    # Existing Columns C-J
    # =====================================================

    regular_columns = []

    for col in df2.columns:

        if col == "Key":
            continue

        if col in df1.columns:
            regular_columns.append(col)

    # =====================================================
    # RPT Pairing using mapping
    # =====================================================

    rpt_pairs = []

    pair_no = 1

    while True:

        exc_key = f"EXCRPT{pair_no}"
        app_key = f"APPRPT{pair_no}"

        if (
            exc_key not in RPT_MAPPINGS
            or app_key not in RPT_MAPPINGS
        ):
            break

        exc_rpt_id = RPT_MAPPINGS[exc_key]
        app_rpt_id = RPT_MAPPINGS[app_key]

        if (
            exc_rpt_id in df1.columns
            and app_rpt_id in df2.columns
        ):

            rpt_pairs.append(
                (
                    exc_rpt_id,
                    app_rpt_id
                )
            )

            print(
                f"Found RPT Pair: "
                f"{exc_rpt_id} <-> {app_rpt_id}"
            )

        else:

            print(
                f"Skipping Pair: "
                f"{exc_rpt_id} <-> {app_rpt_id}"
            )

            print(
                f"EXC Exists = {exc_rpt_id in df1.columns}"
            )

            print(
                f"APP Exists = {app_rpt_id in df2.columns}"
            )

        pair_no += 1

    # =====================================================
    # Build Merge Columns
    # =====================================================

    exc_cols = ["Key"] + regular_columns

    for exc_rpt_id, app_rpt_id in rpt_pairs:
        if exc_rpt_id not in exc_cols:
            exc_cols.append(exc_rpt_id)

    app_cols = ["Key"] + regular_columns

    for exc_rpt_id, app_rpt_id in rpt_pairs:
        if app_rpt_id not in app_cols:
            app_cols.append(app_rpt_id)

    merged = pd.merge(
        df1[exc_cols],
        df2[app_cols],
        on="Key",
        how="inner",
        suffixes=("_File1", "_File2")
    )

    print("\nMerged Columns:")
    print(list(merged.columns))

    mismatch_rows = []

    # =====================================================
    # Compare Existing Columns C-J
    # =====================================================

    for col in regular_columns:

        file1_col = f"{col}_File1"
        file2_col = f"{col}_File2"

        if (
            file1_col not in merged.columns
            or file2_col not in merged.columns
        ):
            continue

        val1_num = pd.to_numeric(
            merged[file1_col],
            errors="coerce"
        )

        val2_num = pd.to_numeric(
            merged[file2_col],
            errors="coerce"
        )

        diff = (val1_num - val2_num).abs()

        mismatch_mask = (
            (diff >= 0.000001)
            |
            (
                (
                    merged[file1_col]
                    .fillna("")
                    .astype(str)
                    !=
                    merged[file2_col]
                    .fillna("")
                    .astype(str)
                )
                &
                (
                    val1_num.isna()
                    |
                    val2_num.isna()
                )
            )
        )

        mismatches = merged[mismatch_mask]

        for idx, row in mismatches.iterrows():

            if (
                pd.notna(val1_num.loc[idx])
                and pd.notna(val2_num.loc[idx])
            ):
                difference = abs(
                    val1_num.loc[idx]
                    - val2_num.loc[idx]
                )
            else:
                difference = ""

            mismatch_rows.append({
                "Key": row["Key"],
                "Column": col,
                "Value_File1": row[file1_col],
                "Value_File2": row[file2_col],
                "Difference": difference
            })

    # =====================================================
    # Compare RPT Columns
    # =====================================================

    for exc_rpt_id, app_rpt_id in rpt_pairs:

        file1_col = exc_rpt_id
        file2_col = app_rpt_id

        if (
            file1_col not in merged.columns
            or file2_col not in merged.columns
        ):
            print(
                f"Skipping comparison for "
                f"{file1_col} vs {file2_col}"
            )
            continue

        val1_num = pd.to_numeric(
            merged[file1_col],
            errors="coerce"
        )

        val2_num = pd.to_numeric(
            merged[file2_col],
            errors="coerce"
        )

        diff = (val1_num - val2_num).abs()

        mismatch_mask = (
            (diff >= 0.000001)
            |
            (
                (
                    merged[file1_col]
                    .fillna("")
                    .astype(str)
                    !=
                    merged[file2_col]
                    .fillna("")
                    .astype(str)
                )
                &
                (
                    val1_num.isna()
                    |
                    val2_num.isna()
                )
            )
        )

        mismatches = merged[mismatch_mask]

        for idx, row in mismatches.iterrows():

            if (
                pd.notna(val1_num.loc[idx])
                and pd.notna(val2_num.loc[idx])
            ):
                difference = abs(
                    val1_num.loc[idx]
                    - val2_num.loc[idx]
                )
            else:
                difference = ""

            mismatch_rows.append({
                "Key": row["Key"],
                "Column": f"{exc_rpt_id} vs {app_rpt_id}",
                "Value_File1": row[file1_col],
                "Value_File2": row[file2_col],
                "Difference": difference
            })

    mismatch_df = pd.DataFrame(mismatch_rows)

    summary = {
        "File 1 Path": str(Path(file1).resolve()),
        "File 2 Path": str(Path(file2).resolve()),
        "Rows in File 1": len(df1),
        "Rows in File 2": len(df2),
        "Common Keys Compared": len(merged),
        "Compared Columns C-J": len(regular_columns),
        "Compared RPT Pairs": len(rpt_pairs),
        "Value Mismatches": len(mismatch_df)
    }

    html = """
    <html>
    <head>
        <title>Excel Comparison Report</title>
        <style>
            body {
                font-family: Arial;
                padding: 20px;
            }

            table {
                border-collapse: collapse;
                width: 100%;
            }

            th, td {
                border: 1px solid #666;
                padding: 8px;
            }

            th {
                background-color: #eeeeee;
            }
        </style>
    </head>
    <body>
    """

    html += "<h1>Excel Comparison Report</h1>"

    html += "<h2>Summary</h2><table>"

    for k, v in summary.items():
        html += f"<tr><th>{k}</th><td>{v}</td></tr>"

    html += "</table>"

    html += "<h2>Mismatches</h2>"

    if mismatch_df.empty:
        html += "<p>No mismatches found.</p>"
    else:
        html += mismatch_df.to_html(index=False)

    html += "</body></html>"

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ Report generated: {report_path}")

    return len(mismatch_df)
# =====================================================
# CONSOLIDATED REPORT
# =====================================================

def generate_consolidated_report(data, output_path):

    df = pd.DataFrame(data)

    if df.empty:
        df = pd.DataFrame(
            columns=[
                "Test Case",
                "Mismatch Count",
                "Status"
            ]
        )

    df.to_excel(output_path, index=False)

    print(
        f"✅ Consolidated report generated: "
        f"{output_path}"
    )


# =====================================================
# TEST CASE RUNNER
# =====================================================

def run_testcases():

    consolidated_data = []

    print("\n▶ Execution Started")

    for i in range(1, MAX_T_FILES + 1):

        print("\n" + "=" * 60)
        print(f"Running Test Case T{i}")

        try:

            exc_file = (
                OUTPUT_RPT_EXC_FOLDER
                / f"T{i}_Exc_Output_RPT.xlsm"
            )
            
            app_file = (
                OUTPUT_RPT_APP_FOLDER
                / f"T{i}_App_Output_RPT.xlsx"
            )

            if not exc_file.exists():
                print(f"⏭ Skipping T{i} (EXC workbook missing)")
                continue

            if not app_file.exists():
                print(f"⏭ Skipping T{i} (APP workbook missing)")
                continue

            report_file = (
                REPORT_OUTPUT_FOLDER
                / f"T{i}_comparison_report.html"
            )

            mismatch_count = compare_excels(
                exc_file,
                app_file,
                report_file
            )

            consolidated_data.append({
                "Test Case": f"T{i}",
                "Mismatch Count": mismatch_count,
                "Status": (
                    "PASS"
                    if mismatch_count == 0
                    else "FAIL"
                )
            })

        except Exception as e:

            print(f"❌ T{i} failed: {e}")

            consolidated_data.append({
                "Test Case": f"T{i}",
                "Mismatch Count": "ERROR",
                "Status": "ERROR"
            })

    consolidated_file = (
        REPORT_OUTPUT_FOLDER
        / "Consolidated_Report.xlsx"
    )

    generate_consolidated_report(
        consolidated_data,
        consolidated_file
    )

    print("\n✅ Execution Completed")


# =====================================================
# MAIN
# =====================================================

if __name__ == "__main__":
    run_testcases()