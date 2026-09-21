import json
import pandas as pd
from pathlib import Path
 
# =====================================================
# CONFIG
# =====================================================
 
APP_FOLDER = Path("data/outputRpt/File1")
 
EXCEL_FOLDER = Path("data/excel_files")
 
OUTPUT_FOLDER = Path("data/outputRpt/Aggregates")
 
MAPPING_FILE = Path("config/mapping.json")
 
APP_SHEET_NAME = "template"
 
EXCEL_SHEET_NAME = "Output| RPT"
 
OUTPUT_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)
 
# =====================================================
# LOAD MAPPING
# =====================================================
 
def load_mapping():
 
    with open(
        MAPPING_FILE,
        "r",
        encoding="utf-8"
    ) as f:
 
        return json.load(f)
 
# =====================================================
# LOAD EXCEL
# =====================================================
 
def load_excel(
        file_path,
        file_type
):
 
    engine = (
        "pyxlsb"
        if str(file_path).lower().endswith(".xlsb")
        else "openpyxl"
    )
 
    sheet_name = (
        APP_SHEET_NAME
        if file_type == "APP"
        else EXCEL_SHEET_NAME
    )
 
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        engine=engine
    )
 
    # Skip first 5 rows
 
    df = (
        df.iloc[5:]
        .reset_index(drop=True)
    )
 
    result = pd.DataFrame()
 
    result["Key"] = df.iloc[:, 0]
 
    # ==========================================
    # Columns before K
    # C to J
    # ==========================================
 
    regular_columns = []
 
    for col_idx in range(
        2,
        min(10, len(df.columns))
    ):
 
        col_name = chr(
            65 + col_idx
        )
 
        result[col_name] = df.iloc[
            :,
            col_idx
        ]
 
        regular_columns.append(
            col_name
        )
 
    # ==========================================
    # K onwards
    # ==========================================
 
    rpt_columns = []
 
    rpt_header_row = 2
 
    for col_idx in range(
        10,
        len(df.columns)
    ):
 
        rpt_id = df.iloc[
            rpt_header_row,
            col_idx
        ]
 
        if pd.isna(rpt_id):
            continue
 
        rpt_id = str(rpt_id).strip()
 
        if rpt_id == "":
            continue
 
        result[rpt_id] = df.iloc[
            :,
            col_idx
        ]
 
        rpt_columns.append(
            rpt_id
        )
 
    result = result[
        result["Key"].notna()
    ]
 
    result["Key"] = (
        result["Key"]
        .astype(str)
        .str.strip()
    )
 
    result = result[
        result["Key"] != ""
    ]
 
    return (
        result,
        regular_columns,
        rpt_columns
    )
 
# =====================================================
# COMPARE TWO SERIES
# =====================================================
 
def compare_series(
        key_series,
        left_series,
        right_series,
        column_name,
        mismatch_rows
):
 
    left_num = pd.to_numeric(
        left_series,
        errors="coerce"
    )
 
    right_num = pd.to_numeric(
        right_series,
        errors="coerce"
    )
 
    diff = (
        left_num
        -
        right_num
    ).abs()
 
    mismatch_mask = (
 
        (diff >= 0.000001)
 
        |
 
        (
            left_series
            .fillna("")
            .astype(str)
 
            !=
 
            right_series
            .fillna("")
            .astype(str)
        )
    )
 
    mismatch_df = pd.DataFrame({
 
        "Key":
            key_series,
 
        "Column":
            column_name,
 
        "Value_File1":
            left_series,
 
        "Value_File2":
            right_series
 
    })
 
    mismatch_df = mismatch_df[
        mismatch_mask
    ]
 
    mismatch_rows.extend(
        mismatch_df.to_dict(
            "records"
        )
    )
 
# =====================================================
# COMPARE EXCELS
# =====================================================
 
def compare_excels(
        excel_file,
        app_file,
        scenario_mapping,
        output_html
):
 
    (
        df_exc,
        regular_exc,
        rpt_exc
    ) = load_excel(
        excel_file,
        "EXC"
    )
 
    (
        df_app,
        regular_app,
        rpt_app
    ) = load_excel(
        app_file,
        "APP"
    )
 
    # ==========================================
    # Build Merge Column Lists
    # ==========================================
 
    exc_cols = (
        ["Key"]
        +
        regular_exc
    )
 
    app_cols = (
        ["Key"]
        +
        regular_app
    )
 
    rpt_pairs = []
 
    for exc_pos, app_pos in scenario_mapping.items():
 
        exc_idx = (
            int(exc_pos)
            - 1
        )
 
        app_idx = (
            int(app_pos)
            - 1
        )
 
        if (
            exc_idx >= len(rpt_exc)
            or
            app_idx >= len(rpt_app)
        ):
            continue
 
        exc_rpt = rpt_exc[
            exc_idx
        ]
 
        app_rpt = rpt_app[
            app_idx
        ]
 
        rpt_pairs.append(
            (
                exc_rpt,
                app_rpt
            )
        )
 
        exc_cols.append(
            exc_rpt
        )
 
        app_cols.append(
            app_rpt
        )
 
    merged = pd.merge(
 
        df_exc[
            exc_cols
        ],
 
        df_app[
            app_cols
        ],
 
        on="Key",
 
        how="inner",
 
        suffixes=(
            "_File1",
            "_File2"
        )
    )
 
    mismatch_rows = []
 
    # ==========================================
    # Compare C-J
    # ==========================================
 
    for col in regular_exc:
 
        compare_series(
 
            merged["Key"],
 
            merged[
                f"{col}_File1"
            ],
 
            merged[
                f"{col}_File2"
            ],
 
            col,
 
            mismatch_rows
        )
 
    # ==========================================
    # Compare mapped RPTs
    # ==========================================
 
    for exc_rpt, app_rpt in rpt_pairs:
 
        compare_series(
 
            merged["Key"],
 
            merged[
                exc_rpt
            ],
 
            merged[
                app_rpt
            ],
 
            f"{exc_rpt} vs {app_rpt}",
 
            mismatch_rows
        )
 
    mismatch_df = pd.DataFrame(
        mismatch_rows
    )
 
    html = """
    <html>
    <head>
    <title>Comparison Report</title>
    </head>
    <body>
    """
 
    html += "<h1>Comparison Report</h1>"
 
    html += (
        f"<p>Total Keys Compared: "
        f"{len(merged)}</p>"
    )
 
    html += (
        f"<p>Total Mismatches: "
        f"{len(mismatch_df)}</p>"
    )
 
    if mismatch_df.empty:
 
        html += (
            "<p>No mismatches found</p>"
        )
 
    else:
 
        html += mismatch_df.to_html(
            index=False
        )
 
    html += """
    </body>
    </html>
    """
 
    with open(
        output_html,
        "w",
        encoding="utf-8"
    ) as f:
 
        f.write(html)
 
    return len(
        mismatch_df
    )
 
# =====================================================
# MAIN
# =====================================================
 
def main():
 
    mappings = load_mapping()
 
    summary = []
 
    for tc, tc_mapping in mappings.items():
 
        print(
            f"\nRunning {tc}"
        )
 
        excel_file = next(
            EXCEL_FOLDER.glob(
                f"*{tc}*.xlsb"
            ),
            None
        )
 
        app_file = next(
            APP_FOLDER.glob(
                f"*{tc}*.xlsx"
            ),
            None
        )
 
        if (
            excel_file is None
            or
            app_file is None
        ):
 
            print(
                f"Skipping {tc}"
            )
 
            continue
 
        report_file = (
            OUTPUT_FOLDER
            /
            f"{tc}_Report.html"
        )
 
        mismatch_count = compare_excels(
 
            excel_file,
 
            app_file,
 
            tc_mapping,
 
            report_file
        )
 
        summary.append({
 
            "Scenario":
                tc,
 
            "Mismatch Count":
                mismatch_count,
 
            "Status":
                (
                    "PASS"
                    if mismatch_count == 0
                    else "FAIL"
                )
        })
 
    pd.DataFrame(
        summary
    ).to_excel(
 
        OUTPUT_FOLDER
        /
        "Consolidated_Report.xlsx",
 
        index=False
    )
 
    print(
        "\nExecution Completed"
    )
 
if __name__ == "__main__":
    main()
 