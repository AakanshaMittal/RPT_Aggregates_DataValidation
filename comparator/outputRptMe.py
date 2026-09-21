import os
import re
import json
import pandas as pd
from pathlib import Path

# =====================================================
# CONFIGURATION
# =====================================================

APP_FOLDER = r"data/outputRpt/File1"

EXCEL_FOLDER = r"data/excel_files"

OUTPUT_FOLDER = r"data/outputRpt/Aggregates"

MAPPING_FILE = r"config/mapping.json"

APP_SHEET_NAME = "template"

EXCEL_SHEET_NAME = "Output| RPT"

Path(
    OUTPUT_FOLDER
).mkdir(
    parents=True,
    exist_ok=True
)

# =====================================================
# LOAD MAPPING
# =====================================================

def load_mapping():

    with open(MAPPING_FILE) as f:
        return json.load(f)


def reverse_mapping(mapping):

    return {
        v: k
        for k, v in mapping.items()
    }


# =====================================================
# BUILD RPT PAIRS FROM CONFIG
# =====================================================
# =====================================================
# LOAD SEQUENCE MAPPING
# =====================================================

SEQUENCE_MAPPING_FILE = (
    r"config/input_for_model_sequence_mapping.json"
)


def load_sequence_mapping():

    with open(
        SEQUENCE_MAPPING_FILE,
        encoding="utf-8"
    ) as f:

        return json.load(f)


# =====================================================
# BUILD RPT PAIRS
# =====================================================

def build_rpt_pairs(
        df1,
        df2,
        tc
):

    sequence_mapping = (
        load_sequence_mapping()
    )

    tc_key = f"TC{tc}"

    rpt_pairs = []

    if tc_key not in sequence_mapping:

        print(
            f"No RPT mapping found for "
            f"{tc_key}"
        )

        return rpt_pairs

    tc_mapping = (
        sequence_mapping[tc_key]
    )

    excel_rpts = [

        col

        for col in df1.columns

        if str(col).startswith(
            "RPT"
        )
    ]

    app_rpts = [

        col

        for col in df2.columns

        if str(col).startswith(
            "RPT"
        )
    ]

    print(
        f"\nExcel RPT Count = "
        f"{len(excel_rpts)}"
    )

    print(
        f"APP RPT Count = "
        f"{len(app_rpts)}"
    )

    for excel_seq, app_seq in tc_mapping.items():

        try:

            excel_index = (
                int(excel_seq) - 1
            )

            app_index = (
                int(app_seq) - 1
            )

            if (
                excel_index >= len(excel_rpts)
                or
                app_index >= len(app_rpts)
            ):
                continue

            excel_rpt = (
                excel_rpts[
                    excel_index
                ]
            )

            app_rpt = (
                app_rpts[
                    app_index
                ]
            )

            rpt_pairs.append(

                (
                    excel_rpt,
                    app_rpt
                )
            )

            print(

                f"EXCEL RPT{excel_seq} "
                f"({excel_rpt}) "

                f"<-> "

                f"APP RPT{app_seq} "
                f"({app_rpt})"
            )

        except Exception as e:

            print(
                f"RPT Pair Error: "
                f"{e}"
            )

    return rpt_pairs


# =====================================================
# FILENAME HELPERS
# =====================================================

def extract_agg_id(filename):

    match = re.search(
        r"(AGG\d+[A-Z]{2})",
        filename.upper()
    )

    return (
        match.group(1)
        if match
        else None
    )


def extract_tc(filename):

    match = re.search(
        r"TC(\d+[A-Z]?)",
        filename.upper()
    )

    return (
        match.group(1)
        if match
        else None
    )

# =====================================================
# FILE PAIRING
# =====================================================

def build_file_maps():

    mapping = load_mapping()

    agg_to_tc = reverse_mapping(
        mapping
    )

    app_map = {}

    excel_map = {}

    app_path = Path(APP_FOLDER)

    excel_path = Path(EXCEL_FOLDER)

    # ------------------------------------
    # APP FILES
    # ------------------------------------

    for f in app_path.glob("*.xlsx"):

        agg_id = extract_agg_id(
            f.name
        )

        if (
            agg_id
            and
            agg_id in agg_to_tc
        ):

            tc = (
                agg_to_tc[agg_id]
                .replace("TC", "")
            )

            app_map[tc] = f

    # ------------------------------------
    # EXCEL FILES
    # ------------------------------------

    for f in excel_path.glob("*.xlsb"):

        tc = extract_tc(
            f.name
        )

        if tc:

            excel_map[tc] = f

    return (
        app_map,
        excel_map
    )

# =====================================================
# LOAD EXCEL
# =====================================================

def load_excel(path, file_type):

    engine = (
        "pyxlsb"
        if str(path).lower().endswith(".xlsb")
        else "openpyxl"
    )

    excel_file = pd.ExcelFile(
        path,
        engine=engine
    )

    print(f"\nWorkbook: {path.name}")
    print("Available Sheets:")

    for sheet in excel_file.sheet_names:
        print(repr(sheet))

    if file_type == "APP":
        sheet_name = APP_SHEET_NAME
    else:
        sheet_name = EXCEL_SHEET_NAME

    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=None,
        engine=engine
    )

    print(
        f"\nWorkbook: {path.name}"
    )

    print(
        "Available Sheets:"
    )

    for sheet in excel_file.sheet_names:

        print(repr(sheet))

    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=None,
        engine=engine
    )

    df = (
        df.iloc[5:]
        .reset_index(drop=True)
    )

    result = pd.DataFrame()

    result["Key"] = df.iloc[:, 0]

    # =====================================================
    # C-J columns
    # =====================================================

    for col_idx in range(
        2,
        min(
            10,
            len(df.columns)
        )
    ):

        col_data = df.iloc[
            :,
            col_idx
        ]

        if col_data.notna().any():

            col_name = chr(
                65 + col_idx
            )

            result[col_name] = (
                col_data
            )

    # =====================================================
    # RPT columns
    # =====================================================

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

        rpt_id = (
            str(rpt_id)
            .strip()
        )

        if rpt_id == "":
            continue

        result[rpt_id] = df.iloc[
            :,
            col_idx
        ]

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

    print(
        f"\nLoaded Columns from "
        f"{path.name}:"
    )

    print(
        list(result.columns)
    )

    return result

# =====================================================
# COMPARE EXCELS
# =====================================================

# =====================================================
# COMPARE EXCELS
# =====================================================

def compare_excels(
        file1,
        file2,
        report_path,
        tc
):

    df1 = load_excel(
        file1,
        "EXC"
    )

    df2 = load_excel(
        file2,
        "APP"
    )

    # =====================================================
    # REGULAR C-J COLUMNS
    # =====================================================

    regular_columns = []

    for col in df2.columns:

        if col == "Key":
            continue

        if col in df1.columns:

            regular_columns.append(
                col
            )

    regular_columns = [

        col

        for col in regular_columns

        if not str(col).startswith(
            "RPT"
        )
    ]

    # =====================================================
    # BUILD RPT PAIRS
    # =====================================================

    rpt_pairs = build_rpt_pairs(
        df1,
        df2,
        tc
    )

    print(
        f"\nTotal RPT Pairs = "
        f"{len(rpt_pairs)}"
    )

    # =====================================================
    # BUILD MERGE COLUMNS
    # =====================================================

    exc_cols = (
        ["Key"]
        + regular_columns
    )

    for excel_rpt, app_rpt in rpt_pairs:

        if excel_rpt not in exc_cols:

            exc_cols.append(
                excel_rpt
            )

    app_cols = (
        ["Key"]
        + regular_columns
    )

    for excel_rpt, app_rpt in rpt_pairs:

        if app_rpt not in app_cols:

            app_cols.append(
                app_rpt
            )

    merged = pd.merge(

        df1[exc_cols],

        df2[app_cols],

        on="Key",

        how="inner",

        suffixes=(
            "_File1",
            "_File2"
        )
    )

    print(
        "\nMerged Columns:"
    )

    print(
        list(
            merged.columns
        )
    )

    mismatch_rows = []

    # =====================================================
    # COMPARE REGULAR C-J COLUMNS
    # =====================================================

    for col in regular_columns:

        file1_col = (
            f"{col}_File1"
        )

        file2_col = (
            f"{col}_File2"
        )

        if (
            file1_col not in merged.columns
            or
            file2_col not in merged.columns
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

        diff = (
            val1_num
            -
            val2_num
        ).abs()

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

        mismatches = merged[
            mismatch_mask
        ]

        for idx, row in mismatches.iterrows():

            if (
                pd.notna(
                    val1_num.loc[idx]
                )
                and
                pd.notna(
                    val2_num.loc[idx]
                )
            ):

                difference = abs(
                    val1_num.loc[idx]
                    -
                    val2_num.loc[idx]
                )

            else:

                difference = ""

            mismatch_rows.append({

                "Key":
                    row["Key"],

                "Column":
                    col,

                "Value_File1":
                    row[file1_col],

                "Value_File2":
                    row[file2_col],

                "Difference":
                    difference
            })

    # =====================================================
    # COMPARE RPT PAIRS
    # =====================================================

    for excel_rpt, app_rpt in rpt_pairs:

        if (
            excel_rpt not in merged.columns
            or
            app_rpt not in merged.columns
        ):
            continue

        val1_num = pd.to_numeric(
            merged[excel_rpt],
            errors="coerce"
        )

        val2_num = pd.to_numeric(
            merged[app_rpt],
            errors="coerce"
        )

        diff = (
            val1_num
            -
            val2_num
        ).abs()

        mismatch_mask = (

            (diff >= 0.000001)

            |

            (

                (
                    merged[excel_rpt]
                    .fillna("")
                    .astype(str)

                    !=

                    merged[app_rpt]
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

        mismatches = merged[
            mismatch_mask
        ]

        for idx, row in mismatches.iterrows():

            if (
                pd.notna(
                    val1_num.loc[idx]
                )
                and
                pd.notna(
                    val2_num.loc[idx]
                )
            ):

                difference = abs(
                    val1_num.loc[idx]
                    -
                    val2_num.loc[idx]
                )

            else:

                difference = ""

            mismatch_rows.append({

                "Key":
                    row["Key"],

                "Column":
                    f"{excel_rpt} vs {app_rpt}",

                "Value_File1":
                    row[excel_rpt],

                "Value_File2":
                    row[app_rpt],

                "Difference":
                    difference
            })

    mismatch_df = pd.DataFrame(
        mismatch_rows
    )

    summary = {

        "File 1 Path":
            str(
                Path(file1)
                .resolve()
            ),

        "File 2 Path":
            str(
                Path(file2)
                .resolve()
            ),

        "Rows in File 1":
            len(df1),

        "Rows in File 2":
            len(df2),

        "Common Keys Compared":
            len(merged),

        "Compared Columns C-J":
            len(
                regular_columns
            ),

        "Compared RPT Pairs":
            len(
                rpt_pairs
            ),

        "Value Mismatches":
            len(
                mismatch_df
            )
    }

    html = """
    <html>
    <head>
        <title>Excel Comparison Report</title>
        <style>

            body{
                font-family:Arial;
                padding:20px;
            }

            table{
                border-collapse:collapse;
                width:100%;
            }

            th,td{
                border:1px solid #666;
                padding:8px;
            }

            th{
                background:#eeeeee;
            }

        </style>
    </head>
    <body>
    """

    html += "<h1>Excel Comparison Report</h1>"

    html += "<h2>Summary</h2><table>"

    for k, v in summary.items():

        html += (
            f"<tr>"
            f"<th>{k}</th>"
            f"<td>{v}</td>"
            f"</tr>"
        )

    html += "</table>"

    html += "<h2>Mismatches</h2>"

    if mismatch_df.empty:

        html += (
            "<p>"
            "No mismatches found."
            "</p>"
        )

    else:

        html += mismatch_df.to_html(
            index=False
        )

    html += "</body></html>"

    with open(
        report_path,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(html)

    print(
        f"✅ Report generated: "
        f"{report_path}"
    )

    return len(
        mismatch_df
    )

# =====================================================
# CONSOLIDATED REPORT
# =====================================================

def generate_consolidated_report(
        summary_rows,
        output_file
):

    df = pd.DataFrame(
        summary_rows
    )

    with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

    print(
        f"✅ Consolidated report generated: "
        f"{output_file}"
    )


# =====================================================
# PROCESS SINGLE TC
# =====================================================

def process_tc(
        tc,
        app_file,
        excel_file
):

    print(
        f"\nProcessing TC{tc}"
    )

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    output_html = os.path.join(
        OUTPUT_FOLDER,
        f"TC{tc}_OutputRpt.html"
    )

    mismatch_count = compare_excels(
        excel_file,
        app_file,
        output_html,
        tc
    )

    return {

        "TC":
            f"TC{tc}",

        "Checks":
            mismatch_count,

        "Failures":
            mismatch_count,

        "Status":
            (
                "PASS"
                if mismatch_count == 0
                else "FAIL"
            )
    }


# =====================================================
# MAIN
# =====================================================

def main():

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    app_map, excel_map = (
        build_file_maps()
    )

    print(
        "\nAPP MAP:"
    )

    print(app_map)

    print(
        "\nEXCEL MAP:"
    )

    print(excel_map)

    common_tcs = sorted(

        set(app_map.keys())

        &

        set(excel_map.keys())
    )

    print(
        f"\nRunning "
        f"{len(common_tcs)} "
        f"scenario(s)\n"
    )

    summary_rows = []

    for tc in common_tcs:

        try:

            summary = process_tc(

                tc,

                app_map[tc],

                excel_map[tc]
            )

            summary_rows.append(
                summary
            )

        except Exception as e:

            print(
                f"❌ TC{tc} failed: {e}"
            )

            summary_rows.append({

                "TC":
                    f"TC{tc}",

                "Checks":
                    "ERROR",

                "Failures":
                    "ERROR",

                "Status":
                    "ERROR"
            })

    consolidated_file = os.path.join(

        OUTPUT_FOLDER,

        "Consolidated_OutputRpt.xlsx"
    )

    generate_consolidated_report(

        summary_rows,

        consolidated_file
    )

    print(
        "\nCompleted."
    )


# =====================================================
# ENTRY
# =====================================================

if __name__ == "__main__":
    main()