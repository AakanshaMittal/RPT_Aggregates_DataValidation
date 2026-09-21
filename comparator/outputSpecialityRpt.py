import pandas as pd
import numpy as np
import os
import json
import re
from pathlib import Path
from collections import defaultdict
from html import escape
 
APP_FOLDER = r"data/outputSpecialityRpt/File1"
EXCEL_FOLDER = r"data/excel_files"
OUTPUT_FOLDER = r"data/outputSpecialityRpt/Aggregate"
mapping_file = r"config/mapping.json"
TARGET_SHEET = "Output| Specialty RPT"
 
 
# APP FILE
APP_HEADER_ROW_1 = 0
APP_HEADER_ROW_2 = 1

# EXCEL FILE
EXCEL_HEADER_ROW_1 = 12
EXCEL_HEADER_ROW_2 = 13


NDC_COL_NAME = "NDC"

TOLERANCE = 0.000001
 
def load_mapping():
    with open(mapping_file) as f:
        return json.load(f)
 
def reverse_mapping(mapping):
    return {v: k for k, v in mapping.items()}
 
def extract_rpt_id(filename):
    match = re.search(r"(AGG\d+[A-Z]{2})", filename.upper())
    return match.group(1) if match else None
 
 
def extract_tc(filename):
    match = re.search(r"TC(\d+)", filename.upper())
    return match.group(1) if match else None 
 
'''def load_and_prepare(file_path):

    #df_raw = pd.read_excel(file_path, header=None)

    
    print(f"Reading sheet '{TARGET_SHEET}' "f"from {os.path.basename(file_path)}")
    df_raw = pd.read_excel(file_path, sheet_name=TARGET_SHEET, header=None, engine="pyxlsb" )

    # Ignore A13
    df_raw.iat[HEADER_ROW_1, 0] = np.nan

    # Ignore D13
    df_raw.iat[HEADER_ROW_1, 3] = np.nan

    header1 = df_raw.iloc[HEADER_ROW_1].ffill()

    header1.iloc[3] = ""

    header2 = df_raw.iloc[HEADER_ROW_2]

    combined_headers = []

    for h1, h2 in zip(header1, header2):

        h1 = str(h1).strip() if pd.notna(h1) else ""

        h2 = str(h2).strip() if pd.notna(h2) else ""

        if h1 and h2:

            combined_headers.append(f"{h1}_{h2}")

        elif h2:

            combined_headers.append(h2)

        else:

            combined_headers.append(h1)

    df = df_raw.iloc[HEADER_ROW_2 + 1:].copy()

    df.columns = combined_headers

    df = df.dropna(how="all")

    df.columns = [normalize_text(c) for c in df.columns]

    ndc_col = find_ndc_column(df.columns)

    if not ndc_col:

        raise Exception("NDC column not found!")

    print("NDC Column:", ndc_col)
    print(df.columns.tolist())

    # Handle duplicate NDC columns by taking the first occurrence
    ndc_data = df.loc[:, ndc_col]

    if isinstance(ndc_data, pd.DataFrame):
        ndc_series = ndc_data.iloc[:, 0]
    else:
        ndc_series = ndc_data

    ndc_series = ndc_series.astype(str).str.strip()


    df = df[ndc_series != "nan"]

    df.index = ndc_series[ndc_series != "nan"]

    print("Duplicate columns:")
    print(df.columns[df.columns.duplicated()].tolist())

    # Remove duplicate column names
    df = df.loc[:, ~df.columns.duplicated()]

    return df '''

def load_app_file(file_path):

    print(
        f"Reading App file: "
        f"{os.path.basename(file_path)}"
    )

    df_raw = pd.read_excel(
        file_path,
        header=None,
        engine="openpyxl"
    )

    return prepare_dataframe(df_raw, APP_HEADER_ROW_1, APP_HEADER_ROW_2)

def load_excel_workbook(file_path):

    print(
        f"Reading workbook sheet '{TARGET_SHEET}' "
        f"from {os.path.basename(file_path)}"
    )

    df_raw = pd.read_excel(
        file_path,
        sheet_name=TARGET_SHEET,
        header=None,
        engine="pyxlsb"
    )

    return prepare_dataframe(df_raw, EXCEL_HEADER_ROW_1, EXCEL_HEADER_ROW_2)

def prepare_dataframe(df_raw, header_row_1, header_row_2):

    # Ignore A13
    df_raw.iat[header_row_1, 0] = np.nan

    # Ignore D13
    df_raw.iat[header_row_1, 3] = np.nan

    header1 = df_raw.iloc[header_row_1].ffill()

    header1.iloc[3] = ""

    header2 = df_raw.iloc[header_row_2]

    combined_headers = []

    for h1, h2 in zip(header1, header2):

        h1 = str(h1).strip() if pd.notna(h1) else ""
        h2 = str(h2).strip() if pd.notna(h2) else ""

        if h1 and h2:
            combined_headers.append(f"{h1}_{h2}")

        elif h2:
            combined_headers.append(h2)

        else:
            combined_headers.append(h1)

    df = df_raw.iloc[header_row_2 + 1:].copy()

    df.columns = combined_headers

    df = df.dropna(how="all")

    df.columns = [normalize_text(c) for c in df.columns]

    ndc_col = find_ndc_column(df.columns)

    if not ndc_col:
        raise Exception("NDC column not found!")

    print("NDC Column:", ndc_col)

    ndc_data = df.loc[:, ndc_col]

    if isinstance(ndc_data, pd.DataFrame):
        ndc_series = ndc_data.iloc[:, 0]
    else:
        ndc_series = ndc_data

    ndc_series = ndc_series.astype(str).str.strip()
    #duplicate_ndcs = ndc_series[ndc_series.duplicated(keep=False)]
    duplicate_ndcs = ndc_series[ndc_series.duplicated(keep=False)]
    duplicate_ndc_list = sorted(duplicate_ndcs.unique())
    if not duplicate_ndcs.empty:
        print("Duplicate NDCs found:")
        print(sorted(duplicate_ndcs.unique()))

    # Remove ALL occurrences of duplicate NDCs from comparison
    df = df[~ndc_series.isin(duplicate_ndc_list)]
    ndc_series = ndc_series[~ndc_series.isin(duplicate_ndc_list)]
    # Set index

    '''df.index = ndc_series

    df = df[ndc_series != "nan"]

    df.index = ndc_series[ndc_series != "nan"]'''

    # Remove blank NDCs
    valid_mask = ndc_series != "nan"
    df = df[valid_mask]
    ndc_series = ndc_series[valid_mask]
    # Find duplicate NDCs
    duplicate_ndcs = ndc_series[ndc_series.duplicated(keep=False)]
    duplicate_ndc_list = sorted(duplicate_ndcs.unique())
    if duplicate_ndc_list:
        print("Duplicate NDCs found:")
        print(duplicate_ndc_list)
        # Remove duplicate NDC rows
    df = df[~ndc_series.isin(duplicate_ndc_list)]
    ndc_series = ndc_series[~ndc_series.isin(duplicate_ndc_list)]
        # Set index
    df.index = ndc_series

    print("Duplicate columns:")
    print(df.columns[df.columns.duplicated()].tolist())

    df = df.loc[:, ~df.columns.duplicated()]

    return df, duplicate_ndc_list
 
def normalize_text(text):

    if pd.isna(text):

        return ""

    return str(text).strip().lower().replace(" ", "").replace("-", "")
 
 
def find_ndc_column(columns):

    for col in columns:

        if "ndc" in col:

            return col

    return None
 
 
def try_float(val):

    if pd.isna(val):

        return None
 
    val = str(val).strip().replace("$", "").replace(",", "")
 
    if val in ("", "-"):

        return None
 
    try:

        return float(val)

    except:

        return None
 
 
def compare_data(df1, df2):

    mismatches = []

    group_summary = defaultdict(int)

    group_ndc_map = defaultdict(set)
 
    all_cols = sorted(set(df1.columns) & set(df2.columns))

    all_rows = sorted(set(df1.index) & set(df2.index))
 
    for col in all_cols:

        group = col.split("_")[0] if "_" in col else "Other"
 
        for row in all_rows:

            v1 = df1.at[row, col]

            v2 = df2.at[row, col]
 
            if pd.isna(v1) and pd.isna(v2):

                continue
 
            f1, f2 = try_float(v1), try_float(v2)
 
            if f1 is not None and f2 is not None:

                if abs(f1 - f2) > TOLERANCE:

                    mismatches.append((row, col, v1, v2))

                    group_summary[group] += 1

                    group_ndc_map[group].add(row)

            else:

                if str(v1).strip() != str(v2).strip():

                    mismatches.append((row, col, v1, v2))

                    group_summary[group] += 1

                    group_ndc_map[group].add(row)
 
    return mismatches, group_summary, group_ndc_map
 
 
def find_missing(df1, df2):

    missing_cols = sorted(set(df1.columns) - set(df2.columns))

    extra_cols = sorted(set(df2.columns) - set(df1.columns))

    missing_rows = sorted(set(df1.index) - set(df2.index))

    extra_rows = sorted(set(df2.index) - set(df1.index))
 
    return missing_cols, extra_cols, missing_rows, extra_rows


def generate_consolidated_report(summary_rows, output_file):

    df = pd.DataFrame(summary_rows)

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
        f"\n Consolidated report generated: "
        f"{output_file}"
    )
 
 
def generate_html(file1, file2, df1, df2, mismatches, group_summary,

                  group_ndc_map, missing_cols, extra_cols, missing_rows, extra_rows, dup_ndcs_file1, dup_ndcs_file2):
 
    file1_name = escape(os.path.basename(file1))

    file2_name = escape(os.path.basename(file2))
 
    html = []
 
    html.append("""
<html>
<head>
<style>

table { border-collapse: collapse; }

th, td { padding: 8px 12px; border: 1px solid black; text-align: left; }

th { background-color: #f2f2f2; }
</style>
</head>
<body>

""")
 
    html.append(f"<h1>Comparison Report</h1>")

    html.append(f"<h2>{file1_name} vs {file2_name}</h2>")

    html.append("<h3>Summary</h3>")

    html.append(f"<p>Rows: {len(df1)} vs {len(df2)}</p>")

    html.append(f"<p>Columns: {len(df1.columns)} vs {len(df2.columns)}</p>")

    html.append(f"<p>Total mismatches: {len(mismatches)}</p>")
 
    html.append("<h3>Mismatch by Group</h3>")

    html.append("<table><tr><th>Group</th><th>Mismatch Count</th><th>NDCs</th></tr>")
 
    for g in sorted(group_summary):

        ndcs = ", ".join(sorted(map(str, group_ndc_map[g])))

        html.append(f"<tr><td>{escape(g)}</td><td>{group_summary[g]}</td><td>{escape(ndcs)}</td></tr>")
 
    html.append("</table>")
 
    html.append("<h3>Detailed Mismatches</h3>")

    html.append("<table>")

    html.append(f"<tr><th>NDC</th><th>Column</th><th>{file1_name}</th><th>{file2_name}</th></tr>")
 
    for row, col, v1, v2 in mismatches:

        html.append(f"<tr><td>{row}</td><td>{escape(col)}</td><td>{v1}</td><td>{v2}</td></tr>")
 
    html.append("</table>")
 
    html.append("<h3>Missing / Extra Columns</h3>")

    html.append(f"<p>Missing Columns: {missing_cols}</p>")

    html.append(f"<p>Extra Columns: {extra_cols}</p>")
 
    html.append("<h3>Missing / Extra Rows (by NDC)</h3>")

    html.append(f"<p>Missing Rows: {missing_rows}</p>")

    html.append(f"<p>Extra Rows: {extra_rows}</p>")

    html.append("<h3>Duplicate NDCs Skipped</h3>")
    html.append(f"<p><b>{file1_name}</b>: "f"{', '.join(map(str, dup_ndcs_file1)) if dup_ndcs_file1 else 'None'}</p>")
    html.append(f"<p><b>{file2_name}</b>: "f"{', '.join(map(str, dup_ndcs_file2)) if dup_ndcs_file2 else 'None'}</p>")
 
    html.append("</body></html>")
 
    return "\n".join(html)
 
 
def main():
 
    app_path = Path(APP_FOLDER)

    excel_path = Path(EXCEL_FOLDER)

    out_path = Path(OUTPUT_FOLDER)

    out_path.mkdir(parents=True, exist_ok=True)
 
    # ===== LOAD MAPPING (same as RA Output) =====

    mapping = load_mapping()

    rpt_to_tc = reverse_mapping(mapping)
 
    app_map = {}

    excel_map = {}
 
    # ===== FILE1 (RPT BASED instead of TC*_App*) =====

    for f in app_path.glob("*.xlsx"):
 
        rpt_id = extract_rpt_id(f.name)
 
        if rpt_id and rpt_id in rpt_to_tc:
 
            tc = rpt_to_tc[rpt_id].replace("TC", "")

            app_map[tc] = f
 
    # ===== FILE2 (KEEP SAME AS WORKING LOGIC BUT FIX PATTERN) =====

    for f in excel_path.glob("*.xlsb"):
 
        tc = extract_tc(f.name)
 
        if tc:

            excel_map[tc] = f
 
    # ===== DEBUG (VERY IMPORTANT) =====

    print("\nChecking File1 mapping:")

    print(app_map)
 
    print("\nChecking File2 mapping:")

    print(excel_map)
 
    print("\nFile1 TC keys:", app_map.keys())

    print("File2 TC keys:", excel_map.keys())
 
    common_tcs = sorted(set(app_map.keys()) & set(excel_map.keys()), key=int)
    summary_rows = []
 
    print(f"\nRunning {len(common_tcs)} scenarios...\n")
 
    for tc in common_tcs:
 
        APP_FILE = app_map[tc]

        EXCEL_FILE = excel_map[tc]
 
        OUTPUT_HTML = out_path / f"TC{tc}_OSRPT_6D_15Sep_Aggregate.html"
 
        print(f"Processing TC{tc}...")
 
        #df1, dup_ndcs_file1 = load_app_file(APP_FILE)
        #df2, dup_ndcs_file2 = load_excel_workbook(EXCEL_FILE)

        try:
            print(f"Reading File1 (App): {APP_FILE}")
            df1, dup_ndcs_file1 = load_app_file(APP_FILE)
        except Exception as e:
            print(f"ERROR in File1 (App): {APP_FILE}")
            print(f"Reason: {e}")
            raise
        try:
            print(f"Reading File2 (Excel): {EXCEL_FILE}")
            df2, dup_ndcs_file2 = load_excel_workbook(EXCEL_FILE)
        except Exception as e:
            print(f"ERROR in File2 (Excel): {EXCEL_FILE}")
            print(f"Reason: {e}")
            raise
        mismatches, group_summary, group_ndc_map = compare_data(df1, df2)
 
        missing_cols, extra_cols, missing_rows, extra_rows = find_missing(df1, df2)

        summary_rows.append( {
        "TC": f"TC{tc}",
        "Mismatch_Count": len(mismatches),
        "Missing_Columns": len(missing_cols),
        "Extra_Columns": len(extra_cols),
        "Missing_Rows": len(missing_rows),
        "Extra_Rows": len(extra_rows),
        "Duplicate_NDC_APP": len(dup_ndcs_file1),
        "Duplicate_NDC_Excel": len(dup_ndcs_file2),
        "Status":
            "PASS"
            if (
                len(mismatches) == 0
                and len(missing_cols) == 0
                and len(extra_cols) == 0
                and len(missing_rows) == 0
                and len(extra_rows) == 0
            )
            else "FAIL" })
 
        html = generate_html(

            APP_FILE, EXCEL_FILE, df1, df2,

            mismatches, group_summary, group_ndc_map,

            missing_cols, extra_cols, missing_rows, extra_rows, dup_ndcs_file1, dup_ndcs_file2

        )
 
        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:

            f.write(html)
 
        print(f"TC{tc} completed\n")

    consolidated_file = (out_path/"Consolidated_OutputSpecialityRpt.xlsx")
    generate_consolidated_report(summary_rows , consolidated_file )
 
    print("All scenarios completed.")
 
 
 
if __name__ == "__main__":

    main()