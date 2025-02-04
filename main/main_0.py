import os
import warnings
import logging
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from tqdm import tqdm

# ----------------------------------------------------------------
# Configure warnings and logging
warnings.simplefilter("ignore", category=UserWarning)
warnings.simplefilter("ignore", category=DeprecationWarning)

logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(message)s'
)

# ----------------------------------------------------------------
def extract_metadata(file_path: str, sheet_name: str) -> dict:
    """
    Extract or build any relevant metadata from a sheet.
    Adjust this to suit your actual metadata needs.
    """
    return {
        "workbook_name": Path(file_path).stem,
        "sheet_name": sheet_name
    }

# ----------------------------------------------------------------
def build_row_data(df: pd.DataFrame) -> list:
    """
    Turn each row in 'df' into a dictionary containing:
      - row index or 'Excel row'
      - category (e.g. 'IS Statement Header', 'Financial Metric', 'Sub‐Header', etc.)
      - other columns from the raw DataFrame you need.

    This example checks if the first column contains 'header' (case-insensitive);
    if so, the category is 'IS Statement Header'. If it contains 'sub', it's 'Sub-Header'.
    Otherwise, it's 'Financial Metric'.
    """
    row_data = []
    for i, row_series in df.iterrows():
        # Example logic: detect category from the first column text
        text_value = str(row_series.iloc[0]).strip().lower()  # or whichever column you rely on
        if "header" in text_value:
            category = "IS Statement Header"
        elif "sub" in text_value:
            category = "Sub-Header"
        else:
            category = "Financial Metric"
        
        row_data.append({
            "excel_row": i + 1,
            "category": category,
            "values": row_series.to_dict()
        })
    return row_data

# ----------------------------------------------------------------
def assign_headers(row_data: list) -> list:
    """
    Traverse row_data in reverse to propagate headers or sub‐headers
    down into each row. Adjust as needed for your data hierarchy.
    """
    current_is_header = None
    current_sub_header = None
    current_main_header = None

    for row in reversed(row_data):
        cat = row["category"]
        if cat == "IS Statement Header":
            current_is_header = row["excel_row"]
            row["IS Statement Header"] = current_is_header
        elif cat == "Sub-Header":
            current_sub_header = row["excel_row"]
            row["Sub Header"] = current_sub_header
        else:
            current_main_header = row["excel_row"]
            row["Main Header"] = current_main_header

        # If you'd like every row to store the nearest IS Statement Header / Sub‐Header / Main Header,
        # uncomment these lines:
        #
        # row["IS Statement Header"] = current_is_header
        # row["Sub Header"] = current_sub_header
        # row["Main Header"] = current_main_header

    return row_data

# ----------------------------------------------------------------
def build_cleaned_dataframe(row_data: list) -> pd.DataFrame:
    """
    Convert row_data (list of dictionaries) back into a DataFrame
    with consistent column names/structure.
    """
    if not row_data:
        return pd.DataFrame()

    flattened = []
    for rd in row_data:
        base = {
            "Excel Row": rd["excel_row"],
            "Category": rd["category"],
            "IS Statement Header": rd.get("IS Statement Header", None),
            "Sub Header": rd.get("Sub Header", None),
            "Main Header": rd.get("Main Header", None)
        }
        # Merge in values from the raw row
        for k, v in rd["values"].items():
            base[str(k)] = v
        flattened.append(base)

    return pd.DataFrame(flattened)

# ----------------------------------------------------------------
def melt_and_parse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Example transformation that 'melts' wide columns into long format
    (e.g. for Quarter/Year columns). If you don't need this, you can remove it.
    """
    id_cols = ["Excel Row", "Category", "IS Statement Header",
               "Sub Header", "Main Header"]

    value_cols = [c for c in df.columns if c not in id_cols]
    if not value_cols:
        return df  # nothing to melt

    melted = df.melt(
        id_vars=id_cols,
        value_vars=value_cols,
        var_name="Quarter/Year",
        value_name="Financial Amount"
    )

    # Simplified, no extra lines for year/quarter parsing
    melted["quarter_year"] = melted["Quarter/Year"].str.strip()
    return melted

# ----------------------------------------------------------------
def process_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Load a single sheet, build row data, assign headers,
    build a cleaned DataFrame, melt/parse, return final.
    Also prints any 'IS Statement Header' rows.
    """
    wb = load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()

    sheet = wb[sheet_name]
    data = sheet.values
    df = pd.DataFrame(data)
    # Optionally skip the first row or remove empty columns, e.g.:
    # df = df.iloc[1:, :]

    # Build row data
    row_data = build_row_data(df)

    # Assign headers in reverse
    row_data = assign_headers(row_data)

    # Build cleaned DataFrame
    cleaned_df = build_cleaned_dataframe(row_data)

    # Print "IS Statement Header" rows
    is_mask = cleaned_df["Category"] == "IS Statement Header"
    if is_mask.any():
        logging.info(f"Workbook: {file_path}, Sheet: {sheet_name} — "
                     f"Found IS Statement Header row(s):")
        for idx, row in cleaned_df[is_mask].iterrows():
            logging.info(f"  Row {row['Excel Row']}: {row['IS Statement Header']}")

    # Melt & parse (if you need to transform wide to long)
    final_df = melt_and_parse(cleaned_df)
    return final_df

# ----------------------------------------------------------------
def process_workbooks(input_path: str, output_file: str) -> None:
    """
    Iterate over .xls/.xlsx files in 'input_path', process each sheet,
    combine results, and write a CSV to 'output_file'.
    """
    all_data = []
    files = list(Path(input_path).glob("*.xls*"))  # matches .xls & .xlsx
    if not files:
        logging.warning("No Excel files found in input path.")

    for file_path in tqdm(files, desc="Processing Excel files", unit="file"):
        wb = load_workbook(file_path, data_only=True)
        for sheet_name in wb.sheetnames:
            meta = extract_metadata(str(file_path), sheet_name)
            df = process_sheet(str(file_path), sheet_name)
            if not df.empty:
                df["Workbook Name"] = meta["workbook_name"]
                df["Sheet Name"] = meta["sheet_name"]
                all_data.append(df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_df.drop_duplicates(inplace=True)
        combined_df.to_csv(output_file, index=False, encoding="utf-8-sig")
        logging.info(f"Combined output saved to: {output_file}")
    else:
        logging.info("No data found in any workbook; no output generated.")

# ----------------------------------------------------------------
if __name__ == "__main__":
    # Example usage:
    input_path = "data/input"           # Folder with your .xls/.xlsx files
    output_file = "data/output/test.csv"
    
    process_workbooks(input_path, output_file)
    logging.info("All processing complete!")
