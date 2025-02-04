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
    metadata = {
        "workbook_name": Path(file_path).stem,
        "sheet_name": sheet_name
        # Add more fields here if needed
    }
    return metadata

# ----------------------------------------------------------------
def build_row_data(df: pd.DataFrame) -> list:
    """
    Turn each row in 'df' into a dictionary containing:
      - row index or 'Excel row'
      - category (e.g. 'IS Statement Header', 'Financial Metric', 'Sub‐Header', etc.)
      - any other columns from the raw DataFrame you need
    """
    row_data = []
    for i, row_series in df.iterrows():
        # Determine row category (dummy example—replace with your own logic)
        # You might detect categories by indentation level, text matching, etc.
        text_value = str(row_series.iloc[0]).strip().lower()  # Example from first column
        if "header" in text_value:
            category = "IS Statement Header"
        elif "sub" in text_value:
            category = "Sub-Header"
        else:
            category = "Financial Metric"
        
        row_data.append({
            "excel_row": i + 1,  # or i, or skip_count + i, etc.
            "category": category,
            "values": row_series.to_dict()  # or pick out columns you care about
        })
    return row_data

# ----------------------------------------------------------------
def assign_headers(row_data: list) -> list:
    """
    Walk through row_data in reverse (or however is appropriate)
    to propagate headers or sub‐headers down into each row dictionary.
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
            # Example: treat everything else as a “main” or “financial metric” row
            current_main_header = row["excel_row"]
            row["Main Header"] = current_main_header

        # If you need to fill them in for every row below, do so:
        # row["IS Statement Header"] = current_is_header
        # row["Sub Header"] = current_sub_header
        # row["Main Header"] = current_main_header

    return row_data

# ----------------------------------------------------------------
def build_cleaned_dataframe(row_data: list) -> pd.DataFrame:
    """
    Convert your list of dictionaries (row_data) back into a DataFrame
    with well‐labeled columns. 
    """
    if not row_data:
        return pd.DataFrame()

    # Flatten each row_data dict as needed
    flattened = []
    for rd in row_data:
        base = {
            "Excel Row": rd["excel_row"],
            "Category": rd["category"],
            "IS Statement Header": rd.get("IS Statement Header", None),
            "Sub Header": rd.get("Sub Header", None),
            "Main Header": rd.get("Main Header", None)
        }
        # Merge in actual data values from row["values"]
        for k, v in rd["values"].items():
            base[str(k)] = v
        flattened.append(base)

    df = pd.DataFrame(flattened)
    return df

# ----------------------------------------------------------------
def melt_and_parse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Example of a 'melt' or pivot step for Quarter/Year columns, etc.
    Adjust to match your actual data transformations.
    """
    # Suppose we identify ID columns vs. numeric columns
    id_cols = ["Excel Row", "Category", "IS Statement Header", 
               "Sub Header", "Main Header"]  # etc.
    
    # If there are columns that look like “Q1 2024” or so, pivot them
    value_cols = [c for c in df.columns if c not in id_cols]
    melted = df.melt(id_vars=id_cols, value_vars=value_cols,
                     var_name="Quarter/Year", value_name="Financial Amount")

    # Simple parse to unify Quarter/Year into separate fields (example)
    melted["quarter_year"] = melted["Quarter/Year"].str.strip()
    # Additional logic to extract numeric year, quarter, etc.
    # ...
    
    return melted

# ----------------------------------------------------------------
def process_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Load a single sheet, build row data, assign headers,
    build a cleaned DataFrame, melt/parse, return final result.
    Also print any 'IS Statement Header' values found.
    """
    wb = load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()  # or None

    sheet = wb[sheet_name]
    # Convert sheet range to DataFrame. Adjust row/col limits to your data
    data = sheet.values
    df = pd.DataFrame(data)  # You might skip the first row if it’s a title, etc.

    # 1. Build row data from raw sheet DataFrame
    row_data = build_row_data(df)

    # 2. Assign headers (propagate them downward)
    row_data = assign_headers(row_data)

    # 3. Build a cleaned structured DataFrame
    cleaned_df = build_cleaned_dataframe(row_data)

    # 4. Print any “IS Statement Header” rows for debugging/logging
    is_mask = cleaned_df["Category"] == "IS Statement Header"
    if is_mask.any():
        logging.info(f"Workbook: {file_path}, Sheet: {sheet_name} — "
                     f"Found IS Statement Header row(s):")
        for idx, row in cleaned_df[is_mask].iterrows():
            logging.info(f"  Row {row['Excel Row']}: {row['IS Statement Header']}")

    # 5. Melt & parse if needed
    final_df = melt_and_parse(cleaned_df)
    return final_df

# ----------------------------------------------------------------
def process_workbooks(input_path: str, output_file: str) -> None:
    """
    Iterate over all .xls/.xlsx files in 'input_path',
    process each sheet, combine the results, and write to CSV.
    """
    all_data = []
    files = list(Path(input_path).glob("*.xls*"))  # .xls or .xlsx
    if not files:
        logging.warning("No Excel files found in input path.")

    for file_path in tqdm(files, desc="Processing Excel files", unit="file"):
        wb = load_workbook(file_path, data_only=True)
        for sheet_name in wb.sheetnames:
            # Optional: skip hidden sheets, etc. if needed
            metadata = extract_metadata(str(file_path), sheet_name)
            df = process_sheet(str(file_path), sheet_name)
            # Attach metadata columns to each row
            if not df.empty:
                df["Workbook Name"] = metadata["workbook_name"]
                df["Sheet Name"] = metadata["sheet_name"]
                all_data.append(df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        # Drop duplicates if needed
        combined_df.drop_duplicates(inplace=True)
        # Write out the final CSV
        combined_df.to_csv(output_file, index=False, encoding="utf-8-sig")
        logging.info(f"Combined output saved to: {output_file}")
    else:
        logging.info("No data to combine. Finished with no output.")

# ----------------------------------------------------------------
if __name__ == "__main__":
    # Example usage:
    input_path = "data/input"
    output_file = "data/output/test_output.csv"
    process_workbooks(input_path, output_file)
