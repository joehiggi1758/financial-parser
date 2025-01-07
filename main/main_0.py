import os
import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

def detect_bold_rows(sheet):
    """Return a set of bold row indices (1-based) in column A."""
    bold_indices = set()
    for row in sheet.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if cell.font and cell.font.bold:
            bold_indices.add(cell.row)
    return bold_indices

def extract_metadata(data, sheet_name, file_path):
    """Extract metadata from the first 8 rows of the sheet."""
    return {
        "Sheet Name": sheet_name,
        "Workbook Name": file_path.split('/')[-1],
        "Meta 1": data[0][0] if len(data) > 0 else None,
        "Meta 2": data[1][0] if len(data) > 1 else None,
        "Meta 3": data[2][0] if len(data) > 2 else None,
        "Meta 4": data[3][0] if len(data) > 3 else None,
        "Meta 5": data[4][0] if len(data) > 4 else None,
        "Meta 6": data[5][0] if len(data) > 5 else None,
        "Meta 7": data[6][0] if len(data) > 6 else None,
        "Meta 8": data[7][0] if len(data) > 7 else None,
    }

def process_sheet(sheet, sheet_name, file_path):
    """Process a single sheet and return a cleaned (melted) DataFrame."""
    data = list(sheet.values)
    skip_count = 8
    if len(data) <= skip_count:
        print(f"Skipping {sheet_name}; not enough rows to reach a header.")
        return None

    metadata = extract_metadata(data, sheet_name, file_path)

    df_data = data[skip_count:]
    tmp_df = pd.DataFrame(df_data)
    if tmp_df.empty:
        print(f"Skipping {sheet_name}; no data after skip_count.")
        return None

    # First row of tmp_df becomes the new column headers
    tmp_df.columns = [
        str(col).strip() if pd.notnull(col) else "Unnamed Column" 
        for col in tmp_df.iloc[0]
    ]
    # Drop the first row (now headers) and any columns completely empty
    tmp_df = tmp_df.iloc[1:].dropna(axis=1, how='all').copy()
    if tmp_df.empty:
        print(f"Skipping {sheet_name}; tmp_df is empty after processing.")
        return None

    bold_indices = detect_bold_rows(sheet)
    row_data = build_row_data(tmp_df, bold_indices, skip_count)

    if not row_data:
        print(f"Skipping {sheet_name}; no data in row_data after processing.")
        return None

    assign_sub_headers(row_data)
    cleaned_df = build_cleaned_df(row_data, tmp_df.columns)
    melted = melt_and_parse(cleaned_df)

    if melted.empty:
        print(f"Skipping {sheet_name}; melted DataFrame is empty.")
        return None

    attach_metadata(melted, metadata)
    return melted

def build_row_data(tmp_df, bold_indices, skip_count):
    """Build a list of row data dictionaries for processing."""
    row_data = []
    col_a_name = tmp_df.columns[0] if tmp_df.columns[0] else "Financial Metric"
    other_cols = tmp_df.columns[1:]

    for i, row_series in tmp_df.iterrows():
        excel_row = i + skip_count + 1
        is_bold = excel_row in bold_indices
        colA_value = row_series[col_a_name].strip() if isinstance(row_series[col_a_name], str) else ""

        row_dict = {
            "excel_row": excel_row,
            "is_bold": is_bold,
            "colA_value": colA_value,
            "other_values": {c: row_series[c] for c in other_cols},
        }
        row_data.append(row_dict)
    return row_data

def assign_sub_headers(row_data):
    """Assign sub-headers to rows based on bold formatting."""
    current_subheader = None
    # Process from bottom to top so each row inherits the closest 'bold' above it
    for row in reversed(row_data):
        if row["is_bold"]:
            current_subheader = row["colA_value"]
        row["sub_header"] = current_subheader

def build_cleaned_df(row_data, columns):
    """Build a cleaned DataFrame from row data."""
    cleaned_rows = []
    for row in row_data:
        new_row = {
            "Sub-Header": row.get("sub_header"), 
            "Financial Metric": row["colA_value"]
        }
        # Merge all other columns from 'other_values'
        new_row.update(row["other_values"])
        cleaned_rows.append(new_row)
    return pd.DataFrame(cleaned_rows)

def melt_and_parse(cleaned_df):
    """
    Melt DataFrame into long format and parse Quarter/Year.

    Preserves the final structure:
      - "Sub-Header"
      - "Financial Metric"
      - "Quarter/Year"
      - "Financial Amount"
      - "Quarter"
      - "Year"
    Then we only drop rows where "Financial Amount" is missing.
    """
    id_vars = ["Sub-Header", "Financial Metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]
    melted = cleaned_df.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Quarter/Year",
        value_name="Financial Amount"
    )

    # Custom parsing to handle "All Periods", "FY29", or "Q1 FY29"
    def parse_qy(qy):
        text = str(qy).strip().lower()
        # 1) All Periods
        if text == "all periods":
            return ("All Periods", None)
        # 2) Q1 FY29 (or Q2-FY2029, etc.)
        match_qfy = re.match(r'^(q\d)\s*[-]?\s*(fy(\d{2,4}))$', text, re.IGNORECASE)
        if match_qfy:
            quarter = match_qfy.group(1).upper()  # e.g. "Q1"
            year_abbrev = match_qfy.group(3)      # e.g. "29" or "2029"
            # Convert 2-digit year to 20xx
            if len(year_abbrev) == 2:
                year_abbrev = "20" + year_abbrev
            return (quarter, year_abbrev)

        # 3) Just FY29 (no quarter)
        match_fy = re.match(r'^fy(\d{2,4})$', text, re.IGNORECASE)
        if match_fy:
            year_abbrev = match_fy.group(1)  # e.g. "29"
            if len(year_abbrev) == 2:
                year_abbrev = "20" + year_abbrev
            return (None, year_abbrev)

        # If we don't match anything, return (None, None)
        return (None, None)

    melted["Quarter"], melted["Year"] = zip(*melted["Quarter/Year"].apply(parse_qy))

    # Keep rows only if there's a valid Financial Amount
    # (We do NOT drop Quarter or Year if they are None — that way 
    #  "All Periods" or just "FY29" won't be discarded)
    melted.dropna(subset=["Financial Amount"], inplace=True)

    return melted

def attach_metadata(melted, metadata):
    """Attach metadata columns to the melted DataFrame."""
    for k, v in metadata.items():
        melted[k] = v

def process_workbook(file_path, output_file_path):
    """Process an Excel workbook and save the combined DataFrame to a CSV file."""
    wb = load_workbook(file_path, data_only=True)
    all_data = []

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        melted = process_sheet(sheet, sheet_name, file_path)
        if melted is not None:
            all_data.append(melted)

    # Combine all sheets
    combined_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

    # --- Fix for "TypeError: unhashable type: 'Series'" ---
    # Convert any unhashable objects (lists, dicts, sets, or pd.Series) to strings before dropping duplicates
    def make_hashable(value):
        if isinstance(value, (dict, list, set, pd.Series)):
            return str(value)
        return value

    combined_df = combined_df.applymap(make_hashable)
    # Drop entire-row duplicates
    combined_df.drop_duplicates(inplace=True)

    # Force UTF-8 encoding on output
    combined_df.to_csv(output_file_path, index=False, encoding='utf-8')
    print(f"\nProcessed workbook saved at: {output_file_path}")

# ------------------ TEST FUNCTIONS ------------------

def test_shapes_match(input_path, output_path):
    """
    Test that the total number of cells (rows × columns) in the input 
    matches the number of rows in the output.
    """
    df_input = pd.read_excel(input_path, None)  # Load all sheets
    # Sum across all sheets: total cells = rows * columns
    total_input_cells = sum(df.shape[0] * df.shape[1] for df in df_input.values())

    df_output = pd.read_csv(output_path)
    total_output_rows = df_output.shape[0]

    assert total_input_cells == total_output_rows, (
        "Mismatch: (input rows × input columns) != output rows!"
    )

def test_aggregates_match(input_path, output_path):
    """Test that aggregate Financial Amounts match."""
    df_input = pd.read_excel(input_path, None)  # Load all sheets
    # Sum up the numeric values in columns (excluding the first, which might be 'Financial Metric')
    total_input_sum = sum([
        pd.DataFrame(sheet).iloc[:, 1:].sum().sum() 
        for sheet in df_input.values()
    ])

    df_output = pd.read_csv(output_path)
    total_output_sum = df_output["Financial Amount"].sum()

    # Round both to 2 decimals to avoid floating-point mismatches
    assert round(total_input_sum, 2) == round(total_output_sum, 2), (
        "Mismatch in Financial Amount totals!"
    )

if __name__ == "__main__":
    # Example usage
    input_folder = "data/input"
    output_folder = "data/output"
    process_workbook(
        os.path.join(input_folder, "example.xlsx"), 
        os.path.join(output_folder, "example.csv")
    )

    # Run unit tests (optional)
    test_shapes_match(
        os.path.join(input_folder, "example.xlsx"), 
        os.path.join(output_folder, "example.csv")
    )
    test_aggregates_match(
        os.path.join(input_folder, "example.xlsx"), 
        os.path.join(output_folder, "example.csv")
    )
