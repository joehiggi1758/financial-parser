import os
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

def process_sheet(sheet, sheet_name, file_path, expected_sizes):
    """Process a single sheet and return a cleaned DataFrame."""
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

    tmp_df.columns = [str(col).strip() if pd.notnull(col) else "Unnamed Column" for col in tmp_df.iloc[0]]
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
    cleaned_df = cleaned_df.drop_duplicates().reset_index(drop=True)  # Ensure no duplicates and reset index
    melted = melt_and_parse(cleaned_df)

    if melted.empty:
        print(f"Skipping {sheet_name}; melted DataFrame is empty.")
        return None

    # Update expected_sizes with the calculated size for this sheet
    expected_sizes[sheet_name] = cleaned_df.shape[0] * cleaned_df.shape[1]

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
            "other_values": {c: row_series[c] if pd.notnull(row_series[c]) else None for c in other_cols},
        }
        row_data.append(row_dict)
    return row_data

def assign_sub_headers(row_data):
    """Assign sub-headers to rows based on bold formatting."""
    current_subheader = None
    for row in reversed(row_data):
        if row["is_bold"]:
            current_subheader = row["colA_value"]
        row["sub_header"] = current_subheader

def build_cleaned_df(row_data, columns):
    """Build a cleaned DataFrame from row data."""
    cleaned_rows = []
    for row in row_data:
        new_row = {"Sub-Header": row.get("sub_header"), "Financial Metric": row["colA_value"]}
        new_row.update(row["other_values"])
        cleaned_rows.append(new_row)
    return pd.DataFrame(cleaned_rows)

def melt_and_parse(cleaned_df):
    """Melt DataFrame into long format and parse Quarter/Year."""
    id_vars = ["Sub-Header", "Financial Metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]
    melted = cleaned_df.melt(id_vars=id_vars, value_vars=value_vars, var_name="Quarter/Year", value_name="Financial Amount")
    
    quarter_year = melted["Quarter/Year"].str.extract(r'^(Q\d)\s*[-]?\s*(FY\d{2,4})$', expand=False)
    melted["Quarter"] = quarter_year[0]
    melted["Year"] = quarter_year[1]

    # Handle cases where only "FY29" is present
    fy_only = melted["Quarter/Year"].str.match(r'^FY\d{2,4}$')
    melted.loc[fy_only, "Year"] = melted.loc[fy_only, "Quarter/Year"].str.extract(r'^FY(\d{2,4})$')[0]
    melted.loc[fy_only, "Quarter"] = "All Periods"

    # Convert Year to full format if necessary
    melted["Year"] = melted["Year"].apply(lambda x: f"20{x[-2:]}" if pd.notnull(x) and len(x) == 2 else x)

    melted.dropna(subset=["Quarter", "Year", "Financial Amount"], inplace=True)
    return melted

def attach_metadata(melted, metadata):
    """Attach metadata to the melted DataFrame."""
    for k, v in metadata.items():
        melted[k] = v

def process_workbook(file_path, output_file_path):
    """Process an Excel workbook and save the combined DataFrame to a CSV file."""
    wb = load_workbook(file_path, data_only=True)
    all_data = []
    expected_sizes = {}  # To track the expected sizes

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        melted = process_sheet(sheet, sheet_name, file_path, expected_sizes)
        if melted is not None:
            all_data.append(melted)

    combined_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    combined_df.drop_duplicates(inplace=True)  # Ensure no duplicates
    combined_df.to_csv(output_file_path, index=False, encoding='utf-8')  # Force UTF-8 encoding
    print(f"\nProcessed workbook saved at: {output_file_path}")

    return expected_sizes

# Test-related functions

def test_shapes_match(output_path, expected_sizes):
    """Test that input cells match output cells in aggregate."""
    df_output = pd.read_csv(output_path, encoding='utf-8')  # Ensure UTF-8 encoding on import
    total_output_cells = df_output.shape[0] * df_output.shape[1]  # Total cells in the output

    total_expected_cells = sum(expected_sizes.values())  # Sum of expected sizes from all sheets

    assert total_output_cells == total_expected_cells, (
        f"Mismatch in cell counts: Expected {total_expected_cells}, but got {total_output_cells}!"
    )

if __name__ == "__main__":
    # Specify input and output folder paths
    input_folder = "data/input"
    output_folder = "data/output"
    input_file = os.path.join(input_folder, "Test-2.xlsx")
    output_file = os.path.join(output_folder, "example.csv")

    expected_sizes = process_workbook(input_file, output_file)

    # Run unit tests
    test_shapes_match(output_file, expected_sizes)
