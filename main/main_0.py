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

def process_sheet(sheet, sheet_name, file_path):
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

    tmp_df.columns = [str(col).strip().lower() if pd.notnull(col) else "unnamed_column" for col in tmp_df.iloc[0]]
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
    col_a_name = tmp_df.columns[0] if tmp_df.columns[0] else "financial_metric"
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
    id_vars = ["sub-header", "financial metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]
    melted = cleaned_df.melt(id_vars=id_vars, value_vars=value_vars, var_name="quarter/year", value_name="financial_amount")
    
    quarter_year = melted["quarter/year"].str.extract(r'^(Q\d)\s*[-]?\s*(FY\d{2,4})$', expand=False)
    melted["quarter"], melted["year"] = quarter_year[0], quarter_year[1]
    melted.dropna(subset=["quarter", "year", "financial_amount"], inplace=True)
    return melted

def attach_metadata(melted, metadata):
    """Attach metadata to the melted DataFrame."""
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

    combined_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    combined_df.drop_duplicates(inplace=True)  # Remove duplicates
    combined_df.to_csv(output_file_path, index=False, encoding="utf-8")  # Ensure UTF-8 encoding
    print(f"\nProcessed workbook saved at: {output_file_path}")

# Test-related functions

def test_shapes_match(input_path, output_path):
    """Test that input shape matches output rows (per sheet: rows * columns = total output rows)."""
    df_input = pd.read_excel(input_path, None)  # Load all sheets
    total_input_cells = sum([df.shape[0] * df.shape[1] for df in df_input.values()])
    df_output = pd.read_csv(output_path)
    total_output_rows = df_output.shape[0]
    assert total_input_cells == total_output_rows, "Mismatch between input cells and output rows!"

def test_aggregates_match(input_path, output_path):
    """Test that aggregate Financial Amounts match."""
    df_input = pd.read_excel(input_path, None)  # Load all sheets
    total_input_sum = sum([pd.DataFrame(sheet).iloc[:, 1:].sum().sum() for sheet in df_input.values()])
    df_output = pd.read_csv(output_path)
    total_output_sum = df_output["financial_amount"].sum()
    assert round(total_input_sum, 2) == round(total_output_sum, 2), "Mismatch in Financial Amount totals!"

if __name__ == "__main__":
    # Specify input and output paths
    input_folder = "data/input"
    output_folder = "data/output"

    input_file = os.path.join(input_folder, "example.xlsx")
    output_file = os.path.join(output_folder, "example.csv")

    process_workbook(input_file, output_file)

    # Run unit tests
    test_shapes_match(input_file, output_file)
    test_aggregates_match(input_file, output_file)
