import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

def detect_bold_rows(sheet):
    """
    Return a set of bold row indices (1-based) in column A.
    Example: {6, 8, 12} if those rows are bold.
    """
    bold_indices = set()
    for row in sheet.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if cell.font and cell.font.bold:
            bold_indices.add(cell.row)
    return bold_indices


# --------------------------------------------------------------------------
# Main Script
# --------------------------------------------------------------------------
file_path = 'data/input/test_0.xlsx'
wb = load_workbook(file_path, data_only=True)

all_data = []

for sheet_name in wb.sheetnames:
    # Only parse if "IS0" in the sheet name (per prior requirement).
    if "IS0" not in sheet_name:
        print(f"Skipping sheet '{sheet_name}' (does not contain 'IS0').")
        continue

    sheet = wb[sheet_name]
    data = list(sheet.values)

    print(f"\n=== DEBUG: Parsing Sheet '{sheet_name}' first 12 rows ===")
    for i, row_vals in enumerate(data[:12]):
        print(f"Row {i}: {row_vals}")

    # ----------------------------------------------------------------------
    # 1) Store the top 8 lines as metadata
    #
    # Adjust these names however you like. If you don't know the exact contents
    # of each row, you can name them generically (e.g. metadata1, metadata2, etc.).
    # ----------------------------------------------------------------------
    metadata = {
        "Sheet Name": sheet_name,
        "Workbook Name": file_path.split('/')[-1]
    }

    # Safely extract up to 8 rows, in case the sheet might have fewer.
    # If you *know* the sheet is always long enough, you can skip the checks.
    if len(data) >= 1: metadata["Last Refresh"]          = data[0][0]
    if len(data) >= 2: metadata["Prod Model"]            = data[1][0]
    if len(data) >= 3: metadata["IS0 Info"]              = data[2][0]
    if len(data) >= 4: metadata["Versions"]              = data[3][0]
    if len(data) >= 5: metadata["Line Items"]            = data[4][0]
    if len(data) >= 6: metadata["Scenarios"]             = data[5][0]
    if len(data) >= 7: metadata["L5 LOB"]                = data[6][0]
    if len(data) >= 8: metadata["E5 LE"]                 = data[7][0]

    # ----------------------------------------------------------------------
    # 2) Skip first 8 lines of metadata + possibly 1 blank row (total 9).
    #    So row 9 in Excel => df_data[0].
    #
    #    If your actual header is at row 8 (with no blank row),
    #    then set skip_count = 8 instead of 9.
    # ----------------------------------------------------------------------
    skip_count = 9  # 8 metadata lines + 1 blank
    df_data = data[skip_count:]
    if not df_data:
        print(f"Skipping {sheet_name} - empty after skipping metadata.")
        continue

    tmp_df = pd.DataFrame(df_data)
    if tmp_df.empty:
        print(f"Skipping {sheet_name} - no data after metadata.")
        continue

    # ----------------------------------------------------------------------
    # 3) The first row of tmp_df is the column headers
    # ----------------------------------------------------------------------
    tmp_df.columns = tmp_df.iloc[0]
    tmp_df = tmp_df.iloc[1:].copy()

    # Clean up columns
    tmp_df.columns = [
        str(col) if pd.notnull(col) else "Unnamed Column"
        for col in tmp_df.columns
    ]
    if tmp_df.empty:
        continue

    # ----------------------------------------------------------------------
    # 4) Detect bold rows for sub-header logic (bottom-up approach)
    # ----------------------------------------------------------------------
    bold_indices = detect_bold_rows(sheet)
    # After skipping 9 lines, row (skip_count) + 1 => Excel row index for tmp_df row 0
    # So if skip_count=9 => row 9 in Excel => df_data[0]
    # => excel_row = i + skip_count + 1 = i + 10
    # (But we have to check how your data lines up)

    row_data = []
    col_a_name = tmp_df.columns[0]  # The first column in the DataFrame
    other_cols = tmp_df.columns[1:]

    for i, row_series in tmp_df.iterrows():
        # This row in Excel is i + skip_count + 1
        excel_row = i + skip_count + 1
        is_bold = (excel_row in bold_indices)

        colA_value = row_series[col_a_name]
        if isinstance(colA_value, str):
            colA_value = colA_value.strip()
        if pd.isnull(colA_value):
            colA_value = ""

        row_dict = {
            "excel_row": excel_row,
            "is_bold": is_bold,
            "colA_value": colA_value,
            "other_values": {c: row_series[c] for c in other_cols}
        }
        row_data.append(row_dict)

    # ----------------------------------------------------------------------
    # 5) Bottom-up pass to assign sub-headers
    #    If row is bold, it becomes the sub-header for the rows ABOVE it
    # ----------------------------------------------------------------------
    current_subheader = None
    for i in range(len(row_data) - 1, -1, -1):
        if row_data[i]["is_bold"]:
            current_subheader = row_data[i]["colA_value"]
        else:
            row_data[i]["sub_header"] = current_subheader

    # Remove bold rows themselves (sub-headers only):
    row_data = [r for r in row_data if not r["is_bold"]]

    if not row_data:
        continue

    # ----------------------------------------------------------------------
    # 6) Build the cleaned DataFrame from row_data
    # ----------------------------------------------------------------------
    cleaned_rows = []
    for r in row_data:
        new_row = {}
        new_row["Sub-Header"] = r.get("sub_header", None)
        new_row["Financial Metric"] = r["colA_value"]
        for c in other_cols:
            new_row[c] = r["other_values"][c]
        cleaned_rows.append(new_row)

    cleaned_df = pd.DataFrame(cleaned_rows)
    if cleaned_df.empty:
        continue

    # ----------------------------------------------------------------------
    # 7) Melt wide columns (e.g. Q1 FY20) to long
    # ----------------------------------------------------------------------
    id_vars = ["Sub-Header", "Financial Metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]

    melted = cleaned_df.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Quarter/Year",
        value_name="Financial Amount"
    )

    # Split something like "Q1 FY20" into Quarter = Q1, Year = FY20
    quarter_year = melted["Quarter/Year"].str.extract(r'^(Q\d)\s*(FY\d{2})$')
    melted["Quarter"] = quarter_year[0]
    melted["Year"] = quarter_year[1]
    melted.drop(columns=["Quarter/Year"], inplace=True)

    # Replace empty or whitespace with NA
    melted.replace(r'^\s*$', pd.NA, regex=True, inplace=True)

    # ----------------------------------------------------------------------
    # 8) Attach metadata & filter rows with null amounts
    # ----------------------------------------------------------------------
    for k, v in metadata.items():
        melted[k] = v

    melted.dropna(subset=["Financial Amount"], how="any", inplace=True)

    # ----------------------------------------------------------------------
    # 9) Reorder columns
    # ----------------------------------------------------------------------
    final_order = [
        "Financial Metric",
        "Financial Amount",
        "Sub-Header",
        "Quarter",
        "Year",
        # Then your metadata columns below, rename or reorder as needed
        "Sheet Name",
        "Workbook Name",
        "Last Refresh",
        "Prod Model",
        "IS0 Info",
        "Versions",
        "Line Items",
        "Scenarios",
        "L5 LOB",
        "E5 LE"
    ]
    existing_cols = [c for c in final_order if c in melted.columns]
    extra_cols = [c for c in melted.columns if c not in existing_cols]
    final_cols = existing_cols + extra_cols

    melted = melted.loc[:, final_cols]

    all_data.append(melted)

# --------------------------------------------------------------------------
# Combine all sheets (that contain "IS0" in name) and save
# --------------------------------------------------------------------------
if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
else:
    combined_df = pd.DataFrame()

output_file_path = 'data/output/test_0.csv'
combined_df.to_csv(output_file_path, index=False)
print(f"\nProcessed workbook saved at: {output_file_path}")
