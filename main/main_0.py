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
    sheet = wb[sheet_name]
    data = list(sheet.values)

    print(f"\n=== DEBUG: Sheet '{sheet_name}' first 10 rows ===")
    for i, row_vals in enumerate(data[:10]):
        print(f"Row {i}: {row_vals}")

    # -----------------------------------------------------
    # 1) Skip the first 5 rows, because:
    #    Row 0..3 => metadata, Row 4 => blank,
    #    Row 5 => actual header
    # -----------------------------------------------------
    if len(data) <= 5:
        print(f"Skipping {sheet_name}; not enough rows to get a header.")
        continue

    metadata = {
        "Sheet Name": sheet_name,
        "Workbook Name": file_path.split('/')[-1]
    }
    # Rows 0..3 contain textual metadata; row 4 is blank (based on your debug).
    # Adjust if you have more or fewer.
    metadata["Model Category"]   = data[0][0]
    metadata["Model Name"]       = data[1][0]
    metadata["Version Name"]     = data[2][0]
    metadata["Forecast Version"] = data[3][0]

    # Row 5 is the header in Excel
    df_data = data[5:]  # row 5 => df_data[0]
    if not df_data:
        continue

    tmp_df = pd.DataFrame(df_data)
    if tmp_df.empty:
        continue

    # The first row in tmp_df => your column headers: (None, 'Q1 FY20', 'Q1 FY22', ...)
    tmp_df.columns = tmp_df.iloc[0]
    tmp_df = tmp_df.iloc[1:].copy()  # real data starts at Excel row 6

    tmp_df.columns = [
        str(col) if pd.notnull(col) else "Unnamed Column"
        for col in tmp_df.columns
    ]

    if tmp_df.empty:
        continue

    # -----------------------------------------------------
    # 2) Collect Rows in a List, with Bold Flags
    #    We'll do sub-header assignment in a bottom-up pass.
    # -----------------------------------------------------
    bold_indices = detect_bold_rows(sheet)
    # Remember: row 6 in Excel => tmp_df row 0, row 7 => tmp_df row 1, etc.
    # So the Excel row index for tmp_df row i = i + 6
    # We'll store them in row_data[] with "is_bold", "excel_row", etc.
    row_data = []
    col_a_name = tmp_df.columns[0]  # first column in the DataFrame
    other_cols = tmp_df.columns[1:]

    for i, row_series in tmp_df.iterrows():
        excel_row = i + 6
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

    if not row_data:
        continue

    # -----------------------------------------------------
    # 3) Bottom-Up Pass to Assign Sub-Headers
    #    If we encounter a bold row, that becomes the current_subheader
    #    for the rows above it (until the next bold row).
    # -----------------------------------------------------
    current_subheader = None
    for i in range(len(row_data)-1, -1, -1):  # from last to first
        if row_data[i]["is_bold"]:
            # This row itself is a sub-header => store the text as sub-header
            current_subheader = row_data[i]["colA_value"]
            # We typically do NOT keep the bold row as a data row,
            # so we can mark it for removal or skip later.
        else:
            # Non-bold => assign the current_subheader
            row_data[i]["sub_header"] = current_subheader

    # Filter out the bold rows themselves (if you do NOT want them in final data)
    # If you do want them, comment out this line and handle them separately.
    row_data = [r for r in row_data if not r["is_bold"]]

    if not row_data:
        continue

    # -----------------------------------------------------
    # 4) Build a "cleaned" DataFrame from row_data
    # -----------------------------------------------------
    cleaned_rows = []
    for r in row_data:
        new_row = {}
        new_row["Sub-Header"] = r.get("sub_header", None)
        new_row["Financial Metric"] = r["colA_value"]
        # Add other columns
        for c in other_cols:
            new_row[c] = r["other_values"][c]
        cleaned_rows.append(new_row)

    cleaned_df = pd.DataFrame(cleaned_rows)
    if cleaned_df.empty:
        continue

    # -----------------------------------------------------
    # 5) Melt wide columns (e.g., Q1 FY20) to long
    # -----------------------------------------------------
    id_vars = ["Sub-Header", "Financial Metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]

    melted = cleaned_df.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Quarter/Year",
        value_name="Financial Amount",
    )

    # Optionally parse Q1 FY20 => Quarter='Q1', Year='FY20'
    quarter_year = melted["Quarter/Year"].str.extract(r'^(Q\d)\s*(FY\d{2})$')
    melted["Quarter"] = quarter_year[0]
    melted["Year"] = quarter_year[1]

    melted.drop(columns=["Quarter/Year"], inplace=True)

    # Replace empty or whitespace with NA
    melted.replace(r'^\s*$', pd.NA, regex=True, inplace=True)

    # -----------------------------------------------------
    # 6) Attach metadata & Filter out rows with no Financial Amount
    # -----------------------------------------------------
    for k, v in metadata.items():
        melted[k] = v

    melted.dropna(subset=["Financial Amount"], how="any", inplace=True)

    # -----------------------------------------------------
    # 7) Reorder columns
    # -----------------------------------------------------
    final_order = [
        "Financial Metric",
        "Financial Amount",
        "Sub-Header",
        "Quarter",
        "Year",
        "Sheet Name",
        "Workbook Name",
        "Model Category",
        "Model Name",
        "Version Name",
        "Forecast Version",
    ]
    existing_cols = [c for c in final_order if c in melted.columns]
    extra_cols = [c for c in melted.columns if c not in existing_cols]
    final_cols = existing_cols + extra_cols

    melted = melted.loc[:, final_cols]

    all_data.append(melted)

# --------------------------------------------------------------------------
# Combine all sheets and save
# --------------------------------------------------------------------------
if all_data:
    combined_df = pd.concat(all_data, ignore_index=True)
else:
    combined_df = pd.DataFrame()

output_file_path = 'data/output/test_0.csv'
combined_df.to_csv(output_file_path, index=False)
print(f"\nProcessed workbook saved at: {output_file_path}")
