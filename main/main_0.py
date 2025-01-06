import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

def detect_bold_rows(sheet):
    """
    Return a set of bold row indices (1-based) in column A.
    Example: {10, 12, 56} if those rows are bold.
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
    # If you only want sheets with "IS0" in the name, uncomment:
    # if "IS0" not in sheet_name:
    #     print(f"Skipping sheet '{sheet_name}' (does not contain 'IS0').")
    #     continue

    sheet = wb[sheet_name]
    data = list(sheet.values)

    # ----------------------------------------------------------------------
    # 1) Skip the first 8 rows (rows 0–7).
    #    => Row 8 in Excel becomes the "header row".
    # ----------------------------------------------------------------------
    skip_count = 8
    if len(data) <= skip_count:
        print(f"Skipping {sheet_name}; not enough rows to reach a header.")
        continue

    # Collect your eight lines of metadata (optional):
    metadata = {
        "Sheet Name": sheet_name,
        "Workbook Name": file_path.split('/')[-1],
        "Meta 1": data[0][0],  # row 0
        "Meta 2": data[1][0],  # row 1
        "Meta 3": data[2][0],  # row 2
        "Meta 4": data[3][0],  # row 3
        "Meta 5": data[4][0],  # row 4
        "Meta 6": data[5][0],  # row 5
        "Meta 7": data[6][0],  # row 6
        "Meta 8": data[7][0],  # row 7
    }

    # ----------------------------------------------------------------------
    # 2) Build a DataFrame from rows after skip_count (row 8 = header)
    # ----------------------------------------------------------------------
    df_data = data[skip_count:]  # row 8 => df_data[0]
    if not df_data:
        continue

    tmp_df = pd.DataFrame(df_data)
    if tmp_df.empty:
        continue

    # The first row in tmp_df => column headers (Excel row 8)
    tmp_df.columns = tmp_df.iloc[0]
    tmp_df = tmp_df.iloc[1:].copy()  # Actual data starts at Excel row 9+

    # Convert columns to strings, fill "Unnamed Column" if needed
    tmp_df.columns = [
        str(col) if pd.notnull(col) else "Unnamed Column"
        for col in tmp_df.columns
    ]
    if tmp_df.empty:
        continue

    # ----------------------------------------------------------------------
    # 3) Detect bold rows for sub-header logic (bottom-up approach)
    # ----------------------------------------------------------------------
    bold_indices = detect_bold_rows(sheet)
    row_data = []
    col_a_name = tmp_df.columns[0]
    other_cols = tmp_df.columns[1:]

    for i, row_series in tmp_df.iterrows():
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

    if not row_data:
        continue

    # ----------------------------------------------------------------------
    # 4) Bottom-Up Pass to assign sub-headers
    # ----------------------------------------------------------------------
    current_subheader = None
    for i in range(len(row_data) - 1, -1, -1):
        if row_data[i]["is_bold"]:
            current_subheader = row_data[i]["colA_value"]
        else:
            row_data[i]["sub_header"] = current_subheader

    # Remove the bold rows themselves (if you do NOT want them in the final data)
    row_data = [r for r in row_data if not r["is_bold"]]
    if not row_data:
        continue

    # ----------------------------------------------------------------------
    # 5) Build a "cleaned" DataFrame from row_data
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
    # 6) Melt wide columns (e.g., Q1 FY20, Q2 FY2024) => long format
    # ----------------------------------------------------------------------
    id_vars = ["Sub-Header", "Financial Metric"]
    value_vars = [c for c in cleaned_df.columns if c not in id_vars]

    melted = cleaned_df.melt(
        id_vars=id_vars,
        value_vars=value_vars,
        var_name="Quarter/Year",
        value_name="Financial Amount"
    )

    # ----------------------------------------------------------------------
    # 7) Parse something like "Q1 FY20" or "Q3 FY2024"
    # ----------------------------------------------------------------------
    # Add debug print to check melted content
    # print("DEBUG: Melted DataFrame before Quarter/Year parsing:\n", melted.head())

    # Ensure Quarter/Year parsing works for all rows
    quarter_year = melted["Quarter/Year"].str.extract(r'^(Q\d)\s*(FY\d{2,4})$')
    melted["Quarter"] = quarter_year[0]
    melted["Year"] = quarter_year[1]

    # Check and fill missing Quarter/Year values
    melted["Quarter"] = melted["Quarter"].fillna(method='ffill')  # Forward-fill missing Quarters
    melted["Year"] = melted["Year"].fillna(method='ffill')  # Forward-fill missing Years

    melted.drop(columns=["Quarter/Year"], inplace=True)

    # ----------------------------------------------------------------------
    # 8) Replace empty or whitespace with NA
    # ----------------------------------------------------------------------
    melted.replace(r'^\s*$', pd.NA, regex=True, inplace=True)

    # ----------------------------------------------------------------------
    # 9) Attach metadata
    # ----------------------------------------------------------------------
    for k, v in metadata.items():
        melted[k] = v

    # ----------------------------------------------------------------------
    # 10) Only keep rows where Financial Amount is non-null
    # ----------------------------------------------------------------------
    melted.dropna(subset=["Financial Amount"], how="any", inplace=True)

    # ----------------------------------------------------------------------
    # 11) Reorder columns
    # ----------------------------------------------------------------------
    final_order = [
        "Financial Metric",
        "Financial Amount",
        "Sub-Header",
        "Quarter",
        "Year",
        "Sheet Name",
        "Workbook Name",
        "Meta 1",
        "Meta 2",
        "Meta 3",
        "Meta 4",
        "Meta 5",
        "Meta 6",
        "Meta 7",
        "Meta 8",
    ]
    existing_cols = [c for c in final_order if c in melted.columns]
    extra_cols = [c for c in melted.columns if c not in existing_cols]
    final_cols = existing_cols + extra_cols

    melted = melted.loc[:, final_cols]

    # ----------------------------------------------------------------------
    # 12) Append to all_data
    # ----------------------------------------------------------------------
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
