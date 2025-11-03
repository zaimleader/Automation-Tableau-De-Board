# build_filter.py
import sys
import os
import json
import math
import pandas as pd

# === File paths ===
main_path = "C:\\Users\\pc\\Desktop\\Documents\\Programming\\Automation Tableau De Board"
excel_name = "TDB_CP_BPILOT_V4 (12).xlsm"

excel_path = f"{main_path}\\excel\\{excel_name}"
data_json_path = f"{main_path}\\main\\result\\data.json"
filter_json_path = f"{main_path}\\main\\result\\filter.json"

def transform_multivalore(text):
    """Apply all text replacements."""
    replacements = {
        "(": "_",
        ")": " =",
        "+": " Y",
        "-": " N",
        ",": " AND "
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def main(column_name, reference_value, sheet_name):
    # Read the sheet into pandas (no header)
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    except Exception as e:
        print("Error reading Excel sheet:", e)
        sys.exit(1)

    # --- Step 1: filter Elemento figlio column (index 4, column E) by reference_value (substring match) ---
    # Values in Elemento figlio begin/end with 00, so we search for the reference_value inside.
    # Ensure we cast to str
    col_elemento_idx = 4    # pandas column index for Elemento figlio (column E)
    col_multivalore_idx = 15 # pandas column index for Caratteristiche Multivalore (column P)

    # Protect against index errors
    if col_elemento_idx not in df.columns or col_multivalore_idx not in df.columns:
        # if columns aren't present as expected, still try to access by position
        maxcol = df.shape[1]

        if col_elemento_idx >= maxcol or col_multivalore_idx >= maxcol:
            print("Sheet does not have expected columns (Elemento figlio or Caratteristiche Multivalore).")
            sys.exit(1)
    
    # print(df[col_elemento_idx])

    mask = df[col_elemento_idx].astype(str).str.contains(str(reference_value), na=False)
    filtered = df[mask]

    # --- Step 2: Build multivalore string ---
    # take values from column index 15 (col_multivalore_idx)
    multivalore_values = filtered[col_multivalore_idx].dropna().astype(str).str.strip().tolist()

    # Join values with " OU " if multiple
    if len(multivalore_values) > 1:
        multivalore = " OU ".join(multivalore_values)
    elif len(multivalore_values) == 1:
        multivalore = multivalore_values[0]
    else:
        multivalore = ""

    # Transform multivalore text
    new_multivalore = transform_multivalore(multivalore)

    # Reverse the order of multivalore values and retry
    vals = multivalore_values[::-1]
    reversed_multivalore = " OU ".join(vals)
    reversed_multivalore = transform_multivalore(reversed_multivalore)

     # Read the sheet into pandas (no header)
    try:
        df_main = pd.read_excel(excel_path, sheet_name="TdeB_OFFI", header=None, engine='openpyxl')
    except Exception as e:
        print("Error reading Excel sheet:", e)
        sys.exit(1)

    # Define column indices
    col_map = {
        "Reference NFC 'FIAT'": 15,  # P
        "HW 'FIAT'": 16,             # Q
        "SW 'FIAT'": 17              # R
    }

    if column_name not in col_map:
        print(f"Invalid column name: {column_name}")
        sys.exit(1)

    target_idx = col_map[column_name]
    pa_usecase_idx = 29  # AD

    # Lists to store results
    list_pa_usecase = []
    list_cels_color = []

    # Filter rows where the reference column matches the value
    matches_idx = []
    col_series = df_main.iloc[:, target_idx].astype(str).str.strip()
    ref = str(reference_value).strip()

    for i, cell in enumerate(col_series):
        if cell == ref:
            matches_idx.append(i)

    matching_rows = df_main.iloc[matches_idx]

    # If no matches
    if matching_rows.empty:
        print("No rows found with that reference value.")
    else:
        # --- Step 5: Compare multivalore and pa_usecase (string comparison) --- # Comparison logic
        for _, row in matching_rows.iterrows():
            pa_usecase = row.iloc[pa_usecase_idx]

            if isinstance(pa_usecase, float) and math.isnan(pa_usecase):
                cel_color = False
            else:
                pa_usecase = str(pa_usecase).strip()
                if pa_usecase == new_multivalore:
                    cel_color = True
                else:
                    cel_color = (pa_usecase == reversed_multivalore)

            list_pa_usecase.append(pa_usecase)
            list_cels_color.append(cel_color)

    # --- Step 6: Write filter.json ---
    filter_data = {
        "column_name": str(column_name),
        "reference_value": str(reference_value),
        "sheet_name": str(sheet_name),
        "multivalore": str(new_multivalore),
        "reversed_multivalore": str(reversed_multivalore),
        "list_pa_usecase": list(list_pa_usecase) if list_pa_usecase else [],
        "list_matches_idx": list(matches_idx) if matches_idx else [],
        "list_cels_color": list(list_cels_color) if list_cels_color else []
    }

    # Ensure output directory exists
    out_dir = os.path.dirname(filter_json_path)
    os.makedirs(out_dir, exist_ok=True)

    with open(filter_json_path, "w", encoding="utf-8") as f:
        json.dump(filter_data, f, indent=4, ensure_ascii=False)

    print("filter.json written successfully:")

if __name__ == "__main__":
    if len(sys.argv) >= 4:
        column_name = sys.argv[1]
        reference_value = sys.argv[2]
        sheet_name = sys.argv[3]

        main(column_name, reference_value, sheet_name)
    else:
        print("Usage: python comparaison.py <column_name> <reference_value> <sheet_name>")