import fitz  # PyMuPDF
import re
import pandas as pd

def extract_all_tolerances_to_df(pdf_path):
    doc = fitz.open(pdf_path)
    all_data = []

    # Define patterns for various tolerance types
    tolerance_patterns = [
        # Format: ⌀83 h7( 0 -0.04 )
        (r"(⌀\d+(?:\.\d+)?)[ ]*([a-zA-Z]+\d+)?\s*\(\s*([+-]?\d*\.?\d+)\s*([+-]?\d*\.?\d+)\s*\)", "Fit Tolerance"),
        
        # Format: 16 P9(-0.02 -0.06)
        (r"(\d+(?:\.\d+)?)\s*([a-zA-Z]+\d+)\s*\(\s*([+-]?\d*\.?\d+)\s*([+-]?\d*\.?\d+)\s*\)", "Fit Tolerance"),
        
        # Format: ⌀85±0.05
        (r"(⌀\d+(?:\.\d+)?)±(\d*\.?\d+)", "Symmetric Tolerance"),
        
        # Format: 85 +0.05 -0.01 or 100 +0.05 +0.01
        (r"(\d+(?:\.\d+)?)\s*\+(\d*\.?\d+)\s*([+-]\d*\.?\d+)", "Asymmetric Tolerance"),
        
        # Format: ⌀75 g6(-0.01 -0.03)
        (r"(⌀\d+(?:\.\d+)?)\s*([a-zA-Z]+\d+)\s*\(\s*([+-]?\d*\.?\d+)\s*([+-]?\d*\.?\d+)\s*\)", "Fit Tolerance")
    ]

    for page in doc:
        text = page.get_text()
        for pattern, tol_type in tolerance_patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                all_data.append((tol_type,) + match)

    # Normalize rows
    columns = ["Type", "Raw Dimension", "Fit/Class", "Upper Tol", "Lower Tol"]
    normalized_data = []
    for row in all_data:
        padded = list(row) + [''] * (5 - len(row))
        normalized_data.append(padded)

    df = pd.DataFrame(normalized_data, columns=columns)

    # Serial number assignment based on order of appearance
    df.insert(0, "Serial No", range(1, len(df) + 1))

    # Parse nominal and identify dimension type
    nominal_vals = []
    dim_types = []
    lower_bounds = []
    upper_bounds = []

    for idx, row in df.iterrows():
        dim_raw = row["Raw Dimension"]
        upper = row["Upper Tol"]
        lower = row["Lower Tol"]

        # Extract numeric part and type
        if dim_raw.startswith("⌀"):
            dim_type = "Diameter"
            raw_num = dim_raw[1:]
        elif dim_raw.startswith("R"):
            dim_type = "Radius"
            raw_num = dim_raw[1:]
        else:
            dim_type = "Linear"
            raw_num = dim_raw

        try:
            nominal = float(raw_num)
        except:
            nominal = 0.0  # Default if parse fails

        try:
            lower_tol = float(lower)
        except:
            lower_tol = 0.0
        try:
            upper_tol = float(upper)
        except:
            upper_tol = 0.0

        lower_bound = nominal + lower_tol
        upper_bound = nominal + upper_tol

        nominal_vals.append(nominal)
        dim_types.append(dim_type)
        lower_bounds.append(round(lower_bound, 3))
        upper_bounds.append(round(upper_bound, 3))

    # Add calculated columns
    df["Nominal"] = nominal_vals
    df["Dimension Type"] = dim_types
    df["Lower Deviation"] = df["Lower Tol"].replace('', 0.0).astype(float)
    df["Upper Deviation"] = df["Upper Tol"].replace('', 0.0).astype(float)
    df["Min Limit"] = lower_bounds
    df["Max Limit"] = upper_bounds

    # Drop duplicates based on key tolerance fields
    df = df.drop_duplicates(subset=["Raw Dimension", "Fit/Class", "Upper Tol", "Lower Tol"])

    # Keep only necessary columns
    df = df[[
        "Serial No", "Nominal", "Type", "Fit/Class",
        "Upper Deviation", "Lower Deviation", "Max Limit", "Min Limit"
    ]]

    return df

def main():
    pdf_path = "D:\\QC\\Drawings\\shaftdraw.pdf"  # <-- Update as needed
    output_excel = "extracted_tolerances_calculated.xlsx"

    df = extract_all_tolerances_to_df(pdf_path)
    df.to_excel(output_excel, index=False)
    print(f" Tolerances with bounds saved to '{output_excel}'")

if __name__ == "__main__":
    main()
