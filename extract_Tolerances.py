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
        (r"(\d+(?:\.\d+)?)\s*(P9|[a-zA-Z]+\d+)\s*\(\s*([+-]?\d*\.?\d+)\s*([+-]?\d*\.?\d+)\s*\)", "Fit Tolerance"),
        
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

    # Normalize row lengths
    columns = ["Type", "Dimension", "Fit/Class", "Upper", "Lower"]
    normalized_data = []
    for row in all_data:
        padded = list(row) + [''] * (5 - len(row))
        normalized_data.append(padded)

    return pd.DataFrame(normalized_data, columns=columns)

def main():
    pdf_path = "disc.pdf"  # Replace with your PDF filename
    output_excel = "extracted_tolerances.xlsx"

    df = extract_all_tolerances_to_df(pdf_path)
    df.to_excel(output_excel, index=False)
    print(f"Tolerances saved to '{output_excel}'")

if __name__ == "__main__":
    main()
