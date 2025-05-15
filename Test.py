import fitz  # PyMuPDF
import re
import pandas as pd

# Input/output files
pdf_path = "disc.pdf"
output_pdf_path = "GeneratedOutput(PDF).pdf"
output_excel_path = "GeneratedOutput(XLS).xlsx"
doc = fitz.open(pdf_path)

# Regex patterns
dimension_pattern = re.compile(r'(⌀\d+(\.\d+)?|R\d+(\.\d+)?|\d+\.\d+|\b\d+\b)')
tolerance_pattern = re.compile(r'(±\d+(\.\d+)?|[+−-]\d+(\.\d+)?\s*[+−-]\d+(\.\d+)?)')
fit_pattern = re.compile(r'([A-Za-z]\d+)\s*\(\s*[+−-]?\d+(\.\d+)?\s*[+−-]?\d+(\.\d+)?\s*\)')

# Collect all text blocks from all pages
all_blocks = []
for page_num, page in enumerate(doc, start=1):
    blocks = page.get_text("blocks")
    for block in blocks:
        text = block[4].strip().replace("−", "-")  # Normalize minus sign
        rect = fitz.Rect(block[:4])
        all_blocks.append((page_num, rect, text))

# Process dimensions and look for tolerances
data = []



for i, (page_num, rect, text) in enumerate(all_blocks):
    dim_matches = dimension_pattern.findall(text)
    if not dim_matches:
        continue

    dim = dim_matches[0][0]
    symbol = ''
    value = ''

    if dim.startswith('R'):
        symbol = 'R'
        value = dim[1:]
    elif dim.startswith('⌀'):
        symbol = '⌀'
        value = dim[1:]
    else:
        value = dim

    try:
        nominal = float(value)
    except ValueError:
        continue  # skip invalid dimension values

    low_tol = ''
    up_tol = ''
    tolerance_type = "No Tolerance"

    # Try to find nearby tolerance information
    for j, (p2, r2, t2) in enumerate(all_blocks):
        if j == i or p2 != page_num:
            continue

        if abs(r2.y0 - rect.y1) < 15 and r2.intersects(rect + (0, 0, 200, 20)):
            t2 = t2.replace("−", "-")

            # ± tolerance
            if "±" in t2:
                try:
                    tol_val = float(t2.replace('±', '').strip())
                    low_tol = up_tol = tol_val
                    tolerance_type = "Symmetric (±)"
                    break
                except:
                    continue

            # + / - tolerance
            parts = re.findall(r'[+-]\d+(\.\d+)?', t2)
            if len(parts) == 2:
                try:
                    up_tol = float(parts[0])
                    low_tol = float(parts[1])
                    tolerance_type = "Asymmetric (+/-)"
                    break
                except:
                    continue

            # Fit notation like h7(-0.04) or P9(-0.02 -0.06)
            fit_match = fit_pattern.search(t2)
            if fit_match:
                tolerance_type = f"Fit ({fit_match.group(1)})"
                low_tol = up_tol = 0
                break

    # Calculate bounds
    if tolerance_type == "Symmetric (±)" and low_tol != '':
        lower_bound = nominal - float(low_tol)
        upper_bound = nominal + float(up_tol)
    elif tolerance_type == "Asymmetric (+/-)" and low_tol != '' and up_tol != '':
        lower_bound = nominal + float(low_tol)
        upper_bound = nominal + float(up_tol)
    else:
        lower_bound = nominal
        upper_bound = nominal

    data.append({
        "Serial Number": len(data) + 1,
        "Symbol": symbol,
        "Value": nominal,
        "Lower Tol": low_tol,
        "Upper Tol": up_tol,
        "Lower Bound": round(lower_bound, 3),
        "Upper Bound": round(upper_bound, 3),
        "Tolerance Type": tolerance_type,
        "Page": page_num,
        "Rect": rect
    })

# Annotate PDF with balloon numbers
for item in data:
    page = doc[item["Page"] - 1]
    rect = item["Rect"]
    number_text = f"#{item['Serial Number']}"
    font_size = 18
    char_width = 0.5 * font_size
    text_width = len(number_text) * char_width
    center_x = rect.x0 + (rect.width / 2) - (text_width / 2)
    below_y = rect.y1 + 2
    insert_point = fitz.Point(center_x, below_y)
    page.insert_text(
        insert_point,
        number_text,
        fontname="helv",
        fontsize=font_size,
        color=(1, 0, 0),  # red
    )

doc.save(output_pdf_path)
print(f" Annotated PDF saved to: {output_pdf_path}")

# Save Excel
df = pd.DataFrame(data)
df = df.drop(columns=["Page", "Rect"])
df.to_excel(output_excel_path, index=False)
print(f" Excel sheet saved to: {output_excel_path}")
