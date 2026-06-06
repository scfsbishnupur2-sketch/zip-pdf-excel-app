import pdfplumber
import pandas as pd
import re

pdf_file = input("Enter PDF path: ")

rows = []

with pdfplumber.open(pdf_file) as pdf:

    for page in pdf.pages:

        text = page.extract_text()

        if not text:
            continue

        lines = text.split("\n")

        for line in lines:

            m = re.match(
                r'^(\d+)\s+(\d+)\s+([A-Z0-9\/]+)\s+(.*?)\s+([A-Z ]+)\s+\((Father|Husband|Mother)\)\s+(Death|Permanently Shifted|Already enrolled)',
                line.strip()
            )

            if m:

                rows.append({
                    "S.No": m.group(1),
                    "Serial No": m.group(2),
                    "EPIC Number": m.group(3),
                    "Elector Name": m.group(4).strip(),
                    "Relative Details": m.group(5).strip(),
                    "Relation Type": m.group(6),
                    "Uncollectable Reason": m.group(7)
                })

df = pd.DataFrame(rows)

df.to_excel(
    "test_output.xlsx",
    index=False
)

print("Done!")
print("Rows Extracted:", len(df))