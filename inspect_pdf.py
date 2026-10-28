import fitz # PyMuPDF
import os

pdf_path = r"c:\Users\Kalai\Downloads\XSLX TO PDF\PDF\731-418_2026-05-27_13-06-22.pdf"

if not os.path.exists(pdf_path):
    print("PDF not found!")
    exit(1)

doc = fitz.open(pdf_path)
print(f"Number of pages: {len(doc)}")
print(f"Metadata: {doc.metadata}")

# Let's check the font list of the first page
print("\nFonts on Page 1:")
for f in doc[0].get_fonts():
    print(f"  {f}")

# Let's print the text of each page to see what sections exist
for i, page in enumerate(doc):
    print(f"\n--- PAGE {i+1} TEXT ---")
    text = page.get_text()
    lines = text.split("\n")
    # Print first 20 lines
    for line in lines[:30]:
        print(f"  {line.strip()}")
    if len(lines) > 30:
        print(f"  ... (+ {len(lines)-30} lines)")
