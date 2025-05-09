import fitz
import sys
from typing import Dict, List, Any


print("Python version:", sys.version)
print("PyMuPDF version:", fitz.__version__)

input_file: str = "test.pdf"
print(f"Processing PDF: {input_file}")

doc: fitz.Document = fitz.open(input_file)
page: fitz.Page = doc[0]
print(f"Page count: {doc.page_count}")
print(f"Page size: {page.rect}")

vector_graphics: List[Dict[str, Any]] = page.get_drawings()
print(f"Found {len(vector_graphics)} drawing paths")

rect: fitz.Rect = None

for i, draw in enumerate(vector_graphics):
    if "rect" not in draw or draw["rect"] is None:
        continue
    if "fill" not in draw or draw["fill"] is None:
        continue

    # not really needed, but just in case
    if draw["fill"][0] > 0.9 and draw["fill"][1] > 0.9 and draw["fill"][2] > 0.9:
        continue
    # -----------------------------------

    rect = draw["rect"]

if rect is None:
    print("No target rectangle found")
    exit(1)

print(f"  Found rect: {rect}")

annot = page.add_rect_annot(rect)
annot.set_colors(stroke=(1, 0, 0))
annot.set_border(width=1)
annot.update()

output_file = input_file.replace(".pdf", "_annotated.pdf")
doc.save(output_file)
print(f"Saved annotated PDF to: {output_file}")
doc.close()
