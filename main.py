import fitz
from typing import Dict, List, Any, Tuple, Optional


def scale_pdf_page(
    doc: fitz.Document, page_num: int = 0, scale_factor: float = 1.05
) -> Tuple[fitz.Document, fitz.Page]:
    """
    Scale a page from a PDF document by creating a new document with enlarged dimensions
    while keeping the content size unchanged (centered in the new page).

    Args:
        doc: The source PDF document
        page_num: The page number to scale (0-based index)
        scale_factor: The scaling factor (e.g., 1.05 for 5% enlargement)

    Returns:
        Tuple containing the new document and the new page
    """
    page: fitz.Page = doc[page_num]
    original_rect = page.rect
    new_width = original_rect.width * scale_factor
    new_height = original_rect.height * scale_factor

    # Calculate margins to center the content
    margin_x = (new_width - original_rect.width) / 2
    margin_y = (new_height - original_rect.height) / 2

    # Create a new page with enlarged dimensions
    new_document = fitz.open()
    new_page = new_document.new_page(width=new_width, height=new_height)

    # Copy content from original page to new page without scaling but centered
    new_page.show_pdf_page(
        fitz.Rect(
            margin_x,
            margin_y,
            margin_x + original_rect.width,
            margin_y + original_rect.height,
        ),
        doc,
        page_num,
    )

    return new_document, new_page


def find_target_rectangle(page: fitz.Page) -> fitz.Rect:
    """
    Find the first non-white filled rectangle in the page.

    Args:
        page: The PDF page to search in

    Returns:
        The rectangle object if found, raises ValueError if not found
    """
    vector_graphics: List[Dict[str, Any]] = page.get_drawings()
    print(f"Found {len(vector_graphics)} drawing paths on the page")

    for i, draw in enumerate(vector_graphics):
        if "rect" not in draw or draw["rect"] is None:
            continue
        if "fill" not in draw or draw["fill"] is None:
            continue

        # Skip white-filled rectangles
        if draw["fill"][0] > 0.9 and draw["fill"][1] > 0.9 and draw["fill"][2] > 0.9:
            continue

        rect = draw["rect"]
        print(f"  Found rectangle at index {i}: {rect}")
        return rect  # Take the first matching rectangle

    raise ValueError("No target rectangle found on the page")


def add_rectangle_annotation(
    page: fitz.Page,
    rect: fitz.Rect,
    color: Tuple[float, float, float] = (1, 0, 0),
    border_width: float = 1,
) -> None:
    """
    Add a rectangle annotation to a PDF page.

    Args:
        page: The PDF page to add annotation to
        rect: The rectangle coordinates for the annotation
        color: RGB tuple for stroke color (default: red)
        border_width: Width of the border line (default: 1)
    """
    annot = page.add_rect_annot(rect)
    annot.set_colors(stroke=color)
    annot.set_border(width=border_width)
    annot.update()


def save_and_close_documents(
    new_doc: fitz.Document, original_doc: fitz.Document, input_filename: str
) -> str:
    """
    Save the new document and close both documents.

    Args:
        new_doc: The new PDF document to save
        original_doc: The original PDF document to close
        input_filename: The original filename to derive the output filename from

    Returns:
        The path to the saved output file
    """
    output_file = input_filename.replace(".pdf", "_scaled_annotated.pdf")
    new_doc.save(output_file)
    original_doc.close()
    new_doc.close()
    return output_file


def process_pdf(input_file: str, scale_factor: float = 1.05) -> Optional[str]:
    """
    Process a PDF file by:
    1. Scaling the page
    2. Finding a target rectangle
    3. Adding an annotation around the rectangle
    4. Saving the result

    Args:
        input_file: Path to the input PDF file
        scale_factor: How much to scale the page by (default: 1.05 = 5%)

    Returns:
        Path to the output file or None if processing failed
    """
    doc = fitz.open(input_file)
    print(f"Processing PDF: {input_file}")
    print(f"Page count: {doc.page_count}")
    print(f"Page size: {doc[0].rect}")

    print(f"Enlarging page by {(scale_factor - 1) * 100:.0f}%")
    new_document, new_page = scale_pdf_page(doc, page_num=0, scale_factor=scale_factor)

    print("Looking for target rectangle in the new scaled page...")
    try:
        rect = find_target_rectangle(new_page)
        print(f"Target rectangle selected: {rect}")
    except ValueError as e:
        print(str(e))
        doc.close()
        new_document.close()
        return None

    print("Adding annotation to the scaled page...")
    add_rectangle_annotation(new_page, rect)
    print("Annotation added successfully")

    output_file = save_and_close_documents(new_document, doc, input_file)
    print(f"Saved scaled and annotated PDF to: {output_file}")
    return output_file


# Run the process if this script is executed directly
if __name__ == "__main__":
    input_file: str = "test2.pdf"
    print("PyMuPDF version:", fitz.__version__)
    process_pdf(input_file)
