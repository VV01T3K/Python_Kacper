import os
import sys
import fitz  # PyMuPDF
from typing import Tuple
from enum import Enum


def get_resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class PDF_Tool:
    class Color(Enum):
        RED = (1, 0, 0)
        GREEN = (0, 1, 0)
        ORANGE = (1, 0.5, 0)
        PINK = (1, 0, 1)

    def __init__(self):
        self.original_pdf: fitz.Document
        self.output_pdf: fitz.Document
        self.netto_format_pts: Tuple[float, float]
        self.bleed_size_pts: float
        self.safe_margin_size_pts: float
        self.additional_margin_pts: float
        self.info_page_added: bool = False
        self.annotation_width: float = 1

    def __del__(self):
        self.close()

    def setNettoFormat(self, width: float, height: float) -> "PDF_Tool":
        self.netto_format = (
            width,
            height,
        )
        return self

    def setBleedSize(self, bleed_size: float) -> "PDF_Tool":
        self.bleed_size = bleed_size
        return self

    def setSafeMarginSize(self, safe_margin_size: float) -> "PDF_Tool":
        self.safe_margin_size = safe_margin_size
        return self

    def setAdditionalMargin(self, margin: float) -> "PDF_Tool":
        self.additional_margin_pts = self.convertMilimetersToPoints(margin)
        return self

    def setAnnotationWidth(self, annotation_width: float) -> "PDF_Tool":
        self.annotation_width = annotation_width
        return self

    def loadPDF(self, input_pdf_path: str) -> "PDF_Tool":
        self.original_pdf = fitz.open(input_pdf_path)
        self.output_pdf = fitz.open()
        print(f"Loaded PDF: {input_pdf_path}")
        print("Output PDF initialized.")
        return self

    def addInfoPage(self, info_page_path: str = "raport.pdf") -> "PDF_Tool":
        # Load the info page PDF
        info_page_path = get_resource_path(info_page_path)
        info_pdf = fitz.open(info_page_path)

        # Insert all pages from the info PDF at the beginning
        for i in range(info_pdf.page_count):
            info_page = info_pdf[i]
            # Create a new page with the same dimensions
            new_page = self.output_pdf.new_page(
                -1,  # Insert at the end initially
                width=info_page.rect.width,
                height=info_page.rect.height,
            )
            # Copy content from the info page
            new_page.show_pdf_page(new_page.rect, info_pdf, i)

        # Move all info pages to the beginning
        page_count = self.output_pdf.page_count
        info_page_count = info_pdf.page_count

        # Move each info page to the beginning (in reverse to maintain order)
        for i in range(info_page_count):
            self.output_pdf.move_page(page_count - 1 - i, 0)

        info_pdf.close()
        self.info_page_added = True

        return self

    def convertMilimetersToPoints(self, value: float) -> float:
        return value * 2.83

    def addPages(self) -> "PDF_Tool":
        for page in self.original_pdf:
            new_page = self.output_pdf.new_page(
                width=page.rect.width, height=page.rect.height
            )
            new_page.show_pdf_page(page.rect, self.original_pdf)
        return self

    def addPagesWithMargin(self) -> "PDF_Tool":
        for page in self.original_pdf:
            original_rect = page.rect
            new_width = original_rect.width + (2 * self.additional_margin_pts)
            new_height = original_rect.height + (2 * self.additional_margin_pts)

            new_page = self.output_pdf.new_page(width=new_width, height=new_height)

            new_page.show_pdf_page(
                fitz.Rect(
                    self.additional_margin_pts,
                    self.additional_margin_pts,
                    self.additional_margin_pts + original_rect.width,
                    self.additional_margin_pts + original_rect.height,
                ),
                self.original_pdf,
                page.number,
            )

        return self

    def __addRectAnnotation(self, width, height, color: Color) -> "PDF_Tool":
        react_size = (
            self.convertMilimetersToPoints(width),
            self.convertMilimetersToPoints(height),
        )

        for i, page in enumerate(self.output_pdf):
            if i == 0 and self.info_page_added:
                continue

            rect = fitz.Rect(
                (page.rect.width - react_size[0]) / 2,
                (page.rect.height - react_size[1]) / 2,
                (page.rect.width + react_size[0]) / 2,
                (page.rect.height + react_size[1]) / 2,
            )
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=color.value)
            annot.set_border(width=self.annotation_width)
            annot.update()

        return self

    def addNettoFormatAnnotation(self) -> "PDF_Tool":
        return self.__addRectAnnotation(
            self.netto_format[0],
            self.netto_format[1],
            self.Color.ORANGE,
        )

    def addBleedSizeAnnotation(self) -> "PDF_Tool":
        return self.__addRectAnnotation(
            self.netto_format[0] + (2 * self.bleed_size),
            self.netto_format[1] + (2 * self.bleed_size),
            self.Color.PINK,
        )

    def addSafeMarginSizeAnnotation(self) -> "PDF_Tool":
        return self.__addRectAnnotation(
            self.netto_format[0] - (2 * self.safe_margin_size),
            self.netto_format[1] - (2 * self.safe_margin_size),
            self.Color.GREEN,
        )

    def savePDF(self, output_pdf_path: str) -> "PDF_Tool":
        self.output_pdf.save(output_pdf_path)
        return self

    def fullProcess(self, input_pdf_path: str, output_pdf_path: str) -> "PDF_Tool":
        return (
            self.loadPDF(input_pdf_path)
            .addInfoPage()
            .addPagesWithMargin()
            .addNettoFormatAnnotation()
            .addBleedSizeAnnotation()
            .addSafeMarginSizeAnnotation()
            .savePDF(output_pdf_path)
        )

    def close(self, keep_output=False) -> None:
        closed_any = False
        if hasattr(self, "original_pdf") and not self.original_pdf.is_closed:
            self.original_pdf.close()
            closed_any = True
        if not keep_output:
            if hasattr(self, "output_pdf") and not self.output_pdf.is_closed:
                self.output_pdf.close()
                closed_any = True
        if closed_any:
            print("PDF files closed.")
