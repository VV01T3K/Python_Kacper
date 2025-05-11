from PDF_Tool import PDF_Tool

input_pdf_path = "Example\\tests\\test2.pdf"
output_pdf_path = "Example\\tests\\output.pdf"

(
    PDF_Tool()
    .setNettoFormat(100, 100)
    .setAdditionalMargin(5)
    .setBleedSize(3)
    .setSafeMarginSize(4)
    .setAnnotationWidth(1)
    .loadPDF(input_pdf_path)
    .addPages()
    .addPagesWithMargin()
    .addNettoFormatAnnotation()
    .addBleedSizeAnnotation()
    .addSafeMarginSizeAnnotation()
    .addInfoPage()
    .savePDF(output_pdf_path)
)

# tool = (
#     PDF_Tool()
#     .setNettoFormat(100, 100)
#     .setAdditionalMargin(5)
#     .setBleedSize(3)
#     .setSafeMarginSize(4)
#     .setAnnotationWidth(1)
# )
# tool.fullProcess(input_pdf_path, output_pdf_path).close()
