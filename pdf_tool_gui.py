import fitz
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import io
import tempfile
from PDF_Tool import PDF_Tool


class PDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Format Tool")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)

        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.status = tk.StringVar()
        self.status.set("Ready")
        self.current_page = tk.IntVar(value=1)
        self.total_pages = tk.IntVar(value=1)
        self.doc = None
        self.output_doc = None
        self.current_page_obj = None
        self.zoom_level = tk.DoubleVar(value=1.0)
        self.image_tk = None
        self.temp_pdf_path = None  # Path to temporary PDF for preview

        # Viewing options
        self.showing_output = tk.BooleanVar(
            value=False
        )  # Flag to track which PDF is displayed

        # PDF Tool settings
        self.netto_width = tk.DoubleVar(value=100.0)
        self.netto_height = tk.DoubleVar(value=100.0)
        self.additional_margin = tk.DoubleVar(value=5.0)
        self.bleed_size = tk.DoubleVar(value=3.0)
        self.safe_margin_size = tk.DoubleVar(value=4.0)
        self.annotation_width = tk.DoubleVar(value=1.0)

        # Options
        self.add_info_page = tk.BooleanVar(value=True)
        self.add_netto_annotation = tk.BooleanVar(value=True)
        self.add_bleed_annotation = tk.BooleanVar(value=True)
        self.add_safe_margin_annotation = tk.BooleanVar(value=True)

        # Bind window closing event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Create UI
        self.create_menu()
        self.create_toolbar()
        self.create_main_area()
        self.create_status_bar()

        # Start with an empty state
        self.update_status("No file opened")

    def create_menu(self):
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Open PDF", command=self.open_file)
        file_menu.add_command(label="Save Processed PDF", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        menu_bar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def create_toolbar(self):
        toolbar = ttk.Frame(self.root)

        # File selection
        ttk.Label(toolbar, text="Input File:").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Entry(
            toolbar, textvariable=self.input_file, width=40, state="readonly"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Browse...", command=self.open_file).pack(
            side=tk.LEFT, padx=5
        )

        # Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, padx=10, fill=tk.Y
        )

        # View toggle
        self.view_toggle_button = ttk.Button(
            toolbar, text="Show Processed", command=self.toggle_view, width=15
        )
        self.view_toggle_button.pack(side=tk.LEFT, padx=10)

        # Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, padx=10, fill=tk.Y
        )

        # Page navigation
        ttk.Label(toolbar, text="Page:").pack(side=tk.LEFT)
        ttk.Button(toolbar, text="◀", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        ttk.Entry(toolbar, textvariable=self.current_page, width=5).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Label(toolbar, text="/").pack(side=tk.LEFT)
        ttk.Label(toolbar, textvariable=self.total_pages, width=5).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="▶", command=self.next_page).pack(side=tk.LEFT, padx=2)

        # Separator
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
            side=tk.LEFT, padx=10, fill=tk.Y
        )

        # Zoom
        ttk.Label(toolbar, text="Zoom:").pack(side=tk.LEFT)
        ttk.Button(toolbar, text="-", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        ttk.Label(toolbar, textvariable=self.zoom_level).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="+", command=self.zoom_in).pack(side=tk.LEFT, padx=2)

        # Process button
        ttk.Button(toolbar, text="Process PDF", command=self.process_pdf).pack(
            side=tk.RIGHT, padx=5
        )

        toolbar.pack(side=tk.TOP, fill=tk.X, pady=5)

    def create_main_area(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create PDF preview panel (left side)
        self.pdf_frame = ttk.LabelFrame(main_frame, text="PDF Preview - Original")
        self.pdf_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Canvas for PDF display with scrollbars
        canvas_frame = ttk.Frame(self.pdf_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg="lightgray")

        h_scroll = ttk.Scrollbar(
            canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview
        )
        v_scroll = ttk.Scrollbar(
            canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview
        )

        self.canvas.config(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Right panel - Settings and Log
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))

        # Settings panel
        settings_frame = ttk.LabelFrame(right_panel, text="Settings")
        settings_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        # Create grid for settings
        # Row 0: Netto Format
        ttk.Label(settings_frame, text="Netto Format (mm):").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Label(settings_frame, text="Width:").grid(
            row=0, column=1, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=1,
            to=1000,
            increment=0.1,
            textvariable=self.netto_width,
            width=8,
        ).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Label(settings_frame, text="Height:").grid(
            row=0, column=3, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=1,
            to=1000,
            increment=0.1,
            textvariable=self.netto_height,
            width=8,
        ).grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)

        # Row 1: Additional Margin
        ttk.Label(settings_frame, text="Additional Margin (mm):").grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=0,
            to=100,
            increment=0.1,
            textvariable=self.additional_margin,
            width=8,
        ).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Row 2: Bleed Size
        ttk.Label(settings_frame, text="Bleed Size (mm):").grid(
            row=2, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=0,
            to=50,
            increment=0.1,
            textvariable=self.bleed_size,
            width=8,
        ).grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Row 3: Safe Margin
        ttk.Label(settings_frame, text="Safe Margin Size (mm):").grid(
            row=3, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=0,
            to=50,
            increment=0.1,
            textvariable=self.safe_margin_size,
            width=8,
        ).grid(row=3, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Row 4: Annotation Width
        ttk.Label(settings_frame, text="Annotation Width (pt):").grid(
            row=4, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Spinbox(
            settings_frame,
            from_=0.1,
            to=10,
            increment=0.1,
            textvariable=self.annotation_width,
            width=8,
        ).grid(row=4, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Row 5-8: Checkboxes
        ttk.Checkbutton(
            settings_frame, text="Add Info Page", variable=self.add_info_page
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(
            settings_frame,
            text="Add Netto Format Annotation (Orange)",
            variable=self.add_netto_annotation,
        ).grid(row=6, column=0, columnspan=5, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(
            settings_frame,
            text="Add Bleed Size Annotation (Pink)",
            variable=self.add_bleed_annotation,
        ).grid(row=7, column=0, columnspan=5, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(
            settings_frame,
            text="Add Safe Margin Annotation (Green)",
            variable=self.add_safe_margin_annotation,
        ).grid(row=8, column=0, columnspan=5, sticky=tk.W, padx=5, pady=5)

        # Row 9: Output file
        ttk.Label(settings_frame, text="Output File:").grid(
            row=9, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Entry(settings_frame, textvariable=self.output_file, width=30).grid(
            row=9, column=1, columnspan=3, sticky=tk.W + tk.E, padx=5, pady=5
        )
        ttk.Button(
            settings_frame, text="Browse...", command=self.browse_output_file
        ).grid(row=9, column=4, sticky=tk.W, padx=5, pady=5)

        # Log panel
        log_frame = ttk.LabelFrame(right_panel, text="Log")
        log_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, width=40)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)

    def create_status_bar(self):
        status_bar = ttk.Label(
            self.root, textvariable=self.status, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def on_closing(self):
        # Clean up temp files and close documents before quitting
        self.cleanup()
        self.root.destroy()

    def cleanup(self):
        # Close any open documents
        if hasattr(self, "doc") and self.doc is not None:
            try:
                if not self.doc.is_closed:
                    self.doc.close()
            except:
                pass

        if hasattr(self, "output_doc") and self.output_doc is not None:
            try:
                if not self.output_doc.is_closed:
                    self.output_doc.close()
            except:
                pass

        # Delete temporary file if it exists
        if self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
            try:
                os.remove(self.temp_pdf_path)
            except:
                pass

    def toggle_view(self):
        # Toggle between original and processed PDF view
        if not self.output_doc or getattr(self.output_doc, "is_closed", True):
            messagebox.showwarning(
                "Warning", "No processed PDF available. Please process the file first."
            )
            return

        self.showing_output.set(not self.showing_output.get())
        if self.showing_output.get():
            self.view_toggle_button.configure(text="Show Original")
            self.pdf_frame.configure(text="PDF Preview - Processed")
        else:
            self.view_toggle_button.configure(text="Show Processed")
            self.pdf_frame.configure(text="PDF Preview - Original")

        # Refresh the display
        self.display_page(self.current_page.get() - 1)

    def log(self, message: str) -> None:
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        print(message)

    def update_status(self, message: str) -> None:
        self.status.set(message)
        self.log(message)

    def open_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Open PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )

        if not file_path:
            return

        try:
            # Close previous document if open
            if hasattr(self, "doc") and self.doc is not None:
                try:
                    if not self.doc.is_closed:
                        self.doc.close()
                except:
                    pass

            if hasattr(self, "output_doc") and self.output_doc is not None:
                try:
                    if not self.output_doc.is_closed:
                        self.output_doc.close()
                except:
                    pass
                self.output_doc = None

            self.input_file.set(file_path)
            self.doc = fitz.open(file_path)
            self.total_pages.set(self.doc.page_count)
            self.current_page.set(1)

            # Reset view to original
            self.showing_output.set(False)
            self.view_toggle_button.configure(text="Show Processed")
            self.pdf_frame.configure(text="PDF Preview - Original")

            # Auto-set output file name
            output_path = file_path.replace(".pdf", "_processed.pdf")
            self.output_file.set(output_path)

            self.update_status(f"Opened: {os.path.basename(file_path)}")
            self.log(f"Page count: {self.doc.page_count}")

            # Display first page
            self.display_page(0)

        except Exception as e:
            self.update_status(f"Error opening file: {str(e)}")
            messagebox.showerror("Error", f"Failed to open PDF: {str(e)}")

    def browse_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Save Processed PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialfile=os.path.basename(self.output_file.get())
            if self.output_file.get()
            else "output.pdf",
        )

        if file_path:
            self.output_file.set(file_path)
            self.log(f"Output file set to: {file_path}")

    def save_file(self) -> None:
        if not self.output_doc or getattr(self.output_doc, "is_closed", True):
            messagebox.showwarning(
                "Warning", "No processed PDF available. Please process the file first."
            )
            return

        output_path = self.output_file.get()
        if not output_path:
            self.browse_output_file()
            output_path = self.output_file.get()
            if not output_path:  # User canceled
                return

        try:
            self.output_doc.save(output_path)
            self.update_status(
                f"Saved processed PDF to: {os.path.basename(output_path)}"
            )
            messagebox.showinfo(
                "Success", f"File saved successfully to:\n{output_path}"
            )
        except Exception as e:
            self.update_status(f"Error saving file: {str(e)}")
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def is_document_valid(self, doc):
        """Check if document is valid and usable"""
        if doc is None:
            return False

        try:
            if getattr(doc, "is_closed", True):
                return False
            # Test if we can access page count - will raise error if document is closed
            _ = doc.page_count
            return True
        except:
            return False

    def display_page(self, page_index: int) -> None:
        # Determine which document to display
        if self.showing_output.get() and self.is_document_valid(self.output_doc):
            doc_to_display = self.output_doc
        elif self.is_document_valid(self.doc):
            doc_to_display = self.doc
        else:
            return  # No valid document to display

        if page_index < 0 or page_index >= doc_to_display.page_count:
            return

        # Update current page
        self.current_page.set(page_index + 1)
        self.current_page_obj = doc_to_display[page_index]
        self.total_pages.set(doc_to_display.page_count)

        try:
            # Get page as image
            zoom = self.zoom_level.get()
            pix = self.current_page_obj.get_pixmap(matrix=fitz.Matrix(zoom, zoom))

            # Convert to PIL image
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            self.image_tk = ImageTk.PhotoImage(image=img)

            # Update canvas
            self.canvas.delete("all")
            self.canvas.config(scrollregion=(0, 0, img.width, img.height))
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.image_tk)

            doc_type = "processed" if self.showing_output.get() else "original"
            self.update_status(
                f"Displaying {doc_type} PDF - page {page_index + 1} of {doc_to_display.page_count}"
            )
        except Exception as e:
            self.update_status(f"Error displaying page: {str(e)}")

    def next_page(self) -> None:
        if self.showing_output.get() and self.is_document_valid(self.output_doc):
            doc_to_display = self.output_doc
        elif self.is_document_valid(self.doc):
            doc_to_display = self.doc
        else:
            return  # No valid document to display

        current = self.current_page.get() - 1
        if current < doc_to_display.page_count - 1:
            self.display_page(current + 1)

    def prev_page(self) -> None:
        if self.showing_output.get() and self.is_document_valid(self.output_doc):
            doc_to_display = self.output_doc
        elif self.is_document_valid(self.doc):
            doc_to_display = self.doc
        else:
            return  # No valid document to display

        current = self.current_page.get() - 1
        if current > 0:
            self.display_page(current - 1)

    def zoom_in(self) -> None:
        new_zoom = min(3.0, self.zoom_level.get() + 0.1)
        self.zoom_level.set(round(new_zoom, 1))
        self.display_page(self.current_page.get() - 1)

    def zoom_out(self) -> None:
        new_zoom = max(0.1, self.zoom_level.get() - 0.1)
        self.zoom_level.set(round(new_zoom, 1))
        self.display_page(self.current_page.get() - 1)

    def process_pdf(self) -> None:
        if not self.is_document_valid(self.doc):
            messagebox.showwarning("Warning", "No PDF file is open")
            return

        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input PDF file first")
            return

        try:
            # Close previous output document if it exists
            if hasattr(self, "output_doc") and self.output_doc is not None:
                try:
                    if not self.output_doc.is_closed:
                        self.output_doc.close()
                except:
                    pass
                self.output_doc = None

            # Clean up any existing temp file
            if self.temp_pdf_path and os.path.exists(self.temp_pdf_path):
                try:
                    os.remove(self.temp_pdf_path)
                except:
                    pass

            # Create a PDF tool with the specified settings
            tool = PDF_Tool()

            # Configure the tool
            tool.setNettoFormat(self.netto_width.get(), self.netto_height.get())
            tool.setAdditionalMargin(self.additional_margin.get())
            tool.setBleedSize(self.bleed_size.get())
            tool.setSafeMarginSize(self.safe_margin_size.get())
            tool.setAnnotationWidth(self.annotation_width.get())

            # Load the PDF
            self.log("Loading PDF...")
            tool.loadPDF(self.input_file.get())

            # Process with selected options
            self.log("Adding pages with margin...")
            tool.addPagesWithMargin()

            if self.add_netto_annotation.get():
                self.log("Adding netto format annotation...")
                tool.addNettoFormatAnnotation()

            if self.add_bleed_annotation.get():
                self.log("Adding bleed size annotation...")
                tool.addBleedSizeAnnotation()

            if self.add_safe_margin_annotation.get():
                self.log("Adding safe margin annotation...")
                tool.addSafeMarginSizeAnnotation()

            if self.add_info_page.get():
                self.log("Adding info page...")
                tool.addInfoPage()

            # Create a temporary file for the processed PDF
            fd, self.temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)  # Close the file descriptor but keep the temp file

            # Save the tool's output to the temp file
            tool.output_pdf.save(self.temp_pdf_path)

            # Clean up the PDF_Tool
            tool.close()

            # Open the saved temporary copy
            self.output_doc = fitz.open(self.temp_pdf_path)

            self.update_status("PDF processed successfully")

            # Automatically switch to showing the processed PDF
            self.showing_output.set(True)
            self.view_toggle_button.configure(text="Show Original")
            self.pdf_frame.configure(text="PDF Preview - Processed")
            self.display_page(0)  # Show the first page of the processed PDF

            # Ask if user wants to save right away
            if messagebox.askyesno(
                "Save File", "PDF processed successfully. Do you want to save it now?"
            ):
                self.save_file()

        except Exception as e:
            self.update_status(f"Error processing PDF: {str(e)}")
            messagebox.showerror("Error", f"Failed to process PDF: {str(e)}")

    def show_about(self) -> None:
        about_text = """PDF Format Tool

Version: 1.0
Python version: {py_ver}
PyMuPDF version: {mupdf_ver}

A tool for adding format annotations to PDF documents."""

        messagebox.showinfo(
            "About",
            about_text.format(
                py_ver=sys.version.split()[0], mupdf_ver=fitz.__version__
            ),
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolGUI(root)
    root.mainloop()
