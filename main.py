import fitz
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import io


class PDFAnnotator:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Annotator")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        # Variables
        self.current_file = tk.StringVar()
        self.status = tk.StringVar()
        self.status.set("Ready")
        self.current_page = tk.IntVar(value=1)
        self.total_pages = tk.IntVar(value=1)
        self.doc = None
        self.current_page_obj = None
        self.zoom_level = tk.DoubleVar(value=1.0)
        self.image_tk = None

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
        file_menu.add_command(label="Save Annotated PDF", command=self.save_file)
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
        ttk.Label(toolbar, text="File:").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Entry(
            toolbar, textvariable=self.current_file, width=40, state="readonly"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Browse...", command=self.open_file).pack(
            side=tk.LEFT, padx=5
        )

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
        ttk.Button(
            toolbar, text="Find & Annotate Rectangle", command=self.process_current_page
        ).pack(side=tk.RIGHT, padx=5)

        toolbar.pack(side=tk.TOP, fill=tk.X, pady=5)

    def create_main_area(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Left panel - PDF display
        pdf_frame = ttk.LabelFrame(main_frame, text="PDF Preview")
        pdf_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Canvas for PDF display with scrollbars
        canvas_frame = ttk.Frame(pdf_frame)
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
        self.canvas.pack(
            side=tk.LEFT, fill=tk.BOTH, expand=True
        )  # Right panel - Log display
        log_frame = ttk.LabelFrame(main_frame, text="Log")
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))
        # Set a minimum width for the log frame
        log_frame.columnconfigure(0, minsize=300)

        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, width=40)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)

    def create_status_bar(self):
        status_bar = ttk.Label(
            self.root, textvariable=self.status, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

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
            if self.doc:
                self.doc.close()

            self.current_file.set(file_path)
            self.doc = fitz.open(file_path)
            self.total_pages.set(self.doc.page_count)
            self.current_page.set(1)

            self.update_status(f"Opened: {os.path.basename(file_path)}")
            self.log(f"Page count: {self.doc.page_count}")

            # Display first page
            self.display_page(0)

        except Exception as e:
            self.update_status(f"Error opening file: {str(e)}")
            messagebox.showerror("Error", f"Failed to open PDF: {str(e)}")

    def save_file(self) -> None:
        if not self.doc:
            messagebox.showwarning("Warning", "No file is currently open")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save Annotated PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialfile=os.path.basename(self.current_file.get()).replace(
                ".pdf", "_annotated.pdf"
            ),
        )

        if not file_path:
            return

        try:
            self.doc.save(file_path)
            self.update_status(f"Saved annotated PDF to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"File saved successfully to:\n{file_path}")

        except Exception as e:
            self.update_status(f"Error saving file: {str(e)}")
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def display_page(self, page_index: int) -> None:
        if not self.doc or page_index < 0 or page_index >= self.doc.page_count:
            return

        # Update current page
        self.current_page.set(page_index + 1)
        self.current_page_obj = self.doc[page_index]

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

        self.update_status(f"Displaying page {page_index + 1} of {self.doc.page_count}")

    def next_page(self) -> None:
        if not self.doc:
            return
        current = self.current_page.get() - 1
        if current < self.doc.page_count - 1:
            self.display_page(current + 1)

    def prev_page(self) -> None:
        if not self.doc:
            return
        current = self.current_page.get() - 1
        if current > 0:
            self.display_page(current - 1)

    def zoom_in(self) -> None:
        new_zoom = min(2.0, self.zoom_level.get() + 0.1)
        self.zoom_level.set(round(new_zoom, 1))
        if self.doc:
            self.display_page(self.current_page.get() - 1)

    def zoom_out(self) -> None:
        new_zoom = max(0.5, self.zoom_level.get() - 0.1)
        self.zoom_level.set(round(new_zoom, 1))
        if self.doc:
            self.display_page(self.current_page.get() - 1)

    def process_current_page(self) -> None:
        if not self.doc or not self.current_page_obj:
            messagebox.showwarning("Warning", "No PDF file is open")
            return

        try:
            page = self.current_page_obj
            vector_graphics = page.get_drawings()
            self.log(f"Found {len(vector_graphics)} drawing paths")

            rect = None

            for i, draw in enumerate(vector_graphics):
                if "rect" not in draw or draw["rect"] is None:
                    continue
                if "fill" not in draw or draw["fill"] is None:
                    continue

                # Skip white rectangles (not really needed, but just in case)
                if (
                    draw["fill"][0] > 0.9
                    and draw["fill"][1] > 0.9
                    and draw["fill"][2] > 0.9
                ):
                    continue

                rect = draw["rect"]
                self.log(f"Found rectangle: {rect}")
                break

            if rect is None:
                self.update_status("No target rectangle found on this page")
                messagebox.showinfo("Results", "No target rectangle found on this page")
                return

            # Add annotation
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=(1, 0, 0))  # Red stroke
            annot.set_border(width=1)
            annot.update()

            # Refresh display to show annotation
            self.display_page(self.current_page.get() - 1)

            self.update_status(
                f"Added rectangle annotation to page {self.current_page.get()}"
            )
            messagebox.showinfo(
                "Success",
                f"Rectangle annotation added to page {self.current_page.get()}",
            )

        except Exception as e:
            self.update_status(f"Error processing page: {str(e)}")
            messagebox.showerror("Error", f"Failed to process page: {str(e)}")

    def show_about(self) -> None:
        about_text = """PDF Annotator

Version: 1.0
Python version: {py_ver}
PyMuPDF version: {mupdf_ver}

A tool for finding and annotating rectangles in PDF documents."""

        messagebox.showinfo(
            "About",
            about_text.format(
                py_ver=sys.version.split()[0], mupdf_ver=fitz.__version__
            ),
        )


def main():
    root = tk.Tk()
    _ = PDFAnnotator(root)
    # Set icon if available
    icon_path = "pdf.ico"
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    root.mainloop()


if __name__ == "__main__":
    main()
