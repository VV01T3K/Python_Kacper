#!/usr/bin/env python

import tkinter as tk
from pdf_tool_gui import PDFToolGUI


def main():
    root = tk.Tk()
    _ = PDFToolGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
