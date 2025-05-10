# Making an exe

```bash
uv run pyinstaller --onefile --name=pdfTool --icon=pdf.ico "main.py"
```
<!-- change python to 3.12 -->
python -m nuitka --standalone --windows-console-mode=disable --include-data-files=pdf.ico=pdf.ico --enable-plugin=tk-inter --output-dir=dist main.py --windows-icon-from-ico=pdf.ico
