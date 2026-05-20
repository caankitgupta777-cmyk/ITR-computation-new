# ITR PDF to Computation of Income Tool

This folder contains a local converter for Income Tax portal ITR PDF downloads.
It extracts the main ITR-4 fields and creates a Word-openable Computation of
Income document in `.doc` format.

## Desktop App

To open the app, double-click:

```powershell
run_itr_converter_app.cmd
```

Or run it from PowerShell:

```powershell
.\run_itr_converter_app.cmd
```

Then select the Income Tax portal PDF, choose the output location if needed, and
click **Convert PDF**.

## Usage

Double-click or run the wrapper:

```powershell
.\convert_itr_pdf.cmd "C:\path\to\income-tax-portal-file.pdf"
```

With a chosen output file:

```powershell
.\convert_itr_pdf.cmd "C:\path\to\income-tax-portal-file.pdf" ".\Computation.doc"
```

Or run the Python script directly:

```powershell
& "C:\Users\caank\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\itr_pdf_to_computation.py "C:\path\to\income-tax-portal-file.pdf"
```

To choose the output file:

```powershell
& "C:\Users\caank\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\itr_pdf_to_computation.py "C:\path\to\income-tax-portal-file.pdf" -o ".\Computation.doc"
```

To inspect extracted fields:

```powershell
& "C:\Users\caank\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\itr_pdf_to_computation.py "C:\path\to\income-tax-portal-file.pdf" --json
```

## Notes

- The current parser is tailored for the ITR-4 PDF layout from the income tax
  portal and the computation style in the sample Word document.
- The output `.doc` is HTML saved with a Word-compatible extension, matching the
  structure of the provided sample computation file.
- Amounts are copied from the PDF where available. The old-regime comparison is
  computed as a basic slab comparison unless more detailed old-regime deduction
  data is available in the source PDF.

## Deploy Online

This project now includes a Flask web app in `web_app.py`. The desktop Tkinter
app remains available in `itr_converter_app.py`, but online hosts should run
the Flask app.

### Render

1. Create a GitHub repository and upload these project files.
2. Go to Render and create a new Web Service from that repository.
3. Use these settings:
   - Runtime: Python
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn web_app:app`
4. Deploy the service.

Render can also read `render.yaml` from this folder and fill most settings
automatically.

### Run locally before deploying

```powershell
pip install -r requirements.txt
python web_app.py
```

Open `http://127.0.0.1:5000`, upload an ITR PDF, and download the generated
Word-openable `.doc` file.
