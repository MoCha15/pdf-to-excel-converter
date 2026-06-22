# PDF → Excel Converter

A standalone desktop tool that extracts tables and text from PDFs and saves them as formatted Excel files. No Python or technical setup required.

---

## Download

Go to the [Releases](https://github.com/MoCha15/pdf-to-excel-converter/releases) page and download the file for your OS:

| Platform | File to download |
|----------|-----------------|
| Windows  | `PDF-to-Excel-Converter.exe` |
| macOS    | `PDF-to-Excel-Converter-macOS.zip` |

**Windows:** double-click the `.exe` to run.

**macOS:** unzip the file, then double-click `PDF-to-Excel-Converter.app`. If macOS blocks it on first launch, right-click → Open.

---

## How to use

1. Launch the app
2. Drop a PDF onto the drop zone, or click to browse for a file
3. Click **Convert**
4. When done, click **Open Result** to view the Excel file

The output file is saved in the same folder as the PDF, named `<original-filename>_tables.xlsx`.

---

## What to expect

**Example:** you have a 5-page bank statement PDF with transaction tables on each page.

- The app detects all 5 tables and checks if they share the same column headers (Date, Particulars, Debit, Credit, Balance)
- It asks: *"All 5 tables have identical headers. Combine them into one sheet?"*
  - **Yes** → one sheet named `Combined` with all rows merged under a single header
  - **No** → five separate sheets named `Page1`, `Page2`, … `Page5`
- A `Summary` sheet is always added as the first tab, showing the source filename, conversion timestamp, and sheet list

The log panel at the bottom shows live progress — pages processed, rows and columns found per table, and the final output path.

---

## Features

- **Drag-and-drop** — drop a PDF directly onto the app window
- **Two extraction strategies** — automatically picks the best method per page:
  - Line-based extraction for tables with visible borders
  - Word-position-based extraction for borderless/text-flow tables
- **Smart header detection** — recognises common financial/tabular column names to find where data starts
- **Auto-combine** — offers to merge multiple tables with identical headers into one sheet
- **Formatted output** — dark header row, alternating row colours, auto-fitted column widths, frozen header row
- **Summary sheet** — always included as the first tab with source file info and sheet index
- **Fallback to raw text** — if no tables are found on a page, raw text lines are saved instead of skipping the page
- **Self-installing** — installs required Python libraries automatically on first run (source version only)
