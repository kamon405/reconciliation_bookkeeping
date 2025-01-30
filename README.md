# QuickBooks vs. Bank Statement Reconciliation Tool

## Overview
This script provides a GUI-based reconciliation tool that allows users to compare transactions between QuickBooks files and bank statements. The tool supports multiple file formats, including CSV, Excel, and PDF. If a PDF contains scanned images, the script employs OCR to extract text data.

## Features
- **Multi-file Handling**: Supports multiple QuickBooks and bank statement files for reconciliation.
- **PDF & OCR Support**: Extracts text from PDFs using `pdfplumber`, and applies OCR (`pytesseract`) when needed.
- **Data Cleansing**: Ensures correct formatting and handles missing or inconsistent data.
- **Fuzzy Matching**: Utilizes `fuzzywuzzy` to identify and match similar transaction descriptions.
- **Duplicate Identification**: Flags potential duplicate transactions.
- **Discrepancy Reporting**: Highlights transactions that exist only in QuickBooks or the bank statement.
- **Automated Report Generation**: Outputs an Excel file with a timestamped filename containing the reconciliation results.

## Requirements
Ensure you have the following dependencies installed:

```sh
pip install pandas fuzzywuzzy pdfplumber pdf2image pytesseract openpyxl
```

Additionally, for OCR functionality:
- Install [Tesseract-OCR](https://github.com/tesseract-ocr/tesseract) and ensure it is added to your system path.
- Install `poppler` for `pdf2image` (varies by OS).

## How to Use
1. **Launch the Application**:
   - Run the script using Python:
     ```sh
     python reconcil-bookkeep.py
     ```

2. **Select QuickBooks Files**:
   - Click the **"Browse"** button to select QuickBooks transaction files (`.csv` or `.xlsx`).

3. **Select Bank Statement Files**:
   - Click the **"Browse"** button to select bank statements (`.csv`, `.xlsx`, or `.pdf`).

4. **Start Reconciliation**:
   - Click **"Reconcile Transactions"** to initiate the comparison.

5. **View Results**:
   - The reconciled report is saved as an Excel file with a timestamp in the same directory as the QuickBooks files.
   - The report includes columns for:
     - **Date, Description, Amount**
     - **Match Status** (`Exact Match`, `Only in QuickBooks`, `Only in Bank Statement`)
     - **Duplicate Flag**

## Troubleshooting
- **No text extracted from PDFs?** Ensure Tesseract and Poppler are installed correctly.
- **Mismatch in amounts or dates?** Verify that transactions are correctly formatted in the source files.
- **Duplicate transactions incorrectly flagged?** Check if the transactions have slight variations in descriptions.

## Future Enhancements
- Improve OCR accuracy using preprocessing techniques.
- Add manual review options for unmatched transactions.
- Enable fuzzy matching for amounts in addition to descriptions.

For any issues or suggestions, feel free to contribute or report them!

