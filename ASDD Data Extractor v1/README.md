ASDD Data Extractor V1
======================

1. Run "0 - Setup.bat" once to install dependencies.
2. Put ASDD PDF files into the input folder.
3. Run "1 - Extract ASDD.bat".
4. The master workbook is created at:
   output\ASDD_Master.xlsx

Workbook sheets:
- All ASDD
- Needs Review

Output columns:
S.No
Serial No
EPIC Number
Elector Name
Relative Details
DOB/Age
Uncollectable Reason
Part No
PDF Page
PDF Box
Source File
EPIC Check
Status Mark
Suspicious Reason
Remark

Notes:
- Built specifically for the ASDD PDF structure.
- Uses native PDF text coordinates, not OCR.
- Handles wrapped text and page-boundary continuations.
- Preserves "Already enrolled (EPIC)" as shown in the PDF.
- A changed source PDF is reprocessed and replaces its previous rows.


V2.0 IMPORTANT FIX
-------------------
ASDD PDFs use native text coordinates where the actual table column starts are around:
S.No 25, Serial 55, EPIC 93.5, Elector 151.6, Relative 282.4, DOB/Age 433.8, Uncollectable Reason 478.8.
V1 incorrectly used 295 / 445 / 490 as the boundaries between Elector, Relative, Age and Reason.
That caused ages such as (25) to be captured inside Relative Details and caused the reason to be captured as DOB/Age, creating hundreds of false Suspicious records.
V2 detects the column starts from the PDF header on each page, with a corrected fallback layout.
PARSER_VERSION=2.0 automatically causes existing PDFs to be reprocessed once, even if their SHA-256 is unchanged.

When upgrading from V1, close ASDD_Master.xlsx before running the extractor.
