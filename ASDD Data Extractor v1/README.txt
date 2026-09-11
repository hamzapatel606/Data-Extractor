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
