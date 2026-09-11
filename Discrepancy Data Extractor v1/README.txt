DISCREPANCY DATA EXTRACTOR V1
=============================

Purpose
-------
Extracts records from ECI Discrepancy PDFs using the fixed Discrepancy table structure.
This is NOT a Draft List extractor.

Input
-----
Put Discrepancy PDFs in the input folder.

Output
------
output\Discrepancy_Master.xlsx

Excel columns (exact order)
---------------------------
S.No
Part Serial Number
EPIC Number
Elector Name
Age
Gender
Reason for discrepancy
Part No
PDF Page
PDF Box
Source File
EPIC Check
Status Mark
Suspicious Reason
Remark

PDF Box
--------
Normalized [x0, y0, x1, y1] coordinates in the PDF page coordinate space, stored as 0..1 values.

Sheets
------
All Discrepancies
Needs Review

Validation
----------
EPIC Check and Status Mark contain only OK or Suspicious.
EPIC Check validates the EPIC format.
Status Mark validates the complete record.

Important
---------
The extractor uses the native PDF text layer and its coordinates first. This is intentional:
for this ECI PDF family, native text gives exact text positions and avoids unnecessary OCR errors.
Do not replace this with a generic table extractor without testing against the source PDFs.

Incremental processing
----------------------
The script stores a SHA-256 fingerprint for each successfully processed PDF. If a PDF changes,
it is processed again. Existing output is preserved when there is nothing new.
