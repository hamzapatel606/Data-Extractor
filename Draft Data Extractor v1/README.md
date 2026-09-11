DRAFT DATA EXTRACTOR V1
=======================

Purpose
-------
First benchmark version for the current Draft electoral-roll PDF format.
Accuracy is prioritized over speed.

Current benchmark PDF
---------------------
input\Part_51.pdf

Output
------
output\Master_Voter_List.xlsx

Excel fields
------------
Sr No
EPIC / Voter ID
EPIC Check
Name
Relation Type
Relative Name
House Number
Age
Gender
Status Mark
Suspicious Reason
Remark
Part No
PDF Page
PDF Box

Important design rules
----------------------
1. Sr No is assigned from physical PDF order, not trusted from OCR.
2. OCR serial is retained internally for mismatch detection.
3. Blank regions are not assumed to be voters.
4. Actual detected boxes determine how many voter records are created.
5. Suspicious records are sent to Needs Review.
6. PDF Page + normalized PDF Box are retained for future exact highlighting.
7. Section is intentionally NOT included.

Run
---
1. First run: 0 - Setup.bat
2. Put Draft PDFs in input\
3. Run 1 - Extract Draft.bat

NOTE
----
The first V1 is based on the proven box-detection/OCR approach we studied.
Do not change OCR settings until Part 51 has been benchmarked against the
known-good developer output.


V1.1 benchmark changes:
- Sr No is authoritative from physical PDF/grid order. OCR serial mismatches do not create Suspicious status.
- EPIC Check is a validation field: OK or Suspicious. OCR-source notes are not treated as failures.
- Output is written to the project-level output folder (not input\output).


V1.2 incremental processing:
- Existing processed PDFs are skipped using file fingerprints.
- Existing Excel rows are preserved, including manual corrections.
- New PDF rows are appended only after successful workbook creation.
- Internal tracking columns remain hidden in Excel for future incremental runs.
- All comments, errors, setup messages, and progress output are written in English.
