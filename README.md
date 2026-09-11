# Data Extractor

A collection of specialized data extraction tools designed to extract, process, clean, and organize electoral and voter roll data from PDF and CSV files.

This repository contains separate extractors built specifically around the different PDF structures and formats used in:

- **ASDD Data** – Extracting and processing ASDD-related voter data
- **Draft Roll Data** – Extracting structured voter information from Draft Roll PDFs
- **Discrepancy Data** – Extracting and organizing discrepancy-related voter records

Each document type follows a different PDF structure, so the extractors are designed independently to handle their respective formats accurately and efficiently.

## Repository Structure

```text
Data-Extractor/
│
├── ASDD Data Extractor v1/
│   └── ASDD data extraction tools
│
├── Draft Data Extractor v1/
│   └── Draft Roll data extraction tools
│
└── Discrepancy Data Extractor v1/
    └── Discrepancy data extraction tools
