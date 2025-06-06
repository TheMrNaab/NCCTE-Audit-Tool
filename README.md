# NCCTE-Audit-Tool

Cross-reference Certiport exam results with NCCTE credential records to catch mismatches, missing data, and reporting errors.

## Description

This Python script generates a structured HTML report comparing Certiport exam results (`Results (264) Plus.csv`) with NC CTE credential records (`Credentials_ByStudentCredential.csv`) by student. It is designed specifically for reports exported from the NCCTE Admin system.

## What It Does

- Loads two CSV files:
  - Certiport results (`Results (264) Plus.csv`)
  - Credential summary by course (`Credentials_ByStudentCredential.csv`)
- Prompts the user to select which exams to analyze.
- Merges and matches student records by ID.
- Lists all exam attempts from Certiport for each student.
- Evaluates whether NCCTE credential status reflects Certiport performance.
- Highlights inconsistencies, such as students marked “Not Attempted” who actually have results.
- Deduplicates identical exam entries to prevent clutter in the report.
- Outputs a print-friendly HTML file organized by student, with optional comment fields.

## Requirements

- Python 3.7 or higher
- Dependencies:
  - `pandas`
  - `tkinter` (standard in most Python distributions)
  - `tabulate`

## Input Files

- `Results (264) Plus.csv`: Exported from Certiport. Skip the first 3 rows when loading.
- `Credentials_ByStudentCredential.csv`: Exported from NCCTE Admin. Skip the first 5 rows when loading.

## Usage

1. Run the script.
2. Select the required CSV files when prompted.
3. Choose one or more exams to include in the comparison.
4. Select a location to save the generated HTML file.
5. Open the HTML report in a browser for review or printing.

## Notes

- This script is designed for credentialing workflows in North Carolina public schools using NCCTE Admin and Certiport.
- Duplicate exam entries (same student, exam, date, score, and result) are automatically removed from the final report.
- Flags are generated for:
  - Students marked “Not Attempted” but with Certiport results.
  - Students who show no attempt data for selected exams.

## Screenshots
[Example Report](example-report.png)

# License
Copyright © 2025 David Naab

This software is provided free of charge for individual, educational use only.

Attribution is required in all derivative works and distributed outputs.

Organization-level, institutional, or district-wide use — including deployment, modification, or redistribution within a school system, company, or agency — is strictly prohibited without prior written permission from the author.

To request licensing or usage approval for school district or institutional use, contact the author directly.
