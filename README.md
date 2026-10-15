# GST Reconciliation Tool (Excel VBA)

An Excel VBA-based GST reconciliation workflow that compares invoice-level data from GST portal records with accounting records exported from Books. It produces a detailed reconciliation report and a summary sheet to help users investigate exceptions.

> \*\*Important:\*\* This repository contains an exported VBA module (`.bas`), not a complete Excel workbook. To run it, import the module into a macro-enabled Excel workbook containing the two required input sheets. The module header attributes the tool to \*\*Chinmay Bhat\*\*; retain that attribution unless you have permission to change it.

## Project purpose

Manual GST reconciliation can take time when invoice numbers are formatted differently, records are missing from one source, duplicate invoices exist, or invoice details and tax amounts differ. This tool automates matching and highlights records requiring review.

## Features

* Compares GST portal and Books data using GSTIN and normalized invoice-number matching.
* Includes alternative invoice-number matching passes for common formatting differences.
* Uses party-name similarity checks when evaluating matched records.
* Flags duplicate invoice numbers within each source before matching.
* Identifies records missing from either source.
* Compares invoice date, party name, GSTIN, taxable value, IGST, CGST, SGST and CESS.
* Produces a detailed `Reco_Result` worksheet with differences and status labels.
* Produces a `GST_Summary` worksheet with reconciliation summaries.
* Uses a configurable amount tolerance; the supplied version sets `GST_TOL = 0` (exact comparison to two decimal places).

## Status categories

* `Matched`
* `Mismatch-Amount`
* `Mismatch-Amt+Other`
* `Mismatch-Other`
* `Missing in Portal`
* `Missing in Books`
* `Duplicate-Portal`
* `Duplicate-Books`

A matched record can include the note `Inv No format differs` when the alternate invoice-number pass identifies a match.

## Input workbook structure

Create an Excel workbook with these exact worksheet names:

1. `As per Portal`
2. `As per Books`

Both sheets must use the following column order in columns A:I. Row 1 should contain headers; data should begin in row 2.

|Column|Required field|Example|
|-|-|-|
|A|GSTIN|`29ABCDE1234F1Z5`|
|B|Party Name|`Sample Components Pvt Ltd`|
|C|Invoice Date|`15/04/2025`|
|D|Invoice Number|`INV/25-26/001`|
|E|Taxable Value|`10000.00`|
|F|IGST|`1800.00`|
|G|CGST|`0.00`|
|H|SGST|`0.00`|
|I|CESS|`0.00`|

Use fictional data for testing. Keep the input column order exactly as specified because the macro reads the source ranges by position rather than locating fields by header text.

## Setup and run

1. Open Excel for Windows and create or open a **macro-enabled workbook** (`.xlsm`).
2. Create the two input worksheets named exactly `As per Portal` and `As per Books`.
3. Paste the corresponding source data into columns A:I, with headers in row 1 and records from row 2.
4. Save a backup copy of the workbook before running the macro.
5. Open the Visual Basic Editor with `Alt + F11`.
6. In the VBA Editor, use **File → Import File…** and select `GST_Reco_v20.bas`.
7. If an earlier version of this module is installed, remove the old copy first to avoid duplicate procedure names. Do not remove unrelated VBA modules without checking what they do.
8. Return to Excel and run `GST_Run` from `Alt + F8`.
9. Review the `GST_Summary` and `Reco_Result` worksheets, then investigate exceptions against source documents before making accounting or tax decisions.

## Matching and review notes

* Matching is primarily GSTIN-scoped and uses normalized invoice-number keys.
* An alternate match for differently formatted invoice numbers requires the same GSTIN, same invoice date and taxable value within the configured tolerance.
* A last-resort cross-GSTIN fallback uses an exact normalized invoice key and taxable value within tolerance; these rows are reported as a GSTIN mismatch rather than a clean match.
* The default amount tolerance is `0`. If you change it, test the impact on fictional data and document the setting.
* Automated matching can produce false positives or false negatives. Review all exceptions and material matches using the underlying invoices and GST records.

## Compatibility and limitations

* Designed for **Windows desktop Excel**; the code uses `Scripting.Dictionary` through `CreateObject`.
* Requires macros to be enabled in a trusted workbook environment.
* The `.bas` file is not a standalone executable and does not itself contain an Excel workbook or sample output.
* The macro recreates the `Reco_Result` and `GST_Summary` sheets when run. Back up any workbook containing prior outputs before execution.
* This tool supports reconciliation work; it does not replace professional review, statutory interpretation, or verification against source records.

## Technology and skills demonstrated

* Microsoft Excel
* Visual Basic for Applications (VBA)
* GST portal vs books reconciliation
* Invoice matching and data normalization
* Duplicate detection and exception reporting
* Accounting-data validation and discrepancy analysis

## License

This project is distributed under the **MIT License**. See the [`LICENSE`](LICENSE) file for the full license text.

The MIT License permits use, copying, modification, distribution, sublicensing and sale of copies, subject to its conditions, including retaining the copyright and permission notice.

## Attribution

The supplied VBA module header identifies **Chinmay Bhat** as the creator. This attribution is retained in the source module and this documentation.

## Confidentiality and responsible use

Use only fictional or appropriately anonymized data in this public repository. Do not publish client workbooks, real GSTINs, invoice details, tax records, credentials or other confidential information. Reconciliation outputs must be reviewed against source documents before accounting or tax decisions are made.
