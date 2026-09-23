# Simplified Class Scheduling App v2.1 — Replace v2.0

This release fixes a Workday Excel compatibility issue discovered with the uploaded Training History report.

## What was wrong
Some Workday-generated `.xlsx` files contain incorrect worksheet dimension metadata (`A1:A1`) even though thousands of rows are present. OpenPyXL read-only mode trusts that metadata and therefore sees only cell A1. Version 2.1 opens Workday report files in normal mode so all rows are available.

## Replace the current GitHub files
Replace the contents of the current Streamlit repository with the contents of this folder, or at minimum replace:

- `scheduler_engine.py`
- `batch_utils.py`
- `app.py`

Commit the changes to `main`. Streamlit should redeploy automatically; reboot the app once if needed.

## Confirm the correct release
Below the app title, confirm you see:

`Simplified stateless build v2.1`

Then re-upload the reports. For the supplied `test history 2026-09-22 13_46 EDT.xlsx`, preflight should report approximately:

- 2,812,000 bytes
- Sheet1: 28,413 rows × 19 columns
- 28,403 data rows

The warning `Workbook contains no default style` is harmless and comes from the Workday export format.
