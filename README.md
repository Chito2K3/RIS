# ARMMC Pharmacy — RIS & RSMI Automation App

Automated web application for **Amang Rodriguez Memorial Medical Center (ARMMC) — Allied Division / Pharmacy** (`RCC: 13 001 14 00012 03 02 33`) that processes raw Hospital Information System (HIS) monthly consumption and issuance exports into clean datasets, finalized **Requisition and Issue Slips (RIS)**, and the consolidated **Report of Supplies and Materials Issued (`RSMI.xlsx`)**.

## Features & 4-Stage Automated Pipeline

1. **Stage 1 — Clean & Combine (`01_CLEANED_DATA/`):**
   - Cleans the 2 raw consumption files (`CONSUMP_ANCILLIARY_<MONTH>_<YEAR>.xls` and `CONS_CONSUMP_<MONTH>_<YEAR>.xls`), unmerging all merged cells and stripping empty rows/columns, and combines them into a single clean file (`CONSUMPTION_COMBINED_CLEANED.xlsx` / `.csv`).
   - Cleans the raw issuance file(s) (`issuance_storage_<month>_<year>.xls`/`.csv` and `issuance_disp_<month>_<year>.xls`/`.csv`) into `ISSUANCE_COMBINED_CLEANED.xlsx`.
2. **Stage 2 — Master Hospital RIS (`02_RIS_FINAL/`):**
   - Generates `RIS - <MONTH>_<YEAR>_POPULATED.xlsx` from `CONSUMPTION_COMBINED_CLEANED.xlsx` with **Unit** populated from `INVENTORY UTILIZATION REPORT_MASTERLIST_2026_V30.xlsx` (`MASTERLIST_2026` Column H), `Lot Number` and `Expiration Date` removed, full `Purpose` banner, and standardized signatories.
3. **Stage 3 — Per-Ward RIS Slips (`02_RIS_FINAL/`):**
   - Generates individual `<WARD>_RIS_<MONTH>.xlsx` files (e.g., `PEDIA WARD_RIS_OCT.xlsx`) for every active ward with issued quantities > 0.
4. **Stage 4 — Consolidated `RSMI.xlsx` (`02_RIS_FINAL/`):**
   - Compiles all generated RIS files into `RSMI.xlsx` with `RSI No:`, `Stock No.`, `Unit`, `Description`, `Quantity`, and a `RIS Count Summary` sheet showing the item count and total quantity per `RSI No:`.

## Quick Start

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Launch the App:
   - Double-click **`run_ris_app.bat`**, or run:
     ```bash
     python app/server.py 5050
     ```
3. Open `http://127.0.0.1:5050` in your browser.
