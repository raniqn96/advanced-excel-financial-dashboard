# Advanced Excel Dashboard & Financial Models

![Excel](https://img.shields.io/badge/Excel-Advanced-217346?logo=microsoftexcel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-5C2D91)
![Power Query](https://img.shields.io/badge/Power%20Query-M-F2C811)
![Python](https://img.shields.io/badge/Python-openpyxl-3776AB?logo=python&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

An executive-level **financial reporting, reconciliation and forecasting workbook** built with Excel, VBA,
Power Query and Pivot Tables — plus a Python generator that rebuilds the whole model from scratch.

---

## 1. Project Overview

Finance teams often work from several disconnected systems. This project shows how to combine them in one
automated Excel model that managers can use without technical help.

| Goal | How this project solves it |
|------|----------------------------|
| **Consolidate multiple data sources** | ERP sales transactions and General Ledger postings are loaded (Power Query or VBA) into structured Excel Tables. |
| **Executive dashboard** | KPI cards (Revenue, Gross Profit, GP Margin, Transactions, Avg Deal Size, Budget Variance, Reconciliation Exceptions), trend, region and category charts. |
| **Automated monthly reconciliation** | ERP vs GL tie-out by month with a configurable tolerance and automatic *Reconciled / Investigate* status. |
| **Performance tracking** | Pivot-style Region × Month summary (SUMIFS) and a native PivotTable built by a macro. |
| **Forecasting** | 12-month rolling forecast driven by growth, seasonality and margin assumptions. |
| **Variance analysis** | Budget vs Actual with $ / % variances, YTD tracking and threshold-based flags. |
| **Usability for non-technical users** | Region drop-down filter, data validation on inputs, conditional formatting, colour-coded cells. |

### Key skills demonstrated
- Advanced formulas: `VLOOKUP`, `HLOOKUP`, `SUMIF`, `SUMIFS`, `COUNTIFS`, `IF`, `TEXT`, `ROUND`, `AVERAGE`
- Power Query (M): typing, cleansing, de-duplication, merges (joins), grouping, reconciliation logic
- VBA: one-click refresh, native PivotTable creation, CSV import with logging, PDF export
- Pivot Tables, Excel Tables (ListObjects), charts (line, bar, pie)
- Conditional formatting (colour scales, data bars, rule-based highlights) and data validation
- Python automation with `openpyxl`

---

## 2. Repository Structure

```text
excel-financial-dashboard/
├── workbook/
│   └── Financial_Dashboard_Model.xlsx     # Ready-to-open financial model
├── data/
│   ├── erp_sales.csv                      # Source 1: ERP sales extract (synthetic)
│   └── gl_revenue.csv                     # Source 2: General Ledger postings (synthetic)
├── src/
│   ├── python/
│   │   └── build_workbook.py              # Automated workbook generator
│   ├── vba/
│   │   ├── modDashboardAutomation.bas     # Refresh, PivotTable, recon review, PDF export
│   │   └── modDataImport.bas              # CSV import, validation, import log, de-dup
│   └── power_query/
│       ├── 01_ERP_Sales.m                 # Clean + enrich ERP sales
│       ├── 02_GL_Revenue.m                # Load GL revenue postings
│       ├── 03_Master_Data.m               # Products & Regions reference queries
│       └── 04_Monthly_Reconciliation.m    # ERP vs GL reconciliation
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

---

## 3. Workbook Tour

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | KPI cards, region selector (cell `C4`), monthly trend, revenue by region, category mix. |
| **Assumptions** | Yellow/blue input cells: growth, Q4 seasonality, GP target, budget growth, recon tolerance, variance threshold. |
| **Lookups** | Product master (`VLOOKUP` source) and horizontal region master (`HLOOKUP` source). |
| **Data_Sales** | 1,200 ERP transactions enriched with lookups, net revenue and gross profit. |
| **Data_GL** | Month-end GL revenue journals per region. |
| **Pivot_Summary** | Region × Month revenue matrix with heat-map, category revenue/GP with data bars. |
| **Reconciliation** | ERP vs GL by month, difference, % and status flag. |
| **Budget_vs_Actual** | Monthly & YTD variance analysis with chart and alert flags. |
| **Forecast** | 12-month forecast of revenue, gross profit and cumulative revenue with chart. |

**Colour convention:** blue text = inputs, black = formulas, green = links to other sheets, yellow fill = key assumptions.

---

## 4. Getting Started

### Option A — Just open the workbook
Open `workbook/Financial_Dashboard_Model.xlsx` in Microsoft Excel (2016+ / Microsoft 365). Formulas calculate
on open. Change the region in **Dashboard!C4** or edit the **Assumptions** sheet to see everything update.

### Option B — Regenerate it with Python
```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python src/python/build_workbook.py                  # default: 1,200 rows, seed 42
python src/python/build_workbook.py --rows 5000 --seed 7
```
This rewrites `workbook/Financial_Dashboard_Model.xlsx` and the CSV files in `data/`.

### Add the VBA macros
1. Open the workbook and press **Alt + F11**.
2. **File → Import File…** and select both `.bas` files in `src/vba/`.
3. Save as **Excel Macro-Enabled Workbook (.xlsm)**.
4. Press **Alt + F8** and run `RefreshAll_Model`, `BuildNativePivot`, `ExportDashboardPDF` or `ImportSalesCSV`.

### Add the Power Query scripts
1. Create a named cell `prmSourceFolder` containing the path to the `data/` folder (ending with `\` or `/`).
2. Name the region block on *Lookups* (`G4:K6`) as `rngRegions`.
3. **Data → Get Data → From Other Sources → Blank Query → Advanced Editor**; paste each `.m` file
   (create `Products` and `Regions` from `03_Master_Data.m` first, then `ERP_Sales`, `GL_Revenue`, `Monthly_Reconciliation`).
4. **Close & Load**. Use **Data → Refresh All** (or the `RefreshAll_Model` macro) each month.

---

## 5. Uploading This Project to GitHub

### Method 1 — GitHub website (no coding needed)
1. Unzip the downloaded file on your computer.
2. Sign in at [github.com](https://github.com) and click **+ → New repository** (top-right).
3. Name it `excel-financial-dashboard`, add a short description, choose **Public**.
   **Do not** tick "Add a README", ".gitignore" or "license" — they are already included.
4. Click **Create repository**.
5. On the new page, click **uploading an existing file**.
6. Open the unzipped `excel-financial-dashboard` folder, select **everything inside it** (including the
   `src`, `workbook` and `data` folders) and drag it into the browser window.
   *Tip:* hidden files like `.gitignore` may not show — on Windows enable *View → Hidden items*; on Mac press
   **Cmd + Shift + .** in Finder.
7. Enter a commit message such as `Initial commit: Excel financial dashboard` and click **Commit changes**.

### Method 2 — Git command line
```bash
cd path/to/excel-financial-dashboard
git init
git add .
git commit -m "Initial commit: Excel financial dashboard & models"
git branch -M main
git remote add origin https://github.com/<your-username>/excel-financial-dashboard.git
git push -u origin main
```
When asked for a password, use a **Personal Access Token** (GitHub → Settings → Developer settings → Personal access tokens).

### Method 3 — GitHub Desktop
1. Install [GitHub Desktop](https://desktop.github.com) and sign in.
2. **File → Add local repository…** → choose the unzipped folder → **create a repository** when prompted.
3. Click **Publish repository**, untick "Keep this code private" if you want it public, and publish.

### After uploading
- Add **topics** in the repo's *About* panel: `excel`, `vba`, `power-query`, `pivot-tables`, `financial-modeling`, `dashboard`, `python`.
- Pin the repository on your GitHub profile and link it from your résumé / LinkedIn.
- Optional: add a dashboard screenshot as `docs/dashboard.png` and reference it at the top of this README.

---

## 6. Data Disclaimer
All data is **synthetic**, generated by `build_workbook.py` for demonstration. No real company or customer data is included.

## 7. License
Released under the [MIT License](LICENSE).

## 8. Author
**Rani Kumari** — Financial / Data Analyst · Excel · VBA · Power Query · Python
# advanced-excel-financial-dashboard
