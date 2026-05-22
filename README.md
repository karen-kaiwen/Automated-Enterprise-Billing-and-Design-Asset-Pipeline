# Automated Enterprise Billing & Design Asset Pipeline

Reduced a 1–2 week manual billing and asset management workflow (requiring overtime) to same-day completion. Built with Google Apps Script on Google Workspace — automating record matching, Drive folder generation, sequential ID assignment, and CSV export for Adobe Illustrator variable data printing.

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat&logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
---

## Table of Contents

- [Overview](#overview)
- [System Architecture](#system-architecture)
- [Technical Challenges & Solutions](#technical-challenges--solutions)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup & Configuration](#setup--configuration)
- [Usage](#usage)
- [CSV Export Schema](#csv-export-schema)
- [Known Limitations](#known-limitations)
- [License](#license)

## Overview

This GAS-based pipeline automates three connected workflows inside Google Workspace: matching records across spreadsheets, auto-generating Drive folder structure with manifest files, and exporting sanitized CSVs formatted for Adobe Illustrator's Data Merge (variable data printing) panel.

It is designed for billing or asset-issuance operations where each record requires a unique sequential ID, a dedicated cloud folder, and a corresponding print-ready data row.

## System Architecture

```
┌─────────────────────────────┐
│   External_Response_Logs    │  (Payment routing: email / address)
│   [External Spreadsheet]    │
└────────────┬────────────────┘
             │ buildLookupMap()
             ▼
┌─────────────────────────────┐        ┌──────────────────────────────┐
│   Central_Master_Database   │────＞  │    Operations_Dashboard      │
│   [Source Sheet]            │        │    [Target Sheet]            │
│   Line-item records         │        │    One row per client        │
└─────────────────────────────┘        └──────────────┬───────────────┘
                                                      │ processCurrentRows()
                                                      ▼
                                 ┌───────────────────────────────────────┐
                                 │  Google Drive                         │
                                 │  123456XXXX_Name/                     │
                                 │  └123456XXXX_Name(XX in total).xlsx   │
                                 └───────────────────────────────────────┘
                                                     │
                                                     ▼
                                        ┌────────────────────────┐
                                        │  Illustrator CSV       │
                                        │ AI_Final_MMDD_HHmm.csv │
                                        └────────────────────────┘
```

> The pipeline is triggered manually via the custom menu in Google Sheets.
> All three data sources must be configured in CONFIG before first use.
---

## Technical Challenges & Solutions

- The Scientific Notation Trap (Data Integrity)
  - **The Problem**: Standard exports of long Account IDs (eg., 123456...) are often misinterpreted by software as numbers, converting them into scientific notation (1.23E+09), which breaks barcode scanning.
  - **The Solution**: Implemented `.getDisplayValues()` throughout the pipeline. This captures data as " What You See Is What You Get" strings, ensuring 100% data fidelity for downstream scanners.
- Cross-Platform Encoding Compatibility (UTF-8 BOM)
  - **The Problem**: Adobe Illustrator is highly sensitive to CSV encoding. Traditional UTF-8 files often cause character corruption or "Invalid Library" errors.
  - **The Solution**: Injected a UTF-8 BOM (\ufeff) into the CSV headers. This acts as a "digital handshake" for Adobe software, ensuring seamless import of multi-byte characters (Traditional Chinese).
- High-Performance Batch Processing
  - **The Problem**: Repeatedly calling Sheets API in loops causes script timeouts during peak billing cycles.
  - **The Solution**: Leveraged In-memory Indexing (Hash Maps). This shifted search complexity from O(n) to O(1), resulting in exponential performance gains for large datasets.
---
## Features 
| Feature | Description | 
|---|---|
| 🚀 **One-Click Sync** | Processes selected rows, assigns serial IDs, and populates all related fields automatically | 
| 🗂️ **Auto Cloud Archiving** | Creates a dedicated Google Drive folder and manifest spreadsheet per record |
| 🔗 **Cross-Sheet Data Matching** | Replaces VLOOKUP chains with an in-memory HashMap for O(1) lookup performance |
| 📮 **ZipCode Lookup Sidebar** | Embedded HTML sidebar for real-time postal code lookup via external API |
| 📥 **Illustrator CSV Export** | Generates sanitized, UTF-8 BOM encoded CSV matched to AI Variables Panel schema |
| 👤 **Operator Tracking** | Auto-identifies the current user by login and maps to a display name |
| 💳 **Payment Route Resolver** | Dynamically routes delivery info (email or physical mail) based on client type |
---

## Prerequisites

- Google Workspace account with access to Google Sheets, Drive, and Apps Script
- Two spreadsheets: an **Operations Dashboard** and an **External Response Logs**  database (can reside in a different spreadsheet file)
- A parent Google Drive folder to receive auto-generated subfolders
- A sidebar HTML file named `Sidebar.html` deployed alongside `Code.gs`  (required for the ZipCode Lookup feature)
- Adobe Illustrator with Data Merge panel (for consuming the exported CSV)
- On first run, Apps Script will request the following OAuth scopes:  Google Sheets, Google Drive, and external URL fetch (for ZipCode lookup)

## Setup & Configuration

1. Open **Extensions → Apps Script** from your Operations Dashboard spreadsheet.
2. Paste the contents of `Code.gs` and create a corresponding `Sidebar.html` file.
3. Update the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  targetSheet:    "Operations_Dashboard",      // Sheet name for the main workflow
  sourceSheet:    "Central_Master_Database",   // Sheet name for source records
  parentFolderId: "YOUR_DRIVE_FOLDER_ID",      // Drive folder ID (from the URL)
  externalSsId:   "YOUR_EXTERNAL_DATABASE_ID", // Spreadsheet ID of the external DB
  externalSheet:  "External_Response_Logs",    // Sheet name inside the external DB
  idPrefix:       "123456",                    // Prefix for generated IDs
  startSuffix:    1000,                        // Starting number for the ID suffix
  exportCols:     [0, 1, 2, 3, 4, 5, 6, 7, 8],// Column indices to include in manifest
  userList: {
    "your.username": "Your Display Name"       // Map: email username → display name
  }
};
```

| Key | How to find it |
|-----|----------------|
| `parentFolderId` | Open the Drive folder → copy the ID from the URL: `.../folders/THIS_PART` |
| `externalSsId` | Open the spreadsheet → copy the ID: `.../spreadsheets/d/THIS_PART` |
| `exportCols` | Zero-indexed column positions from `Central_Master_Database` to include in each manifest |
| `userList` | Maps Google account username (part before `@`) to a display name written into column H |

4. Save and reload the spreadsheet. A custom menu **🛠️ Automation Engine** will
   appear in the menu bar.

---

## Usage

### 🚀 Run: Sync & Generate Assets

1. Navigate to the `Operations_Dashboard` sheet.
2. Select one or more data rows where **column A is empty** and **column B contains
   a client name** matching records in `Central_Master_Database`.
3. Run **🛠️ Automation Engine → Run: Sync & Generate Assets**.

The script will:
- Assign a sequential ID and hyperlink it to the newly created Drive folder
- Inject formulas into columns C–L (cross-sheet lookup, amount totals,
  payment routing, status checkboxes)
- Create a Drive subfolder named `{ID}_{Name}({N} in total)`
- Generate a manifest spreadsheet inside it, with filtered source rows and
  a `Total Amount` summary row

### 📮 Open: ZipCode Search Sidebar

1. Click a cell containing a Taiwanese address string.
2. Run **🛠️ Automation Engine → Open: ZipCode Search Sidebar**.
3. The sidebar will read the active cell's value, query `zipcode.tw`, and write
   the matched 6-digit postal code to the cell immediately to the left.

### 📥 Export: Illustrator CSV (Clean Text)

1. Select the data rows to export — **do not include the header row**.
2. Run **🛠️ Automation Engine → Export: Illustrator CSV (Clean Text)**.
3. A download dialog will appear with a timestamped file
   (e.g., `AI_Final_0420_1430.csv`) ready for Illustrator's Data Merge panel.

---

## CSV Export Schema

The exported file contains 14 fixed columns mapped to Illustrator variable names.
The `Period` and `Issue_Date` columns are intentionally left blank — fill them
manually before importing into Illustrator.

| # | Column name | Source | Notes |
|---|-------------|--------|-------|
| 1–2 | `Client_Name_1`, `Client_Name_2` | Column B | Duplicated for layout flexibility |
| 3 | `Case_ID` | Column C | Info / approval reference |
| 4 | `Capacity` | Column D | |
| 5–6 | `Amount_1`, `Amount_2` | Column E | Duplicated for layout flexibility |
| 7–8 | `Account_ID_1`, `Account_ID_2` | Column A | Duplicated for layout flexibility |
| 9 | `barcode1` | Column A | Last 11 digits of Account ID |
| 10 | `barcode2` | Column E | Amount with commas stripped |
| 11 | `Period` | — | **Blank — fill manually** |
| 12 | `ZipCode` | Column F | |
| 13 | `Shipping_Address` | Column G | |
| 14 | `Issue_Date` | — | **Blank — fill manually** |

---

## Known Limitations

- `getMaxSuffix()` parses the last 4 digits of existing IDs; changing `idPrefix`
  mid-operation may cause numbering conflicts.
- `exportCols` indices must stay in sync with the actual column layout of
  `Central_Master_Database` if the source sheet structure changes.
- The ZipCode lookup relies on `zipcode.tw` — if the site's HTML structure changes,
  the regex in `fetchZipCode()` will need to be updated accordingly.
- `Sidebar.html` is not included in this repository; implement your own HTML sidebar
  to interact with `getSelectedCellInfo()` and `writeZipCode()`.
- Apps Script has a **6-minute execution time limit** per run; processing very large
  batches in a single selection may hit this ceiling.

---

## License

MIT License.

This codebase has been de-identified for public sharing. All sheet names, folder IDs,
spreadsheet IDs, and user references are placeholders — replace them with your own
values before use.

## 🤝 Acknowledgements

Built with AI pair programming assistance from [Claude](https://claude.ai/) (Anthropic) and [Gemini](https://gemini.google.com/) (Google).  
Core system design, logic, and requirements by the author.
