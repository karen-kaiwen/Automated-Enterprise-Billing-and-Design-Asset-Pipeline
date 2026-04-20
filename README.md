# Automated Enterprise Billing & Design Asset Pipeline

An end-to-end Google Apps Script system Google Apps Script system for Google Workspace that automates record matching, Drive folder generation, and CSV export for Adobe Illustrator variable data printing.

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat&logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
---

## Table of Contents

- [Overview](#overview)
- [System Architecture](#system-architecture)
- [Technical Challenges & Solutions](#technical-challenges-&-solutions)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup & Configuration](#setup--configuration)
- [Usage](#usage)
- [CSV Export Schema](#csv-export-schema)
- [Known Limitations](#known-limitations)
- [License](#license)

## Overview

This GAS-based pipeline automates three connected workflows inside Google Workspace: matching records across spreadsheets, auto-generating Drive folder structure with manifest files, and exporting sanitized CSVs formateted for Adobe Illustrator's Data Merge (variable data printing) panel.

It is designed for billing or asset-issuance opperations where each record requires a unique sequential ID, a dedicated cloud folder, and a corresponding print-ready data row.

The codebase has been de=identified for public sharing. All sheet names, folder IDs, and user mappings use placeholder values that you replace during setup.

## System Arichetecture

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
                                         ┌────────────────────────┐
                                         │  Google Drive          │
                                         │  INV-2026-XXXX_Name/   │
                                         │  └── XXXX_Manifest.xlsx│
                                         └────────────────────────┘
                                                     │
                                                     ▼
                                        ┌────────────────────────┐
                                        │  Illustrator CSV       │
                                        │  Export_AI_Variables_  │
                                        │  MMDD_HHmm.csv         │
                                        └────────────────────────┘
```
 
---

## Features 

