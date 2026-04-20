# Automated Enterprise Billing & Design Asset Pipeline

An end-to-end Google Apps Script system that automates record matching, cloud archiving, and sanitized CSV export for high-volume Adobe Illustrator variable data production.

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

This project was built to eliminate manual work in a billing operations workflow. Given a list of client names in an operations dashboard, the script automatically: 
- Generates sequential invoice IDs and hyperlinks them to newly created Drive folders.
- Pulls matching records from a master database and archives them as a manifest spresdsheet.
- Injects VLOOKUP/ SUMIF formulas and payment routing info into the dashboard.

