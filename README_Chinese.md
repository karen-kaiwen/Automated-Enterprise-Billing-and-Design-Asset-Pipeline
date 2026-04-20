# 自動化企業帳單與設計資產流水線

一套基於 Google Apps Script 的端對端自動化系統，整合 Google Workspace 實現跨表資料比對、雲端資料夾建檔，以及輸出適用於 Adobe Illustrator 變數資料印刷的 CSV 檔案。

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat&logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

---

## 目錄

- [專案概述](#專案概述)
- [系統架構](#系統架構)
- [技術挑戰與解決方案](#技術挑戰與解決方案)
- [功能總覽](#功能總覽)
- [使用前置條件](#使用前置條件)
- [設定與部署](#設定與部署)
- [使用說明](#使用說明)
- [CSV 匯出欄位結構](#csv-匯出欄位結構)
- [已知限制](#已知限制)
- [授權條款](#授權條款)

---

## 專案概述

本系統以 Google Apps Script 為核心，自動化處理三個相互串聯的工作流程：跨試算表資料比對、依記錄自動建立 Drive 資料夾與清冊檔案，以及輸出符合 Adobe Illustrator 變數面板格式的 CSV 檔案。

適用於帳單發行或資產核發作業，每筆記錄需要獨立的流水序號、專屬雲端資料夾，以及對應的印刷用資料列。

---

## 系統架構

```
┌─────────────────────────────┐
│   External_Response_Logs    │  (付款路由資訊：電子郵件 / 實體地址)
│   [外部試算表]               │
└────────────┬────────────────┘
             │ buildLookupMap()
             ▼
┌─────────────────────────────┐        ┌──────────────────────────────┐
│   Central_Master_Database   │────＞  │    Operations_Dashboard      │
│   [來源工作表]               │        │    [目標工作表]               │
│   各項明細記錄               │        │    每列對應一位客戶            │
└─────────────────────────────┘        └──────────────┬───────────────┘
                                                      │ processCurrentRows()
                                                      ▼
                                         ┌────────────────────────┐
                                         │  Google 雲端硬碟        │
                                         │  INV-2026-XXXX_Name/   │
                                         │  └── XXXX_Manifest.xlsx│
                                         └────────────────────────┘
                                                     │
                                                     ▼
                                        ┌────────────────────────┐
                                        │  Illustrator CSV       │
                                        │  AI_Asset_Data_        │
                                        │  MMDD_HHmm.csv         │
                                        └────────────────────────┘
```

---

## 技術挑戰與解決方案

### 科學記號陷阱（資料完整性）
- **問題**：長字串帳號 ID（如 `123456789012`）在匯出時常被軟體誤判為數字，自動轉換為科學記號（`1.23E+09`），導致條碼掃描失效。
- **解決方案**：全流程採用 `.getDisplayValues()`，以「所見即所得」的方式擷取格式化字串，確保資料 100% 保真，不受下游掃描器影響。

### 跨平台編碼相容性（UTF-8 BOM）
- **問題**：Adobe Illustrator 對 CSV 編碼格式極為敏感，標準 UTF-8 檔案常造成字元亂碼或「無效程式庫」錯誤。
- **解決方案**：在 CSV 標頭注入 UTF-8 BOM（`\ufeff`），作為 Adobe 軟體的編碼識別信號，確保繁體中文等多位元字元能無縫匯入。

### 高效能批次處理
- **問題**：在迴圈中反覆呼叫 Sheets API，於批次帳單高峰期容易觸發逾時限制。
- **解決方案**：採用記憶體內索引（Hash Map）策略，將查找複雜度從 O(n) 降至 O(1)，大幅提升大量資料集的處理效能。

---

## 功能總覽

| 功能 | 說明 |
|------|------|
| 🚀 **一鍵同步與建檔** | 處理選取列、自動指派流水序號，並填入所有關聯欄位 |
| 🗂️ **自動雲端歸檔** | 每筆記錄自動建立獨立的 Drive 資料夾與清冊試算表 |
| 🔗 **跨表資料比對** | 以記憶體 HashMap 取代 VLOOKUP 鏈，達成 O(1) 查找效能 |
| 📮 **郵遞區號查詢側欄** | 嵌入式 HTML 側欄，透過外部 API 即時查詢台灣郵遞區號 |
| 📥 **Illustrator CSV 匯出** | 輸出經過清理、含 UTF-8 BOM 的 CSV，欄位對應 AI 變數面板結構 |
| 👤 **操作者自動識別** | 自動辨識當前登入帳號並對應顯示名稱 |
| 💳 **付款路由解析** | 依客戶類型動態判斷寄送方式（電子郵件 / 實體郵寄） |

---

## 使用前置條件

- 具備 Google Sheets、Drive 及 Apps Script 存取權限的 Google Workspace 帳號
- 兩份試算表：**Operations Dashboard**（主作業表）與 **External Response Logs**（外部資料庫，可位於不同試算表檔案）
- 一個 Google Drive 父資料夾，用於接收自動建立的子資料夾
- 與 `Code.gs` 並列部署的 `Sidebar.html` 側欄檔案（郵遞區號查詢功能必要）
- Adobe Illustrator（含 Data Merge 面板，用於匯入匯出的 CSV）

---

## 設定與部署

1. 在 Operations Dashboard 試算表中，開啟 **擴充功能 → Apps Script**。
2. 貼上 `Code.gs` 的內容，並建立對應的 `Sidebar.html` 檔案。
3. 修改 `Code.gs` 頂部的 `CONFIG` 物件：

```javascript
const CONFIG = {
  targetSheet:    "Operations_Dashboard",      // 主作業工作表名稱
  sourceSheet:    "Central_Master_Database",   // 來源主資料工作表名稱
  parentFolderId: "YOUR_DRIVE_FOLDER_ID",      // Drive 資料夾 ID（從網址列取得）
  externalSsId:   "YOUR_EXTERNAL_DATABASE_ID", // 外部資料庫的試算表 ID
  externalSheet:  "External_Response_Logs",    // 外部資料庫內的工作表名稱
  idPrefix:       "INV-2026-",                 // 流水序號前綴
  startSuffix:    1000,                        // 流水序號起始數字
  exportCols:     [0, 1, 2, 3, 4, 5, 6, 7, 8],// 清冊匯出的欄位索引（從 0 起算）
  userList: {
    "your.username": "您的顯示名稱"            // 帳號使用者名稱（@ 前段）→ 顯示名稱對應
  }
};
```

| 設定鍵 | 取得方式 |
|--------|---------|
| `parentFolderId` | 開啟 Drive 目標資料夾 → 從網址複製 ID：`.../folders/這一段` |
| `externalSsId` | 開啟外部試算表 → 從網址複製 ID：`.../spreadsheets/d/這一段` |
| `exportCols` | `Central_Master_Database` 中需納入清冊的欄位位置，以 0 為起始索引 |
| `userList` | 將 Google 帳號的使用者名稱（`@` 前段）對應至寫入工作表的顯示名稱 |

4. 儲存後重新整理試算表，選單列將出現 **🛠️ Automation Engine** 自訂選單。

---

## 使用說明

### 🚀 執行：同步與建立資產

1. 切換至 `Operations_Dashboard` 工作表。
2. 選取一列或多列資料，確認 **A 欄為空**（尚未處理）且 **B 欄含有客戶名稱**（須與 `Central_Master_Database` 中的記錄相符）。
3. 執行 **🛠️ Automation Engine → Run: Sync & Generate Assets**。

腳本將自動執行以下操作：
- 指派流水序號，並以超連結形式連結至新建的 Drive 資料夾
- 向 C～L 欄注入公式（跨表查找、金額加總、付款路由、狀態核取方塊）
- 建立名為 `{序號}_{客戶名稱}` 的 Drive 子資料夾
- 在子資料夾內生成清冊試算表，包含篩選後的來源記錄與 `Total Amount` 加總列

### 📮 開啟：郵遞區號查詢側欄

1. 點選含有台灣地址字串的儲存格。
2. 執行 **🛠️ Automation Engine → Open: ZipCode Search Sidebar**。
3. 側欄將讀取選取格的內容，查詢 `zipcode.tw`，並將比對到的 6 位數郵遞區號寫入左側相鄰儲存格。

### 📥 匯出：Illustrator CSV（純文字）

1. 選取要匯出的資料列（**不含標題列**）。
2. 執行 **🛠️ Automation Engine → Export: Illustrator CSV (Clean Text)**。
3. 下載對話框將出現，提供含時間戳記的檔案（例如 `AI_Asset_Data_0420_1430.csv`），可直接匯入 Illustrator 的 Data Merge 面板。

---

## CSV 匯出欄位結構

匯出檔案包含 14 個固定欄位，對應 Illustrator 變數名稱。`Period` 與 `Issue_Date` 欄位刻意留空，請在匯入 Illustrator 前手動填入。

| # | 欄位名稱 | 資料來源 | 備註 |
|---|---------|---------|------|
| 1–2 | `Client_Name_1`, `Client_Name_2` | B 欄 | 重複欄位，供版面配置彈性使用 |
| 3 | `Case_ID` | C 欄 | 資訊 / 核准參考編號 |
| 4 | `Capacity` | D 欄 | |
| 5–6 | `Amount_1`, `Amount_2` | E 欄 | 重複欄位，供版面配置彈性使用 |
| 7–8 | `Account_ID_1`, `Account_ID_2` | A 欄 | 重複欄位，供版面配置彈性使用 |
| 9 | `barcode1` | A 欄 | 帳號 ID 末 11 碼 |
| 10 | `barcode2` | E 欄 | 金額去除千分位逗號 |
| 11 | `Period` | — | **留空，請手動填入** |
| 12 | `ZipCode` | F 欄 | |
| 13 | `Shipping_Address` | G 欄 | |
| 14 | `Issue_Date` | — | **留空，請手動填入** |

---

## 已知限制

- `getMaxSuffix()` 解析現有 ID 的末 4 位數字；若在作業中途變更 `idPrefix`，可能導致序號衝突。
- 若 `Central_Master_Database` 的欄位結構異動，`exportCols` 的索引值須同步更新。
- 郵遞區號查詢依賴 `zipcode.tw`；若該網站 HTML 結構變更，需同步修改 `fetchZipCode()` 中的正規表達式。
- `Sidebar.html` 未收錄於本儲存庫；請自行實作 HTML 側欄，並串接 `getSelectedCellInfo()` 與 `writeZipCode()` 函式。
- Apps Script 單次執行上限為 **6 分鐘**；單次選取大量列進行批次處理時，可能觸及此限制。

---

## 授權條款

MIT License

本程式碼已去識別化後公開分享。所有工作表名稱、資料夾 ID、試算表 ID 及使用者資訊均為佔位符，使用前請替換為您自己的實際值。
