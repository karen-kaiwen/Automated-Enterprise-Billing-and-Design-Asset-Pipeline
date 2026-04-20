/**
 * @Project: Automated Enterprise Billing & Design Asset Pipeline
 * @Description: An End-to-End system for Google Workspace that automates data matching, 
 * cloud filing, and generates sanitized CSVs for Adobe Illustrator variable data.
 */

// ============================================================
//  ⚙️ 全域設定區 (已去識別化：使用通用變數與佔位符)
// ============================================================
const CONFIG = {
  targetSheet:    "Operations_Dashboard",     
  sourceSheet:    "Central_Master_Database",  
  parentFolderId: "YOUR_DRIVE_FOLDER_ID",     
  externalSsId:   "YOUR_EXTERNAL_DATABASE_ID",
  externalSheet:  "External_Response_Logs",   
  idPrefix:       "INV-2026-",               
  startSuffix:    1000,                       
  exportCols:     [0, 1, 2, 3, 4, 5, 6, 7, 8],
  userList: {                                 
    "admin.user1": "Account Manager A",
    "admin.user2": "Project Specialist B",
    "admin.user3": "Operations Lead C",
    "admin.user4": "System Admin D"
  }
};

// ============================================================
//  📌 選單系統
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Automation Engine')
    .addItem('🚀 Run: Sync & Generate Assets', 'processCurrentRows')
    .addItem('📮 Open: ZipCode Search Sidebar', 'showSidebar')
    .addItem('📥 Export: Illustrator CSV (Clean Text)', 'downloadIllustratorCSV')
    .addToUi();
}

// ============================================================
//  🚀 核心執行邏輯 (資料流自動化與雲端歸檔)
// ============================================================
function processCurrentRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.targetSheet) {
    SpreadsheetApp.getUi().alert(`❌ Please run this from the "${CONFIG.targetSheet}" sheet.`);
    return;
  }

  // 預先載入索引 Map (Memory Optimization)
  const lookupMap = buildLookupMap();
  if (!lookupMap) return;

  const userName   = getCurrentUserName();
  const sourceSheet = ss.getSheetByName(CONFIG.sourceSheet);
  if (!sourceSheet) return;
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  const range   = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows  = range.getNumRows();

  // 批次讀取 A、B 欄，優化 API 呼叫效率
  const colAB     = sheet.getRange(startRow, 1, numRows, 2).getValues();
  const colADisp  = sheet.getRange(startRow, 1, numRows, 1).getDisplayValues();

  let maxSuffix   = getMaxSuffix(sheet);
  const tasks     = [];

  // --- 第一階段：過濾待處理清單 ---
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= 1) continue;

    const nameValue = colAB[i][1]?.toString().trim();
    if (!nameValue || colADisp[i][0] !== "") continue;

    maxSuffix++;
    tasks.push({ row: currentRow, name: nameValue, fullId: CONFIG.idPrefix + maxSuffix });
  }

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ No pending rows detected.");
    return;
  }

  // --- 第二階段：資料轉換與自動化操作 ---
  const errors = [];
  const src    = CONFIG.sourceSheet;

  for (const task of tasks) {
    try {
      const { rows, totalAmount } = filterSourceRows(sourceData, sourceHeaders, task.name);
      const folderUrl = createFolderAndSheet(task.fullId, task.name, rows, totalAmount);
      const { k, l }  = resolvePaymentInfo(lookupMap[task.name]);

      // 1. 回填序列號超連結 (連結至 Drive 資料夾)
      sheet.getRange(task.row, 1).setFormula(
        folderUrl ? `=HYPERLINK("${folderUrl}","${task.fullId}")` : String(task.fullId)
      );

      // 2. 跨表資料比對與公式注入 (C～L 欄)
      sheet.getRange(task.row, 3, 1, 10).setValues([[
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:D,3,0)&" and ${"COUNTIF"}"&COUNTIF('${src}'!B:B,TRIM(B${task.row}))&" Items","N/A"))`,
        "(See Detailed List)",
        `=IF(B${task.row}="",0,SUMIF('${src}'!B:B,TRIM(B${task.row}),'${src}'!G:G))`,
        "",
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:I,8,0),""))`,
        userName,
        false, false, // Status Checkboxes
        k, l
      ]]);

      sheet.getRange(task.row, 5).setNumberFormat("#,##0");
      sheet.getRange(task.row, 9, 1, 2).insertCheckboxes();

    } catch (e) {
      errors.push(`Row ${task.row} (${task.name}): ${e.message}`);
    }
  }

  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(`⚠️ Errors occurred:\n\n${errors.join("\n")}`);
  } else {
    SpreadsheetApp.getUi().alert(`✅ Processed ${tasks.length} rows successfully.`);
  }
}

// ============================================================
//  🔧 工具函數區
// ============================================================

function buildLookupMap() {
  try {
    const data = SpreadsheetApp.openById(CONFIG.externalSsId).getSheetByName(CONFIG.externalSheet).getDataRange().getValues();
    return data.slice(1).reduce((map, row) => {
      const name = row[2]?.toString().trim();
      if (name) map[name] = { type: row[7], address: row[8], email: row[9] };
      return map;
    }, {});
  } catch (e) { return null; }
}

function getMaxSuffix(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return CONFIG.startSuffix - 1;
  return sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat()
    .filter(v => v.includes(CONFIG.idPrefix))
    .reduce((max, v) => {
      const n = parseInt(v.replace(/\D/g, '').slice(-4), 10);
      return isNaN(n) ? max : Math.max(max, n);
    }, CONFIG.startSuffix - 1);
}

function filterSourceRows(sourceData, sourceHeaders, nameValue) {
  let totalAmount = 0;
  const header = CONFIG.exportCols.map(i => sourceHeaders[i]);
  const dataRows = sourceData.slice(1)
    .filter(row => row[1]?.toString() === nameValue)
    .map(row => {
      const amt = parseFloat(row[6]);
      if (!isNaN(amt)) totalAmount += amt;
      return CONFIG.exportCols.map(i => row[i]);
    });
  return { rows: [header, ...dataRows], totalAmount };
}

function createFolderAndSheet(fullId, nameValue, filteredRows, totalAmount) {
  const targetFolder = DriveApp.getFolderById(CONFIG.parentFolderId).createFolder(`${fullId}_${nameValue}`);
  if (filteredRows.length > 1) {
    const newFile  = SpreadsheetApp.create(`${fullId}_Manifest`);
    const newSheet = newFile.getSheets()[0];
    newSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
    newSheet.getRange(filteredRows.length + 1, 6).setValue("Total");
    newSheet.getRange(filteredRows.length + 1, 7).setValue(totalAmount).setNumberFormat("#,##0");
    DriveApp.getFileById(newFile.getId()).moveTo(targetFolder);
  }
  return targetFolder.getUrl();
}

function resolvePaymentInfo(info) {
  if (!info) return { k: "", l: "" };
  if (info.type.includes("Digital")) return { k: info.email, l: "" };
  if (info.type.includes("Paper"))   return { k: "Physical Mail", l: info.address };
  return { k: "", l: "" };
}

function getCurrentUserName() {
  const email = Session.getEffectiveUser().getEmail().split('@')[0];
  return CONFIG.userList[email] || email;
}

// ============================================================
//  📥 CSV 匯出引擎 (針對 Adobe Illustrator 變數資料優化)
// ============================================================
function downloadIllustratorCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  if (startRow <= 1 && numRows === 1) {
    SpreadsheetApp.getUi().alert("❌ Please select the data rows you wish to export.");
    return;
  }

  // Illustrator Schema: 14 定義欄位 (需與 AI 變數面板匹配)
  const headers = [
    "Client_Name_1", "Client_Name_2", "Approval_ID", "Capacity", 
    "Amount_1", "Amount_2", "Account_ID_1", "Account_ID_2", 
    "barcode1", "barcode2", "Fiscal_Period", "ZipCode", "Shipping_Address", "Issue_Date"
  ];

  // 使用 getDisplayValues() 確保純文字，防止科學記號 (e.g., 3.75E+09)
  const rawData = sheet.getRange(startRow, 1, numRows, 7).getDisplayValues();
  
  let csvContent = "\ufeff"; // UTF-8 BOM
  csvContent += headers.map(h => `"${h}"`).join(",") + "\n";

  for (let i = 0; i < rawData.length; i++) {
    if (startRow + i === 1) continue; 
    
    const [accId, name, info, cap, amt, zip, addr] = rawData[i];
    const accStr = accId.toString().trim();
    
    // 資料清洗與轉換邏輯
    const b1 = accStr.length >= 11 ? accStr.slice(-11) : accStr; 
    const b2 = amt.replace(/,/g, ''); 

    const mappedRow = [
      name, name, info, cap, amt, amt, accId, accId,
      b1, b2, "FY2026_Services", zip, addr, ""
    ];

    csvContent += mappedRow.map(cell => `"${cell.replace(/"/g, '""')}"`).join(",") + "\n";
  }

  const fileName = "Export_AI_Variables_" + Utilities.formatDate(new Date(), "GMT+8", "MMdd_HHmm") + ".csv";
  const base64 = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  const html = `<script>const a=document.createElement('a');a.href='data:text/csv;base64,${base64}';a.download='${fileName}';a.click();setTimeout(()=>google.script.host.close(),1500);</script><body style="font-family:sans-serif;text-align:center;">Preparing Download...</body>`;
  
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150), "Exporting CSV");
}
