/**
 * @Project: Automated Enterprise Billing & Design Asset Pipeline
 * @Description: An End-to-End system for Google Workspace that automates data matching, 
 * cloud filing, and generates sanitized CSVs for Adobe Illustrator variable data.
 */
// ============================================================
//  ⚙️  Global Steeings (Anonymized: Using generic variables and placeholders)
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
//  📌  Menu system
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
//  🚀  Core Execution Logic (Data Flow Automation & Cloud Archiving)
// ============================================================
function processCurrentRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.targetSheet) {
    SpreadsheetApp.getUi().alert(`❌ Please run this from the "${CONFIG.targetSheet}" sheet.``);
    return;
  }

  // Pre-load index Map (Memory Optimization to replace traditional VLOOKUPs)
  const lookupMap  = buildLookupMap();
  if (!lookupMap) return;

  const userName   = getCurrentUserName();
  const sourceData = ss.getSheetByName(CONFIG.sourceSheet).getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  const range    = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows  = range.getNumRows();

  // Batch read columns A & B to minimize API calls and improve performance
  const colAB    = sheet.getRange(startRow, 1, numRows, 2).getValues();
  const colADisp = sheet.getRange(startRow, 1, numRows, 1).getDisplayValues();

  let maxSuffix  = getMaxSuffix(sheet);
  const tasks    = [];

  // Phase 1: Identify and Filter Pending Rows
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= 1) continue;

    const nameValue = colAB[i][1]?.toString().trim();
    if (!nameValue || colADisp[i][0] !== "") continue;

    maxSuffix++;
    tasks.push({ row: currentRow, name: nameValue, fullId: CONFIG.idPrefix + maxSuffix });
  }

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ No pending rows detected in the selection.");
    return;
  }

  // Phase 2: Execution (Data Transformation & Folder Generation)
  const errors = [];
  const src    = CONFIG.sourceSheet;

  for (const task of tasks) {
    try {
      const { rows, totalAmount } = filterSourceRows(sourceData, sourceHeaders, task.name);
      const folderUrl = createFolderAndSheet(task.fullId, task.name, rows, totalAmount);
      const { k, l }  = resolvePaymentInfo(lookupMap[task.name]);

      // Write Serial ID with Hyperlink (Links directly to the generated Drive folder)
      sheet.getRange(task.row, 1).setFormula(
        folderUrl
          ? `=HYPERLINK("${folderUrl}","${task.fullId}")`
          : String(task.fullId)
      );

      // Cross-sheet data matching and dynamic formula injection (Columns C to L)
      sheet.getRange(task.row, 3, 1, 10).setValues([[
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:D,3,0)&"等共"&COUNTIF('${src}'!B:B,TRIM(B${task.row}))&"案","無資料"))`,
        "(詳見案件清冊)",
        `=IF(B${task.row}="",0,SUMIF('${src}'!B:B,TRIM(B${task.row}),'${src}'!G:G))`,
        "",
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:I,8,0),""))`,
        userName,
        false, false,
        k, l
      ]]);

      sheet.getRange(task.row, 5).setNumberFormat("#,##0");
      sheet.getRange(task.row, 9, 1, 2).insertCheckboxes();

    } catch (e) {
      errors.push(`第 ${task.row} 列（${task.name}）：${e.message}`);
    }
  }

  // Final Execution Report
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(
      `⚠️ Partial processing errors：\n\n${errors.join("\n")}`
    );
  } else {
    SpreadsheetApp.getUi().alert(`✅ Processed ${tasks.length} records successfully.`);
  }
}

// ============================================================
//  🔧  Utility Functions
// ============================================================

// Builds an in-memory Hash Map from external data for O(1) lookup performance
function buildLookupMap() {
  try {
    const data = SpreadsheetApp
      .openById(CONFIG.externalSsId)
      .getSheetByName(CONFIG.externalSheet)
      .getDataRange().getValues();

    return data.slice(1).reduce((map, row) => {
      const name = row[2]?.toString().trim();
      if (name) map[name] = { type: row[7], address: row[8], email: row[9] };
      return map;
    }, {});
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ Failed to read external database：${e.message}`);
    return null;
  }
}

// Scans Column A to determine the current maximum serial suffix
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

// Filters master data based on Client Name and calculates aggregate totals
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

// Creates a dedicated folder and a detailed manifest spreadsheet in Google Drive
function createFolderAndSheet(fullId, nameValue, filteredRows, totalAmount) {
  const totalCount  = filteredRows.length - 1;
  const folderName  = `${fullId}_${nameValue}(${totalCount}in total)`;
  const targetFolder = DriveApp.getFolderById(CONFIG.parentFolderId).createFolder(folderName);

  if (totalCount > 0) {
    const newFile  = SpreadsheetApp.create(folderName);
    const newSheet = newFile.getSheets()[0];
    newSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
    newSheet.getRange(filteredRows.length + 1, 6).setValue("Total Amount");
    newSheet.getRange(filteredRows.length + 1, 7).setValue(totalAmount).setNumberFormat("#,##0");
    DriveApp.getFileById(newFile.getId()).moveTo(targetFolder);
  }

  return targetFolder.getUrl();
}

//Resolves payment method and routing information 

function resolvePaymentInfo(info) {
  if (!info) return { k: "", l: "" };
  if (info.type.includes("email")) return { k: info.email, l: "" };
  if (info.type === "Physical")  return { k: "Physical Mail", l: info.address };
  return { k: "", l: "" };
}

//Identifies the current operator based on user login data
function getCurrentUserName() {
  const email = Session.getEffectiveUser().getEmail().split('@')[0];
  return CONFIG.userList[email] || email;
}

// ============================================================
//  📮  ZipCode Sidebar UI
// ============================================================
function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('📮 Zipcode Lookup Tool')
      .setWidth(320)
  );
}

function getSelectedCellInfo() {
  const r = SpreadsheetApp.getActiveSheet().getActiveRange();
  return { row: r.getRow(), col: r.getColumn(), address: String(r.getValue()) };
}

function writeZipCode(row, col, zipcode) {
  SpreadsheetApp.getActiveSheet().getRange(row, col - 1).setValue(zipcode);
}

function fetchZipCode(address) {
  const url  = 'https://zipcode.tw/' + encodeURIComponent(address);
  const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText('UTF-8');
  const match = html.match(/3\+3[\s\S]*?(\d{6})/);
  return match ? match[1] : null;
}


// ============================================================
//  📥  CSV Export Engine (Optimized for Adobe Illustrator)
// ============================================================
function downloadIllustratorCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  if (startRow <= 1 && numRows === 1) {
    SpreadsheetApp.getUi().alert("❌ Selection Error: Please highlight data rows (excluding headers).");
    return;
  }

  // Illustrator Variables Schema: Must match the variable names in the AI Variables Panel
  const headers = [
    "Client_Name_1", "Client_Name_2", "Case_ID", "Capacity", 
    "Amount_1", "Amount_2", "Account_ID_1", "Account_ID_2", 
    "barcode1", "barcode2", "Period", "ZipCode", "Shipping_Address", "Issue_Date"
  ];

  // 🛡️ Critical Optimization: Use getDisplayValues() to capture formatted strings.
  // This prevents long IDs or large amounts from converting to scientific notation (e.g., 3.75E+09).
  const rawData = sheet.getRange(startRow, 1, numRows, 7).getDisplayValues();
  
  // Initialize CSV with UTF-8 BOM (\ufeff) to resolve character encoding issues in Adobe software
  let csvContent = "\ufeff"; 
  csvContent += headers.map(h => `"${h}"`).join(",") + "\n";

  // Data Transformation & Mapping
  for (let i = 0; i < rawData.length; i++) {
    if (startRow + i === 1) continue; // Skip sheet header
    
    // De-structure original columns: A:Account, B:Name, C:Info, D:Capacity, E:Amount, F:Zip, G:Address
    const [accId, name, info, cap, amt, zip, addr] = rawData[i];
    
    const accStr = accId.toString().trim();
    // Business Logic: Extract last 11 digits for Barcode 1.
    const b1 = accStr.length >= 11 ? accStr.slice(-11) : accStr; 
    // Business Logic: Sanitize amount for Barcode 2
    const b2 = amt.replace(/,/g, ''); 

    // Mapping logic (1-to-Many mapping from Sheets to Illustrator Schema)
    const mappedRow = [
      name,             // Client_Name_1
      name,             // Client_Name_2
      info,             // Case_ID
      cap,              // Capacity
      amt,              // Amount_1
      amt,              // Amount_2
      accId,            // Account_ID_1
      accId,            // Account_ID_2
      b1,               // barcode1
      b2,               // barcode2
      "", // Period
      zip,              // Zipcode
      addr,             // Shipping_Address
      ""                // Issue_Date
    ];

    csvContent += mappedRow.map(cell => `"${cell.replace(/"/g, '""')}"`).join(",") + "\n";
  }

  // Trigger Browser-side Download
  const fileName = "AI_Final_" + Utilities.formatDate(new Date(), "GMT+8", "MMdd_HHmm") + ".csv";
  const base64 = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  const html = `<script>const a=document.createElement('a');a.href='data:text/csv;base64,${base64}';a.download='${fileName}';a.click();setTimeout(()=>google.script.host.close(),1500);</script><body style="font-family:sans-serif;text-align:center;">檔案產生中...</body>`;
  
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150), "正在下載 CSV");
}
