/**
 * @Project: 自動化企業帳單與設計資產流水線
 * @Description: 一套 Google Workspace 端對端系統，自動化處理資料比對、
 * 雲端歸檔，並產出適用於 Adobe Illustrator 變數資料的清理版 CSV 檔案。
 */

// ============================================================
//  ⚙️  全域設定（已去識別化：使用通用變數與佔位符）
// ============================================================
const CONFIG = {
  targetSheet:    "Operations_Dashboard",      // 主作業目標工作表名稱
  sourceSheet:    "Central_Master_Database",   // 來源主資料工作表名稱
  parentFolderId: "YOUR_DRIVE_FOLDER_ID",      // Drive 父資料夾 ID（佔位符）
  externalSsId:   "YOUR_EXTERNAL_DATABASE_ID", // 外部資料庫試算表 ID（佔位符）
  externalSheet:  "External_Response_Logs",    // 外部資料庫內的工作表名稱
  idPrefix:       "123456",                    // 流水序號前綴
  startSuffix:    1000,                        // 流水序號起始數字
  exportCols:     [0, 1, 2, 3, 4, 5, 6, 7, 8],// 清冊匯出的欄位索引（從 0 起算）
  userList: {                                  // 帳號使用者名稱 → 顯示名稱對應表（已遮蔽）
    "admin.user1": "Account Manager A",
    "admin.user2": "Project Specialist B",
    "admin.user3": "Operations Lead C",
    "admin.user4": "System Admin D"
  }
};

// ============================================================
//  📌  選單系統
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ 案件自動化')
    .addItem('🚀 執行：生成流水號與資料夾', 'processCurrentRows')
    .addItem('📮 開啟郵遞區號查詢側邊欄', 'showSidebar')
    .addItem('📥 下載 Illustrator 專用 CSV', 'downloadIllustratorCSV')
    .addToUi();
}

// ============================================================
//  🚀  核心執行邏輯（資料流自動化與雲端歸檔）
// ============================================================
function processCurrentRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 確認使用者在正確的工作表上執行
  if (sheet.getName() !== CONFIG.targetSheet) {
    SpreadsheetApp.getUi().alert(`❌ 請在「${CONFIG.targetSheet}」分頁執行此功能`);
    return;
  }

  // 預先建立索引 Map（記憶體優化，取代傳統 VLOOKUP）
  const lookupMap  = buildLookupMap();
  if (!lookupMap) return;

  const userName   = getCurrentUserName();
  const sourceData = ss.getSheetByName(CONFIG.sourceSheet).getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  const range    = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows  = range.getNumRows();

  // 批次讀取 A、B 欄，減少 API 呼叫次數以提升效能
  const colAB    = sheet.getRange(startRow, 1, numRows, 2).getValues();
  const colADisp = sheet.getRange(startRow, 1, numRows, 1).getDisplayValues();

  let maxSuffix  = getMaxSuffix(sheet);
  const tasks    = [];

  // 第一階段：識別並篩選待處理列
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    if (currentRow <= 1) continue; // 跳過標題列

    const nameValue = colAB[i][1]?.toString().trim();
    // 僅處理 B 欄有客戶名稱且 A 欄尚未填入序號的列
    if (!nameValue || colADisp[i][0] !== "") continue;

    maxSuffix++;
    tasks.push({ row: currentRow, name: nameValue, fullId: CONFIG.idPrefix + maxSuffix });
  }

  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ 沒有需要處理的列。");
    return;
  }

  //第二階段：執行（資料轉換與資料夾建立)

  const errors = [];
  const src    = CONFIG.sourceSheet;

  for (const task of tasks) {
    try {
      const { rows, totalAmount } = filterSourceRows(sourceData, sourceHeaders, task.name);
      const folderUrl = createFolderAndSheet(task.fullId, task.name, rows, totalAmount);
      const { k, l }  = resolvePaymentInfo(lookupMap[task.name]);

      // 1. 回填流水序號超連結（直接連結至已建立的 Drive 資料夾）
      sheet.getRange(task.row, 1).setFormula(
        folderUrl
          ? `=HYPERLINK("${folderUrl}","${task.fullId}")`
          : String(task.fullId)
      );

      // 2. 跨表資料比對與動態公式注入（C 欄至 L 欄）
      sheet.getRange(task.row, 3, 1, 10).setValues([[
        // C欄：VLOOKUP 查找客戶資訊，並統計案件數量
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:D,3,0)&"等共"&COUNTIF('${src}'!B:B,TRIM(B${task.row}))&"案","無資料"))`,
        // D欄：提示查看詳細清冊
        "(詳見案件清冊)",
        // E欄：SUMIF 加總該客戶所有金額
        `=IF(B${task.row}="",0,SUMIF('${src}'!B:B,TRIM(B${task.row}),'${src}'!G:G))`,
        // F欄：保留空白
        "",
        // G欄：VLOOKUP 查找其他客戶欄位資訊
        `=IF(B${task.row}="","",IFERROR(VLOOKUP(TRIM(B${task.row}),'${src}'!B:I,8,0),""))`,
        userName, // H欄：操作者顯示名稱
        false, false, // I、J欄：狀態核取方塊（預設未勾選）
        k, l          // K、L欄：付款路由資訊（Email 或實體地址）
      ]]);

      sheet.getRange(task.row, 5).setNumberFormat("#,##0"); // E欄設定千分位格式
      sheet.getRange(task.row, 9, 1, 2).insertCheckboxes(); // I、J欄插入核取方塊

    } catch (e) {
      errors.push(`第 ${task.row} 列（${task.name}）：${e.message}`);
    }
  }

  // --- 最終執行報告 ---
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(
      `⚠️ 部分列處理失敗：\n\n${errors.join("\n")}`
    );
  } else {
    SpreadsheetApp.getUi().alert(`✅ 處理完成！共處理 ${tasks.length} 筆。`);
  }
}

// ============================================================
//  🔧  工具函式
// ============================================================

// 從外部資料建立記憶體內 Hash Map，達成 O(1) 查找效能
function buildLookupMap() {
  try {
    const data = SpreadsheetApp
      .openById(CONFIG.externalSsId)
      .getSheetByName(CONFIG.externalSheet)
      .getDataRange().getValues();

    // 以第 3 欄（index 2）的客戶名稱為鍵，建立付款路由索引
    return data.slice(1).reduce((map, row) => {
      const name = row[2]?.toString().trim();
      if (name) map[name] = { type: row[7], address: row[8], email: row[9] };
      return map;
    }, {});
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 無法讀取外部資料庫：${e.message}`);
    return null;
  }
}

// 掃描 A 欄，取得目前最大流水序號尾碼
function getMaxSuffix(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return CONFIG.startSuffix - 1;

  return sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat()
    .filter(v => v.includes(CONFIG.idPrefix))
    .reduce((max, v) => {
      // 取末 4 碼轉為整數，與目前最大值比較
      const n = parseInt(v.replace(/\D/g, '').slice(-4), 10);
      return isNaN(n) ? max : Math.max(max, n);
    }, CONFIG.startSuffix - 1);
}

// 依客戶名稱篩選來源資料，並加總金額小計
function filterSourceRows(sourceData, sourceHeaders, nameValue) {
  let totalAmount = 0;
  const header = CONFIG.exportCols.map(i => sourceHeaders[i]);
  const dataRows = sourceData.slice(1)
    .filter(row => row[1]?.toString() === nameValue) // 比對 B 欄客戶名稱
    .map(row => {
      const amt = parseFloat(row[6]); // G 欄為金額欄位
      if (!isNaN(amt)) totalAmount += amt;
      return CONFIG.exportCols.map(i => row[i]);
    });

  return { rows: [header, ...dataRows], totalAmount };
}

// 在 Google Drive 建立專屬資料夾與詳細清冊試算表
function createFolderAndSheet(fullId, nameValue, filteredRows, totalAmount) {
  const totalCount  = filteredRows.length - 1; // 扣除標題列後的實際資料列數
  const folderName  = `${fullId}_${nameValue}(共${totalCount}案)`;
  const targetFolder = DriveApp.getFolderById(CONFIG.parentFolderId).createFolder(folderName);

  if (totalCount > 0) {
    // 建立同名清冊試算表並寫入篩選後的資料
    const newFile  = SpreadsheetApp.create(folderName);
    const newSheet = newFile.getSheets()[0];
    newSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
    // 在資料最後一列下方加入合計列
    newSheet.getRange(filteredRows.length + 1, 6).setValue("總金額");
    newSheet.getRange(filteredRows.length + 1, 7).setValue(totalAmount).setNumberFormat("#,##0");
    // 將試算表移至目標資料夾
    DriveApp.getFileById(newFile.getId()).moveTo(targetFolder);
  }

  return targetFolder.getUrl();
}

// 依客戶類型解析付款方式與路由資訊
function resolvePaymentInfo(info) {
  if (!info) return { k: "", l: "" };
  if (info.type.includes("電子"))  return { k: info.email, l: "" };        // 電子郵件寄送
  if (info.type === "紙本繳費單")     return { k: "紙本繳費單", l: info.address }; // 實體郵寄
  return { k: "", l: "" }; // 無法識別的類型，回傳空值
}

// 依登入帳號識別當前操作者，並對應至顯示名稱
function getCurrentUserName() {
  const email = Session.getEffectiveUser().getEmail().split('@')[0];
  return CONFIG.userList[email] || email; // 若不在名單內，直接使用帳號名稱
}

// ============================================================
//  📮  郵遞區號查詢側欄
// ============================================================
function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('📮 郵遞區號查詢')
      .setWidth(320)
  );
}

// 取得目前選取儲存格的位置與內容，供側欄 JS 呼叫
function getSelectedCellInfo() {
  const r = SpreadsheetApp.getActiveSheet().getActiveRange();
  return { row: r.getRow(), col: r.getColumn(), address: String(r.getValue()) };
}

// 將查詢結果寫入選取儲存格左側相鄰格
function writeZipCode(row, col, zipcode) {
  SpreadsheetApp.getActiveSheet().getRange(row, col - 1).setValue(zipcode);
}

// 透過 zipcode.tw 查詢台灣 6 位數郵遞區號
function fetchZipCode(address) {
  const url  = 'https://zipcode.tw/' + encodeURIComponent(address);
  const html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText('UTF-8');
  // 以正規表達式從頁面 HTML 中比對 3+3 格式的郵遞區號
  const match = html.match(/3\+3[\s\S]*?(\d{6})/);
  return match ? match[1] : null;
}

// ============================================================
//  📥  CSV 匯出引擎（針對 Adobe Illustrator 變數資料優化）
// ============================================================
function downloadIllustratorCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  // 防呆：確認使用者有選取資料列而非僅停在標題
  if (startRow <= 1 && numRows === 1) {
    SpreadsheetApp.getUi().alert("❌ 請先反白選取想要下載的資料列（不含標題）。");
    return;
  }

  // Illustrator 變數面板欄位結構（需與 AI 範本內的變數名稱完全一致）
  const headers = [
    "設置者名稱1", "設置者名稱2", "同意備案編號", "裝置容量", 
    "金額1", "金額2", "繳款帳號1", "繳款帳號2", 
    "barcode1", "barcode2", "期數", "郵遞區號", "收件地址", "繳款日期"
  ];

  // 🛡️ 關鍵優化：使用 getDisplayValues() 擷取格式化字串
  // 防止長 ID 或大金額被轉換為科學記號（例如 1.23E+09），確保條碼資料完整
  const rawData = sheet.getRange(startRow, 1, numRows, 7).getDisplayValues();
  
  // 初始化 CSV，注入 UTF-8 BOM（\ufeff）
  // 解決 Adobe 軟體匯入時的字元編碼問題（繁體中文亂碼）
  let csvContent = "\ufeff"; 
  csvContent += headers.map(h => `"${h}"`).join(",") + "\n";

  // --- 資料轉換與欄位映射 ---
  for (let i = 0; i < rawData.length; i++) {
    if (startRow + i === 1) continue; // 跳過工作表標題列
    
    // 解構原始欄位：A:帳號、B:姓名、C:案件資訊、D:容量、E:金額、F:郵遞區號、G:地址
    const [accId, name, info, cap, amt, zip, addr] = rawData[i];
    
    const accStr = accId.toString().trim();
    // 條碼邏輯 1：取帳號 ID 末 11 碼（不足 11 碼則完整使用）
    const b1 = accStr.length >= 11 ? accStr.slice(-11) : accStr; 
    // 條碼邏輯 2：移除金額中的千分位逗號，確保為純數字字串
    const b2 = amt.replace(/,/g, ''); 

    // 欄位映射（來源 Sheets 欄位 → Illustrator 變數面板，支援一對多映射）
    const mappedRow = [
      name,             // 設置者名稱 1
      name,             // 設置者名稱 2
      info,             // 同意備案編號
      cap,              // 裝置容量
      amt,              // 金額 1 (保留千分位)
      amt,              // 金額 2
      accId,            // 繳款帳號 1
      accId,            // 繳款帳號 2
      b1,               // barcode1 (帳號後11碼)
      b2,               // barcode2 (純數字金額)
      "",               // 固定期數
      zip,              // 郵遞區號
      addr,             // 收件地址
      ""                // 繳款日期 (空白)
    ];

    // 對儲存格內的雙引號進行跳脫處理，確保 CSV 格式正確
    csvContent += mappedRow.map(cell => `"${cell.replace(/"/g, '""')}"`).join(",") + "\n";
  }

  // --- 觸發瀏覽器端下載 ---
  const fileName = "AI_Final_" + Utilities.formatDate(new Date(), "GMT+8", "MMdd_HHmm") + ".csv";
  const base64 = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  const html = `<script>const a=document.createElement('a');a.href='data:text/csv;base64,${base64}';a.download='${fileName}';a.click();setTimeout(()=>google.script.host.close(),1500);</script><body style="font-family:sans-serif;text-align:center;">檔案產生中...</body>`;
  
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150),
    "正在下載 CSV"
  );
}
