// =================================================================
// 【1-A_メイン基盤と共通関数】
// 全体の進行管理（大司令塔からの呼び出し）と、共通ヘルパー関数群
// =================================================================

function onEdit_processProgressSheet(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== SHEET_NAME_PROGRESS || range.getColumn() !== CHECKBOX_COLUMN || range.getValue() !== true) {
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
    if (!archiveSheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i > 0; i--) {
      const isChecked = data[i][CHECKBOX_COLUMN - 1];
      if (isChecked === true) {
        const rowNum = i + 1;
        let rowData = [...data[i]];
        rowData[CHECKBOX_COLUMN - 1] = new Date();
        archiveSheet.appendRow(rowData);
        archiveSheet.getRange(archiveSheet.getLastRow(), 1, 1, archiveSheet.getLastColumn()).setHorizontalAlignment('left');
        safeDeleteRow(sheet, rowNum);
      }
    }
  } catch (error) {
    console.error('onEdit 処理エラー: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

function mainProcessEmails_internal() {
  Logger.log("本番処理：メール処理を開始します。");
  let processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabel) processedLabel = GmailApp.createLabel(PROCESSED_LABEL);
  
  const querySuffix = ` newer_than:2d`;
  const confirmQuerySuffix = ` newer_than:14d`;

  processM3(querySuffix, confirmQuerySuffix, processedLabel);
  Utilities.sleep(1500);
  processMStage(querySuffix, confirmQuerySuffix, processedLabel);
  Utilities.sleep(1500);
  processMRT_Application(querySuffix, processedLabel);
  Utilities.sleep(1500);
  processMinkanIkyoku_Application(querySuffix, processedLabel);
  Utilities.sleep(1500);
  processMediwel_Application(querySuffix, processedLabel);
  Utilities.sleep(1500);
  
  // ★新規追加：ATMSメールの検知・通知（別ファイルの関数を呼び出し）
  if (typeof checkAtmsEmails_internal === 'function') {
    checkAtmsEmails_internal();
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sortSheetByDate_internal(ss.getSheetByName(SHEET_NAME_PROGRESS));
  Logger.log("本番処理：メール処理が完了しました。");
}

// ====================================
// 共通ヘルパー関数群
// ====================================

function safeDeleteRow(sheet, rowNum) {
  try {
    if (sheet.getMaxRows() <= 2) {
      sheet.insertRowAfter(1); 
    }
    sheet.deleteRow(rowNum);
  } catch(e) {
    try {
      sheet.insertRowAfter(1);
      sheet.deleteRow(rowNum);
    } catch(e2) {
      console.error("行の削除に失敗しました (rowNum: " + rowNum + ") - " + e2.message);
    }
  }
}

let mailProcessExistingKeysSet = null;

function initExistingKeys(sheet) {
  mailProcessExistingKeysSet = new Set();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);

  const addDataToSet = (data) => {
    for (let i = 1; i < data.length; i++) {
      const agency = String(data[i][0]).trim();
      if (!agency) continue;
      const dateObj = new Date(data[i][1]);
      const clinic = String(data[i][3]).trim();
      const jobId = String(data[i][5]).trim();
      const doctor = String(data[i][6]).split(/[\/／（(]/)[0].replace(/\s/g, '').trim();

      let dateStr = '';
      if (dateObj instanceof Date && !isNaN(dateObj.getTime()) && dateObj.getFullYear() > 1900) {
        dateStr = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd');
      } else {
        dateStr = String(data[i][1]).trim();
      }

      const key = jobId ? `${agency}_${jobId}` : `${agency}_${dateStr}_${clinic}_${doctor}`;
      mailProcessExistingKeysSet.add(key);
    }
  };

  addDataToSet(sheet.getDataRange().getValues());
  if (archiveSheet) {
    addDataToSet(archiveSheet.getDataRange().getValues());
  }
}

function appendRowAndFormatUnique(sheet, rowData) {
  if (!mailProcessExistingKeysSet) initExistingKeys(sheet);
  
  const agency = String(rowData[0]).trim();
  const dateStr = String(rowData[1]).trim();
  const clinic = String(rowData[3]).trim();
  const jobId = String(rowData[5]).trim();
  const doctor = String(rowData[6]).split(/[\/／（(]/)[0].replace(/\s/g, '').trim();
  
  const key = jobId ? `${agency}_${jobId}` : `${agency}_${dateStr}_${clinic}_${doctor}`;
  
  if (mailProcessExistingKeysSet.has(key)) return false; 
  
  mailProcessExistingKeysSet.add(key);
  appendRowAndFormat(sheet, rowData);
  return true;
}

function appendRowAndFormat(sheet, rowData) {
  const newRowIndex = Math.max(sheet.getLastRow() + 1, 2);
  const newRow = sheet.getRange(newRowIndex, 1, 1, rowData.length);
  newRow.setValues([rowData]);
  sheet.getRange(newRowIndex, 2).setHorizontalAlignment('left');
  sheet.getRange(newRowIndex, CHECKBOX_COLUMN).insertCheckboxes();
  
  const sourceValidation = sheet.getRange("M2").getDataValidation();
  if (sourceValidation) {
    sheet.getRange(newRowIndex, 13).setDataValidation(sourceValidation);
  }
}

function extract(text, regex) { 
  if (!text) return ''; 
  const match = text.match(regex); 
  return match ? match[1].trim() : ''; 
}

function formatDate(dateString, receivedDate) {
  if (!dateString) return '';
  let str = dateString.replace(/\(.\)/g, '').trim();
  let year, month, day;

  if (str.match(/^\d{4}[./年]/)) {
    let parts = str.split(/[./年日月]/).filter(s => s);
    year = parseInt(parts[0], 10); month = parseInt(parts[1], 10); day = parseInt(parts[2], 10);
  } else if (str.match(/^\d{1,2}[./月]/)) {
    let parts = str.split(/[./月日月]/).filter(s => s);
    year = receivedDate.getFullYear(); month = parseInt(parts[0], 10); day = parseInt(parts[1], 10);
    const receivedMonth = receivedDate.getMonth() + 1;
    if (month < receivedMonth) { year++; }
  } else {
    return dateString;
  }

  if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
    return `${year}/${('0' + month).slice(-2)}/${('0' + day).slice(-2)}`;
  }
  return dateString;
}

function findDepartmentFromHistory(ss, doctorName) {
    if (!doctorName) return '';
    const nameToFind = doctorName.replace(/\s/g, '');
    const sheetsToSearch = [ss.getSheetByName(SHEET_NAME_ARCHIVE), ss.getSheetByName(SHEET_NAME_PROGRESS)];
    for (const sheet of sheetsToSearch) {
        if (sheet) {
            const data = sheet.getDataRange().getValues();
            for (let i = data.length - 1; i > 0; i--) {
                const rowDoctorNameRaw = String(data[i][6]);
                const rowDoctorNameClean = rowDoctorNameRaw.split(/[\/／（(]/)[0].replace(/\s/g, '');
                const department = String(data[i][4]);
                if (rowDoctorNameClean === nameToFind && department) return department;
            }
        }
    }
    return '';
}

function extractM3DoctorName(subject) {
    const namePartWithId = subject.match(/C\d+\s(.+?)先生/);
    if (namePartWithId && namePartWithId[1]) return namePartWithId[1].trim();
    const namePartWithoutId = subject.match(/\d+月\d+日\s*(.+?)先生/);
    if (namePartWithoutId && namePartWithoutId[1]) return namePartWithoutId[1].trim();
    const namePart = subject.split('先生')[0];
    const words = namePart.trim().split(/\s+/);
    return words[words.length - 1];
}

function sortSheetByDate_internal(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    range.sort({column: 2, ascending: true});
  }
}