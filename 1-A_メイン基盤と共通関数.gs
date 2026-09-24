// =================================================================
// 【1-A_メイン基盤と共通関数】
// 全体の進行管理（大司令塔からの呼び出し）と、共通ヘルパー関数群
// =================================================================

function onEdit_processProgressSheet(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== SHEET_NAME_PROGRESS) return;

  const saiyoColIndex = getColIndex_internal(sheet, '採用可否');

  if (saiyoColIndex > 0 && range.getColumn() === saiyoColIndex) {
    const val = range.getValue();
    if (val === '採用') range.setBackground('#b6d7a8');
    else if (val === '不採用') range.setBackground('#ffe599');
    else range.setBackground(null);
    return;
  }

  if (range.getColumn() !== CHECKBOX_COLUMN || range.getValue() !== true) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  if (!archiveSheet) return;

  const ui = SpreadsheetApp.getUi();
  const doctorColIndex = getColIndex_internal(sheet, '医師名');
  const rowIndex = range.getRow();
  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  const agency = String(rowData[0]).trim();
  const saiyoStatus = saiyoColIndex > 0 ? String(rowData[saiyoColIndex - 1]).trim() : '';
  const doctorName = doctorColIndex > 0 ? String(rowData[doctorColIndex - 1]).trim() : '';

  if (!saiyoStatus) {
    ui.alert('【エラー】', '採用可否が選択されていません。採用または不採用を選択してください。', ui.ButtonSet.OK);
    range.setValue(false);
    return;
  }
  if (agency === '民間医局' && doctorName.includes('ドクター')) {
    ui.alert('【エラー】', '医師名を正しくいれてください。', ui.ButtonSet.OK);
    range.setValue(false);
    return;
  }
  if (agency === 'MRT' && !doctorName) {
    ui.alert('【エラー】', '医師名を確認していれてください。', ui.ButtonSet.OK);
    range.setValue(false);
    return;
  }

  // ★行ズレ巻き込み事故防止：0秒ロックで重複を弾く
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    return;
  }

  try {
    // ★行ズレ巻き込み事故防止：処理対象の行に本当にチェックが入っているか最終確認
    if (sheet.getRange(rowIndex, CHECKBOX_COLUMN).getValue() !== true) {
      return; 
    }

    rowData[CHECKBOX_COLUMN - 1] = new Date();
    appendRowToArchive(sheet, archiveSheet, rowData, saiyoStatus);
    safeDeleteRow(sheet, rowIndex);
  } catch (error) {
    console.error('onEdit 処理エラー: ' + error.toString());
    range.setValue(false);
  } finally {
    lock.releaseLock();
  }
}

function mainProcessEmails_internal() {
  Logger.log("本番処理：メール処理を開始します。");
  let processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabel) processedLabel = GmailApp.createLabel(PROCESSED_LABEL);
  
  // ★ バグの元凶「after:UNIX時間」を完全撤廃。安全な newer_than:1d に固定
  const querySuffix = ` newer_than:1d`;
  const confirmQuerySuffix = ` newer_than:1d`;

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
  
  if (typeof checkAtmsEmails_internal === 'function') {
    checkAtmsEmails_internal();
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sortSheetByDate_internal(ss.getSheetByName(SHEET_NAME_PROGRESS));
  Logger.log("本番処理：メール処理が完了しました。");
}

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
  
  const saiyoColIndex = getColIndex_internal(sheet, '採用可否');
  if (saiyoColIndex > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['採用', '不採用'], true).build();
    sheet.getRange(newRowIndex, saiyoColIndex).setDataValidation(rule);
  }
  
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

function getColIndex_internal(sheet, headerName) {
  if (!sheet) return -1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(headerName);
  return idx !== -1 ? idx + 1 : -1;
}

function appendRowToArchive(sheet, archiveSheet, rawValues, saiyoStatus) {
  const archSaiyoColIndex = getColIndex_internal(archiveSheet, '採用可否');
  const progSaiyoColIndex = sheet ? getColIndex_internal(sheet, '採用可否') : -1;
  const dateColIndex = getColIndex_internal(archiveSheet, '勤務希望日'); 

  let archiveData = new Array(archiveSheet.getLastColumn()).fill('');
  for (let i = 0; i < rawValues.length; i++) {
    if (i < archiveData.length) archiveData[i] = rawValues[i];
  }
  
  if (archSaiyoColIndex > 0) {
    archiveData[archSaiyoColIndex - 1] = saiyoStatus;
    if (progSaiyoColIndex > 0 && progSaiyoColIndex !== archSaiyoColIndex && progSaiyoColIndex - 1 < archiveData.length) {
      archiveData[progSaiyoColIndex - 1] = ''; 
    }
  }

  archiveSheet.appendRow(archiveData);
  const insertedRow = archiveSheet.getLastRow();
  archiveSheet.getRange(insertedRow, 1, 1, archiveSheet.getLastColumn()).setHorizontalAlignment('left');

  if (archSaiyoColIndex > 0) {
    const bgCell = archiveSheet.getRange(insertedRow, archSaiyoColIndex);
    if (saiyoStatus === '採用') bgCell.setBackground('#b6d7a8');
    else if (saiyoStatus === '不採用') bgCell.setBackground('#ffe599');
  }

  if (insertedRow > 2 && dateColIndex > 0) {
    const sortRange = archiveSheet.getRange(2, 1, insertedRow - 1, archiveSheet.getLastColumn());
    sortRange.sort({column: dateColIndex, ascending: true});
  }
}

// =================================================================
// ▼▼▼ jinjer勤怠 有給申請メール検知・処理スクリプト ▼▼▼
// =================================================================

function processJinjerPaidLeave_internal() {
  console.log('【jinjer有給申請検知】処理を開始します。');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('有給申請');
  if (!sheet) {
    console.error('シート「有給申請」が見つかりません。');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('JINJER_PAIDLEAVE_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch(e) {}

  // ★ API上限回避の最強ストッパー: newer_than:1d で検索し、最新30件のみ取得
  const query = `from:entry@kintai.jinjer.biz subject:"【jinjer勤怠】お知らせ_休日休暇申請が提出されました" newer_than:1d`;
  const threads = GmailApp.search(query, 0, 30);

  if (threads.length === 0) {
    console.log('jinjer勤怠からの新しい有給申請メールはありませんでした。');
    return;
  }

  const colDate = getColIndex_internal(sheet, '受信日');
  const colTarget = getColIndex_internal(sheet, '対象日');
  const colName = getColIndex_internal(sheet, '医師名');
  const colStatus = getColIndex_internal(sheet, '対応有無');

  if (colDate < 0 || colTarget < 0 || colName < 0 || colStatus < 0) {
    console.error('「有給申請」シートに必要なヘッダーが見つかりません。');
    return;
  }

  let isUpdated = false;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const messageId = message.getId();
      
      if (processedIds.includes(messageId)) continue;

      const body = message.getPlainBody();
      const receivedDateStr = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');

      const targetDateMatch = body.match(/申請対象日：\s*([^\n\r]+)/);
      const targetDate = targetDateMatch ? targetDateMatch[1].trim() : '';

      const applicantMatch = body.match(/申請者：\s*([^\n\r]+)/);
      const applicant = applicantMatch ? applicantMatch[1].trim() : '';

      if (targetDate && applicant) {
        const newRowIdx = Math.max(sheet.getLastRow() + 1, 3);
        const maxCol = sheet.getLastColumn();

        const templateRange = sheet.getRange(2, 1, 1, maxCol);
        const targetRange = sheet.getRange(newRowIdx, 1, 1, maxCol);
        
        templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

        sheet.getRange(newRowIdx, colDate).setValue(receivedDateStr);
        sheet.getRange(newRowIdx, colTarget).setValue(targetDate);
        sheet.getRange(newRowIdx, colName).setValue(applicant);
        sheet.getRange(newRowIdx, colStatus).setValue('未着手');

        console.log(`[有給申請追加] ${applicant} : ${targetDate}`);
      }

      processedIds.push(messageId);
      isUpdated = true;
    }
  }

  if (isUpdated) {
    if (processedIds.length > 200) {
      processedIds = processedIds.slice(-200);
    }
    props.setProperty('JINJER_PAIDLEAVE_IDS', JSON.stringify(processedIds));
  }

  console.log('【jinjer有給申請検知】処理が完了しました。');
}

// =================================================================
// ▼ 有給申請シートの自動グレーアウト処理 ▼
// =================================================================

function onEdit_processPaidLeaveSheet(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== '有給申請') return;

  if (range.getRow() < 3) return;

  const statusColIndex = getColIndex_internal(sheet, '対応有無');
  
  if (statusColIndex > 0 && range.getColumn() === statusColIndex) {
    const val = range.getValue();
    const rowRange = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());
    
    if (val === '済') {
      rowRange.setBackground('#d9d9d9'); // グレーアウト
    } else {
      rowRange.setBackground(null); // 色をリセット
    }
  }
}