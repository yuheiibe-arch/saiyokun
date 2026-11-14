// =================================================================
// 【1-C_その他紹介会社処理】（検索上限ガード ＆ ID押し出し防止版）
// エムステージ、MRT、民間医局、メディウェル
// =================================================================

// === エムステージ ===
function processMStage(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const now = new Date();

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('MSTAGE_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch (e) {}
  const processedSet = new Set(processedIds);
  let isIdUpdated = false;

  const markId = (id) => {
    if (!processedSet.has(id)) {
      processedSet.add(id);
      processedIds.push(id);
      isIdUpdated = true;
    }
  };

  // === 1. 応募通知の処理 ===
  const applyQuery = `from:(ra@mstage-corp.jp) subject:("貴院の求人に応募がありました") ${querySuffix}`;
  const applyThreads = GmailApp.search(applyQuery, 0, 50);

  if (applyThreads.length > 0) {
    const applyMessagesAll = GmailApp.getMessagesForThreads(applyThreads);

    for (let i = 0; i < applyThreads.length; i++) {
      const thread = applyThreads[i];
      const messages = applyMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const receivedDate = message.getDate();
        if ((now - receivedDate) / (1000 * 60 * 60) > 48) return;

        const messageId = message.getId();
        if (processedSet.has(messageId)) return;
        markId(messageId);

        const body = message.getPlainBody();
        const dateLine = extract(body, /勤務日[\s ]*[:：][\s ]*(.+)/);
        const doctorName = extract(body, /お名前\(漢字\)[\s ]*[:：][\s ]*(.+)/).replace(/[\s ]/g, '');
        let department = extract(body, /募集科目[\s ]*[:：][\s ]*(.+)/);
        if (!department) department = findDepartmentFromHistory(ss, doctorName) || '';

        const id = extract(body, /(?:求人番号|案件ID|求人ID)[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
        const clinic = extract(body, /キャップスクリニック([^\r\n]+)/).replace(/様$/, '').trim();
        const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})/);

        if (doctorName && id) {
          const added = appendRowAndFormatUnique(sheet, [
            'エムステージ', formatDate(dateLine, receivedDate), workTime,
            clinic, department,
            id,
            doctorName, extract(body, /(https:\/\/www\.doctor-navi\.jp\/medical\/entries.*)/),
            '', '応募通知受信', receivedDate, '', ''
          ]);
          if (added) threadProcessed = true;
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }

  // === 2. 確定・お断り・キャンセル通知の処理 ===
  const mStageConfirmQuery = `from:(ra@mstage-corp.jp) subject:("勤務が確定" OR "キャンセル" OR "書類選考の結果を通知しました(NG)") ${confirmQuerySuffix}`;
  const confirmThreads = GmailApp.search(mStageConfirmQuery, 0, 50);

  if (confirmThreads.length > 0) {
    const confirmMessagesAll = GmailApp.getMessagesForThreads(confirmThreads);
    const archSaiyoCol = getColIndex_internal(archiveSheet, '採用可否');
    const archStatusCol = getColIndex_internal(archiveSheet, '処理状況');

    for (let i = 0; i < confirmThreads.length; i++) {
      const thread = confirmThreads[i];
      const messages = confirmMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const receivedDate = message.getDate();
        if ((now - receivedDate) / (1000 * 60 * 60) > 72) return;

        const messageId = message.getId();
        if (processedSet.has(messageId)) return;
        markId(messageId);

        const subject = message.getSubject();
        const body = message.getPlainBody();

        let type = '';
        let status = '';
        if (subject.includes('勤務が確定')) { type = '確定'; status = '勤務確定'; }
        else if (subject.includes('キャンセル')) { type = 'お断り'; status = 'キャンセル済み'; }
        else if (subject.includes('書類選考の結果を通知しました(NG)')) { type = 'お断り'; status = 'お断り済み'; }

        if (!type) return;

        const jobId = extract(body, /(?:求人番号|案件ID|求人ID)[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
        const doctorNameMatch = extract(body, /お名前\(漢字\)[\s ]*[:：][\s ]*(.+)/) || extract(subject, /\[(.+?)医師\]/);
        const nameToFind = doctorNameMatch ? doctorNameMatch.replace(/[\s ]/g, '') : '';

        let rowNum = -1;
        const data = sheet.getDataRange().getValues();
        for (let j = data.length - 1; j > 0; j--) {
          if (String(data[j][0]).trim() !== 'エムステージ') continue;
          const rowStatus = String(data[j][9]).trim();
          if (!rowStatus.includes('応募通知受信')) continue;

          const rowDocClean = String(data[j][6]).split(/[\/／（(]/)[0].replace(/[\s ]/g, '');
          const rowJobId = String(data[j][5]).trim();

          if ((jobId && rowJobId === jobId) && (!nameToFind || rowDocClean === nameToFind)) {
            rowNum = j + 1;
            break;
          }
        }

        if (rowNum !== -1) {
          const values = data[rowNum - 1].slice();
          values[8] = type; values[9] = status; values[11] = new Date();

          appendRowToArchive(sheet, archiveSheet, values, type === '確定' ? '採用' : '不採用');
          safeDeleteRow(sheet, rowNum);

          threadProcessed = true;
        } else if (type === 'お断り' && jobId && archiveSheet) {
          const archData = archiveSheet.getDataRange().getValues();
          for (let j = archData.length - 1; j > 0; j--) {
            const rowAgency = String(archData[j][0]).trim();
            const rowJobId = String(archData[j][5]).trim();
            const rowDocClean = String(archData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]/g, '');
            if ((rowAgency === 'エムステージ' || rowAgency === '') && rowJobId === jobId && (!nameToFind || rowDocClean === nameToFind)) {
              if (archSaiyoCol > 0) archiveSheet.getRange(j + 1, archSaiyoCol).setValue('不採用').setBackground('#ffe599');
              if (archStatusCol > 0) archiveSheet.getRange(j + 1, archStatusCol).setValue(status);
              threadProcessed = true;
            }
          }
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }

  if (isIdUpdated) {
    if (processedIds.length > 300) processedIds = processedIds.slice(-300);
    props.setProperty('MSTAGE_PROCESSED_IDS', JSON.stringify(processedIds));
  }
}

// === MRT ===
function processMRT_Application(querySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const now = new Date();

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('MRT_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch (e) {}
  const processedSet = new Set(processedIds);
  let isIdUpdated = false;

  const markId = (id) => {
    if (!processedSet.has(id)) {
      processedSet.add(id);
      processedIds.push(id);
      isIdUpdated = true;
    }
  };

  // === 1. 応募通知の処理 ===
  const applyQuery = `subject:("【MRT】") subject:("応募通知") ${querySuffix}`;
  const applyThreads = GmailApp.search(applyQuery, 0, 50);

  if (applyThreads.length > 0) {
    const applyMessagesAll = GmailApp.getMessagesForThreads(applyThreads);

    for (let i = 0; i < applyThreads.length; i++) {
      const thread = applyThreads[i];
      const messages = applyMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const receivedDate = message.getDate();
        if ((now - receivedDate) / (1000 * 60 * 60) > 48) return;

        const messageId = message.getId();
        if (processedSet.has(messageId)) return;
        markId(messageId);

        const body = message.getPlainBody();
        const dateLine = extract(body, /日時[\s:：]+([^\n\r]+)/);
        const id = extract(body, /ID[\s:：]*(\d+)/i);
        const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/);
        const clinic = extract(body, /クリニック[\s:：]*\d+\s*キャップスクリニック([^\n\r]+)/);
        const department = extract(body, /診療科[\s:：]*([^\n\r,]+)/);

        if (id && dateLine) {
          const added = appendRowAndFormatUnique(sheet, [
            'MRT', formatDate(dateLine, receivedDate), workTime,
            clinic, department,
            id,
            '', extract(body, /(https:\/\/medical\.medrt\.com\/customers\/sessions\/new)/),
            '', '応募通知受信（手動更新要）', receivedDate, '', ''
          ]);
          if (added) threadProcessed = true;
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }

  // === 2. 応募取消し（キャンセル）の自動処理（★14d全件取得による負荷を3d＋上限50件へ適正化） ===
  const mrtCancelQuery = `subject:("【MRT】") subject:("応募取消し") newer_than:3d`;
  const cancelThreads = GmailApp.search(mrtCancelQuery, 0, 50);

  if (cancelThreads.length > 0) {
    const cancelMessagesAll = GmailApp.getMessagesForThreads(cancelThreads);
    const archSaiyoCol = getColIndex_internal(archiveSheet, '採用可否');
    const archStatusCol = getColIndex_internal(archiveSheet, '処理状況');

    for (let i = 0; i < cancelThreads.length; i++) {
      const thread = cancelThreads[i];
      const messages = cancelMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        if ((now - message.getDate()) / (1000 * 60 * 60) > 72) return;

        const messageId = message.getId();
        if (processedSet.has(messageId)) return;
        markId(messageId);

        const body = message.getPlainBody();
        const idMatch = body.match(/ID[\s:：]*(\d+)/i);
        const mailJobId = idMatch ? idMatch[1].trim() : "";
        if (!mailJobId) return;

        const docMatch = body.match(/応募者[\s:：]*([^\n\r]+?)先生/);
        const mailDoctorName = docMatch ? docMatch[1].trim() : "";

        let rowNum = -1;
        const sheetData = sheet.getDataRange().getValues();
        for (let j = sheetData.length - 1; j > 0; j--) {
          const rowAgency = String(sheetData[j][0]).trim();
          const rowJobId = String(sheetData[j][5]).trim();
          const rowStatus = String(sheetData[j][9]).trim();

          if ((rowAgency === 'MRT' || rowAgency === '') && rowJobId === mailJobId && rowStatus.includes('応募通知受信')) {
            rowNum = j + 1;
            break;
          }
        }

        if (rowNum !== -1) {
          const values = sheetData[rowNum - 1].slice();
          if (mailDoctorName) values[6] = mailDoctorName;
          values[8] = 'お断り';
          values[9] = 'キャンセル済み';
          values[11] = new Date();

          appendRowToArchive(sheet, archiveSheet, values, '不採用');
          safeDeleteRow(sheet, rowNum);

          threadProcessed = true;
        } else if (archiveSheet) {
          const archData = archiveSheet.getDataRange().getValues();
          const cleanMailDoc = mailDoctorName.replace(/[\s ]/g, '');
          for (let j = archData.length - 1; j > 0; j--) {
            const rowAgency = String(archData[j][0]).trim();
            const rowJobId = String(archData[j][5]).trim();
            const rowDocClean = String(archData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]/g, '');
            if ((rowAgency === 'MRT' || rowAgency === '') && rowJobId === mailJobId && (!cleanMailDoc || !rowDocClean || rowDocClean === cleanMailDoc)) {
              if (archSaiyoCol > 0) archiveSheet.getRange(j + 1, archSaiyoCol).setValue('不採用').setBackground('#ffe599');
              if (archStatusCol > 0) archiveSheet.getRange(j + 1, archStatusCol).setValue('キャンセル済み');
              threadProcessed = true;
            }
          }
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }

  if (isIdUpdated) {
    if (processedIds.length > 300) processedIds = processedIds.slice(-300);
    props.setProperty('MRT_PROCESSED_IDS', JSON.stringify(processedIds));
  }
}

// === 民間医局 ===
function processMinkanIkyoku_Application(querySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const now = new Date();

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('MINKAN_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch (e) {}
  const processedSet = new Set(processedIds);
  let isIdUpdated = false;

  const applyQuery = `subject:("【民間医局ポータル】") subject:("申し込みがありました") ${querySuffix}`;
  const threads = GmailApp.search(applyQuery, 0, 50);

  if (threads.length > 0) {
    const allMessages = GmailApp.getMessagesForThreads(threads);

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      const messages = allMessages[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const receivedDate = message.getDate();
        if ((now - receivedDate) / (1000 * 60 * 60) > 48) return;

        const messageId = message.getId();
        if (processedSet.has(messageId)) return;

        const body = message.getPlainBody();
        const doctorName = extract(body, /医師名[\s ]*[:：][\s ]*(.+)/).replace(/[\s ]/g, '');
        let department = extract(body, /診療科目[\s ]*[:：][\s ]*(.+)/);
        if (!department) department = findDepartmentFromHistory(ss, doctorName) || '';

        const id = extract(body, /(?:求人ID|求人番号|案件番号)[\s ]*[:：][\s ]*(.+)/);
        const dateLine = extract(body, /勤務日程[\s ]*[:：][\s ]*(.+)/);
        const timeLine = extract(body, /勤務時間[\s ]*[:：][\s ]*(.+)/);
        const clinic = extract(body, /^キャップスクリニック([^\r\n]+)/m).replace(/様$/, '').trim();

        if (doctorName && dateLine) {
          const added = appendRowAndFormatUnique(sheet, [
            '民間医局', formatDate(dateLine, receivedDate), timeLine,
            clinic, department,
            id,
            doctorName, extract(body, /(https:\/\/portal\.doctor-agent\.com\/.+)/),
            '', '応募通知受信（手動更新要）', receivedDate, '', ''
          ]);
          if (added) threadProcessed = true;

          processedSet.add(messageId);
          processedIds.push(messageId);
          isIdUpdated = true;
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }

  if (isIdUpdated) {
    if (processedIds.length > 300) processedIds = processedIds.slice(-300);
    props.setProperty('MINKAN_PROCESSED_IDS', JSON.stringify(processedIds));
  }
}

// === メディウェル ===
function processMediwel_Application(querySuffix, processedLabel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROGRESS);
  const processedLabelName = processedLabel.getName();
  const now = new Date();

  const applyQuery = `subject:("メディウェル") subject:("応募") -label:${processedLabelName} ${querySuffix}`;
  const threads = GmailApp.search(applyQuery, 0, 50);

  if (threads.length > 0) {
    const allMessages = GmailApp.getMessagesForThreads(threads);

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      const messages = allMessages[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const receivedDate = message.getDate();
        if ((now - receivedDate) / (1000 * 60 * 60) > 48) return;

        const body = message.getPlainBody();
        const doctorName = extract(body, /医師氏名[\s:：]+([^\n\r]+)/) || extract(body, /医師名[\s:：]+([^\n\r]+)/) || '不明';
        const dateLine = extract(body, /勤務日[\s:：]+([^\n\r]+)/) || extract(body, /勤務日程[\s:：]+([^\n\r]+)/) || Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd');
        const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/) || extract(body, /勤務時間[\s:：]+([^\n\r]+)/) || '';
        const clinic = extract(body, /医療機関名[\s:：]+([^\n\r]+)/) || extract(body, /施設名[\s:：]+([^\n\r]+)/) || '不明';
        const department = extract(body, /募集科目[\s:：]+([^\n\r]+)/) || extract(body, /診療科目[\s:：]+([^\n\r]+)/) || '';
        const id = extract(body, /(?:求人ID|案件ID|求人番号)[\s:：]+([a-zA-Z0-9\-]+)/) || '';

        if (doctorName && dateLine) {
          const added = appendRowAndFormatUnique(sheet, [
            'メディウェル', formatDate(dateLine, receivedDate), workTime,
            clinic, department, id, doctorName, '', '', '応募通知受信（手動更新要）', receivedDate, '', ''
          ]);
          if (added) threadProcessed = true;
        }
      });

      if (threadProcessed) {
        thread.addLabel(processedLabel);
      }
    }
  }
}