/**
 * ★★★【修正】大司令塔から呼び出される内部関数に変更 ★★★
 * (executeAllProcesses -> executeAllProcesses_internal)
 */
function executeAllProcesses_internal() {
  console.log("=============== 応募関連の自動化プロセスを開始 ===============");
  part1_processEmailsToSheet();
  part2_checkForAgencyUrgentApplications();
  part3_postToChatworkFromSheet();
  console.log("\n=============== 応募関連の自動化プロセスが完了 ===============");
}

// =================================================================
// ▼▼▼ part1 関連関数（民間医局くん連携・API最適化・重複すり抜け防止版） ▼▼▼
// =================================================================

function part1_processEmailsToSheet() {
  console.log("\n--- ステップ1: Gmailからの新規応募を処理中 ---");
  
  // ★【追加】処理済みIDをプロパティから取得
  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('URGENT_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch(e) {}
  let isIdUpdated = false;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) { console.error(`シート「${LOG_SHEET_NAME}」が見つかりません。`); return; }

  const agencySheet = spreadsheet.getSheetByName(AGENCY_SOURCE_SHEET_NAME);
  let agencyData = [];
  let agencyStatusColumnIndex = -1;

  if (!agencySheet) {
    console.warn(`シート「${AGENCY_SOURCE_SHEET_NAME}」が見つからないため、紹介会社シートとの重複チェックをスキップします。`);
  } else {
    agencyStatusColumnIndex = AGENCY_STATUS_COLUMN_LETTER.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0);
    const lastAgencyRow = agencySheet.getLastRow();
    if (lastAgencyRow > 1) {
      agencyData = agencySheet.getRange(2, 1, lastAgencyRow - 1, agencySheet.getLastColumn()).getValues();
    }
  }

  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:G1').setValues([['応募日', '応募拠点', '勤務日', '応募医師', '応募時間', 'ユニークキー', '連携ステータス']]);
  }
  const lastRow = sheet.getLastRow();
  const existingUniqueKeys = lastRow > 1 ? sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat() : [];

  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  const dayAfterTomorrow = new Date();
  dayAfterTomorrow.setDate(today.getDate() + 2);

  const getJapaneseDay = (date) => ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
  const todayStr = `${Utilities.formatDate(today, 'JST', 'yyyy年M月d日')} (${getJapaneseDay(today)})`;
  const tomorrowStr = `${Utilities.formatDate(tomorrow, 'JST', 'yyyy年M月d日')} (${getJapaneseDay(tomorrow)})`;
  const dayAfterTomorrowStr = `${Utilities.formatDate(dayAfterTomorrow, 'JST', 'yyyy年M月d日')} (${getJapaneseDay(dayAfterTomorrow)})`;

  // ★★★【重要修正】ラベルによる検索除外を完全撤廃（スレッド吸収によるすり抜け防止）
  let query = GMAIL_QUERY_APPLY;
  query = query.replace(/-label:[^\s]+/g, '').trim();

  // ★★★【究極修正】Gmailバグの原因「after:UNIX時間」を強制削除し、スレッド化対策で14日分取得 ★★★
  query = query.replace(/after:\d+/g, '').trim();
  if (!query.includes('newer_than')) {
      query += ' newer_than:14d'; 
  }
  
  // 第2引数、第3引数で「最新の30スレッドだけを取得」することでAPI上限を絶対に回避します
  const threads = GmailApp.search(query, 0, 30);
  
  if (threads.length === 0) { 
    console.log('対象のメールはありませんでした。'); 
    try { flushSpecialCandidateNotifications(); } catch(e) {}
    return; 
  }
  
  const now = new Date();
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];

    for (const message of messages) {
      const messageId = message.getId();
      // ★【追加】既に処理したメールIDなら、何も読み込まずにスキップ（API消費ゼロ）
      if (processedIds.includes(messageId)) continue;
      processedIds.push(messageId);
      isIdUpdated = true;

      const messageDate = message.getDate();
      const diffHours = (now - messageDate) / (1000 * 60 * 60);
      
      // ★★★【重要】14日分取得した中で、直近24時間以内の「本当の最新メール」だけを処理するストッパー ★★★
      if (diffHours > 24) {
        continue; 
      }

      processSingleMessage_internal(
        message, sheet, existingUniqueKeys,
        todayStr, tomorrowStr, dayAfterTomorrowStr,
        agencyData, agencyStatusColumnIndex,
        messageDate 
      );
    }
  }
  
  // ★【追加】新しく処理したID群をプロパティに保存（最新500件のみ保持）
  if (isIdUpdated) {
    if (processedIds.length > 500) {
      processedIds = processedIds.slice(-500);
    }
    props.setProperty('URGENT_PROCESSED_IDS', JSON.stringify(processedIds));
  }

  try { flushSpecialCandidateNotifications(); } catch(e) {}
}

function processSingleMessage_internal(message, sheet, existingUniqueKeys, todayStr, tomorrowStr, dayAfterTomorrowStr, agencyData, agencyStatusColumnIndex, receivedDateObj) {
  const receivedDate = Utilities.formatDate(receivedDateObj, 'JST', 'yyyy/MM/dd');
  const body = message.getPlainBody();
  let processedFlag = false;

  // ★★★【修正】医師名の抽出ロジックを強化（フォーマット揺れに完全対応）★★★
  let doctorName = "不明な医師";
  const docMatch1 = body.match(/^\s*([^\n\r]+?)\s*先生/);
  const docMatch2 = body.match(/(?:勤務医師名|応募医師名|医師氏名|医師名|氏名)[\s ]*[：:]?[\s ]*([^\n\r]+)/);

  if (docMatch2) {
    // まずラベル付き（勤務医師名：など）を探す。見つかれば「先生」や余分な空白を除去
    doctorName = docMatch2[1].replace(/先生/g, '').trim();
  } else if (docMatch1) {
    // ラベルが無く、いきなり「〇〇先生」で始まっている場合
    doctorName = docMatch1[1].trim();
  }

  // ★★★【重要修正】改行なし・カッコ連続のフォーマットでも確実に抽出できるように正規表現を修正 ★★★
  const shiftRegex = /(\d{4}年\d{1,2}月\d{1,2}日\s*\([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})\s*\(([^)]+?)\)/g;
  let shiftMatch;
  let foundShift = false;
  
  while ((shiftMatch = shiftRegex.exec(body)) !== null) {
    foundShift = true;
    const entryData = { 
      doctorName: doctorName,
      workDate: shiftMatch[1].trim(), 
      workTime: shiftMatch[2].trim(), 
      clinic: shiftMatch[3].trim() 
    };
    
    const status = processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex);
    if (status) processedFlag = true;
    try { accumulateSpecialCandidate(entryData, body); } catch(e) {}
  }
  
  if (!foundShift) {
    const regexListOld = /シフト応募日：\s*(\d{4}年\d{1,2}月\d{1,2}日 \([月火水木金土日]\))(?:\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2}))?[^\n\r]*[\n\r]+応募拠点名：\s*([^\n\r]+)/g;
    while ((shiftMatch = regexListOld.exec(body)) !== null) {
       const entryData = { 
         doctorName: doctorName,
         workDate: shiftMatch[1].trim(), 
         workTime: shiftMatch[2] ? shiftMatch[2].trim() : "（時間記載なし）",
         clinic: shiftMatch[3].trim() 
       };
       const status = processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex);
       if (status) processedFlag = true;
       try { accumulateSpecialCandidate(entryData, body); } catch(e) {}
    }
  }
  
  return processedFlag;
}

function processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex) {
  const { workDate, clinic, doctorName, workTime } = entryData;
  const clinicNormalized = clinic.replace(/\s+/g, '');
  const doctorNameNormalized = doctorName.replace(/\s+/g, '');
  const uniqueKey = `${workDate}_${clinic}_${doctorName}_${workTime}`;

  if (existingUniqueKeys.includes(uniqueKey) || doctorName.includes('テスト') || clinic.toLowerCase().includes('test')) {
    return ''; 
  }

  if (agencyData.length > 0 && agencyStatusColumnIndex !== -1) {
    let gmailWorkDateStr;
    if (workDate.includes('年')) {
      const match = workDate.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
      if (match) { gmailWorkDateStr = match[1]; }
    } else {
      try { gmailWorkDateStr = Utilities.formatDate(new Date(workDate), 'JST', 'yyyy年M月d日'); } catch (e) {}
    }
    if (gmailWorkDateStr) {
      for (const agencyRow of agencyData) {
        const agencyWorkDateObj = new Date(agencyRow[1]);
        if (isNaN(agencyWorkDateObj.getTime())) continue;
        const agencyWorkDateStr = Utilities.formatDate(agencyWorkDateObj, 'JST', 'yyyy年M月d日');
        const agencyClinicRaw = agencyRow[3] || '';
        const agencyDoctorRaw = agencyRow[6] || ''; 
        const agencyStatus = agencyRow[agencyStatusColumnIndex];
        const agencyClinicNormalized = agencyClinicRaw.replace(/\s+/g, '');
        const agencyDoctorNormalized = agencyDoctorRaw.replace(/\s+/g, '');
        if (agencyWorkDateStr === gmailWorkDateStr &&
            agencyClinicNormalized === clinicNormalized &&
            agencyDoctorNormalized === doctorNameNormalized &&
            agencyStatus === AGENCY_STATUS_PROCESSED) {
          return ''; 
        }
      }
    }
  }

  const hour = receivedDateObj.getHours();
  const isNightTime = (hour >= 17 || hour < 9);

  let postStatus = '';
  if (isNightTime && (workDate === tomorrowStr || workDate === dayAfterTomorrowStr)) {
      postStatus = '投稿連携(夜間)';
  } else if (workDate === todayStr || workDate === tomorrowStr) {
      postStatus = '投稿連携';
  }

  if (postStatus) {
    const newRowData = [receivedDate, clinic, workDate, doctorName, workTime, uniqueKey, postStatus];
    sheet.appendRow(newRowData);
    existingUniqueKeys.push(uniqueKey); 
  }

  return postStatus;
}


// =================================================================
// ▼▼▼ 以下、part2 と part3 (Slack/CW 独立送信・ステータス分離版) ▼▼▼
// =================================================================

function sendUrgentAlertToSlack_internal(roomId, message) {
  try {
    const slackPayload = buildSlackPayload(message, roomId); 
    if (slackPayload && slackPayload.channelId) {
      const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify({ "channel": slackPayload.channelId, "text": slackPayload.text }),
        "muteHttpExceptions": true
      };
      const res = UrlFetchApp.fetch(webhookUrl, options);
      return (res.getResponseCode() === 200 || res.getResponseCode() === 201);
    }
  } catch (e) {
    console.error("Slack送信エラー: " + e.message);
  }
  return false;
}

function sendUrgentAlertToChatwork_internal(roomId, message) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey) return false;
  const url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error("Chatwork 通信エラー: " + e.message);
    return false;
  }
}

function part2_checkForAgencyUrgentApplications() {
  console.log("--- ステップ2: 紹介会社シートの直前応募を処理中 ---");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(AGENCY_SOURCE_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);
  const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  
  const archiveSheetName = typeof SHEET_NAME_ARCHIVE !== 'undefined' ? SHEET_NAME_ARCHIVE : '処理済み';
  const archiveSheet = spreadsheet.getSheetByName(archiveSheetName);

  let processedLogKeys = []; 
  if (logSheet && logSheet.getLastRow() > 1) {
    const logValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).getValues();
    logValues.forEach(row => {
      const status = row[6]; 
      const uniqueKey = row[5]; 
      if ((status === '投稿連携' || status === '投稿連携(夜間)' || status === '投稿済み' || String(status).includes('済')) && typeof uniqueKey === 'string' && uniqueKey.includes('_')) {
        const parts = uniqueKey.split('_');
        if (parts.length < 3) return; 

        const logWorkDateRaw = parts[0]; 
        const logClinic = parts[1] || '';
        const logDoctor = parts[2] || '';
        
        let logWorkDateNorm = logWorkDateRaw;
        if (logWorkDateRaw.includes('年')) {
          const match = logWorkDateRaw.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
          if (match) logWorkDateNorm = match[1];
        } else {
          try { logWorkDateNorm = Utilities.formatDate(new Date(logWorkDateRaw), 'JST', 'yyyy年M月d日'); } catch (e) {}
        }
        processedLogKeys.push(`${logWorkDateNorm}_${logClinic.replace(/\s+/g, '')}_${logDoctor.replace(/\s+/g, '')}`);
      }
    });
  }

  if (!destinationSheet) { return; }
  const statusColumnIndex = typeof AGENCY_STATUS_COLUMN_LETTER !== 'undefined' ? 
                            AGENCY_STATUS_COLUMN_LETTER.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1 : 13;

  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'); 
  const tomorrow = new Date(today.getTime());
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');
  const dayAfterTomorrow = new Date(today.getTime());
  dayAfterTomorrow.setDate(today.getDate() + 2);
  const dayAfterTomorrowStr = Utilities.formatDate(dayAfterTomorrow, 'JST', 'yyyy/MM/dd');

  let combinedData = [];

  if (sourceSheet && sourceSheet.getLastRow() >= 2) {
    const sData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
    sData.forEach((row, i) => combinedData.push({ row: row, rowNumber: i + 2, sheet: sourceSheet }));
  }

  if (archiveSheet && archiveSheet.getLastRow() >= 2) {
    const aHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    
    let taiouColIndex = aHeaders.findIndex(h => String(h).includes('対応時間') || String(h).includes('処理日時'));
    if (taiouColIndex === -1 && typeof CHECKBOX_COLUMN !== 'undefined') {
      taiouColIndex = CHECKBOX_COLUMN - 1; 
    }

    const aData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archiveSheet.getLastColumn()).getValues();
    
    const todayMidnight = new Date();
    todayMidnight.setHours(0, 0, 0, 0);
    const yesterdayMidnight = new Date(todayMidnight.getTime() - 24 * 60 * 60 * 1000);

    aData.forEach((row, i) => {
      if (taiouColIndex !== -1 && row[taiouColIndex]) {
        const tDate = new Date(row[taiouColIndex]);
        if (!isNaN(tDate.getTime()) && tDate.getTime() >= yesterdayMidnight.getTime()) {
          combinedData.push({ row: row, rowNumber: i + 2, sheet: archiveSheet });
        }
      }
    });
  }

  const processedStatusStr = typeof AGENCY_STATUS_PROCESSED !== 'undefined' ? AGENCY_STATUS_PROCESSED : '処理済';
  
  for (let i = 0; i < combinedData.length; i++) {
    const item = combinedData[i];
    const row = item.row;
    const rowNumber = item.rowNumber;
    const currentSheet = item.sheet;

    const agencyName = row[0];
    const workDateObj = new Date(row[1]); 
    let workTime = row[2] ? (isNaN(new Date(row[2]).getTime()) ? row[2] : Utilities.formatDate(new Date(row[2]), 'JST', 'HH:mm')) : '記載なし';
    const clinic = row[3] || '記載なし';
    const doctorName = row[6] || '記載なし';
    const receivedDate = row[10] ? new Date(row[10]) : new Date(); 
    const status = row[statusColumnIndex - 1];

    if (isNaN(workDateObj.getTime())) continue;

    const workDateStr = Utilities.formatDate(workDateObj, 'JST', 'yyyy/MM/dd'); 
    const hour = receivedDate.getHours();
    const isNightTime = (hour >= 17 || hour < 9);

    let isUrgent = false;
    let isNightUrgent = false;
    if (isNightTime && (workDateStr === tomorrowStr || workDateStr === dayAfterTomorrowStr)) {
        isUrgent = true;
        isNightUrgent = true;
    } else if (workDateStr === todayStr || workDateStr === tomorrowStr) {
        isUrgent = true;
    }

    if (isUrgent && status !== processedStatusStr && status !== 'Slack済' && status !== 'CW済') {
      const agencyWorkDateNorm = Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日');
      const agencyNormalizedKey = `${agencyWorkDateNorm}_${clinic.replace(/\s+/g, '')}_${doctorName.replace(/\s+/g, '')}`;
      
      if (processedLogKeys.includes(agencyNormalizedKey)) {
        try { currentSheet.getRange(rowNumber, statusColumnIndex).setValue(processedStatusStr); } catch(e) {}
        continue; 
      }

      let messageBody = isNightUrgent ? '直近の応募が夜間にありました。（紹介会社経由）\n確認お願いします。' : '紹介会社経由で直前応募がありました。\n確認お願いします。';
      const message = `[info][title]直前応募通知（紹介会社経由）[/title]\n${messageBody}\n\n紹介会社： ${agencyName}\n勤務日： ${Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日 (E)')}\n応募時間： ${workTime}\n応募拠点： ${clinic}\n応募医師： ${doctorName}\n[/info]`;

      let slackSuccess = sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      if (slackSuccess && cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue(processedStatusStr);
        const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd');
        const uniqueKey = `${agencyName}_${workDateStr}_${doctorName}`;
        const newRowData = [receivedDateStr, clinic, Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日 (E)'), doctorName, workTime, uniqueKey, '紹介会社経由', agencyName];
        destinationSheet.appendRow(newRowData);
      } else if (slackSuccess && !cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue('Slack済');
      } else if (!slackSuccess && cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue('CW済');
      }
    }
  }
}

function part3_postToChatworkFromSheet() {
  console.log("--- ステップ3: Chatworkへの投稿処理を処理中 ---");
  const STATUS_TO_POST = '投稿連携';
  const STATUS_TO_POST_NIGHT = '投稿連携(夜間)';
  const STATUS_POSTED = '投稿済み';
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);
  
  if (!sourceSheet || !destinationSheet) { return; }
  if (destinationSheet.getRange('A1').getValue() === '') {
    destinationSheet.getRange('A1:G1').setValues(sourceSheet.getRange('A1:G1').getValues());
  }
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) { return; }
  
  const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, 7);
  const values = dataRange.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const currentStatus = values[i][6];
    
    const isTarget = (currentStatus === STATUS_TO_POST || currentStatus === STATUS_TO_POST_NIGHT || String(currentStatus).includes('Slack済') || String(currentStatus).includes('CW済'));
    
    if (isTarget) {
      const rowNumber = i + 2;
      let isNight = String(currentStatus).includes('夜間');
      
      const messageBody = isNight ? '直近の応募が夜間にありました。\n確認お願いします。' : '直前応募がありました。\n確認お願いします。';
      const message = `[info][title]直前応募通知[/title]${messageBody}\n\n応募拠点： ${values[i][1]}\n勤務日： ${values[i][2]}\n応募時間： ${values[i][4]}\n応募医師： ${values[i][3]}[/info]`;

      let isSlackDone = String(currentStatus).includes('Slack済');
      let isCwDone = String(currentStatus).includes('CW済');

      let slackSuccess = isSlackDone ? true : sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = isCwDone ? true : sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      if (slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(STATUS_POSTED);
        values[i][6] = STATUS_POSTED;
        destinationSheet.appendRow(values[i]);
      } else if (slackSuccess && !cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'Slack済(夜間)' : 'Slack済');
      } else if (!slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'CW済(夜間)' : 'CW済');
      }
    }
  }
}