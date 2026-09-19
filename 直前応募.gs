/**
 * ★★★【修正】大司令塔から呼び出される内部関数に変更 ★★★
 * (executeAllProcesses -> executeAllProcesses_internal)
 */
function executeAllProcesses_internal() {
  // ★★★【修正】ロック処理を削除（大司令塔が担当） ★★★

  console.log("=============== 応募関連の自動化プロセスを開始 ===============");
  part1_processEmailsToSheet();
  part2_checkForAgencyUrgentApplications();
  part3_postToChatworkFromSheet();
  console.log("\n=============== 応募関連の自動化プロセスが完了 ===============");

  // ★★★【修正】ロック処理を削除 ★★★
}

// =================================================================
// ▼▼▼ 【修正版】part1 関連関数（民間医局くん連携・API最適化済み） ▼▼▼
// =================================================================

// === ステップ1：Gmailからの新規応募を処理 ===
function part1_processEmailsToSheet() {
  console.log("\n--- ステップ1: Gmailからの新規応募を処理中 ---");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) { console.error(`シート「${LOG_SHEET_NAME}」が見つかりません。`); return; }

  // ▼▼▼【追加】紹介会社シートの情報を取得 ▼▼▼
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
  // ▲▲▲【追加】紹介会社シートの情報を取得 ▲▲▲

  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:G1').setValues([['応募日', '応募拠点', '勤務日', '応募医師', '応募時間', 'ユニークキー', '連携ステータス']]);
  }
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_APPLY);
  if (!label) { label = GmailApp.createLabel(PROCESSED_LABEL_APPLY); }
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

  // ▼▼▼【修正】全体設定の GMAIL_QUERY_APPLY に「-label:(処理済み)」を連結して無駄な検索を削減
  const query = `${GMAIL_QUERY_APPLY} -label:${label.getName()}`;
  const threads = GmailApp.search(query);
  
  if (threads.length === 0) { 
    console.log('対象のメールはありませんでした。'); 
    // ★★★【追加】メールがない場合でも、バッファに残っていれば放出
    try { flushSpecialCandidateNotifications(); } catch(e) {}
    return; 
  }
  
  // ★★★【追加】現在時刻（24時間チェック用） ★★★
  const now = new Date();

  // ★★★【修正】一括取得でAPIリクエストを1回に削減 ★★★
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadWasProcessed = false;

    for (const message of messages) {
      // ★★★【重要：二重ロック】★★★
      // 受信から24時間以上経過しているメールは、どんな理由があろうと絶対に無視する
      const messageDate = message.getDate();
      const diffHours = (now - messageDate) / (1000 * 60 * 60);
      if (diffHours > 24) {
        continue; 
      }
      // ★★★★★★★★★★★★★★★★★

      const messageProcessed = processSingleMessage_internal(
        message, sheet, existingUniqueKeys,
        todayStr, tomorrowStr, dayAfterTomorrowStr,
        agencyData, agencyStatusColumnIndex,
        messageDate 
      );
      // 1件でも処理（シート書き込み）されればフラグを立てる
      if (messageProcessed) {
        threadWasProcessed = true;
      }
    }

    // ▼▼▼【修正】API消費の原因だった thread.getLabels() のチェックを削除（検索時点で除外しているため不要）
    if (threadWasProcessed) {
      thread.addLabel(label);
    }
  }
  
  // ★★★【追加】全てのメール処理が終わったら、特定医師（民間医局くん）の通知をまとめて放出 ★★★
  try { flushSpecialCandidateNotifications(); } catch(e) {}
}

/**
 * ▼▼▼ processSingleMessage_internal (★超堅牢フォーマット対応版★) ▼▼▼
 * 新フォーマット・旧フォーマットのどちらでも確実に日付と時間を分離します
 */
function processSingleMessage_internal(message, sheet, existingUniqueKeys, todayStr, tomorrowStr, dayAfterTomorrowStr, agencyData, agencyStatusColumnIndex, receivedDateObj) {
  const receivedDate = Utilities.formatDate(receivedDateObj, 'JST', 'yyyy/MM/dd');
  const body = message.getPlainBody();
  let processedFlag = false;

  // 1. 医師名の取得
  let doctorName = "不明な医師";
  const docMatch1 = body.match(/^\s*([^\n\r]+?)\s*先生/);
  const docMatch2 = body.match(/勤務医師名：\s*(.*?先生)/);
  if (docMatch1) {
    doctorName = docMatch1[1].trim();
  } else if (docMatch2) {
    doctorName = docMatch2[1].replace('先生', '').trim();
  }

  // 2. 新フォーマットの抽出（連続形式）
  const shiftRegex = /(\d{4}年\d{1,2}月\d{1,2}日 \([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})\s*\(([^)]+)\)/g;
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
  
  // 3. 旧フォーマットの抽出（リスト形式：超堅牢版）
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

/**
 * ▼▼▼ processEntry_internal (変更なし) ▼▼▼
 */
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
      if (match) {
        gmailWorkDateStr = match[1];
      }
    } else {
      try {
        gmailWorkDateStr = Utilities.formatDate(new Date(workDate), 'JST', 'yyyy年M月d日');
      } catch (e) {
        console.warn(`不明な日付形式のため紹介会社シートとの比較をスキップ: ${workDate}`);
      }
    }
    if (gmailWorkDateStr) {
      for (const agencyRow of agencyData) {
        const agencyWorkDateObj = new Date(agencyRow[1]);
        if (isNaN(agencyWorkDateObj.getTime())) continue;
        const agencyWorkDateStr = Utilities.formatDate(agencyWorkDateObj, 'JST', 'yyyy年M月d日');
        const agencyClinicRaw = agencyRow[3] || '';
        // ★★★【修正】列ズレ対応: 医師名を [5] から [6] に変更 ★★★
        const agencyDoctorRaw = agencyRow[6] || ''; 
        const agencyStatus = agencyRow[agencyStatusColumnIndex];
        const agencyClinicNormalized = agencyClinicRaw.replace(/\s+/g, '');
        const agencyDoctorNormalized = agencyDoctorRaw.replace(/\s+/g, '');
        if (agencyWorkDateStr === gmailWorkDateStr &&
            agencyClinicNormalized === clinicNormalized &&
            agencyDoctorNormalized === doctorNameNormalized &&
            agencyStatus === AGENCY_STATUS_PROCESSED) {
          console.log(`紹介会社シートで処理済みのためスキップ (Gmail): ${workDate}, ${clinic}, ${doctorName}`);
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
  } else if (workDate === todayStr) {
      postStatus = '投稿連携';
  } else if (workDate === tomorrowStr) {
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

/**
 * 独立動作用のSlack送信ヘルパー
 */
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

/**
 * 独立動作用のChatwork送信ヘルパー
 */
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

/**
 * === ステップ2：紹介会社シートの直前応募を処理 ===
 */
function part2_checkForAgencyUrgentApplications() {
  console.log("--- ステップ2: 紹介会社シートの直前応募を処理中 ---");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(AGENCY_SOURCE_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);

  const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  let processedLogKeys = []; 
  if (logSheet && logSheet.getLastRow() > 1) {
    const logValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).getValues();
    logValues.forEach(row => {
      const status = row[6]; 
      const uniqueKey = row[5]; 
      if ((status === '投稿連携' || status === '投稿連携(夜間)' || status === '投稿済み' || status.includes('済')) && typeof uniqueKey === 'string' && uniqueKey.includes('_')) {
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
          try {
            logWorkDateNorm = Utilities.formatDate(new Date(logWorkDateRaw), 'JST', 'yyyy年M月d日');
          } catch (e) { /* 形式が不正ならそのまま */ }
        }
        const normalizedKey = `${logWorkDateNorm}_${logClinic.replace(/\s+/g, '')}_${logDoctor.replace(/\s+/g, '')}`;
        processedLogKeys.push(normalizedKey);
      }
    });
  }

  if (!sourceSheet || !destinationSheet) { console.error(`シート「${AGENCY_SOURCE_SHEET_NAME}」または「${URGENT_SHEET_NAME}」が見つかりません。`); return; }

  const statusColumnIndex = AGENCY_STATUS_COLUMN_LETTER.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) { console.log('紹介会社シートにチェック対象データがありません。'); return; }

  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  
  const yesterday17 = new Date(today.getTime()); 
  yesterday17.setDate(yesterday17.getDate() - 1); 
  yesterday17.setHours(17, 0, 0, 0); 

  const today9 = new Date(today.getTime());
  today9.setHours(9, 0, 0, 0); 

  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'); 
  const tomorrow = new Date(today.getTime());
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');
  const dayAfterTomorrow = new Date(today.getTime());
  dayAfterTomorrow.setDate(today.getDate() + 2);
  const dayAfterTomorrowStr = Utilities.formatDate(dayAfterTomorrow, 'JST', 'yyyy/MM/dd');

  const sourceData = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  let processCount = 0;
  for (let i = 0; i < sourceData.length; i++) {
    const row = sourceData[i];
    const rowNumber = i + 2;

    const agencyName  = row[0];
    const workDateObj = new Date(row[1]); 

    let workTime = '記載なし';
    if (row[2]) {
      const timeObj = new Date(row[2]);
      if (timeObj.getFullYear() === 1899) {
        workTime = Utilities.formatDate(timeObj, 'JST', 'HH:mm');
      } else if (!isNaN(timeObj.getTime())) {
        workTime = Utilities.formatDate(timeObj, 'JST', 'HH:mm');
      } else {
        workTime = row[2];
      }
    }

    const clinic      = row[3] || '記載なし';
    // ★★★【修正】列ズレ対応: 医師名を [5] から [6] に変更 ★★★
    const doctorName  = row[6] || '記載なし';
    // ★★★【修正】列ズレ対応: 受信日時を [9] から [10] に変更 ★★★
    const receivedDate = row[10] ? new Date(row[10]) : new Date(); 
    const status      = row[statusColumnIndex - 1];

    if (isNaN(workDateObj.getTime())) continue;

    const workDateStr = Utilities.formatDate(workDateObj, 'JST', 'yyyy/MM/dd'); 

    let isUrgent = false;
    let isNightUrgent = false;

    const hour = receivedDate.getHours();
    const isNightTime = (hour >= 17 || hour < 9);

    if (isNightTime && (workDateStr === tomorrowStr || workDateStr === dayAfterTomorrowStr)) {
        isUrgent = true;
        isNightUrgent = true;
    } else if (workDateStr === todayStr) {
        isUrgent = true;
    } else if (workDateStr === tomorrowStr) {
        isUrgent = true;
    }

    if (isUrgent && status !== AGENCY_STATUS_PROCESSED) {

      const agencyWorkDateNorm = Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日');
      const agencyClinicNorm = clinic.replace(/\s+/g, '');
      const agencyDoctorNorm = doctorName.replace(/\s+/g, '');
      const agencyNormalizedKey = `${agencyWorkDateNorm}_${agencyClinicNorm}_${agencyDoctorNorm}`;
      
      if (processedLogKeys.includes(agencyNormalizedKey)) {
        console.log(`part2: LOG_SHEETで処理済みのためスキップ: ${agencyWorkDateNorm}, ${clinic}, ${doctorName}`);
        try {
          sourceSheet.getRange(rowNumber, statusColumnIndex).setValue(AGENCY_STATUS_PROCESSED);
          console.log(`part2: スキップした案件のステータスを「通知済み」に更新しました (Row: ${rowNumber})`);
        } catch(e) {
          console.error(`part2: スキップした案件のステータス更新に失敗 (Row: ${rowNumber}): ${e.message}`);
        }
        continue; 
      }

      let messageBody = isNightUrgent ? '直近の応募が夜間にありました。（紹介会社経由）\n確認お願いします。' : '紹介会社経由で直前応募がありました。\n確認お願いします。';

      const message = `[info][title]直前応募通知（紹介会社経由）[/title]
${messageBody}

紹介会社： ${agencyName}
勤務日： ${Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日 (E)')}
応募時間： ${workTime}
応募拠点： ${clinic}
応募医師： ${doctorName}
[/info]`;

      // ★ 独立管理：すでに片方が済んでいればスキップし、未済の方だけを送信する
      let isSlackDone = (status === 'Slack済');
      let isCwDone = (status === 'CW済');

      let slackSuccess = isSlackDone ? true : sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = isCwDone ? true : sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      if (slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, statusColumnIndex).setValue(AGENCY_STATUS_PROCESSED);
        const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd');
        const uniqueKey = `${agencyName}_${workDateStr}_${doctorName}`;
        const newRowData = [receivedDateStr, clinic, Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日 (E)'), doctorName, workTime, uniqueKey, '紹介会社経由', agencyName];
        destinationSheet.appendRow(newRowData);
        processCount++;
      } else if (slackSuccess && !cwSuccess) {
        sourceSheet.getRange(rowNumber, statusColumnIndex).setValue('Slack済');
      } else if (!slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, statusColumnIndex).setValue('CW済');
      }
    }
  }
  if(processCount === 0) console.log('紹介会社経由の新しい直前応募はありませんでした。');
}


/**
 * === ステップ3：「投稿連携」を検知し、Chatworkへ投稿 ===
 */
function part3_postToChatworkFromSheet() {
  console.log("--- ステップ3: Chatworkへの投稿処理を処理中 ---");
  const STATUS_TO_POST = '投稿連携';
  const STATUS_TO_POST_NIGHT = '投稿連携(夜間)';
  const STATUS_POSTED = '投稿済み';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);
  
  if (!sourceSheet || !destinationSheet) { console.error(`シートが見つかりません。`); return; }
  if (destinationSheet.getRange('A1').getValue() === '') {
    destinationSheet.getRange('A1:G1').setValues(sourceSheet.getRange('A1:G1').getValues());
  }
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) { console.log('投稿対象データはありませんでした。'); return; }
  
  const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, 7);
  const values = dataRange.getValues();
  let processCount = 0;
  
  for (let i = 0; i < values.length; i++) {
    const currentStatus = values[i][6];
    
    // ★ 未送信、または片方だけ送信済みのステータスを対象にする
    const isTarget = (currentStatus === STATUS_TO_POST || currentStatus === STATUS_TO_POST_NIGHT || currentStatus.includes('Slack済') || currentStatus.includes('CW済'));
    
    if (isTarget) {
      const rowNumber = i + 2;
      let isNight = currentStatus.includes('夜間');
      
      const messageBody = isNight ? '直近の応募が夜間にありました。\n確認お願いします。' : '直前応募がありました。\n確認お願いします。';
      const message = `[info][title]直前応募通知[/title]${messageBody}\n\n応募拠点： ${values[i][1]}\n勤務日： ${values[i][2]}\n応募時間： ${values[i][4]}\n応募医師： ${values[i][3]}[/info]`;

      // ★ 独立管理：すでに片方が済んでいればスキップし、未済の方だけを送信する
      let isSlackDone = currentStatus.includes('Slack済');
      let isCwDone = currentStatus.includes('CW済');

      let slackSuccess = isSlackDone ? true : sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = isCwDone ? true : sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      if (slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(STATUS_POSTED);
        // 両方成功したタイミングで1回だけ URGENT_SHEET に転記
        values[i][6] = STATUS_POSTED;
        destinationSheet.appendRow(values[i]);
        processCount++;
      } else if (slackSuccess && !cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'Slack済(夜間)' : 'Slack済');
      } else if (!slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'CW済(夜間)' : 'CW済');
      }
    }
  }
  if (processCount === 0) console.log('Gmail経由の新しい直前応募はありませんでした。');
}