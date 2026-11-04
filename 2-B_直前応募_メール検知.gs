// =================================================================
// 【2-B_直前応募_メール検知 (Part 1)】
// Gmailからのシフト応募通知を解析し、「通知履歴」シートへ記録する
// （あわせて民間医局の特定医師スキャンも呼び出す）
// =================================================================

function part1_processEmailsToSheet() {
  console.log("\n--- ステップ1: Gmailからの新規応募を処理中 ---");
  
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

  const query = GMAIL_QUERY_APPLY;
  const threads = GmailApp.search(query, 0, 100);
  
  if (threads.length === 0) { 
    console.log('対象のメールはありませんでした。'); 
    try { flushSpecialCandidateNotifications(); } catch(e) {}
    return; 
  }
  
  const now = new Date();
  const allMessages = GmailApp.getMessagesForThreads(threads);
  const labelName = typeof PROCESSED_LABEL_APPLY !== 'undefined' ? PROCESSED_LABEL_APPLY : '処理済み-医師シフト応募';

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];
    const thread = threads[i];
    let threadHasNewData = false;
    
    for (const message of messages) {
      const messageDate = message.getDate();
      const diffHours = (now - messageDate) / (1000 * 60 * 60);
      
      if (diffHours > 24) continue; 

      const isProcessed = processSingleMessage_internal(
        message, sheet, existingUniqueKeys,
        todayStr, tomorrowStr, dayAfterTomorrowStr,
        agencyData, agencyStatusColumnIndex,
        messageDate 
      );
      if (isProcessed) threadHasNewData = true;
    }
    
    if (threadHasNewData) {
      try {
        const label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
        thread.addLabel(label);
      } catch(e) {}
    }
  }

  try { flushSpecialCandidateNotifications(); } catch(e) {}
}

function processSingleMessage_internal(message, sheet, existingUniqueKeys, todayStr, tomorrowStr, dayAfterTomorrowStr, agencyData, agencyStatusColumnIndex, receivedDateObj) {
  const receivedDate = Utilities.formatDate(receivedDateObj, 'JST', 'yyyy/MM/dd');
  let processedFlag = false;

  // HTMLからクリーンテキストを生成（ノイズ排除）
  const rawHtml = message.getBody();
  let body = rawHtml
    .replace(/<head[^>]*>[\s\S]*?<\/head>/gi, '') 
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') 
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>|<\/div>|<p>|<div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/^[ \t]+/gm, '') 
    .replace(/[\n\r]+/g, '\n'); 

  let doctorName = "不明な医師";
  const docMatch1 = body.match(/^\s*([^\n\r]+?)\s*先生/);
  const docMatch2 = body.match(/(?:勤務医師名|応募医師名|医師氏名|医師名|氏名)[\s ]*[：:]?[\s ]*([^\n\r]+)/);

  if (docMatch2) {
    doctorName = docMatch2[1].replace(/先生/g, '').trim();
  } else if (docMatch1) {
    doctorName = docMatch1[1].trim();
  }

  const shiftRegex = /(\d{4}年\d{1,2}月\d{1,2}日\s*\([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})\s*\(([^)]+?)\)/g;
  let shiftMatch;
  let foundShift = false;
  
  // パターンA: 共通フォーマット
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
  
  // パターンB: 民間医局などのフォーマット
  if (!foundShift) {
    const regexListOld = /シフト応募日：\s*(\d{4}年\d{1,2}月\d{1,2}日 \([月火水木金土日]\))(?:\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2}))?[^\n\r]*[\n\r]+応募拠点名：\s*([^\n\r]+)/g;
    while ((shiftMatch = regexListOld.exec(body)) !== null) {
      foundShift = true;
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

  // パターンC: MRT対応フォーマット
  if (!foundShift) {
    const mrtDateRegex = /日時\s*:\s*(\d{4}\/\d{1,2}\/\d{1,2}\s*\([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/;
    const mrtClinicRegex = /クリニック\s*:\s*(?:\d+\s+)?([^\n\r]+)/;
    
    const dateMatch = body.match(mrtDateRegex);
    const clinicMatch = body.match(mrtClinicRegex);
    
    if (dateMatch && clinicMatch) {
      foundShift = true;
      const formattedDate = dateMatch[1].replace(/(\d{4})\/(\d{1,2})\/(\d{1,2})/, "$1年$2月$3日");
      
      if (doctorName === "不明な医師") {
        doctorName = "未記載(システム確認)";
      }
      
      const entryData = { 
        doctorName: doctorName,
        workDate: formattedDate, 
        workTime: dateMatch[2].trim(),
        clinic: clinicMatch[1].trim() 
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
  const isNightTime = (hour >= 19 || hour < 7);

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