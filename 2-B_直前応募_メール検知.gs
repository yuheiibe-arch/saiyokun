// =================================================================
// 【2-B_直前応募_メール検知 (Part 1)】
// Gmailからのシフト応募通知を解析し、「通知履歴」シートへ記録する
// （あわせて民間医局の特定医師スキャンも呼び出す）
// ★修正：一度応募して取り下げ、再応募したケース（別メール）も確実に検知・通知するよう改善
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
  const existingUniqueKeys = lastRow > 1 ? sheet.getRange(2, 6, lastRow - 1, 1).getValues().flat().map(String) : [];

  // ★追加：各シフト枠（勤務日_拠点_医師_時間）の「最後に処理したメール受信時刻」を管理するマップ
  // （1回目の応募メールと、取り下げ後の2回目の再応募メールを正確に区別するため）
  const props = PropertiesService.getScriptProperties();
  let lastMailTimeMap = {};
  try {
    const storedMap = props.getProperty('URGENT_APPLY_LAST_TIME_MAP');
    if (storedMap) lastMailTimeMap = JSON.parse(storedMap);
  } catch (e) {}
  let isMapUpdated = false;

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

  // ★重要：スレッド横断で全ての対象メッセージを「受信時刻の古い順（昇順）」に並べてから処理する
  // （1回目の応募 → 2回目の再応募 の順序を確実に保証するため）
  const targetMessages = [];
  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];
    const thread = threads[i];
    for (const message of messages) {
      const messageDate = message.getDate();
      const diffHours = (now - messageDate) / (1000 * 60 * 60);
      if (diffHours <= 24) {
        targetMessages.push({ message: message, thread: thread, messageDate: messageDate });
      }
    }
  }

  // 古い順にソート
  targetMessages.sort((a, b) => a.messageDate.getTime() - b.messageDate.getTime());

  const threadsToLabel = new Set();

  for (const item of targetMessages) {
    const result = processSingleMessage_internal(
      item.message, sheet, existingUniqueKeys,
      todayStr, tomorrowStr, dayAfterTomorrowStr,
      agencyData, agencyStatusColumnIndex,
      item.messageDate,
      lastMailTimeMap
    );
    if (result.processed) {
      threadsToLabel.add(item.thread);
    }
    if (result.mapUpdated) {
      isMapUpdated = true;
    }
  }

  // 新規データがあったスレッドにラベルを付与
  if (threadsToLabel.size > 0) {
    try {
      const label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
      threadsToLabel.forEach(t => t.addLabel(label));
    } catch(e) {}
  }

  // プロパティの更新（古いゴミデータが溜まらないよう直近7日分のみ保持）
  if (isMapUpdated) {
    const cutoff = now.getTime() - (7 * 24 * 60 * 60 * 1000);
    for (const k in lastMailTimeMap) {
      if (lastMailTimeMap[k] < cutoff) {
        delete lastMailTimeMap[k];
      }
    }
    try {
      props.setProperty('URGENT_APPLY_LAST_TIME_MAP', JSON.stringify(lastMailTimeMap));
    } catch (e) {}
  }

  try { flushSpecialCandidateNotifications(); } catch(e) {}
}

function processSingleMessage_internal(message, sheet, existingUniqueKeys, todayStr, tomorrowStr, dayAfterTomorrowStr, agencyData, agencyStatusColumnIndex, receivedDateObj, lastMailTimeMap) {
  const receivedDate = Utilities.formatDate(receivedDateObj, 'JST', 'yyyy/MM/dd');
  const messageId = message.getId();
  let processedFlag = false;
  let mapUpdatedFlag = false;

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
      clinic: shiftMatch[3].trim(),
      messageId: messageId
    };
    
    const res = processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex, lastMailTimeMap);
    if (res.status) processedFlag = true;
    if (res.mapUpdated) mapUpdatedFlag = true;
    try { accumulateSpecialCandidate(entryData, body); } catch(e) {}
  }
  
  // パターンB: 民間医局・ドクターサイト運営宛てフォーマット
  if (!foundShift) {
    const regexListOld = /シフト応募日：\s*(\d{4}年\d{1,2}月\d{1,2}日 \([月火水木金土日]\))(?:\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2}))?[^\n\r]*[\n\r]+応募拠点名：\s*([^\n\r]+)/g;
    while ((shiftMatch = regexListOld.exec(body)) !== null) {
      foundShift = true;
      const entryData = { 
        doctorName: doctorName,
        workDate: shiftMatch[1].trim(), 
        workTime: shiftMatch[2] ? shiftMatch[2].trim() : "（時間記載なし）",
        clinic: shiftMatch[3].trim(),
        messageId: messageId
      };
      const res = processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex, lastMailTimeMap);
      if (res.status) processedFlag = true;
      if (res.mapUpdated) mapUpdatedFlag = true;
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
        clinic: clinicMatch[1].trim(),
        messageId: messageId
      };
      const res = processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex, lastMailTimeMap);
      if (res.status) processedFlag = true;
      if (res.mapUpdated) mapUpdatedFlag = true;
      try { accumulateSpecialCandidate(entryData, body); } catch(e) {}
    }
  }
  
  return { processed: processedFlag, mapUpdated: mapUpdatedFlag };
}

function processEntry_internal(sheet, existingUniqueKeys, receivedDate, entryData, todayStr, tomorrowStr, dayAfterTomorrowStr, receivedDateObj, agencyData, agencyStatusColumnIndex, lastMailTimeMap) {
  const { workDate, clinic, doctorName, workTime, messageId } = entryData;
  const clinicNormalized = clinic.replace(/\s+/g, '');
  const doctorNameNormalized = doctorName.replace(/\s+/g, '');

  // 基本キー（旧フォーマットと同じ：勤務日_拠点_医師名_時間）
  const baseKey = `${workDate}_${clinic}_${doctorName}_${workTime}`;
  // 新ユニークキー（末尾にMessage-IDを付与し、再応募メールを別レコードとして扱う）
  const uniqueKeyWithMsgId = `${baseKey}_${messageId}`;

  if (doctorName.includes('テスト') || clinic.toLowerCase().includes('test')) {
    return { status: '', mapUpdated: false };
  }

  // 1. すでにこのMessage-IDで記録済みなら確実にスキップ
  if (existingUniqueKeys.includes(uniqueKeyWithMsgId)) {
    return { status: '', mapUpdated: false };
  }

  const msgTime = receivedDateObj.getTime();
  let mapUpdated = false;

  // 2. 同じシフト枠（baseKey）が既にシートに存在するかチェック（旧キー または 新キーの前方一致）
  const hasExistingShiftRecord = existingUniqueKeys.some(k => k === baseKey || k.startsWith(baseKey + '_'));

  if (hasExistingShiftRecord) {
    const prevTime = lastMailTimeMap[baseKey];

    if (!prevTime) {
      // ★初回移行ガード：スクリプト改修前に旧キーで記録済みだった1回目のメール（例：15:31のメール）が
      // 最初にスキャンされた時は、その時刻をマップに記憶してスキップする（過去メールの一斉再通知暴発を防止）。
      // そうすることで、時系列で後ろにある2回目のメール（例：17:22のメール）が来た時に下の else に入り、確実に通知される！
      lastMailTimeMap[baseKey] = msgTime;
      return { status: '', mapUpdated: true };
    } else {
      // すでに1回目のメール時刻が記録されている場合：
      // 1回の応募で同時に複数メールが届いた場合の重複を防ぐため、前回のメールから3分（180秒）以内のものはスキップし、
      // 3分以上経過して届いた新しいメール（取り下げ後の再応募）は新規通知として通過させる！
      const diffMinutes = (msgTime - prevTime) / (1000 * 60);
      if (diffMinutes < 3) {
        return { status: '', mapUpdated: false };
      }
    }
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
          return { status: '', mapUpdated: mapUpdated }; 
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
    const newRowData = [receivedDate, clinic, workDate, doctorName, workTime, uniqueKeyWithMsgId, postStatus];
    sheet.appendRow(newRowData);
    existingUniqueKeys.push(uniqueKeyWithMsgId);
    lastMailTimeMap[baseKey] = msgTime;
    mapUpdated = true;
    console.log(`[直前応募検知・記録] ${doctorName} / ${clinic} / ${workDate} ${workTime} (Key: ${uniqueKeyWithMsgId})`);
  }

  return { status: postStatus, mapUpdated: mapUpdated };
}