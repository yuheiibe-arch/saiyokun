/**
 * ★★★【復活】大司令塔 (3/9) から呼び出される内部関数 ★★★
 * 勤務2週間以内・3日以内の緊急キャンセル申請メールを検知し、
 * 「キャンセルログ」シートへの記録とChatwork(2ルーム)/Slackへの通知を行う
 */
const EXTRA_EMERGENCY_CANCEL_CW_ROOM_ID = '449648673'; // ★追加投下先のChatworkルームID

function checkAndLogEmergencyCancel_internal() {
  console.log('【緊急キャンセル通知】処理を開始します。');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(CANCEL_LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(CANCEL_LOG_SHEET_NAME);
    logSheet.getRange('A1:G1').setValues([['番号', '勤務日', '勤務拠点', '診療科', '勤務者', '勤務時間', '受信日']]);
    logSheet.setFrozenRows(1);
  } else if (logSheet.getRange('A1').getValue() === '') {
    logSheet.getRange('A1:G1').setValues([['番号', '勤務日', '勤務拠点', '診療科', '勤務者', '勤務時間', '受信日']]);
  }

  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_CANCEL);
  if (!label) {
    label = GmailApp.createLabel(PROCESSED_LABEL_CANCEL);
  }

  // スレッド内の二重処理・重複行増殖を防ぐため、処理済みMessage-IDをプロパティで管理
  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('EMERGENCY_CANCEL_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch (e) {}
  let isIdUpdated = false;

  // シート側の既存データからも重複防止キーを作成（万が一プロパティがリセットされても二重記録しない防衛線）
  const lastRow = logSheet.getLastRow();
  const existingRecordKeys = new Set();
  if (lastRow > 1) {
    const data = logSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      const wDate = String(data[i][1]).trim();
      const wClinic = String(data[i][2]).trim();
      const wDoc = String(data[i][4]).replace(/\s+/g, '').trim();
      const wTime = String(data[i][5]).trim();
      if (wDate && wDoc) {
        existingRecordKeys.add(`${wDate}_${wClinic}_${wDoc}_${wTime}`);
      }
    }
  }

  // 検索クエリ（全体設定の GMAIL_QUERY_CANCEL を使用し、直近2日以内を対象）
  const query = `${GMAIL_QUERY_CANCEL} newer_than:2d`;
  const threads = GmailApp.search(query, 0, 30);

  if (threads.length === 0) {
    console.log('対象の緊急キャンセルメールはありませんでした。');
    return;
  }

  const now = new Date();
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadProcessed = false;

    for (const message of messages) {
      const messageId = message.getId();
      if (processedIds.includes(messageId)) continue;

      const from = message.getFrom();
      // ステージング環境（テストメール）は除外
      if (typeof GMAIL_SENDER_EXCLUDE !== 'undefined' && GMAIL_SENDER_EXCLUDE && from.includes(GMAIL_SENDER_EXCLUDE)) {
        processedIds.push(messageId);
        isIdUpdated = true;
        continue;
      }

      const receivedDate = message.getDate();
      // 48時間以上前の古いメールはスキップ
      if ((now.getTime() - receivedDate.getTime()) / (1000 * 60 * 60) > 48) {
        processedIds.push(messageId);
        isIdUpdated = true;
        continue;
      }

      const subject = message.getSubject();
      const body = message.getPlainBody();

      // 本文から各項目を抽出
      const docMatch = body.match(/医師名[\s ]*[:：][\s ]*([^\n\r]+)/);
      const clinicDeptMatch = body.match(/勤務拠点名[\s ]*[:：][\s ]*([^\n\r]+)/);
      const dateMatch = body.match(/キャンセル日付[\s ]*[:：][\s ]*([^\n\r]+)/);
      const timeMatch = body.match(/勤務時間[\s ]*[:：][\s ]*([^\n\r]+)/);

      if (!docMatch || !dateMatch) continue;

      const doctorName = docMatch[1].replace(/先生$/, '').trim();
      const workDate = dateMatch[1].trim();
      const workTime = timeMatch ? timeMatch[1].trim() : '';
      const rawClinicDept = clinicDeptMatch ? clinicDeptMatch[1].replace(/キャップスクリニック/g, '').trim() : '';

      // 「相模大野 小児科」を「勤務拠点（相模大野）」と「診療科（小児科）」に分割
      let clinic = rawClinicDept;
      let department = '小児科';
      const parts = rawClinicDept.split(/[\s ]+/);
      if (parts.length >= 2) {
        department = parts.pop();
        clinic = parts.join(' ');
      } else if (rawClinicDept.includes('内科')) {
        department = '内科';
        clinic = rawClinicDept.replace('内科', '').trim();
      } else if (rawClinicDept.includes('小児科')) {
        department = '小児科';
        clinic = rawClinicDept.replace('小児科', '').trim();
      }

      // シート上の重複チェック（2行ダブりを防止）
      const recordKey = `${workDate}_${clinic}_${doctorName.replace(/\s+/g, '')}_${workTime}`;
      if (existingRecordKeys.has(recordKey)) {
        processedIds.push(messageId);
        isIdUpdated = true;
        continue;
      }

      const currentLastRow = logSheet.getLastRow();
      let nextNum = 1;
      if (currentLastRow >= 2) {
        const lastNumVal = parseInt(logSheet.getRange(currentLastRow, 1).getValue(), 10);
        nextNum = isNaN(lastNumVal) ? currentLastRow : lastNumVal + 1;
      }

      const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd H:mm');

      // キャンセルログシートへ追記（A:番号, B:勤務日, C:勤務拠点, D:診療科, E:勤務者, F:勤務時間, G:受信日）
      logSheet.appendRow([nextNum, workDate, clinic, department, doctorName, workTime, receivedDateStr]);
      existingRecordKeys.add(recordKey);

      // Chatwork & Slack への通知作成（共通関数.gs の「直近緊急キャンセル」判定と連動）
      const cancelTypeTitle = subject.includes('3日以内') ? '直近緊急キャンセル（3日以内）' : '緊急キャンセル（2週間以内）';
      const cwMessage = `[toall]\n[info][title]${cancelTypeTitle}[/title]\n下記の医師からキャンセル申請が届きました。\n担当者は内容を至急ご確認ください。\n\n医師名：${doctorName} 先生\n勤務拠点名：${clinic} ${department}\nキャンセル日付：${workDate}\n勤務時間：${workTime}\n受信時刻：${receivedDateStr}\n[/info]`;

      // ① 既存のキャンセルログルーム（344086172）＋ Slack へ送信
      sendToChatwork(CANCEL_LOG_CHATWORK_ROOM_ID, cwMessage);
      Utilities.sleep(1000);

      // ② ★追加：指定ルーム（449648673）へも送信（Slack二重送信を防ぐためChatworkのみへ投下）
      sendChatworkOnly_internal(EXTRA_EMERGENCY_CANCEL_CW_ROOM_ID, cwMessage);
      Utilities.sleep(1500);

      processedIds.push(messageId);
      isIdUpdated = true;
      threadProcessed = true;
    }

    if (threadProcessed) {
      thread.addLabel(label);
    }
  }

  if (isIdUpdated) {
    if (processedIds.length > 250) {
      processedIds = processedIds.slice(-250);
    }
    props.setProperty('EMERGENCY_CANCEL_PROCESSED_IDS', JSON.stringify(processedIds));
  }

  console.log('【緊急キャンセル通知】処理が完了しました。');
}

/**
 * Slackへの二重通知を防ぎ、Chatworkの指定ルームにのみメッセージを投下する内部ヘルパー関数
 */
function sendChatworkOnly_internal(roomId, message) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey || !roomId) return false;

  const url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  const options = {
    method: 'post',
    headers: { 'X-ChatWorkToken': apiKey },
    payload: { body: message },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error('Chatwork追加ルーム送信エラー: ' + e.message);
    return false;
  }
}

/**
 * ★★★【新規】大司令塔 (6/9) から呼び出される内部関数 ★★★
 * 非定型キャンセル（ドタキャン）メールを検知して通知する
 * ★★★【API制限対策・最適化版】★★★
 */
function checkInformalCancel_internal() {
  console.log('【非定型キャンセル通知】処理を開始します。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const configSheet = ss.getSheetByName(DOTAKYAN_CONFIG_SHEET_NAME);
  const logSheet = ss.getSheetByName(DOTAKYAN_LOG_SHEET_NAME);

  if (!configSheet) { console.error(`シート「${DOTAKYAN_CONFIG_SHEET_NAME}」が見つかりません。`); return; }
  if (!logSheet) { console.error(`シート「${DOTAKYAN_LOG_SHEET_NAME}」が見つかりません。`); return; }
  
  if (logSheet.getRange('A1').getValue() === '') {
    logSheet.getRange('A1:G1').setValues([['番号', '受信時刻', '差出人', '送信先', '件名', '内容', 'ユニークキー']]);
  }

  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_DOTAKYAN);
  if (!label) { label = GmailApp.createLabel(PROCESSED_LABEL_DOTAKYAN); }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) { console.log(`シート「${DOTAKYAN_CONFIG_SHEET_NAME}」にキーワードがありません。`); return; }
  
  const configData = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const timeKeywords = configData.map(row => row[0]).filter(String); 
  const actionKeywords = configData.map(row => row[1]).filter(String); 

  if (timeKeywords.length === 0 || actionKeywords.length === 0) {
    console.log('A列またはB列のキーワードが空のため、処理をスキップします。');
    return;
  }

  // ★★★【改善】除外する送信元をGmailの検索クエリ（-from:）に直接組み込み、GASでの無駄なループを排除
  const userEmail = Session.getActiveUser().getEmail();
  const excludeDomains = [
    'caps365.jp', 'mns.jp', 'mnys.jp', 'm3career.com', 'gemini-notes@google.com', 
    'asiantms.com', 'mediwel.net', 'nicho.co.jp', 'medrt.com', userEmail
  ];
  const excludeQuery = excludeDomains.map(d => `-from:${d}`).join(' ');

  const timeQuery = `(${timeKeywords.join(' OR ')})`;
  const actionQuery = `(${actionKeywords.join(' OR ')})`;
  
  // ★★★【改善】検索期間を14日から実運用レベルの2日間に短縮（過去の巨大なメール群の読み込みを防止）
  const gmailQuery = `${timeQuery} ${actionQuery} ${excludeQuery} newer_than:2d`;

  console.log(`Gmail検索クエリ: ${gmailQuery}`);

  SpreadsheetApp.flush();
  const logLastRow = logSheet.getLastRow();
  let existingMessageIDs = [];
  if (logLastRow > 1) {
    existingMessageIDs = logSheet.getRange(2, 7, logLastRow - 1, 1).getValues().flat();
  }

  const timeRegex = new RegExp(timeKeywords.join('|'), 'i');
  const actionRegex = new RegExp(actionKeywords.join('|'), 'i');
  const now = new Date();

  const threads = GmailApp.search(gmailQuery);
  if (threads.length === 0) { console.log('対象のメールはありませんでした。'); return; }

  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let processedThisThread = false; 

    for (const message of messages) {
      const messageId = message.getId();

      if (existingMessageIDs.includes(messageId)) continue; 

      const receivedDate = message.getDate();
      if ((now.getTime() - receivedDate.getTime()) / (1000 * 60 * 60) > 24) continue; 

      const fullBody = message.getPlainBody();
      const regexQuoteMarkers = /\n(>|On .*> wrote:|.* <.*@.*> wrote:|\d{4}年\d{1,2}月\d{1,2}日.*:|From: .*|Sent: .*|-{3,}|={3,}|#{3,}|_+\s*$)/i;
      const bodyOnly = fullBody.split(regexQuoteMarkers)[0].trim();

      const hasTimeKeyword = timeRegex.test(bodyOnly);
      const hasActionKeyword = actionRegex.test(bodyOnly);

      if (!hasTimeKeyword || !hasActionKeyword) {
        continue;
      }
      
      console.log(`★検知成功 [メールID: ${messageId}]`);
      processedThisThread = true;

      const receivedTime = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');
      const to = message.getTo();
      const subject = message.getSubject();
      const bodySnippet = bodyOnly.substring(0, 300); 

      const newId = logSheet.getLastRow();
      logSheet.appendRow([newId, receivedTime, message.getFrom(), to, subject, bodySnippet, messageId]);
      
      existingMessageIDs.push(messageId); 

      const chatworkMessage = `[toall]\n[info][title]特定アラート[/title]\nメールに特定のフレーズを含む内容が届きました。\n内容の確認をお願いします。\n※本アラートには関係のないメールも含まれる可能性があります。\n\n受信時刻： ${receivedTime}\n件名： ${subject}\n内容：\n${bodySnippet}...\n[/info]`;

      sendToChatwork(DOTAKYAN_CHATWORK_ROOM_ID, chatworkMessage);
      Utilities.sleep(1500);
    }
    
    if (processedThisThread) {
      thread.addLabel(label);
    }
  } 
  console.log('【非定型キャンセル通知】処理が完了しました。');
}