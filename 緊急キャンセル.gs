/**
 * ★★★【修正】大司令塔から呼び出される内部関数に変更 ★★★
 * (checkAndLogEmergencyCancel -> checkAndLogEmergencyCancel_internal)
 *
 * ★★★【2025/11/02 修正】★★★
 * 2種類のメール形式（「★★緊急★★」と「【要対応】3日以内」）に対応
 *
 * ★★★【API制限対策・最適化版】★★★
 * API消費を1回にまとめる getMessagesForThreads を導入し、
 * 重複チェックをプロパティ(Message-ID)でも行うことで通信エラーを防止。
 */
function checkAndLogEmergencyCancel_internal() {
  console.log('【緊急キャンセル通知】処理を開始します。');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CANCEL_LOG_SHEET_NAME);
  if (!sheet) {
    console.error(`シート「${CANCEL_LOG_SHEET_NAME}」が見つかりません。全体設定.gs とシート名を確認してください。`);
    return;
  }

  // 1. 処理済みラベル（PROCESSED_LABEL_CANCEL）を取得・なければ作成
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_CANCEL);
  if (!label) {
    label = GmailApp.createLabel(PROCESSED_LABEL_CANCEL);
  }

  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:G1').setValues([['番号', '勤務日', '勤務拠点', '診療科', '勤務者', '勤務時間', '受信時刻']]);
  }
  
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy年M月d日');
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy年M月d日');
  
  // ★ API通信削減: すでに処理済みラベルがついているものを検索段階で弾く
  const query = `${GMAIL_QUERY_CANCEL} newer_than:2d`;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    console.log('緊急キャンセルの対象メールはありませんでした。');
    return;
  }

  // 重複チェック用のデータ取得（シートから）
  SpreadsheetApp.flush();
  const latestSheetData = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];

  // ★ API通信1回で全メッセージを一括取得
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadProcessed = false;

    for (const message of messages) {
      // ストッパー：24時間以上古いメールは無視
      if ((today.getTime() - message.getDate().getTime()) / (1000 * 60 * 60) > 24) continue;

      const subject = message.getSubject();
      if (message.getFrom().includes(GMAIL_SENDER_EXCLUDE)) {
        console.log(`除外SENDERのためスキップ: ${message.getFrom()}`);
        threadProcessed = true; // ラベルを貼るためにフラグを立てる
        continue;
      }
      const body = message.getPlainBody();

      let workDate, department, clinic, time, name;

      // --- 1. 件名に応じて処理を分岐 ---
      if (subject.includes('★★緊急★★')) {
        const mainRegex = /(\d{4}年\d{1,2}月\d{1,2}日)\s+(.+?)\s+(.+?)、\s*(\d{1,2}:\d{2}~\d{1,2}:\d{2})\s+(.+?)のキャンセルが申請されました/;
        const match = body.match(mainRegex);
        if (!match) {
          console.warn(`「★★緊急★★」メールの本文解析に失敗: ${subject}`);
          continue;
        }
        [ , workDate, department, clinic, time, name] = match.map(item => item.trim());

      } else if (subject.includes('【要対応】勤務3日以内')) {
        const nameMatch = body.match(/医師名：(.+)/);
        const clinicMatch = body.match(/勤務拠点名：(.+)/);
        const dateMatch = body.match(/キャンセル日付：(.+)/);
        const timeMatch = body.match(/勤務時間：(.+)/);

        if (!nameMatch || !clinicMatch || !dateMatch || !timeMatch) {
          console.warn(`「3日以内」メールの本文解析に失敗: ${subject}`);
          continue;
        }

        name = nameMatch[1].trim();
        const clinicRaw = clinicMatch[1].trim(); 
        workDate = dateMatch[1].trim(); 
        time = timeMatch[1].trim().replace('~', '～'); 

        const parts = clinicRaw.split(/\s+/); 
        if (parts.length >= 2) {
          department = parts.pop(); 
          clinic = parts.join(' '); 
        } else {
          clinic = clinicRaw;
          department = '（診療科不明）';
        }

      } else {
        console.warn(`不明なキャンセルメール形式です: ${subject}`);
        continue;
      }

      // --- 2. 共通の通知処理 ---
      const receivedTime = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');
      
      if (clinic === '【関東】バックアップシフト' || (workDate !== todayStr && workDate !== tomorrowStr)) {
        console.log(`通知対象外（バックアップ or 日付）のためスキップ: ${workDate} ${clinic}`);
        threadProcessed = true;
        continue;
      }
      
      const isDuplicate = latestSheetData.some(row => String(row[1]).trim() === workDate && String(row[2]).trim() === clinic && String(row[4]).trim() === name && String(row[5]).trim() === time);
      
      if (!isDuplicate) {
        console.log(`新規キャンセルを検知、シートに書き込みます: ${workDate} ${clinic} ${name}`);
        const newId = sheet.getLastRow();
        sheet.appendRow([newId, workDate, clinic, department, name, time, receivedTime]);
        
        const chatworkMessage = `[toall]\n[info][title]直近緊急キャンセル[/title]勤務予定医師より緊急キャンセルが申請されました。\n詳細を確認してください。\n\n勤務日： ${workDate}\n勤務拠点： ${clinic}\n診療科： ${department}\n勤務者： ${name}\n勤務時間： ${time}\n受信時刻： ${receivedTime}[/info]`;
        sendToChatwork(CANCEL_LOG_CHATWORK_ROOM_ID, chatworkMessage);
        Utilities.sleep(1500); // 制限回避
      } else {
        console.log(`処理済み（シートに重複あり）のためスキップ: ${workDate} ${clinic} ${name}`);
      }

      threadProcessed = true;
    } 

    if (threadProcessed) {
      thread.addLabel(label);
    }
  } 
  console.log('【緊急キャンセル通知】処理が完了しました。');
}


/**
 * ★★★【修正】時間外着信の検知（独立・API最適化版）★★★
 * API制限を回避し、かつ他のフィルターに巻き込まれないように専用処理化
 */
function checkTimegaiCallWithLabel_internal() {
  console.log('【時間外着信通知】処理を開始します。');

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('MINDERU_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch(e) {}

  // ★ 修正: ハードコードされていた検索条件を撤廃し、全体設定の GMAIL_QUERY_TIMEGAI を正しく使用する
  const query = `${GMAIL_QUERY_TIMEGAI} newer_than:1d`;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    console.log('時間外着信の対象メールはありませんでした。');
    return;
  }

  const now = new Date();
  let isUpdated = false;
  
  // ★ API通信1回で全メッセージを一括取得
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const messages = allMessages[i];

    for (const message of messages) {
      const messageId = message.getId();
      const receivedDate = message.getDate();
      
      if (processedIds.includes(messageId)) continue;
      if ((now.getTime() - receivedDate.getTime()) / (1000 * 60 * 60) > 24) continue;

      const body = message.getPlainBody();
      
      const receivedTime = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');
      const chatworkBody = `[toall]\n[info][title]時間外着信のお知らせ[/title]受信時刻： ${receivedTime}\n時間外応答がありました。\nマイページにログインして内容をご確認ください。\nhttps://minderu.com/mypage/telephones/[/info]`;
      
      sendToChatwork(TIMEGAI_CALL_CHATWORK_ROOM_ID, chatworkBody);
      Utilities.sleep(1500); 
      
      processedIds.push(messageId);
      isUpdated = true;
    }
  }

  if (isUpdated) {
    if (processedIds.length > 200) processedIds = processedIds.slice(-200);
    props.setProperty('MINDERU_PROCESSED_IDS', JSON.stringify(processedIds));
  }

  console.log('【時間外着信通知】処理が完了しました。');
}


/**
 * ★★★【新規】大司令塔から呼び出される内部関数 ★★★
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

  // ★ API通信削減: 検索段階で処理済みを除外
  const timeQuery = `(${timeKeywords.join(' OR ')})`;
  const actionQuery = `(${actionKeywords.join(' OR ')})`;
  const gmailQuery = `${timeQuery} ${actionQuery} -label:${label.getName()} newer_than:1d`;

  console.log(`Gmail検索クエリ: ${gmailQuery}`);

  SpreadsheetApp.flush();
  const logLastRow = logSheet.getLastRow();
  let existingMessageIDs = [];
  if (logLastRow > 1) {
    existingMessageIDs = logSheet.getRange(2, 7, logLastRow - 1, 1).getValues().flat();
  }

  const userEmail = Session.getActiveUser().getEmail();
  const excludeSenders = [
    'caps365.jp', 
    'mns.jp', 
    'mnys.jp', 
    'm3career.com', 
    'gemini-notes@google.com', 
    'asiantms.com', 
    'mediwel.net',  
    'nicho.co.jp',  
    userEmail
  ];

  const timeRegex = new RegExp(timeKeywords.join('|'), 'i');
  const actionRegex = new RegExp(actionKeywords.join('|'), 'i');
  const now = new Date();

  const threads = GmailApp.search(gmailQuery);
  if (threads.length === 0) { console.log('対象のメールはありませんでした。'); return; }

  // ★ API通信1回で全メッセージを一括取得
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

      const from = message.getFrom();
      const isInternal = excludeSenders.some(senderKeyword => from.includes(senderKeyword));
      
      if (isInternal) {
        console.log(`[メールID: ${messageId}] 除外対象の送信元のためスキップ: ${from}`);
        processedThisThread = true;
        continue; 
      }

      const fullBody = message.getPlainBody();
      const regexQuoteMarkers = /\n(>|On .*> wrote:|.* <.*@.*> wrote:|\d{4}年\d{1,2}月\d{1,2}日.*:|From: .*|Sent: .*|-{3,}|={3,}|#{3,}|_+\s*$)/i;
      const bodyOnly = fullBody.split(regexQuoteMarkers)[0].trim();

      const hasTimeKeyword = timeRegex.test(bodyOnly);
      const hasActionKeyword = actionRegex.test(bodyOnly);

      if (!hasTimeKeyword || !hasActionKeyword) {
        console.log(`[メールID: ${messageId}] 引用文除去後の本文にキーワード(A+B)が含まれないためスキップ`);
        continue;
      }
      
      console.log(`★検知成功 [メールID: ${messageId}]`);
      processedThisThread = true;

      const receivedTime = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');
      const to = message.getTo();
      const subject = message.getSubject();
      const bodySnippet = bodyOnly.substring(0, 300); 

      const newId = logSheet.getLastRow();
      logSheet.appendRow([newId, receivedTime, from, to, subject, bodySnippet, messageId]);
      
      existingMessageIDs.push(messageId); 

      const chatworkMessage = `[toall]\n[info][title]特定アラート[/title]
メールに特定のフレーズを含む内容が届きました。
内容の確認をお願いします。
※本アラートには関係のないメールも含まれる可能性があります。

受信時刻： ${receivedTime}
件名： ${subject}
内容：
${bodySnippet}...
[/info]`;

      sendToChatwork(DOTAKYAN_CHATWORK_ROOM_ID, chatworkMessage);
      Utilities.sleep(1500); // 制限回避
    }
    
    if (processedThisThread) {
      thread.addLabel(label);
    }
  } 
  console.log('【非定型キャンセル通知】処理が完了しました。');
}