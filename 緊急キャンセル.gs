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

  // ★★★【修正】スレッド巻き込みによる見落としを防ぐため、-label: の除外を撤廃
  const timeQuery = `(${timeKeywords.join(' OR ')})`;
  const actionQuery = `(${actionKeywords.join(' OR ')})`;
  const gmailQuery = `${timeQuery} ${actionQuery} newer_than:14d`;

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
      
      // ★ 14日分取得した中で、直近24時間以内の「本当の最新メール」だけを処理するストッパー
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

      const chatworkMessage = `[toall]\n[info][title]特定アラート[/title]\nメールに特定のフレーズを含む内容が届きました。\n内容の確認をお願いします。\n※本アラートには関係のないメールも含まれる可能性があります。\n\n受信時刻： ${receivedTime}\n件名： ${subject}\n内容：\n${bodySnippet}...\n[/info]`;

      sendToChatwork(DOTAKYAN_CHATWORK_ROOM_ID, chatworkMessage);
      Utilities.sleep(1500); // 制限回避
    }
    
    if (processedThisThread) {
      thread.addLabel(label);
    }
  } 
  console.log('【非定型キャンセル通知】処理が完了しました。');
}