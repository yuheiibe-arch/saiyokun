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