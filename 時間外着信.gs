// =============================================================
// 【時間外着信（みんなにでんわ転送）検知・通知スクリプト】
// =============================================================

function checkTimegaiCallWithLabel_internal() {
  console.log('【時間外着信】検知処理を開始します。');

  const labelName = '処理済み-時間外着信';
  let processedLabel = GmailApp.getUserLabelByName(labelName);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(labelName);
  }

  // 検索条件
  const query = 'from:info@minderu.com ("時間外応答" OR "時間外着信" OR "不在着信") -label:' + labelName + ' newer_than:1d';
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    console.log('対象の時間外着信メールはありませんでした。');
    return;
  }

  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadProcessed = false;
    
    for (const message of messages) {
      const subject = message.getSubject();
      const receivedDateStr = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');
      
      const fullBody = message.getPlainBody();
      
      // 「-----」以降の登録メールアドレスや署名などの不要部分をカット
      const cleanBody = fullBody.split("-----")[0].trim();

      // ★過去ログと全く同じ「医師からのメール」フォーマットに完全復元
      // （共通関数.gs がこれを読み取り、Slackでは自動的にコードブロックで整形してくれます）
      const chatworkMessage = `[toall]\n[info][title]医師からのメール[/title]\n医師からメールが届きました。担当者は確認してください。\n受信時刻：${receivedDateStr}\n件名：${subject}\n[info][title]みんなにでんわ転送 先生[/title]\n${cleanBody}\n[/info]\n[/info]`;

      const roomId = '415381235'; 
      
      sendToChatwork(roomId, chatworkMessage);
      threadProcessed = true;
      console.log('時間外着信を検知・通知しました');
    }
    
    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }
  
  console.log('【時間外着信】検知処理が完了しました。');
}