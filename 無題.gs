function debugRealM3SlackFormat() {
  console.log("=== 実際のエムスリーメールでSlack変換テスト開始 ===");

  // 1. 実際のGmailからエムスリーのメッセージ受信・問い合わせメールを1件検索
  const query = 'from:career_spot@m3career.com subject:("メッセージを受信しました" OR "お問い合わせがありました")';
  const threads = GmailApp.search(query, 0, 1);

  if (threads.length === 0) {
    console.log("❌ 該当のエムスリーメールが見つかりませんでした。");
    return;
  }

  const messages = threads[0].getMessages();
  const message = messages[messages.length - 1];
  const subject = message.getSubject();
  const rawBody = message.getPlainBody();
  const receivedDate = message.getDate();

  console.log("▼ 取得したメール件名: " + subject);

  // 2. 『エムスリー処理.gs』の関数をそのまま使ってChatwork文面を生成
  const info = extractM3Info_Final(subject, rawBody);
  const type = subject.includes('メッセージを受信しました') ? 'メッセージ受信' : 'お問い合わせ';
  const typeTitle = type === 'メッセージ受信' ? 'メッセージ受信' : '問い合わせ';
  const cleanBody = extractM3CoreMessage_Final(rawBody, type);

  const dateObj = normalizeDate_Final(info.workDate, receivedDate);
  const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
  const formattedDate = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd') + ` (${dayMap[dateObj.getDay()]})`;

  let doctorLine = type === 'お問い合わせ' ? '' : `医師名： ${info.doctorName}\n`;
  const cwMessage = `[info][title]${typeTitle}[/title]\n${doctorLine}勤務日： ${formattedDate}\n拠点名： ${info.clinic}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n${cleanBody}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nリンク：https://career.m3.com/admin/hospital/message_threads\n[/info]`;

  console.log("\n▼ 1. 変換前のChatwork用テキスト:");
  console.log(cwMessage);

  // 3. 差し替えた『共通関数.gs』の関数でSlack用に変換
  const slackPayload = buildSlackPayload(cwMessage, '419888887');

  console.log("\n▼ 2. 変換後のSlack投稿用テキスト:");
  console.log("送信先チャンネルID: " + slackPayload.channelId);
  console.log("-----------------------------------------");
  console.log(slackPayload.text);
  console.log("-----------------------------------------");

  // 4. 自動チェック判定
  if (slackPayload.text.includes('[<!') || slackPayload.text.includes('|>')) {
    console.log("❌ まだ文字化けが残っています。『共通関数.gs』が正しく上書き保存されているか確認してください。");
  } else {
    console.log("✅ 文字化けや異常な分割は一切なく、正常な体裁に直っています！");
  }
}