// =================================================================
// ▼▼▼ 共通で使う関数 ▼▼▼
// =================================================================

/**
 * APIキーをスクリプトプロパティに設定するための関数（初回のみ実行）
 * ★★★【注意】★★★
 * このプロジェクトは2種類のAPIキーを使っています。
 * 1. CHATWORK_API_KEY （直前応募通知用）
 * 2. CHATWORK_API_TOKEN（採用報告通知用）
 * この関数は 1. のみ設定します。2. は別途設定が必要です。
 */
function setUp() {
  // ★★★★★↓ここに「直前応募通知用」のAPIキーを入力し、一度だけ実行してください↓★★★★★
  const apiKey = '';
  // ★★★★★↑実行後、この行のAPIキーは必ず削除してください↑★★★★★
  PropertiesService.getScriptProperties().setProperty('CHATWORK_API_KEY', apiKey);
  console.log('CHATWORK_API_KEY の設定が完了しました。入力したAPIキーはコードから削除してください。');
}

/**
 * 共通のChatwork送信関数（主に直前応募通知(part2, part3)が使用）
 */
function sendToChatwork(roomId, message) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey) { console.error('APIキー(CHATWORK_API_KEY)が設定されていません。'); return false; }
  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      console.log(`Chatwork(RoomID: ${roomId})への通知に成功しました。`);
      return true;
    } else {
      console.error(`Chatwork投稿失敗: ${response.getContentText()}`);
      return false;
    }
  } catch (e) {
    console.error(`Chatwork通信エラー: ${e.message}`);
    return false;
  }
}