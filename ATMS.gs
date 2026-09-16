// =================================================================
// ▼▼▼ ATMSメール検知・通知スクリプト（単独ファイル版） ▼▼▼
// ★修正内容：API制限対策として getMessagesForThreads による一括取得に変更
// =================================================================

const ATMS_CHATWORK_ROOM_ID = '421840785'; 

function checkAtmsEmails_internal() {
  console.log('【ATMSメール検知】処理を開始します。');

  // スプレッドシートを使わず、GASの裏側に「処理済みメッセージID」を記憶・取得する
  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('ATMS_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch(e) {}

  // ★ここを修正：yuhei.ibe@... と doctor-support@... の両方を対象（OR条件）にする
  const query = `from:assist@asiantms.com (to:yuhei.ibe@caps365.jp OR to:doctor-support@caps365.jp) newer_than:1d`;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    console.log('ATMSからのメールはありませんでした。');
    return;
  }

  const mentions = "[To:8440783]Gibran ~RGS~さん\n[To:11351957]丸山 幹枝さん\n[To:9935353]伊瓶さん\n";
  let isUpdated = false; // プロパティを更新する必要があるかのフラグ

  // ★ API通信1回で全メッセージを一括取得
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];

    for (const message of messages) {
      const messageId = message.getId();
      
      // ① すでに通知済みのメッセージ（過去のメール）はスキップ
      if (processedIds.includes(messageId)) continue;

      // ② 送信元(From)が ATMS (assist@asiantms.com) 以外ならスキップ（自チームの返信を除外）
      const from = message.getFrom();
      if (!from.includes('assist@asiantms.com')) {
        processedIds.push(messageId); // 自チームの返信も「処理済み」として記録し、次回以降スルー
        isUpdated = true;
        continue;
      }

      // --- ここからATMSからの「新規メール」のみの処理 ---
      
      // 本文を取得し、過去の引用履歴をカット
      const fullBody = message.getPlainBody();
      const regexQuoteMarkers = /\n(>|On .*> wrote:|.* <.*@.*> wrote:|\d{4}[年\/]\d{1,2}[月\/]\d{1,2}日?.*(?::|のメール:)|From: .*|Sent: .*|-{3,}|={3,}|#{3,}|_+\s*$|(?:iPhone|iPad|スマートフォン)から送信)/i;
      const cleanBody = fullBody.split(regexQuoteMarkers)[0].trim();

      const receivedDateStr = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');
      const subject = message.getSubject();

      const cwMessage = `${mentions}[info][title]ATMSメール[/title]\n受信時刻：${receivedDateStr}\n件名：${subject}\n\n${cleanBody}\n[/info]`;

      // 通知を送信
      sendToChatwork(ATMS_CHATWORK_ROOM_ID, cwMessage);
      Utilities.sleep(1500); // API制限回避

      // 処理が完了したメールIDを記憶リストに追加
      processedIds.push(messageId);
      isUpdated = true;
    }
  }

  // 新しい処理があった場合のみ、IDリストを保存（容量オーバーを防ぐため最新200件のみ保持）
  if (isUpdated) {
    if (processedIds.length > 200) {
      processedIds = processedIds.slice(-200);
    }
    props.setProperty('ATMS_PROCESSED_IDS', JSON.stringify(processedIds));
  }

  console.log('【ATMSメール検知】処理が完了しました。');
}