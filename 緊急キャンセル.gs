/**
 * ★★★【修正】大司令塔から呼び出される内部関数に変更 ★★★
 * (checkAndLogEmergencyCancel -> checkAndLogEmergencyCancel_internal)
 *
 * ★★★【2025/11/02 修正】★★★
 * 2種類のメール形式（「★★緊急★★」と「【要対応】3日以内」）に対応
 *
 * ★★★【2025/11/02 修正2】★★★
 * 「is:unread」を廃止し、Gmailラベル「PROCESSED_LABEL_CANCEL」での重複排除に変更。
 * これにより既読メールも（1回だけ）処理対象となる。
 */
function checkAndLogEmergencyCancel_internal() {
  // ★★★【修正】ロック処理を削除（大司令塔が担当） ★★★

  console.log('【緊急キャンセル通知】処理を開始します。');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CANCEL_LOG_SHEET_NAME);
  // ★★★【重要】シートが見つかるかここでチェック ★★★
  if (!sheet) {
    console.error(`シート「${CANCEL_LOG_SHEET_NAME}」が見つかりません。全体設定.gs とシート名を確認してください。`);
    return;
  }
  // ★★★★★★★★★★★★★★★★★★★★★★★★★★★

  // ▼▼▼ 【ラベル修正】 ▼▼▼
  // 1. 処理済みラベル（PROCESSED_LABEL_CANCEL）を取得・なければ作成
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_CANCEL);
  if (!label) {
    label = GmailApp.createLabel(PROCESSED_LABEL_CANCEL);
  }
  // ▲▲▲ 【ラベル修正】 ▲▲▲

  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange('A1:G1').setValues([['番号', '勤務日', '勤務拠点', '診療科', '勤務者', '勤務時間', '受信時刻']]);
  }
  const today = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy年M月d日');
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy年M月d日');
  const threads = GmailApp.search(GMAIL_QUERY_CANCEL);

  for (const thread of threads) {
    // ▼▼▼ 【ラベル修正】 ▼▼▼
    // ループの最後に thread.addLabel(label) を実行するため、try...finally を使用
    try {
      const message = thread.getMessages()[0];
      const subject = message.getSubject(); // ★件名を取得
      if (message.getFrom().includes(GMAIL_SENDER_EXCLUDE)) {
        console.log(`除外SENDERのためスキップ: ${message.getFrom()}`);
        continue; // continue の前に finally が実行される
      }
      const body = message.getPlainBody();

      // ★ 抽出した情報を格納する変数を定義
      let workDate, department, clinic, time, name;

      // --- 1. 件名に応じて処理を分岐 ---
      if (subject.includes('★★緊急★★')) {
        // --- 形式1（従来）：「★★緊急★★＜勤務2週間以内のキャンセル申請＞」
        const mainRegex = /(\d{4}年\d{1,2}月\d{1,2}日)\s+(.+?)\s+(.+?)、\s*(\d{1,2}:\d{2}~\d{1,2}:\d{2})\s+(.+?)のキャンセルが申請されました/;
        const match = body.match(mainRegex);
        if (!match) {
          console.warn(`「★★緊急★★」メールの本文解析に失敗: ${subject}`);
          continue;
        }
        [ , workDate, department, clinic, time, name] = match.map(item => item.trim());

      } else if (subject.includes('【要対応】勤務3日以内')) {
        // --- 形式2（新規）：「【要対応】勤務3日以内」のキャンセル申請」
        const nameMatch = body.match(/医師名：(.+)/);
        const clinicMatch = body.match(/勤務拠点名：(.+)/);
        const dateMatch = body.match(/キャンセル日付：(.+)/);
        const timeMatch = body.match(/勤務時間：(.+)/);

        // 4つの情報が揃わなければ解析失敗
        if (!nameMatch || !clinicMatch || !dateMatch || !timeMatch) {
          console.warn(`「3日以内」メールの本文解析に失敗: ${subject}`);
          continue;
        }

        name = nameMatch[1].trim();
        const clinicRaw = clinicMatch[1].trim(); // 例: "国立 小児科"
        workDate = dateMatch[1].trim(); // 例: "2025年11月2日"
        time = timeMatch[1].trim().replace('~', '～'); // チルダを全角に統一

        // 「勤務拠点名」から拠点と診療科を分離（例: "国立 小児科" -> "国立" と "小児科"）
        const parts = clinicRaw.split(/\s+/); // スペースで分割
        if (parts.length >= 2) {
          department = parts.pop(); // 最後の要素を診療科とする ( "小児科" )
          clinic = parts.join(' '); // 残りを拠点とする ( "国立" )
        } else {
          clinic = clinicRaw;
          department = '（診療科不明）';
        }

      } else {
        // GMAIL_QUERY_CANCEL に一致したが、上記のどちらでもない場合
        console.warn(`不明なキャンセルメール形式です: ${subject}`);
        continue;
      }
      // --- ここまで分岐処理 ---


      // --- 2. 共通の通知処理 ---
      const receivedTime = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');
      
      // バックアップシフトか、当日・翌日以外は除外
      if (clinic === '【関東】バックアップシフト' || (workDate !== todayStr && workDate !== tomorrowStr)) {
        console.log(`通知対象外（バックアップ or 日付）のためスキップ: ${workDate} ${clinic}`);
        continue;
      }
      
      // 重複チェック
      SpreadsheetApp.flush();
      const latestSheetData = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
      const isDuplicate = latestSheetData.some(row => String(row[1]).trim() === workDate && String(row[2]).trim() === clinic && String(row[4]).trim() === name && String(row[5]).trim() === time);
      
      if (!isDuplicate) {
        // シートに書き込み
        console.log(`新規キャンセルを検知、シートに書き込みます: ${workDate} ${clinic} ${name}`);
        const newId = sheet.getLastRow();
        sheet.appendRow([newId, workDate, clinic, department, name, time, receivedTime]);
        
        // Chatworkに通知
        const chatworkMessage = `[toall]\n[info][title]直近緊急キャンセル[/title]勤務予定医師より緊急キャンセルが申請されました。\n詳細を確認してください。\n\n勤務日： ${workDate}\n勤務拠点： ${clinic}\n診療科： ${department}\n勤務者： ${name}\n勤務時間： ${time}\n受信時刻： ${receivedTime}[/info]`;
        // ★★★【重要】共通関数 sendToChatwork を呼び出す ★★★
        sendToChatwork(CANCEL_LOG_CHATWORK_ROOM_ID, chatworkMessage);
      } else {
        console.log(`処理済み（シートに重複あり）のためスキップ: ${workDate} ${clinic} ${name}`);
      }

    } finally {
      // ▼▼▼ 【ラベル修正】 ▼▼▼
      // スキップされた場合も、通知・重複検知が完了した場合も、
      // 必ず「処理済みラベル」をスレッドに付けて、次回以降の検索対象から除外する
      thread.addLabel(label);
      // ▲▲▲ 【ラベル修正】 ▲▲▲
    }
  } // end for loop
}


/**
 * ★★★【修正】大司令塔から呼び出される内部関数に変更 ★★★
 * (checkTimegaiCallWithLabel -> checkTimegaiCallWithLabel_internal)
 */
function checkTimegaiCallWithLabel_internal() {
  // ★★★【修正】ロック処理を削除（大司令塔が担当） ★★★

  console.log('【時間外着信通知】処理を開始します。');
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_TIMEGAI);
  if (!label) { label = GmailApp.createLabel(PROCESSED_LABEL_TIMEGAI); }

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const searchDate = Utilities.formatDate(yesterday, 'JST', 'yyyy/MM/dd');
  const finalQuery = `${GMAIL_QUERY_TIMEGAI} -label:${PROCESSED_LABEL_TIMEGAI} after:${searchDate}`;
  const threads = GmailApp.search(finalQuery);
  if (threads.length === 0) { console.log('処理対象のメールはありませんでした。'); return; }
  for (const thread of threads) {
    const message = thread.getMessages()[0];
    const receivedTime = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');
    const chatworkBody = `[toall]\n[info][title]時間外着信のお知らせ[/title]受信時刻： ${receivedTime}\n時間外応答がありました。\nマイページにログインして内容をご確認ください。\nhttps://minderu.com/mypage/telephones/[/info]`;
    // ★★★【重要】共通関数 sendToChatwork を呼び出す ★★★
    sendToChatwork(TIMEGAI_CALL_CHATWORK_ROOM_ID, chatworkBody);
    thread.addLabel(label);
  }
}


/**
 * ★★★【新規】大司令塔から呼び出される内部関数 ★★★
 * 非定型キャンセル（ドタキャン）メールを検知して通知する
 * (★ nicho.co.jp など除外対象を追加 ★)
 */
function checkInformalCancel_internal() {
  // ★★★ ロック処理は「大司令塔」が担当 ★★★

  console.log('【非定型キャンセル通知】処理を開始します。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 設定シートとログシートの取得
  const configSheet = ss.getSheetByName(DOTAKYAN_CONFIG_SHEET_NAME);
  const logSheet = ss.getSheetByName(DOTAKYAN_LOG_SHEET_NAME);

  if (!configSheet) { console.error(`シート「${DOTAKYAN_CONFIG_SHEET_NAME}」が見つかりません。`); return; }
  if (!logSheet) { console.error(`シート「${DOTAKYAN_LOG_SHEET_NAME}」が見つかりません。`); return; }
  
  // 2. ログシートのヘッダー確認 (G列まで)
  if (logSheet.getRange('A1').getValue() === '') {
    logSheet.getRange('A1:G1').setValues([['番号', '受信時刻', '差出人', '送信先', '件名', '内容', 'ユニークキー']]);
  }

  // 3. Gmailラベルの準備
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_DOTAKYAN);
  if (!label) { label = GmailApp.createLabel(PROCESSED_LABEL_DOTAKYAN); }

  // 4. キーワードの読み込み (A列: 時間, B列: 内容)
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) { console.log(`シート「${DOTAKYAN_CONFIG_SHEET_NAME}」にキーワードがありません。`); return; }
  
  const configData = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const timeKeywords = configData.map(row => row[0]).filter(String); // A列
  const actionKeywords = configData.map(row => row[1]).filter(String); // B列

  if (timeKeywords.length === 0 || actionKeywords.length === 0) {
    console.log('A列またはB列のキーワードが空のため、処理をスキップします。');
    return;
  }

  // 5. Gmail検索クエリの構築 (キーワードのみで検索)
  const timeQuery = `(${timeKeywords.join(' OR ')})`;
  const actionQuery = `(${actionKeywords.join(' OR ')})`;
  const otherFilters = `-label:${PROCESSED_LABEL_DOTAKYAN} newer_than:1d`;
  const gmailQuery = `${timeQuery} ${actionQuery} ${otherFilters}`;

  console.log(`Gmail検索クエリ: ${gmailQuery}`);

  // 6. 既存ログの読み込み (G列のメッセージIDのみ)
  SpreadsheetApp.flush();
  const logLastRow = logSheet.getLastRow();
  let existingMessageIDs = [];
  if (logLastRow > 1) {
    existingMessageIDs = logSheet.getRange(2, 7, logLastRow - 1, 1).getValues().flat();
  }

  // 7. 除外リストの定義（ドメインまたはメールアドレスの一部）
  const userEmail = Session.getActiveUser().getEmail();
  const excludeSenders = [
    'caps365.jp', 
    'mns.jp', 
    'mnys.jp', 
    'm3career.com', 
    'gemini-notes@google.com', 
    'asiantms.com', // ★追加：ATMS
    'mediwel.net',  // ★追加：メディウェル
    'nicho.co.jp',  // ★追加：日本調剤
    userEmail
  ];

  // 8. ★キーワードを正規表現オブジェクトに変換（本文チェック用）
  const timeRegex = new RegExp(timeKeywords.join('|'), 'i');
  const actionRegex = new RegExp(actionKeywords.join('|'), 'i');

  // 9. ★24時間前のDateオブジェクトを定義
  const now = new Date();
  const twentyFourHoursAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));

  // 10. Gmail検索と処理
  const threads = GmailApp.search(gmailQuery);
  if (threads.length === 0) { console.log('対象のメールはありませんでした。'); return; }

  for (const thread of threads) {
    let processedThisThread = false; // このスレッドで通知・ログ記録したか
    try {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        // --- 4つの条件でメールを厳密にチェック ---
        const messageId = message.getId();

        // 【条件1】ログに重複がないか (二重通知防止)
        if (existingMessageIDs.includes(messageId)) {
          continue; 
        }

        // 【条件2】受信日時が過去24時間以内か (古いメールを除外)
        const receivedDate = message.getDate();
        if (receivedDate < twentyFourHoursAgo) {
          continue; 
        }

        // 【条件3】送信元が社外か (除外リストのチェック)
        const from = message.getFrom();
        const isInternal = excludeSenders.some(senderKeyword => from.includes(senderKeyword));
        
        if (isInternal) {
          console.log(`[メールID: ${messageId}] 除外対象の送信元のためスキップ: ${from}`);
          continue; 
        }

        // 引用文を削除 (強化版)
        const fullBody = message.getPlainBody();
        const regexQuoteMarkers = /\n(>|On .*> wrote:|.* <.*@.*> wrote:|\d{4}年\d{1,2}月\d{1,2}日.*:|From: .*|Sent: .*|-{3,}|={3,}|#{3,}|_+\s*$)/i;
        const bodyOnly = fullBody.split(regexQuoteMarkers)[0].trim();

        // 【条件4】「引用文を除いた本文」にキーワード(A+B)が両方含まれるか
        const hasTimeKeyword = timeRegex.test(bodyOnly);
        const hasActionKeyword = actionRegex.test(bodyOnly);

        if (!hasTimeKeyword || !hasActionKeyword) {
          console.log(`[メールID: ${messageId}] 引用文除去後の本文にキーワード(A+B)が含まれないためスキップ`);
          continue;
        }
        
        // --- すべての条件をクリア ---
        console.log(`★検知成功 [メールID: ${messageId}]`);
        processedThisThread = true;

        // 11. 新規処理
        const receivedTime = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');
        const to = message.getTo();
        const subject = message.getSubject();
        const bodySnippet = bodyOnly.substring(0, 300); // 引用文を除いた本文のスニペット

        // 12. ログシートに書き込み
        const newId = logSheet.getLastRow();
        logSheet.appendRow([newId, receivedTime, from, to, subject, bodySnippet, messageId]);
        
        existingMessageIDs.push(messageId); // メモリ内リストにも追加

        // 13. Chatwork通知
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

      } // --- メッセージループ終了 ---
      
    } catch (e) {
      console.error(`メール処理中にエラー: ${e.message}`);
    } finally {
      // 14. 処理済みラベルをスレッドに付与
      thread.addLabel(label);
    }
  } // --- スレッドループ終了 ---
}