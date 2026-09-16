// =============================================================
// 【ファイル5：エムスリーメッセージ処理】(オファー希望 本文完全一致対応版)
// 対象：当日受信分のみ
// 通知先：419888887
// ★修正内容：API制限対策として getMessagesForThreads による一括取得に変更
// =============================================================

const M3_MSG_CONFIG = {
  SHEET_NAME: 'エムスリーメッセージ',
  CHATWORK_ROOM_ID: '419888887', // 本番用ルームID
  // ★修正：検索条件から「へ」を削除し、「お問い合わせがありました」で確実にヒットさせます
  SEARCH_QUERY: `from:career_spot@m3career.com subject:("メッセージを受信しました" OR "お問い合わせがありました" OR "キャンセルいたしました" OR "キャンセルが申請されました") -subject:("ご勤務希望がありました" OR "辞退されました") newer_than:1d`
};

/**
 * ★ エムスリーメッセージ処理のメイン関数
 * (大司令塔から呼び出されます)
 */
function processM3Messages_internal() {
  console.log("--- エムスリーメッセージ処理を開始 ---");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(M3_MSG_CONFIG.SHEET_NAME);

  // 1. シート初期化（なければ作成）
  if (!sheet) {
    sheet = ss.insertSheet(M3_MSG_CONFIG.SHEET_NAME);
    // 指定の9項目
    sheet.getRange('A1:I1').setValues([['受信日時', '分類', '医師名', '勤務日', '拠点名', '件名', '本文', 'Message-ID', '年月']]);
    sheet.setFrozenRows(1);
    console.log(`シート「${M3_MSG_CONFIG.SHEET_NAME}」を新規作成しました。`);
  }

  // 2. 既存のMessage-ID（重複チェック用）を取得
  const lastRow = sheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    // H列(8列目)にあるMessage-IDを取得
    existingIds = sheet.getRange(2, 8, lastRow - 1, 1).getValues().flat();
  }

  // 3. Gmail検索実行
  const threads = GmailApp.search(M3_MSG_CONFIG.SEARCH_QUERY);
  if (threads.length === 0) {
    console.log("処理対象のメールはありませんでした。");
    return;
  }

  // ★「当日」の判定用（0時0分0秒）
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const newRows = []; // 書き込み用バッファ

  // ★ API通信1回で全メッセージを一括取得
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    
    for (const message of messages) {
      // 日付チェック：今日以前のメールは無視
      if (message.getDate() < todayStart) {
        continue;
      }

      const messageId = message.getId();

      // 重複チェック
      if (existingIds.includes(messageId)) continue;
      // 送信元チェック（念のため）
      if (!message.getFrom().includes('career_spot@m3career.com')) continue;

      // --- 解析処理 ---
      const subject = message.getSubject();
      const rawBody = message.getPlainBody();
      const receivedDate = message.getDate();
      const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');

      // ★追加：オファー希望判定用（改行や空白の揺らぎを完全に吸収）
      const cleanBodyForCheck = rawBody.replace(/[\r\n\s ]+/g, "");
      const targetKeyword = "■オファーを出される前にご確認ください・医師のご専門や可能／不可能な手技など必要に応じて直接先生にご確認ください。";

      // 分類判定
      let type = '';
      let typeTitle = ''; 

      if (subject.includes('メッセージを受信しました')) {
        type = 'メッセージ受信';
        typeTitle = 'メッセージ受信';
      } else if (subject.includes('お問い合わせがありました')) {
        // ★修正：本文の空白・改行を無視したテキストでキーワードを完全一致チェック
        if (cleanBodyForCheck.includes(targetKeyword)) {
           type = 'オファー希望';
           typeTitle = 'オファー希望';
        } else {
           type = 'お問い合わせ';
           typeTitle = '問い合わせ';
        }
      } else if (subject.includes('キャンセルいたしました') || subject.includes('キャンセルが申請されました')) {
        type = '応募キャンセル';
        typeTitle = 'キャンセル';
      } else {
        continue; 
      }

      // 情報抽出（オファー希望専用の抽出と、通常の抽出を分岐）
      let doctorName = '不明';
      let workDate = '不明';
      let clinic = '不明';
      let cleanBody = '';
      let url = 'https://career.m3.com/admin/hospital/message_threads';

      if (type === 'オファー希望') {
         // --- オファー希望専用の抽出ロジック ---
         const docMatch = rawBody.match(/氏名[\s:：]*([^\n\r]+)/);
         doctorName = docMatch ? docMatch[1].trim() : "不明";

         const clinicMatch = rawBody.match(/勤務地[\s:：]*([^\n\r]+)/);
         if (clinicMatch) {
             const rawClinic = clinicMatch[1].trim();
             // 最初の数字（半角・全角）の手前まででカット
             const shortClinicMatch = rawClinic.match(/^([^\d０-９]+)/);
             clinic = shortClinicMatch ? shortClinicMatch[1].trim() : rawClinic;
         }

         const urlMatch = rawBody.match(/(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/);
         if (urlMatch) url = urlMatch[1].trim();
         
         const dateMatch = rawBody.match(/勤務時間[\s:：]*(\d{1,2}\/\d{1,2}\([月火水木金土日]\))/);
         if (dateMatch) workDate = String(receivedDate.getFullYear()) + "年" + dateMatch[1].replace('/', '月').replace('(', '日 (');
         
         cleanBody = `${doctorName}先生から【${clinic}】へのオファー希望がありました。内容を確認してください。`;

      } else {
         // --- 通常の抽出ロジック ---
         const info = extractM3Info_Final(subject, rawBody);
         doctorName = info.doctorName;
         workDate = info.workDate;
         clinic = info.clinic;
         cleanBody = extractM3CoreMessage_Final(rawBody, type);
      }

      // ★日付の正規化（D列、I列用）
      const dateObj = normalizeDate_Final(workDate, receivedDate);
      
      // 日本語曜日への変換
      const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
      const dayStr = dayMap[dateObj.getDay()];
      const formattedDate = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd') + ` (${dayStr})`;
      
      const yearMonth = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM');

      console.log(`[新規検知] ${type}: ${doctorName} / ${formattedDate}`);

      // シート用データ作成 (9列)
      newRows.push([
        receivedDateStr,
        type,
        doctorName,
        formattedDate, // D列：yyyy/MM/dd (曜日)
        clinic,
        subject,
        cleanBody,     // 整形済み本文
        messageId,
        yearMonth      // I列：yyyy/MM
      ]);

      existingIds.push(messageId); // 重複リストに追加

      // Chatwork通知の作成と送信
      let cwMessage = "";

      if (type === 'オファー希望') {
         // オファー希望専用のフォーマット
         cwMessage = `[toall]\n[info][title]オファー希望[/title]\n${cleanBody}\n▼医師情報・メッセージ詳細ページはこちら\n${url}\n[/info]`;
      } else {
         // 通常のフォーマット（問い合わせ時は医師名カット）
         let doctorLine = `医師名： ${doctorName}\n`;
         if (type === 'お問い合わせ') doctorLine = ''; 

         cwMessage = `[info][title]${typeTitle}[/title]\n${doctorLine}勤務日： ${formattedDate}\n拠点名： ${clinic}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n${cleanBody}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nリンク：https://career.m3.com/admin/hospital/message_threads\n[/info]`;
      }

      sendToChatwork(M3_MSG_CONFIG.CHATWORK_ROOM_ID, cwMessage);
    }
  }

  // 4. まとめてシートに書き込み
  if (newRows.length > 0) {
    // 最終行の次から書き込み
    const targetRange = sheet.getRange(lastRow + 1, 1, newRows.length, 9);
    targetRange.setValues(newRows);
    targetRange.setHorizontalAlignment('left'); // 左詰め
    console.log(`${newRows.length} 件のメッセージを処理・記録しました。`);
  }
}

/**
 * 日付文字列を解析し、正しい年のDateオブジェクトを返す
 */
function normalizeDate_Final(dateStr, receivedDate) {
  // 既に年がある場合
  let match = dateStr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (match) return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  
  match = dateStr.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (match) return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));

  // 年がない場合（年またぎ判定）
  match = dateStr.match(/(\d{1,2})月(\d{1,2})日/);
  if (match) {
    const month = parseInt(match[1]);
    const day = parseInt(match[2]);
    let year = receivedDate.getFullYear();
    const receivedMonth = receivedDate.getMonth() + 1;

    if (receivedMonth >= 11 && month <= 3) year++;     // 年越し
    else if (receivedMonth <= 2 && month >= 10) year--; // 年戻り

    return new Date(year, month - 1, day);
  }
  return receivedDate;
}

/**
 * 件名と本文から情報を抽出
 */
function extractM3Info_Final(subject, body) {
  let doctorName = '不明';
  let workDate = '不明';
  let clinic = '不明';

  // 1. 医師名（日付除去）
  const nameMatch = subject.match(/(?:^\d+月\d+日\s*)?(.+?)先生[よりが]/);
  if (nameMatch) {
    doctorName = nameMatch[1].trim() + ' 先生';
  } else {
    const nameInBody = body.match(/医師名:\s*(.+?)[\(\n]/);
    if (nameInBody) doctorName = nameInBody[1].trim() + ' 先生';
  }

  // 2. 勤務日
  const jobMatch = body.match(/(\d{4}年\d{1,2}月\d{1,2}日.*?)[／C]/);
  if (jobMatch) {
    workDate = jobMatch[1].trim();
  } else {
    const dateInSubject = subject.match(/スポット\s*(\d{1,2}月\d{1,2}日.*?)\s/);
    if (dateInSubject) workDate = dateInSubject[1].trim();
  }

  // 3. 拠点名
  const clinicMatch = body.match(/^医療法人社団.*?(キャップスクリニック.+?)様/m);
  if (clinicMatch) {
    clinic = clinicMatch[1].replace(/キャップスクリニック/g, '').trim();
  }

  return { doctorName, workDate, clinic };
}

/**
 * 本文の核心部分切り出し
 */
function extractM3CoreMessage_Final(body, type) {
  let extracted = "";

  if (type === 'メッセージ受信') {
    const match = body.match(/(タイトル：[\s\S]*?)(?=━{5,})/);
    if (match) extracted = match[1].trim();

  } else if (type === 'お問い合わせ') {
    const match = body.match(/(【医師からのメッセージ】[\s\S]*?)(?=━{5,})/);
    if (match) extracted = match[1].trim();

  } else if (type === '応募キャンセル') {
    const match = body.match(/様\s*\n+([\s\S]*?)(?=━{5,})/);
    if (match) {
      extracted = match[1].trim();
      extracted = extracted.replace(/^いつも.*?ありがとうございます。[\r\n]*/, '');
      extracted = extracted.replace(/^平素は.*?ありがとうございます。[\r\n]*/, '');
    }
  }

  if (!extracted) {
    extracted = "(本文の自動抽出に失敗しました。リンク先をご確認ください)";
  }

  return extracted;
}