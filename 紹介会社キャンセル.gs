/**
 * 【本番用】紹介会社キャンセルメールの検知・自動反映スクリプト
 * 大司令塔から定期的に呼び出されます。
 */
function processAgencyCancelEmails_internal() {
  console.log("=============== 紹介会社キャンセル自動反映 開始 ===============");

  // 1. 設定値
  const TARGET_CW_ROOM_ID = '415529974'; // 指定のChatworkルーム
  const LABEL_NAME = '処理済み-紹介会社キャンセル';
  
  // シート取得（アーカイブシートを対象）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // ※環境によってシート名が違う場合は '処理済み' や 'アーカイブ' に変更してください
  const archiveSheet = ss.getSheetByName('アーカイブ') || ss.getSheetByName('処理済み');
  if (!archiveSheet) {
    console.error("アーカイブ（処理済み）シートが見つかりません。");
    return;
  }

  // 2. Gmailの検索（直近2日間のメールに限定してAPI消費を抑える）
  const query = `(subject:"取消のご連絡" OR subject:"【キャンセル】") (from:medical-principle.co.jp OR from:medrt.com OR from:mnys.jp) newer_than:2d -label:${LABEL_NAME}`;
  const threads = GmailApp.search(query, 0, 20);
  
  if (threads.length === 0) {
    console.log("対象のキャンセルメールはありませんでした。");
    return;
  }

  // ラベルの準備
  let processedLabel = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!processedLabel) { processedLabel = GmailApp.createLabel(LABEL_NAME); }

  const allMessages = GmailApp.getMessagesForThreads(threads);
  const processedJobIds = new Set(); // 重複・増殖防止ストッパー

  // 3. シートの列設定（必要に応じて調整）
  const colAgency = 1, colDate = 2, colTime = 3, colClinic = 4, colDept = 5;
  const colId = 6, colDoctor = 7, colUrl = 8, colStatus = 9;
  const colTime1 = 10, colTime2 = 11, colPerson = 12, colSaiyo = 13;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadProcessed = false;
    
    for (const message of messages) {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      
      let agency = '', jobId = '', clinic = '', doctorName = '', workDate = '', workTime = '';

      // 【民間医局】の抽出
      if (subject.includes('取消のご連絡') && subject.includes('民間医局')) {
        agency = '民間医局';
        const jobIdMatch = body.match(/案件番号[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
        const clinicMatch = body.match(/勤務先[\s ]*[:：][\s ]*([^\n\r]+)/);
        const docMatch = body.match(/ドクター名[\s ]*[:：][\s ]*([^\n\r]+)/);
        const dateMatch = body.match(/勤務日程[\s ]*[:：][\s ]*([^\n\r]+)/);
        const timeMatch = body.match(/勤務時間[\s ]*[:：][\s ]*([^\n\r]+)/);

        if (jobIdMatch) jobId = jobIdMatch[1].trim();
        if (clinicMatch) clinic = clinicMatch[1].replace(/キャップスクリニック/g, '').trim();
        if (docMatch) doctorName = docMatch[1].trim();
        if (dateMatch) workDate = dateMatch[1].match(/(\d{4}年\d{1,2}月\d{1,2}日)/) ? dateMatch[1].match(/(\d{4}年\d{1,2}月\d{1,2}日)/)[1] : dateMatch[1].trim();
        if (timeMatch) workTime = timeMatch[1].trim();

      // 【MRT】の抽出
      } else if (subject.includes('【キャンセル】') && subject.includes('MRT')) {
        agency = 'MRT';
        const jobIdMatch = body.match(/ID\s*:\s*(\d+)/);
        const clinicMatch = body.match(/クリニック\s*:\s*\d+\s*([^\n\r]+)/);
        const dateMatch = body.match(/日時\s*:\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
        const timeMatch = body.match(/日時\s*:.*?\)\s*(\d{1,2}:\d{2}[〜~～\-]\d{1,2}:\d{2})/);
        const docMatch = body.match(/応募者\s*:\s*([^\n\r]+?)先生/);

        if (jobIdMatch) jobId = jobIdMatch[1].trim();
        if (clinicMatch) clinic = clinicMatch[1].replace(/キャップスクリニック/g, '').trim();
        if (docMatch) doctorName = docMatch[1].trim();
        if (dateMatch) workDate = dateMatch[1].trim();
        if (timeMatch) workTime = timeMatch[1].trim();
      }

      if (!jobId) continue; 

      // 重複ストッパー（スレッド内の重複を排除）
      if (processedJobIds.has(jobId)) continue;
      processedJobIds.add(jobId);

      // 最新のシートデータを取得（差し込みがあった場合も考慮して毎回取得）
      const sheetData = archiveSheet.getDataRange().getValues();
      let foundRowIndex = -1;

      // 第1検索：識別番号(F列)で検索
      for (let r = 0; r < sheetData.length; r++) {
        if (String(sheetData[r][colId - 1]).trim() === jobId) {
          foundRowIndex = r;
          break;
        }
      }

      // 第2検索：識別番号が無い場合、会社＋日付＋拠点＋名前 で検索
      if (foundRowIndex === -1) {
        let normalizedMailDate = workDate.replace(/年|月/g, '/').replace(/日/g, '');
        let cleanMailName = doctorName.replace(/\s+/g, '');

        for (let r = 0; r < sheetData.length; r++) {
          let sheetAgency = String(sheetData[r][colAgency - 1]).trim();
          let sheetDate = String(sheetData[r][colDate - 1]).trim();
          let sheetClinic = String(sheetData[r][colClinic - 1]).trim();
          let sheetName = String(sheetData[r][colDoctor - 1]).replace(/\s+/g, '');

          if (sheetAgency === agency && sheetDate === normalizedMailDate && sheetClinic === clinic && sheetName === cleanMailName) {
            foundRowIndex = r;
            break;
          }
        }
      }

      let targetRowNumber;
      let isInserted = false;

      if (foundRowIndex !== -1) {
        // 【存在するパターン】対象行のみをピンポイント更新
        targetRowNumber = foundRowIndex + 1;
        archiveSheet.getRange(targetRowNumber, colStatus).setValue('キャンセル済み');
        archiveSheet.getRange(targetRowNumber, colSaiyo).setValue('不採用');
        // A列からM列（13列）まで背景色を黄色に変更
        archiveSheet.getRange(targetRowNumber, 1, 1, 13).setBackground('#ffe599');
      } else {
        // 【存在しないパターン】新規差し込み作成
        isInserted = true;
        const newRow = new Array(13).fill('');
        newRow[colAgency - 1] = agency;
        newRow[colDate - 1] = workDate;
        newRow[colTime - 1] = workTime;
        newRow[colClinic - 1] = clinic;
        newRow[colDept - 1] = '小児科'; // 仮置き
        newRow[colId - 1] = jobId;
        newRow[colDoctor - 1] = doctorName;
        newRow[colStatus - 1] = 'キャンセル済み';
        const nowStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
        newRow[colTime1 - 1] = nowStr;
        newRow[colTime2 - 1] = nowStr;
        newRow[colPerson - 1] = '自動連携';
        newRow[colSaiyo - 1] = '不採用';

        archiveSheet.appendRow(newRow);
        targetRowNumber = archiveSheet.getLastRow();
        archiveSheet.getRange(targetRowNumber, 1, 1, 13).setBackground('#ffe599');
      }

      // URL生成
      const targetRowUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${archiveSheet.getSheetId()}&range=A${targetRowNumber}`;

      // 通知メッセージ生成
      const chatworkMessage = `[info][title]紹介会社キャンセル自動反映[/title]
紹介会社からのキャンセル連絡を検知し、シートの該当行を「不採用」「キャンセル済み」に自動更新${isInserted ? '（※該当行なしのため新規追加）' : ''}しました。

[code]
紹介会社： ${agency}
案件番号： ${jobId}
勤務時間： ${workTime}
医師名 ： ${doctorName} 先生
勤務日 ： ${workDate}
勤務先 ： ${clinic}
[/code]

▼ 処理済みシートの該当行はこちら（クリックで直接飛びます）
${targetRowUrl}
[/info]`;

      // 通知送信
      sendToChatwork(TARGET_CW_ROOM_ID, chatworkMessage);
      threadProcessed = true;
      Utilities.sleep(1500); // 連続送信時のAPIエラー回避
    }

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }

  console.log("=============== 紹介会社キャンセル自動反映 完了 ===============");
}