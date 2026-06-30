// =================================================================
// 【1-C_その他紹介会社処理】
// エムステージ、MRT（応募・取消し自動処理対応）、民間医局、メディウェル
// =================================================================

// === エムステージ（スレッド展開・堅牢抽出対応）===
function processMStage(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const processedLabelName = processedLabel.getName();
  
  GmailApp.search(`from:(ra@mstage-corp.jp) subject:("貴院の求人に応募がありました") ${querySuffix}`).forEach(thread => {
    const messages = thread.getMessages();
    let threadProcessed = false;

    messages.forEach(message => {
      const body = message.getPlainBody();
      const receivedDate = message.getDate();
      
      const dateLine = extract(body, /勤務日[\s　]*[:：][\s　]*(.+)/);
      const doctorName = extract(body, /お名前\(漢字\)[\s　]*[:：][\s　]*(.+)/).replace(/[\s　]/g, '');
      let department = extract(body, /募集科目[\s　]*[:：][\s　]*(.+)/);
      if (!department) department = findDepartmentFromHistory(ss, doctorName) || '';
      
      const id = extract(body, /(?:求人番号|案件ID|求人ID)[\s　]*[:：][\s　]*([A-Za-z0-9\-]+)/);
      const clinic = extract(body, /キャップスクリニック([^\r\n]+)/).replace(/様$/, '').trim();
      const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})/);

      if (doctorName && id) {
        const added = appendRowAndFormatUnique(sheet, [
          'エムステージ', formatDate(dateLine, receivedDate), workTime, 
          clinic, department, 
          id, 
          doctorName, extract(body, /(https:\/\/www\.doctor-navi\.jp\/medical\/entries.*)/), 
          '', '応募通知受信', receivedDate, '', ''
        ]);
        if (added) threadProcessed = true;
      }
    });
    
    if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
       thread.addLabel(processedLabel);
    }
  });

  const mStageConfirmQuery = `from:(ra@mstage-corp.jp) subject:("勤務が確定" OR "キャンセル" OR "書類選考の結果を通知しました(NG)") ${confirmQuerySuffix}`;
  GmailApp.search(mStageConfirmQuery).forEach(thread => {
      const messages = thread.getMessages();
      let threadProcessed = false;

      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        
        let type = '';
        let status = '';
        if (subject.includes('勤務が確定')) { type = '確定'; status = '勤務確定'; } 
        else if (subject.includes('キャンセル')) { type = 'お断り'; status = 'キャンセル済み'; } 
        else if (subject.includes('書類選考の結果を通知しました(NG)')) { type = 'お断り'; status = 'お断り済み'; }
        
        if (!type) return;

        const jobId = extract(body, /(?:求人番号|案件ID|求人ID)[\s　]*[:：][\s　]*([A-Za-z0-9\-]+)/);
        const doctorNameMatch = extract(body, /お名前\(漢字\)[\s　]*[:：][\s　]*(.+)/) || extract(subject, /\[(.+?)医師\]/);
        const nameToFind = doctorNameMatch ? doctorNameMatch.replace(/\s/g, '') : '';
        
        let rowNum = -1;
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i > 0; i--) {
            if (String(data[i][0]).trim() !== 'エムステージ') continue;
            const rowStatus = String(data[i][9]).trim();
            if (!rowStatus.includes('応募通知受信')) continue;
            
            const rowDocClean = String(data[i][6]).split(/[\/／（(]/)[0].replace(/\s/g, '');
            const rowJobId = String(data[i][5]).trim();
            
            if (nameToFind && rowDocClean === nameToFind && jobId && rowJobId === jobId) {
                rowNum = i + 1; 
                break;
            }
        }
        
        if (rowNum !== -1) {
            let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
            values[8] = type; values[9] = status; values[11] = new Date();
            archiveSheet.appendRow(values);
            archiveSheet.getRange(archiveSheet.getLastRow(), 1, 1, archiveSheet.getLastColumn()).setHorizontalAlignment('left');
            safeDeleteRow(sheet, rowNum);
            threadProcessed = true;
        }
      });
      
      if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
         thread.addLabel(processedLabel);
      }
  });
}

// === MRT（スレッド対応・応募取消し自動処理搭載版） ===
function processMRT_Application(querySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const processedLabelName = processedLabel.getName();
  
  // === 1. 応募通知の処理 ===
  GmailApp.search(`subject:("【MRT】") subject:("応募通知") ${querySuffix}`).forEach(thread => {
    if (thread.getLabels().some(l => l.getName() === processedLabelName)) return; 

    let threadProcessed = false;
    // スレッド化対策：すべてのメールを展開して処理
    thread.getMessages().forEach(message => {
      const body = message.getPlainBody();
      const receivedDate = message.getDate();
      
      const dateLine = extract(body, /日時[\s:：]+([^\n\r]+)/);
      const id = extract(body, /ID[\s:：]*(\d+)/i);
      const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[〜~〜\-]\s*\d{1,2}:\d{2})/);
      const clinic = extract(body, /クリニック[\s:：]*\d+\s*キャップスクリニック([^\n\r]+)/);
      const department = extract(body, /診療科[\s:：]*([^\n\r,]+)/);

      if (id && dateLine) {
        const added = appendRowAndFormatUnique(sheet, [
          'MRT', formatDate(dateLine, receivedDate), workTime, 
          clinic, department, 
          id, 
          '', extract(body, /(https:\/\/medical\.medrt\.com\/customers\/sessions\/new)/), 
          '', '応募通知受信（手動更新要）', receivedDate, '', ''
        ]);
        if (added) threadProcessed = true;
      }
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  });

  // === 2. 応募取消し（キャンセル）の自動処理 ===
  // ※キャンセルの遅延を拾うため、ここだけ検索期間を14日間に広げています
  const mrtCancelQuery = `subject:("【MRT】") subject:("応募取消し") newer_than:14d`;
  GmailApp.search(mrtCancelQuery).forEach(thread => {
    if (thread.getLabels().some(l => l.getName() === processedLabelName)) return;

    let threadProcessed = false;
    thread.getMessages().forEach(message => {
      const body = message.getPlainBody();
      
      // IDの抽出（テスト済みの確実な正規表現）
      const idMatch = body.match(/ID[\s:：]*(\d+)/i);
      const mailJobId = idMatch ? idMatch[1].trim() : "";
      
      if (!mailJobId) return;

      // 医師名の抽出 (テスト済みの確実な正規表現)
      const docMatch = body.match(/応募者[\s:：]*([^\n\r]+?)先生/);
      const mailDoctorName = docMatch ? docMatch[1].trim() : "";

      // 進捗シートから合致するIDの行を下から探す
      let rowNum = -1;
      const sheetData = sheet.getDataRange().getValues();
      for (let i = sheetData.length - 1; i > 0; i--) {
        const rowAgency = String(sheetData[i][0]).trim();
        const rowJobId = String(sheetData[i][5]).trim();
        const rowStatus = String(sheetData[i][9]).trim();
        
        // MRTかつIDが一致し、まだ処理待ち（応募通知受信）のもの
        if (rowAgency === 'MRT' && rowJobId === mailJobId && rowStatus.includes('応募通知受信')) {
          rowNum = i + 1;
          break;
        }
      }

      // 進捗シートに見つかった場合、アーカイブへ移動
      if (rowNum !== -1) {
        let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // 空欄だった医師名を補完
        if (mailDoctorName) values[6] = mailDoctorName; 
        
        values[8] = 'お断り'; 
        values[9] = 'キャンセル済み'; 
        values[11] = new Date();
        
        archiveSheet.appendRow(values);
        archiveSheet.getRange(archiveSheet.getLastRow(), 1, 1, archiveSheet.getLastColumn()).setHorizontalAlignment('left');
        safeDeleteRow(sheet, rowNum);
        
        threadProcessed = true;
      }
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  });
}

// === 民間医局（識別番号・揺らぎ対応済） ===
function processMinkanIkyoku_Application(querySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const processedLabelName = processedLabel.getName();
  
  GmailApp.search(`subject:("【民間医局ポータル】") subject:("申し込みがありました") ${querySuffix}`).forEach(thread => {
    const messages = thread.getMessages();
    let threadProcessed = false;

    messages.forEach(message => {
      const body = message.getPlainBody();
      const receivedDate = message.getDate();
      
      const doctorName = extract(body, /医師名[\s　]*[:：][\s　]*(.+)/).replace(/\s/g, '');
      let department = extract(body, /診療科目[\s　]*[:：][\s　]*(.+)/);
      if (!department) department = findDepartmentFromHistory(ss, doctorName) || '';
      
      const id = extract(body, /(?:求人ID|求人番号|案件番号)[\s　]*[:：][\s　]*(.+)/);
      const dateLine = extract(body, /勤務日程[\s　]*[:：][\s　]*(.+)/);
      const timeLine = extract(body, /勤務時間[\s　]*[:：][\s　]*(.+)/);
      const clinic = extract(body, /^キャップスクリニック([^\r\n]+)/m).replace(/様$/, '').trim();

      if (doctorName && dateLine) {
        const added = appendRowAndFormatUnique(sheet, [
          '民間医局', formatDate(dateLine, receivedDate), timeLine, 
          clinic, department, 
          id, 
          doctorName, extract(body, /(https:\/\/portal\.doctor-agent\.com\/.+)/), 
          '', '応募通知受信（手動更新要）', receivedDate, '', ''
        ]);
        if (added) threadProcessed = true;
      }
    });

    if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
      thread.addLabel(processedLabel);
    }
  });
}

// === メディウェル ===
function processMediwel_Application(querySuffix, processedLabel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROGRESS);
  const processedLabelName = processedLabel.getName();
  
  GmailApp.search(`subject:("メディウェル") subject:("応募") ${querySuffix}`).forEach(thread => {
    if (thread.getLabels().some(l => l.getName() === processedLabelName)) return;

    const message = thread.getMessages()[0];
    const body = message.getPlainBody();
    const receivedDate = message.getDate();

    const doctorName = extract(body, /医師氏名[\s:：]+([^\n\r]+)/) || extract(body, /医師名[\s:：]+([^\n\r]+)/) || '不明';
    const dateLine = extract(body, /勤務日[\s:：]+([^\n\r]+)/) || extract(body, /勤務日程[\s:：]+([^\n\r]+)/) || Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd');
    const workTime = extract(dateLine, /(\d{1,2}:\d{2}\s*[〜~〜\-]\s*\d{1,2}:\d{2})/) || extract(body, /勤務時間[\s:：]+([^\n\r]+)/) || '';
    const clinic = extract(body, /医療機関名[\s:：]+([^\n\r]+)/) || extract(body, /施設名[\s:：]+([^\n\r]+)/) || '不明';
    const department = extract(body, /募集科目[\s:：]+([^\n\r]+)/) || extract(body, /診療科目[\s:：]+([^\n\r]+)/) || '';
    const id = extract(body, /(?:求人ID|案件ID|求人番号)[\s:：]+([a-zA-Z0-9\-]+)/) || '';

    if (doctorName && dateLine) {
      appendRowAndFormatUnique(sheet, [
        'メディウェル', formatDate(dateLine, receivedDate), workTime, 
        clinic, department, id, doctorName, '', '', '応募通知受信（手動更新要）', receivedDate, '', ''
      ]);
    }
    thread.addLabel(processedLabel);
  });
}