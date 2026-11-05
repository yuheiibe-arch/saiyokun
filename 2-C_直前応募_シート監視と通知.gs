// =================================================================
// 【2-C_直前応募_シート監視と通知 (Part 2 & Part 3)】
// ・Part 2: 「紹介会社進捗」「処理済み」シートから直前応募を検知・通知
// ・Part 3: 「通知履歴」シートから未送信の直前応募を通知し「直前通知」へ転記
// ★修正：449648673 への通知はキューに貯め、2-A側で拠点ごとにまとめて一括送信
// =================================================================

function part2_checkForAgencyUrgentApplications() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(AGENCY_SOURCE_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);
  const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  
  const archiveSheetName = typeof SHEET_NAME_ARCHIVE !== 'undefined' ? SHEET_NAME_ARCHIVE : '処理済み';
  const archiveSheet = spreadsheet.getSheetByName(archiveSheetName);

  const sourceUrgentColIdx = sourceSheet ? getColIndex_internal(sourceSheet, '直前応募通知') : -1;
  const archiveUrgentColIdx = archiveSheet ? getColIndex_internal(archiveSheet, '直前応募通知') : -1;

  let processedLogKeys = []; 
  if (logSheet && logSheet.getLastRow() > 1) {
    const logValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).getValues();
    logValues.forEach(row => {
      const status = row[6]; 
      const uniqueKey = row[5]; 
      if ((status === '投稿連携' || status === '投稿連携(夜間)' || status === '投稿済み' || String(status).includes('済')) && typeof uniqueKey === 'string' && uniqueKey.includes('_')) {
        const parts = uniqueKey.split('_');
        if (parts.length < 3) return; 

        const logWorkDateRaw = parts[0]; 
        const logClinic = parts[1] || '';
        const logDoctor = parts[2] || '';
        
        let logWorkDateNorm = logWorkDateRaw;
        if (logWorkDateRaw.includes('年')) {
          const match = logWorkDateRaw.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
          if (match) logWorkDateNorm = match[1];
        } else {
          try { logWorkDateNorm = Utilities.formatDate(new Date(logWorkDateRaw), 'JST', 'yyyy年M月d日'); } catch (e) {}
        }
        processedLogKeys.push(`${logWorkDateNorm}_${logClinic.replace(/\s+/g, '')}_${logDoctor.replace(/\s+/g, '')}`);
      }
    });
  }

  const statusColumnIndex = typeof AGENCY_STATUS_COLUMN_LETTER !== 'undefined' ? 
                            AGENCY_STATUS_COLUMN_LETTER.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1 : 13;

  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const todayStr = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'); 
  const tomorrow = new Date(today.getTime());
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');
  const dayAfterTomorrow = new Date(today.getTime());
  dayAfterTomorrow.setDate(today.getDate() + 2);
  const dayAfterTomorrowStr = Utilities.formatDate(dayAfterTomorrow, 'JST', 'yyyy/MM/dd');

  let combinedData = [];

  if (sourceSheet && sourceSheet.getLastRow() >= 2) {
    const sData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
    sData.forEach((row, i) => combinedData.push({ row: row, rowNumber: i + 2, sheet: sourceSheet, urgentColIdx: sourceUrgentColIdx }));
  }

  if (archiveSheet && archiveSheet.getLastRow() >= 2) {
    const aHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    
    let taiouColIndex = aHeaders.findIndex(h => String(h).includes('対応時間') || String(h).includes('処理日時'));
    if (taiouColIndex === -1 && typeof CHECKBOX_COLUMN !== 'undefined') {
      taiouColIndex = CHECKBOX_COLUMN - 1; 
    }

    const aData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archiveSheet.getLastColumn()).getValues();
    
    const todayMidnight = new Date();
    todayMidnight.setHours(0, 0, 0, 0);
    const yesterdayMidnight = new Date(todayMidnight.getTime() - 24 * 60 * 60 * 1000);

    aData.forEach((row, i) => {
      if (taiouColIndex !== -1 && row[taiouColIndex]) {
        const tDate = new Date(row[taiouColIndex]);
        if (!isNaN(tDate.getTime()) && tDate.getTime() >= yesterdayMidnight.getTime()) {
          combinedData.push({ row: row, rowNumber: i + 2, sheet: archiveSheet, urgentColIdx: archiveUrgentColIdx });
        }
      }
    });
  }

  const processedStatusStr = typeof AGENCY_STATUS_PROCESSED !== 'undefined' ? AGENCY_STATUS_PROCESSED : '処理済';
  const getJpDay = (d) => ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  
  for (let i = 0; i < combinedData.length; i++) {
    const item = combinedData[i];
    const row = item.row;
    const rowNumber = item.rowNumber;
    const currentSheet = item.sheet;
    const urgentColIdx = item.urgentColIdx;

    if (urgentColIdx > 0) {
      const urgentStatus = String(row[urgentColIdx - 1]).trim();
      if (urgentStatus !== '') continue; 
    }

    const agencyName = row[0];
    const workDateObj = new Date(row[1]); 
    let workTime = row[2] ? (isNaN(new Date(row[2]).getTime()) ? row[2] : Utilities.formatDate(new Date(row[2]), 'JST', 'HH:mm')) : '記載なし';
    const clinic = row[3] || '記載なし';
    const doctorName = row[6] || '記載なし';
    const receivedDate = row[10] ? new Date(row[10]) : new Date(); 
    const status = row[statusColumnIndex - 1];

    if (isNaN(workDateObj.getTime())) continue;

    // 着信日から24時間以上経過している古いデータは通知しないストッパー
    const nowTime = new Date().getTime();
    if ((nowTime - receivedDate.getTime()) > (24 * 60 * 60 * 1000)) {
      continue; 
    }

    const workDateStr = Utilities.formatDate(workDateObj, 'JST', 'yyyy/MM/dd'); 
    const hour = receivedDate.getHours();
    
    const isNightTime = (hour >= 19 || hour < 7);

    let isUrgent = false;
    let isNightUrgent = false;
    if (isNightTime && (workDateStr === tomorrowStr || workDateStr === dayAfterTomorrowStr)) {
      isUrgent = true;
      isNightUrgent = true;
    } else if (workDateStr === todayStr || workDateStr === tomorrowStr) {
      isUrgent = true;
    }

    if (isUrgent && status !== processedStatusStr && status !== 'Slack済' && status !== 'CW済') {
      const agencyWorkDateNorm = Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日');
      const agencyNormalizedKey = `${agencyWorkDateNorm}_${clinic.replace(/\s+/g, '')}_${doctorName.replace(/\s+/g, '')}`;
      
      if (processedLogKeys.includes(agencyNormalizedKey)) {
        try { currentSheet.getRange(rowNumber, statusColumnIndex).setValue(processedStatusStr); } catch(e) {}
        continue; 
      }

      const displayWorkDateJp = `${Utilities.formatDate(workDateObj, 'JST', 'yyyy年M月d日')} (${getJpDay(workDateObj)})`;

      // ① 既存ルーム（415381235）＆ Slack への送信
      let messageBody = isNightUrgent ? '直近の応募が夜間にありました。（紹介会社経由）\n確認お願いします。' : '紹介会社経由で直前応募がありました。\n確認お願いします。';
      const message = `[info][title]直前応募通知（紹介会社経由）[/title]\n${messageBody}\n\n紹介会社： ${agencyName}\n勤務日： ${displayWorkDateJp}\n応募時間： ${workTime}\n応募拠点： ${clinic}\n応募医師： ${doctorName}\n[/info]`;

      let slackSuccess = sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      // ② ★449648673 ルーム用のまとめ送信キューに追加
      if (cwSuccess) {
        queueUrgentApplyForExtraRoom_internal(clinic, displayWorkDateJp, workTime, doctorName, isNightUrgent);
      }

      if (slackSuccess || cwSuccess) {
        if (urgentColIdx > 0) {
          const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
          currentSheet.getRange(rowNumber, urgentColIdx).setValue(timestamp);
        }
      }

      if (slackSuccess && cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue(processedStatusStr);
        const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd');
        const uniqueKey = `${agencyName}_${workDateStr}_${doctorName}`;
        const newRowData = [receivedDateStr, clinic, displayWorkDateJp, doctorName, workTime, uniqueKey, '紹介会社経由', agencyName];
        if (destinationSheet) destinationSheet.appendRow(newRowData);
      } else if (slackSuccess && !cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue('Slack済');
      } else if (!slackSuccess && cwSuccess) {
        currentSheet.getRange(rowNumber, statusColumnIndex).setValue('CW済');
      }
    }
  }
}

function part3_postToChatworkFromSheet() {
  const STATUS_TO_POST = '投稿連携';
  const STATUS_TO_POST_NIGHT = '投稿連携(夜間)';
  const STATUS_POSTED = '投稿済み';
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  const destinationSheet = spreadsheet.getSheetByName(URGENT_SHEET_NAME);
  
  if (!sourceSheet) { return; } 
  if (destinationSheet && destinationSheet.getRange('A1').getValue() === '') {
    destinationSheet.getRange('A1:G1').setValues(sourceSheet.getRange('A1:G1').getValues());
  }
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) { return; }
  
  const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, 7);
  const values = dataRange.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const currentStatus = values[i][6];
    
    const isTarget = (currentStatus === STATUS_TO_POST || currentStatus === STATUS_TO_POST_NIGHT || String(currentStatus).includes('Slack済') || String(currentStatus).includes('CW済'));
    
    if (isTarget) {
      const rowNumber = i + 2;
      let isNight = String(currentStatus).includes('夜間');
      
      const clinic = values[i][1];
      const workDate = values[i][2];
      const doctorName = values[i][3];
      const workTime = values[i][4];

      // ① 既存ルーム（415381235）＆ Slack への送信
      const messageBody = isNight ? '直近の応募が夜間にありました。\n確認お願いします。' : '直前応募がありました。\n確認お願いします。';
      const message = `[info][title]直前応募通知[/title]${messageBody}\n\n応募拠点： ${clinic}\n勤務日： ${workDate}\n応募時間： ${workTime}\n応募医師： ${doctorName}[/info]`;

      let isSlackDone = String(currentStatus).includes('Slack済');
      let isCwDone = String(currentStatus).includes('CW済');

      let slackSuccess = isSlackDone ? true : sendUrgentAlertToSlack_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);
      let cwSuccess = isCwDone ? true : sendUrgentAlertToChatwork_internal(URGENT_APPLY_CHATWORK_ROOM_ID, message);

      // ② ★449648673 ルーム用のまとめ送信キューに追加（まだCW送信が完了していなかった場合のみ）
      if (!isCwDone && cwSuccess) {
        queueUrgentApplyForExtraRoom_internal(clinic, workDate, workTime, doctorName, isNight);
      }

      if (slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(STATUS_POSTED);
        values[i][6] = STATUS_POSTED;
        if (destinationSheet) destinationSheet.appendRow(values[i]); 
      } else if (slackSuccess && !cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'Slack済(夜間)' : 'Slack済');
      } else if (!slackSuccess && cwSuccess) {
        sourceSheet.getRange(rowNumber, 7).setValue(isNight ? 'CW済(夜間)' : 'CW済');
      }
    }
  }
}