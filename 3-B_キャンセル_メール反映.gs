// ============================================================================
// 【3-B_キャンセル_メール反映】（長大スレッドID保護 ＆ 省API版）
// ・Part 1: 民間医局・MRT キャンセルメールの自動反映
// ・Part 2: ドクターサイト「お問い合わせスタッフ送信（キャンセル申請承認）」自動反映
// ============================================================================

function processMinkanAndMrtCancels_internal_(ss, archiveSheet, progressSheet, colMap, props) {
  const TARGET_CW_ROOM_ID = '415529974';
  const LABEL_NAME = '処理済み-紹介会社キャンセル';

  let agencyProcessedIds = [];
  try {
    const stored = props.getProperty('AGENCY_CANCEL_PROCESSED_IDS');
    if (stored) agencyProcessedIds = JSON.parse(stored);
  } catch (e) {}
  const agencyProcessedSet = new Set(agencyProcessedIds);
  let isAgencyIdUpdated = false;

  const markAgencyId = (id) => {
    if (!agencyProcessedSet.has(id)) {
      agencyProcessedSet.add(id);
      agencyProcessedIds.push(id);
      isAgencyIdUpdated = true;
    }
  };

  const query = `(subject:"取消のご連絡" OR subject:"【キャンセル】") (from:medical-principle.co.jp OR from:medrt.com OR from:mnys.jp) newer_than:2d`;
  const threads = GmailApp.search(query, 0, 30);

  if (threads.length === 0) {
    console.log("対象の紹介会社キャンセルメール（民間医局・MRT）はありませんでした。");
    return;
  }

  let processedLabel = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!processedLabel) processedLabel = GmailApp.createLabel(LABEL_NAME);

  const now = new Date();
  const allMessages = GmailApp.getMessagesForThreads(threads);
  const processedJobIdsThisRun = new Set();

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadProcessed = false;

    for (const message of messages) {
      // ★72時間超の過去メールはIDリストに入れず即スキップ
      if ((now - message.getDate()) / (1000 * 60 * 60) > 72) continue;

      const msgId = message.getId();
      if (agencyProcessedSet.has(msgId)) continue;

      const subject = message.getSubject();
      const body = message.getPlainBody();

      let agency = '', jobId = '', clinic = '', doctorName = '', workDate = '', workTime = '';

      if (subject.includes('取消のご連絡') && subject.includes('民間医局')) {
        agency = '民間医局';
        const jobIdMatch = body.match(/案件番号[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
        const clinicMatch = body.match(/勤務先[\s ]*[:：][\s ]*([^\n\r]+)/);
        const docMatch = body.match(/ドクター名[\s ]*[:：][\s ]*([^\n\r]+)/);
        const dateMatch = body.match(/勤務日程[\s ]*[:：][\s ]*([^\n\r]+)/);
        const timeMatch = body.match(/勤務時間[\s ]*[:：][\s ]*([^\n\r]+)/);

        if (jobIdMatch) jobId = jobIdMatch[1].trim();
        if (clinicMatch) clinic = clinicMatch[1].replace(/キャップスクリニック/g, '').trim();
        if (docMatch) doctorName = docMatch[1].replace(/ドクター$|先生$/, '').trim();
        if (dateMatch) {
          const dm = dateMatch[1].match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
          workDate = dm ? dm[1] : dateMatch[1].trim();
        }
        if (timeMatch) workTime = timeMatch[1].trim();

      } else if (subject.includes('【キャンセル】') && subject.includes('MRT')) {
        agency = 'MRT';
        const jobIdMatch = body.match(/ID\s*[:：]\s*(\d+)/);
        const clinicMatch = body.match(/クリニック\s*[:：]\s*\d+\s*([^\n\r]+)/);
        const dateMatch = body.match(/日時\s*[:：]\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
        const timeMatch = body.match(/日時\s*[:：].*?\)\s*(\d{1,2}:\d{2}[〜~～\-]\d{1,2}:\d{2})/);
        const docMatch = body.match(/応募者\s*[:：]\s*([^\n\r]+?)先生/);

        if (jobIdMatch) jobId = jobIdMatch[1].trim();
        if (clinicMatch) clinic = clinicMatch[1].replace(/キャップスクリニック/g, '').trim();
        if (docMatch) doctorName = docMatch[1].trim();
        if (dateMatch) workDate = dateMatch[1].trim();
        if (timeMatch) workTime = timeMatch[1].trim();
      } else {
        markAgencyId(msgId);
        continue;
      }

      markAgencyId(msgId);

      if (!jobId) continue;
      if (processedJobIdsThisRun.has(jobId)) continue;
      processedJobIdsThisRun.add(jobId);

      let movedFromProgress = false;
      if (progressSheet && progressSheet.getLastRow() > 1) {
        const pData = progressSheet.getDataRange().getValues();
        for (let r = pData.length - 1; r >= 1; r--) {
          if (String(pData[r][5]).trim() === jobId) {
            const vals = pData[r].slice();
            if (doctorName && !vals[6]) vals[6] = doctorName;
            vals[8] = 'お断り';
            vals[9] = 'キャンセル済み';
            vals[11] = new Date();
            if (typeof appendRowToArchive === 'function') {
              appendRowToArchive(progressSheet, archiveSheet, vals, '不採用');
              if (typeof safeDeleteRow === 'function') safeDeleteRow(progressSheet, r + 1);
              movedFromProgress = true;
            }
          }
        }
      }

      const sheetData = archiveSheet.getDataRange().getValues();
      const matchedRows = [];

      for (let r = 1; r < sheetData.length; r++) {
        if (String(sheetData[r][colMap.colId - 1]).trim() === jobId) {
          matchedRows.push(r + 1);
        }
      }

      if (matchedRows.length === 0) {
        const normMailDate = normalizeCancelDateStr_(workDate);
        const cleanMailName = doctorName.replace(/[\s ]+/g, '');
        const normMailClinic = identifyCancelClinic_internal_(clinic).formal;

        for (let r = 1; r < sheetData.length; r++) {
          const sheetAgency = String(sheetData[r][colMap.colAgency - 1]).trim();
          const sheetDate = normalizeCancelDateStr_(sheetData[r][colMap.colDate - 1]);
          const sheetClinic = identifyCancelClinic_internal_(sheetData[r][colMap.colClinic - 1]).formal;
          const sheetName = String(sheetData[r][colMap.colDoctor - 1]).replace(/[\s ]+/g, '');

          if ((!sheetAgency || sheetAgency === agency) && sheetDate === normMailDate && sheetClinic === normMailClinic && sheetName === cleanMailName) {
            matchedRows.push(r + 1);
          }
        }
      }

      let targetRowNumber = 0;
      let isInserted = false;

      if (matchedRows.length > 0) {
        matchedRows.forEach(rowNum => {
          archiveSheet.getRange(rowNum, colMap.colType).setValue('お断り');
          archiveSheet.getRange(rowNum, colMap.colStatus).setValue('キャンセル済み');
          archiveSheet.getRange(rowNum, colMap.colSaiyo).setValue('不採用');
          archiveSheet.getRange(rowNum, 1, 1, colMap.totalCols).setBackground('#ffe599');
          targetRowNumber = rowNum;
        });
      } else if (!movedFromProgress) {
        isInserted = true;
        const newRow = new Array(colMap.totalCols).fill('');
        const nowStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
        const normDate = normalizeCancelDateStr_(workDate) || workDate;

        newRow[colMap.colAgency - 1] = agency;
        newRow[colMap.colDate - 1]   = normDate;
        newRow[colMap.colTime - 1]   = workTime;
        newRow[colMap.colClinic - 1] = clinic;
        newRow[colMap.colDept - 1]   = '小児科';
        newRow[colMap.colId - 1]     = jobId;
        newRow[colMap.colDoctor - 1] = doctorName;
        newRow[colMap.colType - 1]   = 'お断り';
        newRow[colMap.colStatus - 1] = 'キャンセル済み';
        newRow[colMap.colRecv - 1]   = nowStr;
        newRow[colMap.colDone - 1]   = nowStr;
        newRow[colMap.colPerson - 1] = '自動連携';
        newRow[colMap.colUKey - 1]   = jobId;
        newRow[colMap.colSaiyo - 1]  = '不採用';

        archiveSheet.appendRow(newRow);
        targetRowNumber = archiveSheet.getLastRow();
        archiveSheet.getRange(targetRowNumber, 1, 1, colMap.totalCols).setBackground('#ffe599');
      } else {
        targetRowNumber = archiveSheet.getLastRow();
      }

      const targetRowUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + archiveSheet.getSheetId() + '&range=A' + targetRowNumber;
      const currentHour = new Date().getHours();
      const mentionTag = (currentHour >= 7 && currentHour < 19) ? "[toall]\n" : "";

      const chatworkMessage = `${mentionTag}[info][title]紹介会社キャンセル自動反映[/title]
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

      sendToChatwork(TARGET_CW_ROOM_ID, chatworkMessage);
      threadProcessed = true;
      Utilities.sleep(1500);
    }

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }

  if (isAgencyIdUpdated) {
    if (agencyProcessedIds.length > 300) agencyProcessedIds = agencyProcessedIds.slice(-300);
    props.setProperty('AGENCY_CANCEL_PROCESSED_IDS', JSON.stringify(agencyProcessedIds));
  }
}

/**
 * Part 2: ドクターサイトの「お問い合わせスタッフ送信」内のキャンセル申請承認メール自動反映
 */
function processSystemApprovedCancels_internal_(archiveSheet, progressSheet, colMap, props, newerThanClause) {
  let sysProcessedIds = [];
  try {
    const stored = props.getProperty('SYS_CANCEL_PROCESSED_IDS');
    if (stored) sysProcessedIds = JSON.parse(stored);
  } catch (e) {}
  const sysProcessedSet = new Set(sysProcessedIds);
  let isSysIdUpdated = false;

  const sysQuery = `subject:"お問い合わせスタッフ送信" "のキャンセル申請が承認されました" ${newerThanClause}`;
  const sysThreads = GmailApp.search(sysQuery, 0, 50);
  if (sysThreads.length === 0) return 0;

  const now = new Date();
  const cancelEvents = [];
  const allMsgs = GmailApp.getMessagesForThreads(sysThreads);

  for (let i = 0; i < sysThreads.length; i++) {
    for (const msg of allMsgs[i]) {
      // ★超重要：1スレッド100通の「お問い合わせスタッフ送信」で72時間超の古いメールをID配列に入れない！
      if ((now - msg.getDate()) / (1000 * 60 * 60) > 72) continue;

      const msgId = msg.getId();
      if (sysProcessedSet.has(msgId)) continue;
      sysProcessedSet.add(msgId);
      sysProcessedIds.push(msgId);
      isSysIdUpdated = true;

      const body = msg.getPlainBody();
      if (!body.includes('のキャンセル申請が承認されました')) continue;

      const parsed = parseSystemCancelApprovalBody_(body);
      if (parsed) cancelEvents.push(parsed);
    }
  }

  let updatedCount = 0;
  if (cancelEvents.length > 0) {
    updatedCount = applySystemCancelEventsToSheets_(archiveSheet, progressSheet, colMap, cancelEvents);
  }

  if (isSysIdUpdated) {
    if (sysProcessedIds.length > 350) sysProcessedIds = sysProcessedIds.slice(-350);
    props.setProperty('SYS_CANCEL_PROCESSED_IDS', JSON.stringify(sysProcessedIds));
  }
  return updatedCount;
}

function parseSystemCancelApprovalBody_(body) {
  const docMatch = body.match(/【医師名】[\s ]*[:：][\s ]*([^\r\n]+)/);
  const contentMatch = body.match(/(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}日?)[\s ]+([^\s ]+)[\s ]+(小児科|内科|[^\s \d]+)?[\s ]*(\d{1,2}:\d{2}\s*[~〜～\-]\s*\d{1,2}:\d{2})?[^\r\n]*?のキャンセル申請が承認されました/);
  if (!docMatch || !contentMatch) return null;

  const docClean = docMatch[1].replace(/先生$/, '').replace(/[\s ]+/g, '');
  const dateStr = normalizeCancelDateStr_(contentMatch[1]);
  const cInfo = identifyCancelClinic_internal_(contentMatch[2]);
  const deptStr = String(contentMatch[3] || '').trim();
  const timeRange = parseCancelTimeRange_(contentMatch[4] || '');

  if (!docClean || !dateStr || !cInfo.formal) return null;
  return { dateStr, normClinic: cInfo.formal, deptStr, timeRange, docClean };
}

function applySystemCancelEventsToSheets_(archiveSheet, progressSheet, colMap, cancelEvents) {
  let updatedCount = 0;

  if (progressSheet && progressSheet.getLastRow() > 1) {
    const pData = progressSheet.getDataRange().getValues();
    for (let r = pData.length - 1; r >= 1; r--) {
      const rDate = normalizeCancelDateStr_(pData[r][1]);
      const rDoc = String(pData[r][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');
      const rClinicInfo = identifyCancelClinic_internal_(pData[r][3]);
      const rDept = String(pData[r][4] || '').includes('内科') ? '内科' : '小児科';
      const rTime = parseCancelTimeRange_(pData[r][2]);

      const hit = cancelEvents.find(ev => {
        if (ev.dateStr !== rDate || ev.docClean !== rDoc || ev.normClinic !== rClinicInfo.formal) return false;
        if ((ev.normClinic.includes('亀有') || ev.normClinic.includes('北葛西')) && ev.deptStr) {
          const evDept = ev.deptStr.includes('内科') ? '内科' : '小児科';
          if (evDept !== rDept) return false;
        }
        if (ev.timeRange && rTime) {
          return (rTime.startMin < ev.timeRange.endMin && rTime.endMin > ev.timeRange.startMin);
        }
        return true;
      });

      if (hit && typeof appendRowToArchive === 'function') {
        const vals = pData[r].slice();
        vals[8] = 'お断り';
        vals[9] = 'キャンセル承認済(システム)';
        vals[11] = new Date();
        appendRowToArchive(progressSheet, archiveSheet, vals, '不採用');
        if (typeof safeDeleteRow === 'function') safeDeleteRow(progressSheet, r + 1);
        updatedCount++;
      }
    }
  }

  if (archiveSheet.getLastRow() > 1) {
    const aData = archiveSheet.getDataRange().getValues();
    for (let r = 1; r < aData.length; r++) {
      const curSaiyo = String(aData[r][colMap.colSaiyo - 1] || '').trim();
      if (curSaiyo !== '採用') continue;

      const rDate = normalizeCancelDateStr_(aData[r][colMap.colDate - 1]);
      const rDoc = String(aData[r][colMap.colDoctor - 1] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');
      const rRawClinic = String(aData[r][colMap.colClinic - 1] || '');
      const rClinicInfo = identifyCancelClinic_internal_(rRawClinic);
      const rRawDept = String(aData[r][colMap.colDept - 1] || '');
      const rDept = (rRawDept.includes('内科') || rRawClinic.includes('内科')) ? '内科' : '小児科';
      const rTime = parseCancelTimeRange_(aData[r][colMap.colTime - 1]);

      const hit = cancelEvents.find(ev => {
        if (ev.dateStr !== rDate || ev.docClean !== rDoc || ev.normClinic !== rClinicInfo.formal) return false;
        if ((ev.normClinic.includes('亀有') || ev.normClinic.includes('北葛西')) && ev.deptStr) {
          const evDept = ev.deptStr.includes('内科') ? '内科' : '小児科';
          if (evDept !== rDept) return false;
        }
        if (ev.timeRange && rTime) {
          return (rTime.startMin < ev.timeRange.endMin && rTime.endMin > ev.timeRange.startMin);
        }
        return true;
      });

      if (hit) {
        const rowNum = r + 1;
        archiveSheet.getRange(rowNum, colMap.colType).setValue('お断り');
        archiveSheet.getRange(rowNum, colMap.colStatus).setValue('キャンセル承認済(システム)');
        archiveSheet.getRange(rowNum, colMap.colSaiyo).setValue('不採用');
        archiveSheet.getRange(rowNum, 1, 1, colMap.totalCols).setBackground('#ffe599');
        console.log(`[システム承認キャンセル反映] 行${rowNum}: ${rDate} ${rClinicInfo.formal} ${rDoc} -> 不採用`);
        updatedCount++;
      }
    }
  }

  return updatedCount;
}