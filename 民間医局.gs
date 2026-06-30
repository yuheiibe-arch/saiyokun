// =============================================================
// 【ファイル4：特定医師（民間医局）対応スクリプト（スレッド統合解析版）】
// ★修正済み：テキスト版で消滅する給与情報をHTMLから回収
// ★修正済み：Chatwork通知から (※要確認) を削除
// ★今回修正：スレッド内の「時間情報」と「給与情報」を統合し、正確な時間を取得（15:00固定廃止）
// =============================================================

const SPECIAL_CONFIG = {
  SETTING_SHEET_NAME: '民間医局',
  LOG_SHEET_NAME: '民間医局くん',
  CHATWORK_ROOM_ID: '417015705',
  TEMPLATE_CELL: 'D2'
};

let specialCandidateBuffer = [];
let cachedSpecialUniqueKeys = null;

// ★エントリーポイント
function accumulateSpecialCandidate(ignoredEntryData, ignoredBody) {
  // 1回の実行で何度もスキャンしないようガード
  if (globalScannerHasRun) return; 
  globalScannerHasRun = true;

  scanForSalaryEmails();
}

let globalScannerHasRun = false;

/**
 * 「掲載日給額」が含まれるスレッドを検索し、
 * スレッド全体から「時間」を、事務局メールから「給与」を取得して合体させる
 */
function scanForSalaryEmails() {
  console.log("=== 民間医局：スレッド統合解析モード開始 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 設定・ログ読み込み
  const settingSheet = ss.getSheetByName(SPECIAL_CONFIG.SETTING_SHEET_NAME);
  if (!settingSheet) return;
  const targetDoctors = settingSheet.getRange(2, 1, settingSheet.getLastRow() - 1, 1).getValues()
    .flat().filter(String).map(n => n.replace(/[\s　]+/g, ''));

  if (cachedSpecialUniqueKeys === null) {
    const logSheet = ss.getSheetByName(SPECIAL_CONFIG.LOG_SHEET_NAME);
    if (logSheet && logSheet.getLastRow() > 1) {
      cachedSpecialUniqueKeys = logSheet.getRange(2, 13, logSheet.getLastRow() - 1, 1).getValues().flat();
    } else {
      cachedSpecialUniqueKeys = [];
    }
  }

  // 2. 検索実行（直近2日間 × "掲載日給額"）
  const now = new Date();
  const yesterday = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000); 
  const afterStr = Utilities.formatDate(yesterday, 'JST', 'yyyy/MM/dd');
  
  const query = `after:${afterStr} "掲載日給額"`;
  const threads = GmailApp.search(query);
  console.log(`ヒットしたスレッド数: ${threads.length}`);

  threads.forEach(thread => {
    const messages = thread.getMessages();

    // =================================================================
    // ▼ STEP 1: スレッド予備調査 (時間を確保して地図を作る) ▼
    // =================================================================
    // スレッド内の全メールを見て、「日付」に対する「時間」のマッピングを作る
    const threadTimeMap = {}; // Key: "2026年2月15日 (日)", Value: {start:"09:00", end:"13:00"}

    // 時間抽出用の正規表現 (日付 + 時間)
    // ハイフン(-)とチルダ(～)の両方に対応
    const shiftRegex = /(\d{4}年\d{1,2}月\d{1,2}日 \([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})/g;

    messages.forEach(msg => {
      const b = msg.getPlainBody();
      let m;
      // 文中のすべての「日付 時間～時間」パターンを拾う
      while ((m = shiftRegex.exec(b)) !== null) {
        const dateKey = m[1].trim(); // 例: 2026年2月15日 (日)
        const timeStr = m[2].trim(); // 例: 09:00 ～ 13:00
        
        // 時間を分割
        const timeParts = timeStr.match(/(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/);
        if (timeParts) {
           threadTimeMap[dateKey] = {
             start: timeParts[1],
             end: timeParts[2]
           };
           // 念のため、曜日なしの日付もキーとして登録しておく（表記ゆれ対策）
           const dateOnly = dateKey.split(' ')[0]; // 2026年2月15日
           threadTimeMap[dateOnly] = { start: timeParts[1], end: timeParts[2] };
        }
      }
    });
    // =================================================================


    // ▼ STEP 2: 給与メール(事務局メール)を探して処理 ▼
    messages.forEach(message => {
      let body = message.getPlainBody();
      let isHtmlMode = false;

      // HTML解析が必要な場合の処理
      if (!body.includes("掲載日給額") && !body.includes("日給")) {
         const html = message.getBody();
         // HTMLタグを除去してテキスト化
         body = html.replace(/<br\s*\/?>/gi, '\n')
                    .replace(/<[^>]+>/g, ' ')
                    .replace(/&nbsp;/g, ' ');
         isHtmlMode = true;
      }

      if (!body.includes("掲載日給額") && !body.includes("日給")) return;

      // 対象医師チェック
      let foundDoctorName = "";
      const cleanBody = body.replace(/[\s　]+/g, '');
      for (const doctor of targetDoctors) {
        if (cleanBody.includes(doctor)) {
          foundDoctorName = doctor;
          break;
        }
      }
      if (!foundDoctorName) return;

      // 医師名抽出
      let displayDoctorName = foundDoctorName;
      const nameMatch = body.match(/勤務医師名：\s*(.*?先生)/);
      if (nameMatch) {
        displayDoctorName = nameMatch[1].replace('先生', '').trim();
      }

      // 日付ごとのループ（給与情報がある箇所を特定）
      const dateRegex = /シフト応募日：\s*(.+)/g;
      let dateMatch;
      
      while ((dateMatch = dateRegex.exec(body)) !== null) {
        const workDate = dateMatch[1].trim(); // 例: 2026年2月15日 (日)
        const remainingText = body.substring(dateRegex.lastIndex);
        
        const clinicMatch = remainingText.match(/応募拠点名：\s*(.+)/);
        const clinic = clinicMatch ? clinicMatch[1].trim() : "不明";

        // 給与抽出
        const salaryMatch = remainingText.match(/(?:掲載日給額|日給)[^0-9]*([0-9,]+)/);
        let salaryRaw = 0;
        let salaryText = "要確認";
        if (salaryMatch) {
          salaryRaw = parseInt(salaryMatch[1].replace(/,/g, ''), 10);
          salaryText = salaryMatch[1] + "円";
        }

        // ▼▼▼【重要】時間の決定ロジック ▼▼▼
        let startTime = "15:00"; 
        let endTime = "18:00"; 
        let timeSource = "default";

        // 1. まず、STEP 1で作った「スレッド全体の時間マップ」に、この日付があるか確認
        // (キー完全一致 または 日付部分一致)
        let foundTime = threadTimeMap[workDate] || threadTimeMap[workDate.split(' ')[0]];

        if (foundTime) {
            startTime = foundTime.start;
            endTime = foundTime.end;
            timeSource = "thread_scan"; // 成功：スレッド内の別メールから補完
        } else {
            // 2. マップになければ、このメール本文内を再度探す（バックアップシフト用など）
            const localTimeRegex = /(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/;
            const m = remainingText.match(localTimeRegex); 
            if (m) {
                startTime = m[1];
                endTime = m[2];
                timeSource = "local_scan";
            } else {
                console.warn(`[Time Warning] 時間特定不可。デフォルト適用。Date: ${workDate}`);
            }
        }
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        // 所要時間の計算
        let durationStr = "3:00:00";
        try {
            const d = new Date();
            const y = d.getFullYear(); const m = d.getMonth(); const day = d.getDate();
            const startD = new Date(y, m, day, parseInt(startTime.split(':')[0]), parseInt(startTime.split(':')[1]));
            const endD = new Date(y, m, day, parseInt(endTime.split(':')[0]), parseInt(endTime.split(':')[1]));
            let diff = endD.getTime() - startD.getTime();
            if (diff < 0) diff += 24 * 60 * 60 * 1000;
            if (diff > 0) {
                 const diffHours = Math.floor(diff / (1000 * 60 * 60));
                 const diffMin = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
                 durationStr = `${diffHours}:${('0' + diffMin).slice(-2)}:00`;
            }
        } catch(e) {}

        // ユニークキー
        const uniqueKey = `${workDate}_${clinic}_${displayDoctorName}_SALARY_OK`;

        if (cachedSpecialUniqueKeys.includes(uniqueKey)) continue;

        // 登録
        const commission = Math.floor(salaryRaw * 0.2); 
        let monthStr = "";
        const mMatch = workDate.match(/(\d{4})年(\d{1,2})月/);
        if (mMatch) monthStr = `${mMatch[1]}/${('00' + mMatch[2]).slice(-2)}`;

        console.log(`[登録成功] ${displayDoctorName} / ${workDate} / ${salaryText} / ${startTime}-${endTime} (src:${timeSource})`);

        specialCandidateBuffer.push({
          doctorName: displayDoctorName,
          workDate: workDate,
          month: monthStr,
          startTime: startTime,
          endTime: endTime,
          clinic: clinic,
          salary: salaryRaw,
          salaryText: salaryText,
          commission: commission,
          workType: "スポット",
          duration: durationStr,
          department: "小児科",
          uniqueKey: uniqueKey
        });

        if (salaryRaw > 0) {
           cachedSpecialUniqueKeys.push(uniqueKey);
        }
      }
    });
  });
}

function flushSpecialCandidateNotifications() {
  if (specialCandidateBuffer.length === 0) {
    console.log("民間医局くん: 新規の通知対象なし");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName(SPECIAL_CONFIG.SETTING_SHEET_NAME);
  const logSheet = ss.getSheetByName(SPECIAL_CONFIG.LOG_SHEET_NAME);

  if (!logSheet) return;

  const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
  const newRows = [];

  specialCandidateBuffer.forEach(item => {
    newRows.push([
      item.doctorName, item.workDate, item.month, item.startTime, item.endTime,
      item.clinic, item.salary, item.commission, item.workType, item.duration,
      todayStr, false, item.uniqueKey
    ]);
  });

  const startRow = logSheet.getLastRow() + 1;
  if (newRows.length > 0) {
    const range = logSheet.getRange(startRow, 1, newRows.length, newRows[0].length);
    range.setValues(newRows);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['スポット'], true).build();
    logSheet.getRange(startRow, 9, newRows.length, 1).setDataValidation(rule);
    logSheet.getRange(startRow, 12, newRows.length, 1).insertCheckboxes();
  }
  
  // Chatwork通知
  const templateRaw = settingSheet.getRange(SPECIAL_CONFIG.TEMPLATE_CELL).getValue();
  const groupedData = {};
  
  specialCandidateBuffer.forEach(data => {
    if (!groupedData[data.doctorName]) groupedData[data.doctorName] = [];
    groupedData[data.doctorName].push(data);
  });

  for (const doctorName in groupedData) {
    const list = groupedData[doctorName];
    let shiftListText = "";
    list.forEach(item => {
      shiftListText += `----------------------\n`;
      shiftListText += `${item.workDate}\n`;
      shiftListText += `勤務時間：${item.startTime}-${item.endTime}\n`; 
      shiftListText += `【${item.clinic}】\n`;
      shiftListText += `給与：${item.salaryText}\n`;
      shiftListText += `診療科目：${item.department}\n`;
      shiftListText += `医師名：${item.doctorName}ドクター\n`;
      shiftListText += `----------------------\n`;
    });

    let message = "[toall]\n" + templateRaw
      .replace(/{{DOCTOR_NAME}}/g, doctorName)
      .replace('{{SHIFT_LIST}}', shiftListText);

    sendToChatwork(SPECIAL_CONFIG.CHATWORK_ROOM_ID, message);
    console.log(`通知完了: ${doctorName}`);
  }
  specialCandidateBuffer = [];
}