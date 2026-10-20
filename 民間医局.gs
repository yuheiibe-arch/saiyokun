// =============================================================
// 【ファイル4：特定医師（民間医局）対応スクリプト（スレッド統合解析版）】
// ★修正済み：テキスト版で消滅する給与情報をHTMLから回収
// ★修正済み：Chatwork通知から (※要確認) を削除
// ★今回修正1：同日複数コマ（午前＋午後など）応募時の時間上書きバグを解消し、日付と時間を完全分離
// ★今回修正2：シート記録・通知書き出し時に「日付＋開始時間」の昇順（古い順）で自動ソート
// ★今回修正3：テンプレート・署名の法人名を「医療法人社団ナイズ」へ自動統一
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
 * 「2026年11月3日 (火)」と「09:00」からソート用の数値（タイムスタンプ）を計算するヘルパー関数
 */
function getShiftSortTimestamp_internal(workDateStr, startTimeStr) {
  const dMatch = String(workDateStr).match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/);
  if (!dMatch) return Number.MAX_SAFE_INTEGER;
  const year = parseInt(dMatch[1], 10);
  const month = parseInt(dMatch[2], 10) - 1;
  const day = parseInt(dMatch[3], 10);

  let hour = 0, min = 0;
  const tMatch = String(startTimeStr || '').match(/(\d{1,2}):(\d{2})/);
  if (tMatch) {
    hour = parseInt(tMatch[1], 10);
    min = parseInt(tMatch[2], 10);
  }
  return new Date(year, month, day, hour, min).getTime();
}

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
    .flat().filter(String).map(n => n.replace(/[\s ]+/g, ''));

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
    // ▼ STEP 1: スレッド予備調査 (同日複数コマも順番通りに配列で確保) ▼
    // =================================================================
    const threadTimeMap = {}; // Key: "2026年12月6日 (日)", Value: [{start:"09:00", end:"13:00"}, {start:"15:00", end:"18:00"}]

    const addTimeToMap = (key, start, end) => {
      if (!key) return;
      if (!threadTimeMap[key]) threadTimeMap[key] = [];
      const exists = threadTimeMap[key].some(t => t.start === start && t.end === end);
      if (!exists) {
        threadTimeMap[key].push({ start: start, end: end });
      }
    };

    const shiftRegex = /(\d{4}年\d{1,2}月\d{1,2}日\s*\([月火水木金土日]\))\s*(\d{1,2}:\d{2}\s*[～~〜\-]\s*\d{1,2}:\d{2})/g;

    messages.forEach(msg => {
      const b = msg.getPlainBody();
      let m;
      while ((m = shiftRegex.exec(b)) !== null) {
        const dateKey = m[1].replace(/\s+/g, ' ').trim(); // 例: 2026年12月6日 (日)
        const timeStr = m[2].trim(); // 例: 09:00 ～ 13:00
        
        const timeParts = timeStr.match(/(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/);
        if (timeParts) {
          addTimeToMap(dateKey, timeParts[1], timeParts[2]);
          const dateOnly = dateKey.split(' ')[0]; // 例: 2026年12月6日
          addTimeToMap(dateOnly, timeParts[1], timeParts[2]);
        }
      }
    });
    // =================================================================


    // ▼ STEP 2: 給与メール(事務局メール)を探して処理 ▼
    messages.forEach(message => {
      let body = message.getPlainBody();

      // HTML解析が必要な場合の処理
      if (!body.includes("掲載日給額") && !body.includes("日給")) {
         const html = message.getBody();
         body = html.replace(/<br\s*\/?>/gi, '\n')
                    .replace(/<\/p>|<\/div>|<\/tr>|<\/li>/gi, '\n')
                    .replace(/<[^>]+>/g, ' ')
                    .replace(/&nbsp;/g, ' ');
      }

      if (!body.includes("掲載日給額") && !body.includes("日給")) return;

      // 対象医師チェック
      let foundDoctorName = "";
      const cleanBody = body.replace(/[\s ]+/g, '');
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

      // 「シフト応募日：」ごとにブロックを切り出して解析（次の「シフト応募日：」の手前までを1ブロックとする）
      const dateRegex = /シフト応募日：\s*([^\n\r]+)/g;
      const matches = [];
      let m;
      while ((m = dateRegex.exec(body)) !== null) {
        matches.push({ rawLine: m[1].trim(), startIndex: dateRegex.lastIndex, matchIndex: m.index });
      }

      // 同じ日付が複数コマ出てきた際に、threadTimeMap の何番目を使うかを管理するカウンター
      const dateUsageIndexMap = {};

      for (let idx = 0; idx < matches.length; idx++) {
        const rawDateLine = matches[idx].rawLine; // 例: "2026年12月6日 (日) 09:00 ～ 13:00" または "2026年12月6日 (日)"
        const nextMatchIndex = (idx + 1 < matches.length) ? matches[idx + 1].matchIndex : body.length;
        const blockText = body.substring(matches[idx].startIndex, nextMatchIndex);

        // ★重要修正1：rawDateLine から「純粋な日付」と「行内の時間」を綺麗に分離する
        let workDate = rawDateLine;
        const cleanDateMatch = rawDateLine.match(/(\d{4}年\d{1,2}月\d{1,2}日\s*\([月火水木金土日]\))/) || rawDateLine.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
        if (cleanDateMatch) {
          workDate = cleanDateMatch[1].replace(/\s+/g, ' ').trim();
        }

        const inlineTimeMatch = rawDateLine.match(/(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/);

        const clinicMatch = blockText.match(/応募拠点名：\s*([^\n\r]+)/);
        const clinic = clinicMatch ? clinicMatch[1].trim() : "不明";

        // 給与抽出
        const salaryMatch = blockText.match(/(?:掲載日給額|日給)[^0-9]*([0-9,]+)/);
        let salaryRaw = 0;
        let salaryText = "要確認";
        if (salaryMatch) {
          salaryRaw = parseInt(salaryMatch[1].replace(/,/g, ''), 10);
          salaryText = salaryMatch[1] + "円";
        }

        // ▼▼▼【重要修正2】時間の決定ロジック（2コマ応募完全対応） ▼▼▼
        let startTime = "15:00"; 
        let endTime = "18:00"; 
        let timeSource = "default";

        const dateOnlyKey = workDate.split(' ')[0];
        const useIdx = dateUsageIndexMap[dateOnlyKey] || 0;
        dateUsageIndexMap[dateOnlyKey] = useIdx + 1;

        if (inlineTimeMatch) {
          // 1. 「シフト応募日：」の同じ行に時間が書かれている場合はそれを最優先で採用（一番確実）
          startTime = inlineTimeMatch[1];
          endTime = inlineTimeMatch[2];
          timeSource = "inline_date_line";
        } else {
          // 2. 行内に時間がない場合、ブロック内（次のシフト応募日までの間）に時間があるか探す
          const blockTimeMatch = blockText.match(/(\d{1,2}:\d{2})\s*[～~〜\-]\s*(\d{1,2}:\d{2})/);
          if (blockTimeMatch) {
            startTime = blockTimeMatch[1];
            endTime = blockTimeMatch[2];
            timeSource = "block_scan";
          } else {
            // 3. スレッド全体の時間マップ（配列）から、この日付の useIdx コマ目の時間を取得
            const timeList = threadTimeMap[workDate] || threadTimeMap[dateOnlyKey];
            if (timeList && timeList.length > 0) {
              const targetTimeObj = timeList[Math.min(useIdx, timeList.length - 1)];
              startTime = targetTimeObj.start;
              endTime = targetTimeObj.end;
              timeSource = "thread_scan";
            } else {
              console.warn(`[Time Warning] 時間特定不可。デフォルト適用。Date: ${workDate}`);
            }
          }
        }
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        // 所要時間の計算
        let durationStr = "3:00:00";
        try {
          const d = new Date();
          const y = d.getFullYear(); const m = d.getMonth(); const day = d.getDate();
          const startD = new Date(y, m, day, parseInt(startTime.split(':')[0], 10), parseInt(startTime.split(':')[1], 10));
          const endD = new Date(y, m, day, parseInt(endTime.split(':')[0], 10), parseInt(endTime.split(':')[1], 10));
          let diff = endD.getTime() - startD.getTime();
          if (diff < 0) diff += 24 * 60 * 60 * 1000;
          if (diff > 0) {
            const diffHours = Math.floor(diff / (1000 * 60 * 60));
            const diffMin = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
            durationStr = `${diffHours}:${('0' + diffMin).slice(-2)}:00`;
          }
        } catch(e) {}

        // ★重要修正3：同日の午前・午後2コマが別々に登録されるよう、ユニークキーに時間を含める
        // （過去に旧キーで登録済みの1コマ案件が再通知されないよう、旧キー群もあわせてチェック）
        const uniqueKey = `${workDate}_${startTime}-${endTime}_${clinic}_${displayDoctorName}_SALARY_OK`;
        const legacyKeyClean = `${workDate}_${clinic}_${displayDoctorName}_SALARY_OK`;
        const legacyKeyRaw = `${rawDateLine}_${clinic}_${displayDoctorName}_SALARY_OK`;

        if (cachedSpecialUniqueKeys.includes(uniqueKey) || cachedSpecialUniqueKeys.includes(legacyKeyRaw)) {
          continue;
        }
        // 1コマ目かつ旧クリーンキーが既に存在する場合もスキップ
        if (useIdx === 0 && cachedSpecialUniqueKeys.includes(legacyKeyClean)) {
          continue;
        }

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
          uniqueKey: uniqueKey,
          sortTime: getShiftSortTimestamp_internal(workDate, startTime)
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

  // ★追加：シート記録・Chatwork通知の前に、全体のバッファを「日付＋開始時間」の昇順（古い順）にソート
  specialCandidateBuffer.sort((a, b) => a.sortTime - b.sortTime);

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
  
  // Chatwork通知テンプレートの取得と法人名（「医療法人社団ナイズ」）への自動補正
  let templateRaw = String(settingSheet.getRange(SPECIAL_CONFIG.TEMPLATE_CELL).getValue() || '');
  const updatedTemplate = templateRaw
    .replace(/キャップス株式会社/g, '医療法人社団ナイズ')
    .replace(/医療社団法人ナイズ/g, '医療法人社団ナイズ');

  // もしシートのD2セルに古い社名が残っていたら、ついでにシート側も自動で書き換えておく💛
  if (templateRaw !== updatedTemplate) {
    try {
      settingSheet.getRange(SPECIAL_CONFIG.TEMPLATE_CELL).setValue(updatedTemplate);
    } catch (e) {}
    templateRaw = updatedTemplate;
  }

  const groupedData = {};
  
  specialCandidateBuffer.forEach(data => {
    if (!groupedData[data.doctorName]) groupedData[data.doctorName] = [];
    groupedData[data.doctorName].push(data);
  });

  for (const doctorName in groupedData) {
    const list = groupedData[doctorName];
    // 念のため医師ごとのリスト内でも日付・開始時間の昇順（古い順）を保証
    list.sort((a, b) => a.sortTime - b.sortTime);

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