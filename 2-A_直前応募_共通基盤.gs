// =================================================================
// 【2-A_直前応募_共通基盤】
// 大司令塔からのエントリーポイント、拠点メンション取得、通知送信ヘルパー
// ★修正1：夜間（19:00〜翌07:00）は拠点長をメンションせず「名前記載のみ」に配慮
// ★修正2：449648673 ルームへの直前応募通知を「拠点ごとに1投稿にまとめる」バッファ機構を追加
// =================================================================

const EXTRA_URGENT_AND_CANCEL_CW_ROOM_ID = '449648673';
const CLINIC_MENTION_MASTER_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
const CLINIC_MENTION_MASTER_SHEET_NAME = '拠点名';

// 449648673 ルームへ拠点ごとにまとめて送信するための一時バッファ
let extraRoomUrgentApplyBuffer = [];

/**
 * ★★★ 大司令塔 (2/9) から呼び出されるメイン関数 ★★★
 */
function executeAllProcesses_internal() {
  console.log("=============== 応募関連の自動化プロセスを開始 ===============");
  extraRoomUrgentApplyBuffer = []; // バッファ初期化

  part1_processEmailsToSheet();              // 2-B で定義
  part2_checkForAgencyUrgentApplications();  // 2-C で定義
  part3_postToChatworkFromSheet();           // 2-C で定義

  // ★ 貯まった直前応募を「拠点ごと」にまとめて 449648673 へ一括送信
  flushExtraRoomUrgentApplyNotifications_internal();

  console.log("\n=============== 応募関連の自動化プロセスが完了 ===============");
}

/**
 * 「拠点名」シートから【正規記載・表記揺れ・共用アカウント・拠点長ID・拠点長表示名】を読み込む（キャッシュ付き）
 */
function getClinicAndManagerMentionList_internal() {
  if (getClinicAndManagerMentionList_internal.cache) {
    return getClinicAndManagerMentionList_internal.cache;
  }

  try {
    const masterSs = SpreadsheetApp.openById(CLINIC_MENTION_MASTER_SS_ID);
    const masterSheet = masterSs.getSheetByName(CLINIC_MENTION_MASTER_SHEET_NAME) || masterSs.getSheets()[0];
    if (!masterSheet) return [];

    const data = masterSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headerRow = data[0].map(v => String(v).trim());
    const sharedAccCol = headerRow.indexOf('共用アカウント') !== -1 ? headerRow.indexOf('共用アカウント') : 9;
    const mgrNameCol = headerRow.indexOf('拠点長氏名') !== -1 ? headerRow.indexOf('拠点長氏名') : 12;
    const mgrIdCol = headerRow.indexOf('拠点長ID') !== -1 ? headerRow.indexOf('拠点長ID') : 13;

    // 日中用のメンションタグ（例: "[To:10266450]関田さん"）と、夜間用の名前のみ（例: "関田さん"）の両方を生成
    const parseManagerInfo = (rawId, rawName) => {
      const idStr = String(rawId || '').trim();
      const nameStr = String(rawName || '').trim();
      const surname = nameStr ? nameStr.split(/[\s ]+/)[0] : '';

      let fullTag = '';
      let nameOnly = '';

      if (idStr.includes('[To:')) {
        const afterBracket = idStr.replace(/^\[To:\d+\]\s*/, '').trim();
        if (afterBracket) {
          nameOnly = afterBracket.endsWith('さん') ? afterBracket : `${afterBracket}さん`;
          fullTag = idStr;
        } else if (surname) {
          nameOnly = surname.endsWith('さん') ? surname : `${surname}さん`;
          fullTag = `${idStr}${nameOnly}`;
        } else {
          fullTag = idStr;
        }
      } else {
        const cleanId = idStr.replace(/[^\d]/g, '');
        if (surname) {
          nameOnly = surname.endsWith('さん') ? surname : `${surname}さん`;
        }
        if (cleanId) {
          fullTag = `[To:${cleanId}]${nameOnly}`;
        }
      }

      if (!nameOnly && surname) {
        nameOnly = surname.endsWith('さん') ? surname : `${surname}さん`;
      }

      return { mgrTag: fullTag, mgrNameOnly: nameOnly };
    };

    const list = [];
    for (let i = 1; i < data.length; i++) {
      const aliases = [];
      for (let c = 0; c <= 4; c++) {
        const val = String(data[i][c] || '').replace(/キャップスクリニック/g, '').trim();
        if (val && !aliases.includes(val)) aliases.push(val);
      }
      if (aliases.length === 0) continue;

      const sharedTag = String(data[i][sharedAccCol] || '').trim();
      const mgrName = String(data[i][mgrNameCol] || '').trim();
      const mgrId = String(data[i][mgrIdCol] || '').trim();
      const { mgrTag, mgrNameOnly } = parseManagerInfo(mgrId, mgrName);

      list.push({
        canonicalName: aliases[0],
        aliases: aliases,
        sharedTag: sharedTag,
        mgrTag: mgrTag,
        mgrNameOnly: mgrNameOnly
      });
    }

    getClinicAndManagerMentionList_internal.cache = list;
    return list;
  } catch (e) {
    console.error('拠点メンションマスタの読み込みエラー: ' + e.message);
    return [];
  }
}

/**
 * 拠点名からマスタの該当拠点オブジェクトを検索して返す内部ヘルパー
 */
function findClinicMasterItem_internal(rawClinicName) {
  if (!rawClinicName) return null;
  const str = String(rawClinicName);
  if (str.includes('バックアップ')) return null;

  const cleanInput = str
    .replace(/キャップスクリニック/g, '')
    .replace(/[（(].*?[）)]/g, '')
    .replace(/(小児科|内科|皮膚科|耳鼻咽喉科)/g, '')
    .replace(/[\s ]+/g, '')
    .trim();

  if (!cleanInput) return null;

  const clinicMentionList = getClinicAndManagerMentionList_internal();

  // 1. 完全一致
  for (const item of clinicMentionList) {
    if (item.aliases.some(alias => alias.replace(/[\s ]+/g, '') === cleanInput)) {
      return item;
    }
  }

  // 2. 部分一致（文字数が長い表記揺れから優先）
  let matchedItem = null;
  let bestLen = 0;
  for (const item of clinicMentionList) {
    for (const alias of item.aliases) {
      const normAlias = alias.replace(/[\s ]+/g, '');
      if (normAlias.length >= 2 && (cleanInput.includes(normAlias) || normAlias.includes(cleanInput))) {
        if (normAlias.length > bestLen) {
          bestLen = normAlias.length;
          matchedItem = item;
        }
      }
    }
  }

  return matchedItem;
}

/**
 * 拠点名から「共用アカウント＋拠点長」のメンション文字列（改行付き）を取得する
 * ★修正：夜間（19:00〜翌06:59 または forceNightMode=true）は、拠点長をメンションせず「名前記載のみ」にする
 */
function getClinicAndManagerMentionText_internal(rawClinicName, forceNightMode) {
  const matchedItem = findClinicMasterItem_internal(rawClinicName);
  if (!matchedItem) return '';

  const currentHour = parseInt(Utilities.formatDate(new Date(), 'JST', 'H'), 10);
  const isNightNow = (currentHour >= 19 || currentHour < 7);
  const suppressManagerMention = Boolean(forceNightMode || isNightNow);

  const lines = [];
  if (matchedItem.sharedTag) {
    lines.push(matchedItem.sharedTag);
  }

  // 夜間は名前のみ（例: "関田さん"）、日中はメンションタグ（例: "[To:10266450]関田さん"）
  const mgrLine = suppressManagerMention ? matchedItem.mgrNameOnly : matchedItem.mgrTag;
  if (mgrLine && !lines.includes(mgrLine)) {
    lines.push(mgrLine);
  }

  return lines.length > 0 ? lines.join('\n') + '\n' : '';
}

/**
 * 449648673 用の直前応募データをバッファに追加する関数（2-C から呼び出される）
 */
function queueUrgentApplyForExtraRoom_internal(clinic, workDateStr, workTime, doctorName, isNight) {
  if (!extraRoomUrgentApplyBuffer) extraRoomUrgentApplyBuffer = [];

  const isBackup = String(clinic || '').includes('バックアップ');
  const matchedItem = isBackup ? null : findClinicMasterItem_internal(clinic);
  // 同じ拠点（例：「柏の葉」と「柏の葉（小児科）」）が綺麗に1つにまとまるよう、正規拠点名をグループキーにする
  const groupKey = isBackup
    ? `BACKUP_${String(clinic).trim()}`
    : (matchedItem ? matchedItem.canonicalName : String(clinic).trim());

  extraRoomUrgentApplyBuffer.push({
    groupKey: groupKey,
    clinic: clinic,
    workDateStr: workDateStr,
    workTime: workTime,
    doctorName: doctorName,
    isNight: Boolean(isNight),
    isBackup: isBackup
  });
}

/**
 * バッファに貯まった直前応募データを「拠点ごと」に1つの投稿にまとめて 449648673 へ送信する関数
 */
function flushExtraRoomUrgentApplyNotifications_internal() {
  if (!extraRoomUrgentApplyBuffer || extraRoomUrgentApplyBuffer.length === 0) {
    return;
  }

  const currentHour = parseInt(Utilities.formatDate(new Date(), 'JST', 'H'), 10);
  const isNightNow = (currentHour >= 19 || currentHour < 7);

  // 拠点キー（groupKey）ごとにグルーピング
  const grouped = {};
  for (const item of extraRoomUrgentApplyBuffer) {
    if (!grouped[item.groupKey]) {
      grouped[item.groupKey] = [];
    }
    grouped[item.groupKey].push(item);
  }

  for (const key in grouped) {
    const items = grouped[key];
    const first = items[0];
    const isBackup = first.isBackup;

    // グループ内に1件でも夜間フラグがあるか、現在時刻が夜間（19:00〜翌06:59）なら夜間モードとする
    const isNightGroup = isNightNow || items.some(x => x.isNight);

    // メンション行の取得（バックアップは無し、夜間は拠点長を名前記載のみに）
    const mentionLines = isBackup ? '' : getClinicAndManagerMentionText_internal(first.clinic, isNightGroup);
    const leadText = isNightGroup
      ? '直近の応募が夜間にありました。\n確認お願いします。'
      : '直前応募がありました。\n確認お願いします。';

    // 各応募明細ブロックを組み立て（複数件ある場合は区切り線で連結）
    const entryBlocks = items.map(entry => {
      return [
        `応募拠点： ${entry.clinic}`,
        `勤務日： ${entry.workDateStr}`,
        `応募時間： ${entry.workTime}`,
        `応募医師： ${entry.doctorName}`
      ].join('\n');
    }).join('\n----------------------\n');

    const importantNoteBlock = isBackup ? '' : `\n＜重要＞
こちらの通知は速報通知になります。
採用の場合は、改めて確定連絡をいたします。
２診目の場合は、拠点移動を打診の可能性があり、必ずしも採用とは限りませんので、ご理解のほどお願いいたします。
医師不在の枠への応募の場合は、原則採用いたします。`;

    const extraMessage = `${mentionLines}[info][title]
直前応募通知
[/title]
${leadText}

${entryBlocks}${importantNoteBlock}
[/info]`;

    sendUrgentAlertToChatwork_internal(EXTRA_URGENT_AND_CANCEL_CW_ROOM_ID, extraMessage);
    Utilities.sleep(1500); // API制限回避
  }

  extraRoomUrgentApplyBuffer = [];
}

/**
 * （互換維持用）単発呼び出しされた場合もバッファ経由で即時送信するラッパー
 */
function sendUrgentApplyToExtraRoom_internal(clinic, workDateStr, workTime, doctorName, isNight) {
  queueUrgentApplyForExtraRoom_internal(clinic, workDateStr, workTime, doctorName, isNight);
  flushExtraRoomUrgentApplyNotifications_internal();
  return true;
}

/**
 * 直前応募のSlack送信ヘルパー関数
 */
function sendUrgentAlertToSlack_internal(roomId, message) {
  try {
    const slackPayload = buildSlackPayload(message, roomId); 
    if (slackPayload && slackPayload.channelId) {
      const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify({ "channel": slackPayload.channelId, "text": slackPayload.text }),
        "muteHttpExceptions": true
      };
      const res = UrlFetchApp.fetch(webhookUrl, options);
      return (res.getResponseCode() === 200 || res.getResponseCode() === 201);
    }
  } catch (e) {
    console.error("Slack送信エラー: " + e.message);
  }
  return false;
}

/**
 * 直前応募のChatwork送信ヘルパー関数
 */
function sendUrgentAlertToChatwork_internal(roomId, message) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey) return false;
  const url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error("Chatwork 通信エラー: " + e.message);
    return false;
  }
}