// =================================================================
// 【2-A_直前応募_共通基盤】
// 大司令塔からのエントリーポイント、拠点メンション取得、通知送信ヘルパー
// ※ 緊急キャンセル.gs からもここのメンション取得関数を共通利用します
// =================================================================

const EXTRA_URGENT_AND_CANCEL_CW_ROOM_ID = '449648673';
const CLINIC_MENTION_MASTER_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
const CLINIC_MENTION_MASTER_SHEET_NAME = '拠点名';

/**
 * ★★★ 大司令塔 (2/9) から呼び出されるメイン関数 ★★★
 */
function executeAllProcesses_internal() {
  console.log("=============== 応募関連の自動化プロセスを開始 ===============");
  part1_processEmailsToSheet();              // 2-B で定義
  part2_checkForAgencyUrgentApplications();  // 2-C で定義
  part3_postToChatworkFromSheet();           // 2-C で定義
  console.log("\n=============== 応募関連の自動化プロセスが完了 ===============");
}

/**
 * 「拠点名」シートから【正規記載・表記揺れ・共用アカウント・拠点長ID】を読み込む（キャッシュ付き）
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

    const buildManagerTag = (rawId, rawName) => {
      const idStr = String(rawId || '').trim();
      const nameStr = String(rawName || '').trim();
      if (!idStr) return '';
      if (idStr.includes('[To:')) {
        if (idStr.endsWith(']') && nameStr) {
          return `${idStr}${nameStr.endsWith('さん') ? nameStr : nameStr + 'さん'}`;
        }
        return idStr;
      }
      const cleanId = idStr.replace(/[^\d]/g, '');
      if (!cleanId) return '';
      const displayName = nameStr ? (nameStr.endsWith('さん') ? nameStr : `${nameStr}さん`) : '';
      return `[To:${cleanId}]${displayName}`;
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
      const mgrTag = buildManagerTag(mgrId, mgrName);

      list.push({
        canonicalName: aliases[0],
        aliases: aliases,
        sharedTag: sharedTag,
        mgrTag: mgrTag
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
 * 拠点名から「共用アカウント＋拠点長」のメンション文字列（改行付き）を取得する
 * ※【関東】バックアップシフト等の該当なし拠点は空文字（メンションなし）を返す
 */
function getClinicAndManagerMentionText_internal(rawClinicName) {
  if (!rawClinicName) return '';
  const str = String(rawClinicName);

  // バックアップシフト等はメンションなし
  if (str.includes('バックアップ')) return '';

  const cleanInput = str
    .replace(/キャップスクリニック/g, '')
    .replace(/[（(].*?[）)]/g, '')
    .replace(/(小児科|内科|皮膚科|耳鼻咽喉科)/g, '')
    .replace(/[\s ]+/g, '')
    .trim();

  if (!cleanInput) return '';

  const clinicMentionList = getClinicAndManagerMentionList_internal();
  let matchedItem = null;

  // 1. 完全一致
  for (const item of clinicMentionList) {
    if (item.aliases.some(alias => alias.replace(/[\s ]+/g, '') === cleanInput)) {
      matchedItem = item;
      break;
    }
  }

  // 2. 部分一致（文字数が長い表記揺れから優先）
  if (!matchedItem) {
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
  }

  if (!matchedItem) return '';

  const tags = [];
  if (matchedItem.sharedTag) tags.push(matchedItem.sharedTag);
  if (matchedItem.mgrTag && !tags.includes(matchedItem.mgrTag)) tags.push(matchedItem.mgrTag);

  return tags.length > 0 ? tags.join('\n') + '\n' : '';
}

/**
 * 449648673 専用の直前応募メッセージを作成して送信するヘルパー関数
 * ★修正：バックアップシフトの場合はメンションも＜重要＞の注意書きも付けない
 */
function sendUrgentApplyToExtraRoom_internal(clinic, workDateStr, workTime, doctorName, isNight) {
  const isBackup = String(clinic || '').includes('バックアップ');
  const mentionLines = isBackup ? '' : getClinicAndManagerMentionText_internal(clinic);
  const leadText = isNight ? '直近の応募が夜間にありました。\n確認お願いします。' : '直前応募がありました。\n確認お願いします。';

  const importantNoteBlock = isBackup ? '' : `\n＜重要＞
こちらの通知は速報通知になります。
採用の場合は、改めて確定連絡をいたします。
２診目の場合は、拠点移動を打診の可能性があり、必ずしも採用とは限りませんので、ご理解のほどお願いいたします。
医師不在の枠への応募の場合は、原則採用いたします。`;

  const extraMessage = `${mentionLines}[info][title]
直前応募通知
[/title]
${leadText}

応募拠点： ${clinic}
勤務日： ${workDateStr}
応募時間： ${workTime}
応募医師： ${doctorName}${importantNoteBlock}
[/info]`;

  return sendUrgentAlertToChatwork_internal(EXTRA_URGENT_AND_CANCEL_CW_ROOM_ID, extraMessage);
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