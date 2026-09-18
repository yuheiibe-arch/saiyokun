var NOTIFY_SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
var NOTIFY_SLACK_MENTION_SHEET_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';

function setUp() {
  const apiKey = '';
  PropertiesService.getScriptProperties().setProperty('CHATWORK_API_KEY', apiKey);
}

function getSlackMentionMap() {
  if (getSlackMentionMap.cache) return getSlackMentionMap.cache;
  try {
    const ss = SpreadsheetApp.openById(NOTIFY_SLACK_MENTION_SHEET_ID);
    
    // ★ 修正：一番左のシートではなく「Slack変換」シートを名指しで取得する
    const sheet = ss.getSheetByName('Slack変換');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const map = [];
    for (let i = 1; i < data.length; i++) {
      const cwName = String(data[i][0]).trim();
      const slackId = String(data[i][1]).trim();
      const displayName = String(data[i][2]).trim();
      if (cwName) map.push({ cwName: cwName, slackId: slackId, displayName: displayName });
    }
    getSlackMentionMap.cache = map;
    return map;
  } catch (e) {
    return [];
  }
}

function buildSlackPayload(cwMessage, cwRoomId) {
  let slackMsg = cwMessage;
  let channelId = '';
  let customTitle = '';
  const mentionMap = getSlackMentionMap();

  // ==============================================================
  // ▼▼▼ 医師からのメール（個人）専用の Slack綺麗化パース処理 ▼▼▼
  // ==============================================================
  if (slackMsg.includes('[title]医師からのメール[/title]')) {
    channelId = 'C0BV5NT4TLY';
    
    // データ抽出
    const dateMatch = slackMsg.match(/受信時刻：([^\n]+)/);
    const subjMatch = slackMsg.match(/件名：([^\n]+)/);
    const dateStr = dateMatch ? dateMatch[1].trim() : "";
    const subjStr = subjMatch ? subjMatch[1].trim() : "";
    
    // 1. 医師名の抽出 ("先生"を含む最初のタイトルを確実に狙う)
    const titleRegex = /\[title\](.+?)\s*先生\[\/title\]/;
    const docMatch = slackMsg.match(titleRegex);
    const realDocName = docMatch ? docMatch[1].trim() : "不明な";

    // 2. 本文の抽出
    let bodyContent = "";
    const bodyMatch = slackMsg.match(/\[title\].+? 先生\[\/title\]\n([\s\S]*?)\[\/info\]/);
    if (bodyMatch) {
        bodyContent = bodyMatch[1].trim();
    }

    // 3. 誤判定で挿入されてしまった冒頭の「〇〇 先生」を完全除去
    const firstLineRegex = /^.+?\s*先生\n/;
    if (firstLineRegex.test(bodyContent)) {
        bodyContent = bodyContent.replace(firstLineRegex, '').trim();
    }

    // 4. Slack用メンションの確実な変換
    let mentionText = "";
    const tags = slackMsg.match(/\[To:\d+\][^\s\n]*/g);
    if (tags) {
        tags.forEach(tag => {
            let foundMatch = false;
            const idMatch = tag.match(/\[To:(\d+)\]/);
            if (idMatch) {
                const cwId = idMatch[1];
                // 実際のシートデータ (mentionMap) から同じ「数字ID」を持つものを探す
                const exactItem = mentionMap.find(item => item.cwName.includes(cwId));
                if (exactItem) {
                    mentionText += exactItem.slackId ? `<@${exactItem.slackId}> ` : `${exactItem.displayName} `;
                    foundMatch = true;
                }
            }
            if (!foundMatch) mentionText += tag + " ";
        });
    }
    
    if (slackMsg.toLowerCase().includes('[toall]')) {
        mentionText += "<!channel> ";
    }
    
    if (!mentionText.trim()) {
        mentionText = "`@担当者`";
    }

    // 5. コードブロック（```）によるフォーマット
    const formattedSlackMsg = `\`医師からメール\`\n${mentionText.trim()}\n受信時刻：${dateStr}\n件名：${subjStr}\n\`${realDocName} 先生\`\n\n\`\`\`\n${bodyContent}\n\`\`\``;
    
    return { channelId: channelId, text: formattedSlackMsg };
  }

  // ==============================================================
  // ▼▼▼ 以下、既存の処理（個人メール以外の汎用フォーマット） ▼▼▼
  // ==============================================================
  if (slackMsg.includes('エムスリー') || slackMsg.includes('オファー希望') || slackMsg.includes('メッセージ受信') || slackMsg.includes('（オファー経由）勤務確定')) {
    channelId = 'C0BV5NT4TLY';
    customTitle = '*【エムスリー医師のお問い合わせ】*';
  } else if (slackMsg.includes('ATMSメール')) {
    channelId = 'C0BU7687HDK';
  } else if (slackMsg.includes('民間医局くん') || slackMsg.includes('給与：')) {
    channelId = 'C09TRHGU64D';
    customTitle = '*【民間医局 直接応募通知】*';
  } else if (slackMsg.includes('直前応募通知')) {
    channelId = 'C09TRHGU64D';
    customTitle = '*【直前応募 紹介会社】*';
  } else if (slackMsg.includes('直近緊急キャンセル') || slackMsg.includes('時間外不在着信') || slackMsg.includes('時間外着信') || slackMsg.includes('非定型キャンセル')) {
    channelId = 'C09TRHGU64D';
  } else if (slackMsg.includes('〈採用報告〉') || slackMsg.includes('〈不採用報告〉') || slackMsg.includes('〈DS承認報告〉')) {
    channelId = 'C0BTW2070UF';
  } else {
    channelId = 'C09TRHGU64D';
  }
  
  mentionMap.forEach(item => {
    if (slackMsg.includes(item.cwName)) {
      const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
      slackMsg = slackMsg.split(item.cwName).join(replacement);
    } else {
      const toMatch = item.cwName.match(/\[To:\d+\]/);
      if (toMatch) {
        const toTag = toMatch[0];
        if (slackMsg.includes(toTag)) {
          const escapedTag = toTag.replace(/\[/g, '\\[').replace(/\]/g, '\\]');
          const regex = new RegExp(escapedTag + '[^\\s ]*', 'g');
          const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
          slackMsg = slackMsg.replace(regex, replacement);
        }
      }
    }
  });
  
  slackMsg = slackMsg.replace(/\[To:\d+\][^\s<]*/g, '');
  slackMsg = slackMsg.replace(/\[toall\]/ig, '<!channel>');
  slackMsg = slackMsg.replace(/\[info\]/g, '').replace(/\[\/info\]/g, '').replace(/\[hr\]/g, '\n---------------------------------------\n');
  slackMsg = slackMsg.replace(/\[title\]([\s\S]*?)\[\/title\]/g, '*$1*\n');
  
  if (customTitle) {
    if (slackMsg.includes('<!channel>')) {
      slackMsg = slackMsg.replace('<!channel>', `<!channel>\n${customTitle}\n`);
    } else {
      slackMsg = `${customTitle}\n\n${slackMsg}`;
    }
  }
  
  return { channelId: channelId, text: slackMsg.trim() };
}

function sendToChatwork(roomId, message) {
  try {
    const slackPayload = buildSlackPayload(message, roomId);
    if (slackPayload && slackPayload.channelId) {
      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify({ "channel": slackPayload.channelId, "text": slackPayload.text })
      };
      UrlFetchApp.fetch(NOTIFY_SLACK_WEBHOOK_URL, options);
    }
  } catch (e) {}
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey) return false;
  const url = `[https://api.chatwork.com/v2/rooms/$](https://api.chatwork.com/v2/rooms/$){roomId}/messages`;
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}