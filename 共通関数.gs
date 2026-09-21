var NOTIFY_SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
var NOTIFY_SLACK_MENTION_SHEET_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';

function setUp() {
  const apiKey = ''; 
  if (apiKey !== '') {
    PropertiesService.getScriptProperties().setProperty('CHATWORK_API_KEY', apiKey);
    console.log('ChatworkのAPIキーを登録・更新しました。');
  } else {
    console.log('apiKeyが空欄のため、誤作動防止で登録をスキップしました。');
  }
}

function getSlackMentionMap() {
  if (getSlackMentionMap.cache) return getSlackMentionMap.cache;
  try {
    const ss = SpreadsheetApp.openById(NOTIFY_SLACK_MENTION_SHEET_ID);
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

  // ★夜間（21:00〜07:00）の <!channel> 制限フラグ
  const currentHour = new Date().getHours();
  const isNightTime = (currentHour >= 21 || currentHour < 7);
  const isEmergencyCancel = slackMsg.includes('直近緊急キャンセル') || slackMsg.includes('緊急キャンセル');
  const shouldSuppressChannel = isNightTime && !isEmergencyCancel; 

  // ==============================================================
  // ▼▼▼ 1. 医師からのメール（個人） 専用の処理 ▼▼▼
  // ==============================================================
  if (slackMsg.includes('[title]医師からのメール[/title]')) {
    channelId = 'C0BV5NT4TLY';
    
    const dateMatch = slackMsg.match(/受信時刻：([^\n]+)/);
    const subjMatch = slackMsg.match(/件名：([^\n]+)/);
    const dateStr = dateMatch ? dateMatch[1].trim() : "";
    const subjStr = subjMatch ? subjMatch[1].trim() : "";
    
    const titleRegex = /\[title\](.+?)\s*先生\[\/title\]/;
    const docMatch = slackMsg.match(titleRegex);
    const realDocName = docMatch ? docMatch[1].trim() : "不明な";

    let bodyContent = "";
    const bodyMatch = slackMsg.match(/\[title\].+? 先生\[\/title\]\n([\s\S]*?)\[\/info\]/);
    if (bodyMatch) {
        bodyContent = bodyMatch[1].trim();
    }

    const firstLineRegex = /^.+?\s*先生\n/;
    if (firstLineRegex.test(bodyContent)) {
        bodyContent = bodyContent.replace(firstLineRegex, '').trim();
    }

    let mentionText = "";
    const tags = slackMsg.match(/\[To:\d+\][^\s\n]*/g);
    if (tags) {
        tags.forEach(tag => {
            let foundMatch = false;
            const idMatch = tag.match(/\[To:(\d+)\]/);
            if (idMatch) {
                const cwId = idMatch[1];
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
        if (!shouldSuppressChannel) mentionText += "<!channel> ";
    }
    
    // ★修正：無意味な `@担当者` の捏造付与を完全に削除しました

    const formattedSlackMsg = `\`医師からメール\`\n${mentionText.trim()}\n受信時刻：${dateStr}\n件名：${subjStr}\n\`${realDocName} 先生\`\n\n\`\`\`\n${bodyContent}\n\`\`\``;
    return { channelId: channelId, text: formattedSlackMsg };
  }

  // ==============================================================
  // ▼▼▼ 2. 特定アラート 専用の処理 ▼▼▼
  // ==============================================================
  if (slackMsg.includes('特定アラート')) {
    channelId = 'C0BV5NT4TLY'; 
    
    const dateMatch = slackMsg.match(/受信時刻：([^\n]+)/);
    const subjMatch = slackMsg.match(/件名：([^\n]+)/);
    const dateStr = dateMatch ? dateMatch[1].trim() : "不明";
    const subjStr = subjMatch ? subjMatch[1].trim() : "不明";
    
    let mentionText = "";
    const tags = slackMsg.match(/\[To:\d+\][^\s\n]*/g);
    if (tags) {
        tags.forEach(tag => {
            let foundMatch = false;
            const idMatch = tag.match(/\[To:(\d+)\]/);
            if (idMatch) {
                const cwId = idMatch[1];
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
        if (!shouldSuppressChannel) mentionText += "<!channel> ";
    }

    // ★修正：無意味な `@担当者` の捏造付与を完全に削除しました

    let bodyContent = slackMsg;
    bodyContent = bodyContent.replace(/\[To:\d+\][^\s\n]*/g, '');
    bodyContent = bodyContent.replace(/\[toall\]/ig, '');
    bodyContent = bodyContent.replace(/\[info\]/ig, '').replace(/\[\/info\]/ig, '');
    bodyContent = bodyContent.replace(/\[title\].*?\[\/title\]/ig, '');
    bodyContent = bodyContent.replace(/受信時刻：[^\n]+/g, '');
    bodyContent = bodyContent.replace(/件名：[^\n]+/g, '');
    bodyContent = bodyContent.replace(/^\s*\n/gm, ''); 
    bodyContent = bodyContent.trim();

    const formattedSlackMsg = `\`特定アラート\`\n${mentionText.trim()}\n受信時刻：${dateStr}\n件名：${subjStr}\n\n\`\`\`\n${bodyContent}\n\`\`\``;
    return { channelId: channelId, text: formattedSlackMsg };
  }

  // ==============================================================
  // ▼▼▼ 3. 直前応募 専用の処理 ▼▼▼
  // ==============================================================
  if (slackMsg.includes('直前応募通知')) {
    channelId = 'C09TRHGU64D';
    
    let customTitle = slackMsg.includes('紹介会社経由') ? '*【直前応募 紹介会社】*' : '*【直前応募】*';
    
    let bodyContent = slackMsg;
    bodyContent = bodyContent.replace(/\[To:\d+\][^\s\n]*/g, '');
    bodyContent = bodyContent.replace(/\[toall\]/ig, '');
    bodyContent = bodyContent.replace(/\[info\]/ig, '').replace(/\[\/info\]/ig, '');
    bodyContent = bodyContent.replace(/\[title\].*?\[\/title\]/ig, '');
    bodyContent = bodyContent.replace(/^\s*\n/gm, ''); 
    bodyContent = bodyContent.trim();

    const channelTag = shouldSuppressChannel ? "" : "<!channel>\n";
    const formattedSlackMsg = `${customTitle}\n${channelTag}\`\`\`\n${bodyContent}\n\`\`\``;
    return { channelId: channelId, text: formattedSlackMsg };
  }

  // ==============================================================
  // ▼▼▼ 4. 採用・不採用・DS承認報告 専用の処理 ▼▼▼
  // ==============================================================
  if (slackMsg.includes('〈採用報告〉') || slackMsg.includes('〈不採用報告〉') || slackMsg.includes('〈DS承認報告〉')) {
    channelId = 'C0BTW2070UF'; 

    let reportType = '';
    let splitKey = '';
    if (slackMsg.includes('〈採用報告〉')) { reportType = '【採用報告】'; splitKey = '〈採用報告〉'; }
    else if (slackMsg.includes('〈不採用報告〉')) { reportType = '【不採用報告】'; splitKey = '〈不採用報告〉'; }
    else if (slackMsg.includes('〈DS承認報告〉')) { reportType = '【DS承認報告】'; splitKey = '〈DS承認報告〉'; }
    
    let parts = slackMsg.split(splitKey);
    let headerPart = parts[0] || '';
    let bodyPart = parts[1] || '';

    let tantoMatch = bodyPart.match(/担当：([^\n]+)/);
    let tanto = tantoMatch ? tantoMatch[1].trim() : '';
    if (tanto) {
        bodyPart = bodyPart.replace(tantoMatch[0], ''); 
    }
    
    let customTitle = `*${reportType}*`;
    if (tanto) {
        customTitle += ` （担当：${tanto}）`;
    }

    // 名前なしタグも救済するメンション変換処理
    mentionMap.forEach(item => {
      const toMatch = item.cwName.match(/\[To:\d+\]/);
      const toTag = toMatch ? toMatch[0] : null;

      if (headerPart.includes(item.cwName)) {
        const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
        headerPart = headerPart.split(item.cwName).join(replacement);
      } else if (toTag && headerPart.includes(toTag)) {
        const escapedTag = toTag.replace(/\[/g, '\\[').replace(/\]/g, '\\]');
        const regex = new RegExp(escapedTag + '[^\\s<]*', 'g');
        const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
        headerPart = headerPart.replace(regex, replacement);
      }
      
      if (bodyPart.includes(item.cwName)) {
        const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
        bodyPart = bodyPart.split(item.cwName).join(replacement);
      } else if (toTag && bodyPart.includes(toTag)) {
        const escapedTag = toTag.replace(/\[/g, '\\[').replace(/\]/g, '\\]');
        const regex = new RegExp(escapedTag + '[^\\s<]*', 'g');
        const replacement = item.slackId ? `<@${item.slackId}>` : item.displayName;
        bodyPart = bodyPart.replace(regex, replacement);
      }
    });

    headerPart = headerPart.replace(/\[To:\d+\][^\s<]*/g, '');
    
    if (shouldSuppressChannel) {
        headerPart = headerPart.replace(/\[toall\]/ig, '');
    } else {
        headerPart = headerPart.replace(/\[toall\]/ig, '<!channel>');
    }
    
    let headerTokens = headerPart.replace(/[\n\r]+/g, ' ').split(' ').filter(t => t.trim() !== '');
    headerTokens = [...new Set(headerTokens)]; 
    let headerText = headerTokens.join(' ');

    bodyPart = bodyPart.replace(/\[info\]/ig, '').replace(/\[\/info\]/ig, '');
    bodyPart = bodyPart.replace(/\[title\]([\s\S]*?)\[\/title\]/g, '*$1*');
    bodyPart = bodyPart.replace(/\[hr\]/ig, '---------------------------------------');
    bodyPart = bodyPart.replace(/^\s*\n/gm, '').trim();

    let formattedSlackMsg = '';
    if (headerText) {
        formattedSlackMsg += `${headerText}\n`;
    }
    formattedSlackMsg += `${customTitle}\n\`\`\`\n${bodyPart}\n\`\`\``;
    
    return { channelId: channelId, text: formattedSlackMsg.trim() };
  }

  // ==============================================================
  // ▼▼▼ 5. その他（既存の汎用フォーマット） ▼▼▼
  // ==============================================================
  if (slackMsg.includes('エムスリー') || slackMsg.includes('オファー希望') || slackMsg.includes('メッセージ受信') || slackMsg.includes('（オファー経由）勤務確定')) {
    channelId = 'C0BV5NT4TLY';
    customTitle = '*【エムスリー医師のお問い合わせ】*';
  } else if (slackMsg.includes('ATMSメール')) {
    channelId = 'C0BU7687HDK';
  } else if (slackMsg.includes('民間医局くん') || slackMsg.includes('給与：')) {
    channelId = 'C09TRHGU64D';
    customTitle = '*【民間医局 直接応募通知】*';
  } else if (slackMsg.includes('直近緊急キャンセル') || slackMsg.includes('時間外不在着信') || slackMsg.includes('時間外着信') || slackMsg.includes('非定型キャンセル')) {
    channelId = 'C09TRHGU64D';
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
  
  if (shouldSuppressChannel) {
      slackMsg = slackMsg.replace(/\[toall\]/ig, '');
  } else {
      slackMsg = slackMsg.replace(/\[toall\]/ig, '<!channel>');
  }

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
  
  const url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}