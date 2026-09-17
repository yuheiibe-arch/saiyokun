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
const sheet = ss.getSheets()[0];
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
if (slackMsg.includes('エムスリー') || slackMsg.includes('オファー希望') || slackMsg.includes('メッセージ受信') || slackMsg.includes('（オファー経由）勤務確定')) {
channelId = 'C0BV5NT4TLY';
customTitle = '*【エムスリー医師のお問い合わせ】*';
} else if (slackMsg.includes('医師からのメール')) {
channelId = 'C0BV5NT4TLY';
const docMatch = slackMsg.match(/\[title\](.+?) 先生\[\/title\]/);
const docName = docMatch ? docMatch[1].trim() : "不明な";
customTitle = `*【${docName}医師からのメール】*`;
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
const mentionMap = getSlackMentionMap();
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
slackMsg = slackMsg.replace(/\[To:\d+\][^\s ]*/g, '');
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
const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
const options = { method: 'post', headers: { 'X-ChatWorkToken': apiKey }, payload: { body: message }, muteHttpExceptions: true };
try {
const response = UrlFetchApp.fetch(url, options);
return response.getResponseCode() === 200;
} catch (e) {
return false;
}
}