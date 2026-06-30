// =============================================================
// 【ファイル2：フォーム・URL処理用スクリプト】
// ※こだわりの書式そのままの完全復元版
// =============================================================

/**
 * Googleフォームから回答が送信された時に実行される（大司令塔から呼ばれます）
 */
function processFormResponse(e) {
  const values = e.namedValues;
  const sheetName = e.range.getSheet().getName();
  const submitterName = getValue(values, '担当者名');
  const doctorName = getValue(values, '氏名');
  const fileUploadQuestionTitle = '履歴書';
  let fileInfoBlock = '';
  if (values[fileUploadQuestionTitle] && values[fileUploadQuestionTitle][0]) {
    const driveUrl = values[fileUploadQuestionTitle][0];
    fileInfoBlock = createFileAttachmentBlock(driveUrl, doctorName);
  }
  let chatworkMessage = '';
  switch (sheetName) {
    case '〈採用報告〉':
      chatworkMessage = createSaiyoChatMessage(values, fileInfoBlock, submitterName);
      break;
    case '＜不採用報告＞':
      chatworkMessage = createFusaiyoChatMessage(values, fileInfoBlock, submitterName);
      break;
    case '＜DS承認報告＞':
      chatworkMessage = createDsChatMessage(values, fileInfoBlock, submitterName);
      break;
    default:
      return;
  }
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('URL');
  if (targetSheet) {
    targetSheet.getRange('C2').setValue(chatworkMessage);
    targetSheet.getRange('C3').setValue(submitterName);
  }
}

// ====================================
// ▼▼▼ 各種テキスト生成・ヘルパー関数群 ▼▼▼
// ====================================

const FORM_URLS = {
  '採用報告': 'https://docs.google.com/forms/d/e/1FAIpQLSdCNS0K_P_rTX6e1g_i0bdz4Xrr9T1XOPs7AbJOD7TVcrTXmA/viewform',
  'DS承認報告': 'https://docs.google.com/forms/d/e/1FAIpQLSfE-KQZHXMJ2LdmnASN0RSIV_t1oClmDsj5QT4luK6Snqdh8Q/viewform',
  '不採用報告': 'https://docs.google.com/forms/d/e/1FAIpQLSfSq62bUdDnvnfRFAZvP7A88I9C8oqYxTt0HrIpfk0BJHX8pQ/viewform'
};

function displayFormUrl() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('URL');
  if (!sheet) {
    spreadsheet.toast('シート「URL」が見つかりません。');
    return;
  }
  const formType = sheet.getRange('A2').getValue();
  const url = FORM_URLS[formType];
  if (url) {
    sheet.getRange('B2').setValue(url);
    spreadsheet.toast(`「${formType}」のURLをB2に表示しました。`);
  } else {
    sheet.getRange('B2').clearContent();
    spreadsheet.toast(`A2で選択された「${formType}」に対応するURLが見つかりません。`);
  }
}

function getValue(values, key) {
  const value = values[key] ? values[key][0] : '';
  return String(value).trim();
}

function formatHistoryText(rawText) {
  if (!rawText) return '';
  return rawText.split('\n').map(line => {
    const trimmedLine = line.trim();
    const match = trimmedLine.match(/^(\d{4})\s+(\d{1,2})\s*(.*)/);
    if (match) {
      const year = match[1];
      const month = match[2];
      const rest = match[3];
      return `　${year}年${month}月 ${rest}`;
    }
    return '　' + trimmedLine;
  }).join('\n');
}

function createFileAttachmentBlock(url, doctorName) {
  if (!url) return '';
  try {
    const fileId = url.includes('id=') ? url.split('id=')[1] : url.split('/d/')[1].split('/')[0];
    if (!fileId) return `（URLからファイルIDを特定できませんでした: ${url}）`;
    const file = DriveApp.getFileById(fileId);
    const cleanDoctorName = doctorName.replace(/（.*）/g, '').replace(/先生/g, '').trim();
    const title = `（${cleanDoctorName}医師履歴書）`;
    const description = `【${file.getUrl()}】`;
    return `${title}\n${description}`;
  } catch (e) {
    return `（ファイルの処理に失敗しました: ${url}）`;
  }
}

function createSaiyoChatMessage(values, fileInfoBlock, submitterName) {
  const kinmuJikan = getValue(values, '勤務時間');
  const hosokuInfo = getValue(values, '補足情報（任意）');
  const jikanBlock = hosokuInfo ? `${kinmuJikan}（${hosokuInfo}）` : kinmuJikan;
  const kinmu_date_raw = getValue(values, '勤務日');
  let kinmu_date = '';
  if (kinmu_date_raw) {
    const dateObj = new Date(kinmu_date_raw);
    const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const day = dateObj.getDate().toString().padStart(2, '0');
    const dayOfWeekJp = ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
    kinmu_date = `${month}/${day}（${dayOfWeekJp}）`;
  }
  const gakurekiRaw = getValue(values, '学歴・職歴');
  const gakurekiFormatted = formatHistoryText(gakurekiRaw);
  const shimeiRaw = getValue(values, '氏名');
  const shimeiFormatted = shimeiRaw.endsWith('先生') ? shimeiRaw : `${shimeiRaw}先生`;
  const mainInfo = [ `【${getValue(values, '拠点')}／${getValue(values, '診療科')}】${kinmu_date} ${jikanBlock}`, `紹介会社：${getValue(values, '紹介会社')}`, `氏名：${shimeiFormatted}`, `性別：${getValue(values, '性別')}`, `①65歳以下⇒${getValue(values, '年齢')}歳`, `②診療科の一致⇒${getValue(values, '診療科')}`, `③クレーム歴なし⇒${getValue(values, '紹介元でのクレーム歴')} ${getValue(values, '勤務状況の詳細（任意）クレームなど懸念事項')}`.trim(), `④当法人で勤務歴⇒${getValue(values, '当法人での勤務歴')} ${getValue(values, '勤務歴の詳細（任意）')} ${getValue(values, '勤務歴等補足があれば')}`.trim(), `専門科目：${getValue(values, '専門科目')}`, `学歴・職歴：`, `${gakurekiFormatted}` ].join('\n');
  const tantoSha = submitterName ? `担当者：${submitterName}` : '';
  const bodyParts = [ '〈採用報告〉', tantoSha, `[info]${mainInfo}[/info]`, getValue(values, '参考URL（勤務先etc）'), getValue(values, '所感・申し送り事項'), fileInfoBlock ];
  return bodyParts.filter(String).join('\n\n');
}

function createFusaiyoChatMessage(values, fileInfoBlock, submitterName) {
  const baseMessage = createSaiyoChatMessage(values, fileInfoBlock, submitterName);
  const fusaiyoRiyuu = getValue(values, '不採用理由');
  let finalMessage = baseMessage.replace('〈採用報告〉', '〈不採用報告〉');
  if (fusaiyoRiyuu) {
    finalMessage += `\n\n[hr]不採用理由：\n${fusaiyoRiyuu}`;
  }
  return finalMessage;
}

function createDsChatMessage(values, fileInfoBlock, submitterName) {
  const birthYear = getValue(values, '生年');
  const age = birthYear ? `${new Date().getFullYear() - parseInt(birthYear, 10)}歳` : '';
  const shimeiRaw = getValue(values, '氏名');
  const shimeiFormatted = shimeiRaw.endsWith('先生') ? shimeiRaw : `${shimeiRaw}先生`;
  const summaryParagraph = [ `${birthYear}年生`, age, `${getValue(values, '卒業大学')}卒`, `同年医師免許取得` ].filter(String).join('.') + `。` + `${getValue(values, '専門医資格（任意）')}。` + `現在、${getValue(values, '現在の勤務先')}。 ` + `${getValue(values, '居住地')}在住。` + `最寄り駅は${getValue(values, '最寄り駅')}駅`;
  const summaryBlock = `[info]${shimeiFormatted}\n${summaryParagraph}[/info]`;
  const detailsRaw = getValue(values, '詳細情報');
  const simpleName = `${getValue(values, '氏名').split(/[\s　]/)[0]}先生`;
  const detailsBlock = detailsRaw ? `[info][title]${simpleName}詳細情報[/title]\n${detailsRaw}[/info]` : '';
  const hrBlock = detailsRaw ? '[hr]' : '';
  const nextActionSensei = getValue(values, '【Next Action】先生へのご案内');
  const nextActionDs = getValue(values, '【Next Action】DS部内のタスク');
  let nextActionBlock = '';
  if (nextActionSensei || nextActionDs) {
    const senseiPart = nextActionSensei ? `〈先生〉\n${nextActionSensei}` : '';
    const dsPart = nextActionDs ? `〈DS部〉\n${nextActionDs}` : '';
    const nextActionBody = [senseiPart, dsPart].filter(String).join('\n');
    nextActionBlock = `[info]【Next action】\n${nextActionBody}\n[/info][/info]`;
  }
  const keirekiRaw = getValue(values, '職務経歴');
  const keirekiFormatted = formatHistoryText(keirekiRaw);
  const keirekiBlock = keirekiFormatted ? `[info][title]職務経歴[/title]\n${keirekiFormatted}[/info]` : '';
  const tantoSha = submitterName ? `担当者：${submitterName}` : '';
  const bodyParts = [ '〈DS承認報告〉', tantoSha, summaryBlock, detailsBlock, hrBlock, nextActionBlock, getValue(values, '参考URL（勤務先etc）'), keirekiBlock, getValue(values, '所感・申し送り事項'), fileInfoBlock ];
  return bodyParts.filter(String).join('\n\n');
}