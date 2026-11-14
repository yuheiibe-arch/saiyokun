// ============================================================================
// 【3-A_キャンセル_共通基盤】
// ・大司令塔 (4/9) から呼び出されるメイン関数 processAgencyCancelEmails_internal
// ・「処理済み」シートの動的列マッピング取得ヘルパー
// ・拠点名・住所表記の100%正規同定エンジン ＆ 日付・時間パーサー
// ============================================================================

const CANCEL_CLINIC_MASTER_SS_ID_ = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';

/**
 * ★ 大司令塔 (4/9) から定期実行されるエントリーポイント関数
 */
function processAgencyCancelEmails_internal() {
  console.log("=============== 紹介会社・システム承認キャンセル自動反映 開始 ===============");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('処理済み') || ss.getSheetByName('アーカイブ');
  const progressSheet = typeof SHEET_NAME_PROGRESS !== 'undefined'
    ? ss.getSheetByName(SHEET_NAME_PROGRESS)
    : ss.getSheetByName('紹介会社進捗');

  if (!archiveSheet) {
    console.error("処理済み（アーカイブ）シートが見つかりません。");
    return;
  }

  const colMap = getArchiveColMap_internal_(archiveSheet);
  const props = PropertiesService.getScriptProperties();

  // Part 1: 民間医局・MRT キャンセルメールの自動反映 [3-B]
  processMinkanAndMrtCancels_internal_(ss, archiveSheet, progressSheet, colMap, props);

  // Part 2: ドクターサイト「お問い合わせスタッフ送信（キャンセル申請承認）」自動反映 [3-B]
  processSystemApprovedCancels_internal_(archiveSheet, progressSheet, colMap, props, 'newer_than:2d');

  console.log("=============== 紹介会社・システム承認キャンセル自動反映 完了 ===============");
}

/**
 * 「処理済み」シートのヘッダーから各列番号(1始まり)を動的に取得するヘルパー
 */
function getArchiveColMap_internal_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 15);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').replace(/\s+/g, '').trim());

  const findCol = (keywords, fallback1Based) => {
    for (const kw of keywords) {
      const idx = headers.findIndex(h => h === kw || h.includes(kw));
      if (idx !== -1) return idx + 1;
    }
    return fallback1Based;
  };

  return {
    totalCols: lastCol,
    colAgency: findCol(['紹介会社名', '紹介会社'], 1),
    colDate:   findCol(['勤務希望日', '勤務日'], 2),
    colTime:   findCol(['該当時間', '勤務時間'], 3),
    colClinic: findCol(['拠点名', '拠点'], 4),
    colDept:   findCol(['診療科'], 5),
    colId:     findCol(['識別番号'], 6),
    colDoctor: findCol(['医師名', '氏名'], 7),
    colUrl:    findCol(['URL'], 8),
    colType:   9, // I列（区分）
    colStatus: findCol(['処理状況'], 10),
    colRecv:   findCol(['着信日'], 11),
    colDone:   findCol(['対応時間', '処理時間'], 12),
    colPerson: findCol(['対応者', '担当者'], 13),
    colUKey:   findCol(['ユニークキー'], 14),
    colSaiyo:  findCol(['採用可否'], 15)
  };
}

/**
 * 拠点名・エムスリー住所表記を拠点マスタの「正規記載」へ100%同定するヘルパー（キャッシュ付き）
 */
function identifyCancelClinic_internal_(rawInput) {
  if (!identifyCancelClinic_internal_.cache) {
    const nameToFormalMap = new Map();
    const addressEntries = [];
    const formalList = [];

    const normalizeAddr = (str) => {
      if (!str) return '';
      let s = String(str)
        .replace(/[【】\(（][^【】\(（\)）]*(内科|小児科)[^【】\(（\)）]*[\)）]/g, '')
        .replace(/[【】\(（]?(一般内科|内科|小児科)[\)）]?/g, '')
        .replace(/[\s ]+/g, '')
        .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
        .replace(/[ヶｹ]/g, 'ケ')
        .replace(/[−ー―‐]/g, '-');
      const kanjiMap = { '〇': 0, '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
      s = s.replace(/[〇零一二三四五六七八九十]+(?=丁目|番|号)/g, match => {
        if (match === '十') return '10';
        let total = 0, cur = 0;
        for (const ch of match) {
          if (ch === '十') { total += (cur === 0 ? 1 : cur) * 10; cur = 0; }
          else { cur = kanjiMap[ch]; }
        }
        return String(total + cur);
      });
      return s;
    };

    try {
      const locSs = SpreadsheetApp.openById(CANCEL_CLINIC_MASTER_SS_ID_);
      const locSheet = locSs.getSheetByName('拠点名') || locSs.getSheets()[0];
      const locData = locSheet.getDataRange().getDisplayValues();
      const locHeaders = locData[0].map(h => String(h).replace(/[\s ]+/g, ''));
      const idxFormal = locHeaders.indexOf('正規記載');
      const idxAddress = locHeaders.findIndex(h => h.includes('住所'));
      const aliasIndices = locHeaders.map((h, i) => (h.includes('表記揺れ') || h.includes('クリニックNo')) ? i : -1).filter(i => i !== -1);

      for (let i = 1; i < locData.length; i++) {
        const formal = String(locData[i][idxFormal] || '').replace(/[\s ]+/g, '');
        if (!formal) continue;
        formalList.push({ formal, key: formal.replace(/[ヶｹ]/g, 'ケ') });
        const cleanKey = (v) => String(v).replace(/[\s 【】]+/g, '').replace(/[ヶｹ]/g, 'ケ');
        nameToFormalMap.set(formal, formal);
        nameToFormalMap.set(cleanKey(formal), formal);

        aliasIndices.forEach(idx => {
          const alias = String(locData[i][idx] || '').replace(/[\s ]+/g, '');
          if (alias) {
            nameToFormalMap.set(alias, formal);
            nameToFormalMap.set(cleanKey(alias), formal);
          }
        });

        if (idxAddress !== -1 && locData[i][idxAddress]) {
          const rawAddr = String(locData[i][idxAddress]).trim();
          const normAddr = normalizeAddr(rawAddr);
          const hyphenAddr = normAddr.replace(/(\d+)丁目(\d+)?/g, (_, c1, c2) => c2 ? `${c1}-${c2}` : `${c1}-`);
          const townMatch = normAddr.match(/^(.+?[都道府県].+?[市区町村].+?\d+丁目)/);
          const townPrefix = townMatch ? townMatch[1] : '';
          if (normAddr) addressEntries.push({ formal, normAddr, hyphenAddr, townPrefix });
        }
      }
      formalList.sort((a, b) => b.key.length - a.key.length);
    } catch (e) {
      console.error('拠点マスタ読込エラー(紹介会社キャンセル): ' + e.message);
    }

    identifyCancelClinic_internal_.cache = { nameToFormalMap, addressEntries, formalList, normalizeAddr };
  }

  const { nameToFormalMap, addressEntries, formalList, normalizeAddr } = identifyCancelClinic_internal_.cache;
  const raw = String(rawInput || '').trim();
  if (!raw) return { resolved: false, formal: '' };

  const stripped = raw
    .replace(/キャップスクリニック|CAPSクリニック|CAPS|クリニック/gi, '')
    .replace(/[【】\(（][^【】\(（\)）]*(内科|小児科)[^【】\(（\)）]*[\)）]/g, '')
    .replace(/[【】\(（]?(一般内科|内科|小児科)[\)）]?/g, '')
    .replace(/[\s ]+/g, '');
  const rawClean = raw.replace(/[\s 【】]+/g, '');

  for (const c of [rawClean, rawClean.replace(/[ヶｹ]/g, 'ケ'), stripped, stripped.replace(/[ヶｹ]/g, 'ケ')]) {
    if (nameToFormalMap.has(c)) return { resolved: true, formal: nameToFormalMap.get(c) };
  }

  const normRaw = normalizeAddr(raw);
  const hyphenRaw = normRaw.replace(/(\d+)丁目(\d+)?/g, (_, c1, c2) => c2 ? `${c1}-${c2}` : `${c1}-`);
  const addrHit = addressEntries.find(e =>
    e.normAddr.includes(normRaw) || normRaw.includes(e.normAddr) ||
    e.hyphenAddr.includes(hyphenRaw) || hyphenRaw.includes(e.hyphenAddr) ||
    (e.townPrefix && normRaw.includes(e.townPrefix))
  );
  if (addrHit) return { resolved: true, formal: addrHit.formal };

  const strippedKe = stripped.replace(/[ヶｹ]/g, 'ケ');
  const partialHit = formalList.find(item => item.key.length >= 2 && strippedKe.includes(item.key));
  if (partialHit) return { resolved: true, formal: partialHit.formal };

  return { resolved: false, formal: raw };
}

/**
 * 日付文字列・Dateオブジェクトを「yyyy/MM/dd」に正規化するヘルパー
 */
function normalizeCancelDateStr_(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'JST', 'yyyy/MM/dd');
  }
  const s = String(val).replace(/\s*[（(].*?[）)]/g, '').replace(/[年月]/g, '/').replace(/日/g, '').replace(/-/g, '/').trim();
  const m = s.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) return '';
  return `${m[1]}/${String(parseInt(m[2], 10)).padStart(2, '0')}/${String(parseInt(m[3], 10)).padStart(2, '0')}`;
}

/**
 * 時間帯文字列（例: "15:00〜18:00"）を開始分・終了分にパースするヘルパー
 */
function parseCancelTimeRange_(timeStr) {
  if (!timeStr) return null;
  const m = String(timeStr).match(/(\d{1,2}):(\d{2})\s*[〜~～\-]\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return {
    startMin: parseInt(m[1], 10) * 60 + parseInt(m[2], 10),
    endMin:   parseInt(m[3], 10) * 60 + parseInt(m[4], 10)
  };
}