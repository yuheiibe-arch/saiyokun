// ============================================================================
// 【4-A_即決アーカイブ_共通基盤】
// ・『次回即決アーカイブ』シートのヘッダー定義・レイアウト初期化・書式設定
// ・〈採用報告〉＜DS承認報告＞の横断名寄せ辞書 ＆ その後の新規登録判定エンジン
// ・氏名正規化（空白・異体字・旧字体・敬称・括弧なしカナ吸収）・旧姓抽出ヘルパー
// ============================================================================

const SOKKETSU_ARCHIVE_HEADERS_ = [
  '判定区分',
  'その後の登録状況',
  '登録詳細（採用・DS承認）',
  '報告日',
  '応募勤務日',
  'エリア',
  '応募拠点',
  '応募診療科',
  '紹介会社',
  '氏名',
  '年齢・性別',
  '専門科目',
  '不採用理由（種別）',
  'コメント抜粋',
  '履歴書URL',
  '元シート行・キー'
];

/**
 * アーカイブシートを取得（存在しなければ新規作成して初期レイアウトを適用）
 */
function ensureNextSokketsuSheet_(ss) {
  const targetName = typeof NEXT_SOKKETSU_SHEET_NAME !== 'undefined' ? NEXT_SOKKETSU_SHEET_NAME : '次回即決アーカイブ';
  let sheet = ss.getSheetByName(targetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetName);
    setupSokketsuSheetHeaderAndLayout_(sheet);
  }
  return sheet;
}

/**
 * アーカイブシートを白紙リセットして最新ヘッダーで再構築（4-C 一括遡及用）
 */
function resetAndEnsureNextSokketsuSheet_(ss) {
  const targetName = typeof NEXT_SOKKETSU_SHEET_NAME !== 'undefined' ? NEXT_SOKKETSU_SHEET_NAME : '次回即決アーカイブ';
  let sheet = ss.getSheetByName(targetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetName);
  } else {
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.clear();
  }
  setupSokketsuSheetHeaderAndLayout_(sheet);
  return sheet;
}

/**
 * ヘッダー装飾・列幅・フィルタの初期設定
 */
function setupSokketsuSheetHeaderAndLayout_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SOKKETSU_ARCHIVE_HEADERS_.length);
  headerRange.setValues([SOKKETSU_ARCHIVE_HEADERS_]);
  headerRange.setBackground('#1c4587').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const widths = [130, 155, 220, 95, 95, 70, 105, 85, 95, 165, 90, 135, 115, 360, 170, 140];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(1, 1, 1, SOKKETSU_ARCHIVE_HEADERS_.length).createFilter();
}

/**
 * 1行分の書き込み配列（全16列）を組み立てる
 */
function buildSokketsuArchiveRow_(p) {
  const reportDate = formatSokketsuDate_(p.timestampRaw);
  const workDate = formatSokketsuDate_(p.workDateRaw);

  let dept = String(p.deptRaw || '').trim();
  if (dept.includes('小児科') && !dept.includes('麻酔')) dept = '小児科';
  else if (dept.includes('内科')) dept = '内科';

  const ageStr = p.age ? `${String(p.age).replace(/歳$/, '')}歳` : '';
  const ageGender = [ageStr, p.gender].filter(Boolean).join('・');

  let reasonCategory = 'その他（充足等）';
  const rRaw = String(p.reasonRaw || '');
  if (rRaw === '充足' || rRaw.includes('充足')) reasonCategory = '充足';
  else if (rRaw.includes('タイムアウト')) reasonCategory = 'タイムアウト';
  else if (rRaw.includes('1時間枠') || rRaw.includes('オリエンテーション')) reasonCategory = '初回枠制限';
  else if (rRaw.includes('取り下げ') || String(p.shokanRaw || '').includes('取り下げ')) reasonCategory = '本人辞退(即決可)';
  else if (rRaw.includes('４診目') || String(p.shokanRaw || '').includes('４診目')) reasonCategory = '体制上限(4診目)';

  let excerpt = stripChatworkQuotes_(p.shokanRaw);
  if (!excerpt && rRaw.length > 15) {
    excerpt = rRaw;
  }
  excerpt = excerpt.replace(/https?:\/\/www\.chatwork\.com\/[^\s]+/g, '').replace(/\s+/g, ' ').trim();
  if (excerpt.length > 220) {
    excerpt = excerpt.slice(0, 220) + '…';
  }

  return [
    p.rankLabel,
    p.regStatus || '未登録（声掛け対象）',
    p.regDetail || '',
    reportDate,
    workDate,
    p.area || '',
    p.clinic || '',
    dept,
    p.agency || '',
    p.doctorName || '',
    ageGender,
    p.specialty || '',
    reasonCategory,
    excerpt || '（特記事項なし・充足お断り）',
    p.resumeUrl || '',
    p.sourceKey || ''
  ];
}

/**
 * アーカイブシートへ行を一括追記し、背景色スタイルを適用する
 */
function appendSokketsuRowWithFormat_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  const numCols = SOKKETSU_ARCHIVE_HEADERS_.length;
  const range = sheet.getRange(startRow, 1, rows.length, numCols);
  range.setValues(rows);
  range.setVerticalAlignment('middle');
  sheet.getRange(startRow, 14, rows.length, 1).setWrap(true); // N列: コメント抜粋の折り返し

  for (let i = 0; i < rows.length; i++) {
    applySokketsuRowVisualStyle_(sheet, startRow + i, rows[i][0], rows[i][1]);
  }
}

/**
 * 1行の視覚スタイル（未登録＝鮮やか表示 ／ 登録済＝グレーアウト）を適用する
 */
function applySokketsuRowVisualStyle_(sheet, rowNum, rankLabel, regStatus) {
  const numCols = SOKKETSU_ARCHIVE_HEADERS_.length;
  const isRegistered = String(regStatus).includes('登録済');

  if (isRegistered) {
    sheet.getRange(rowNum, 1, 1, numCols).setBackground('#efefef').setFontColor('#666666');
    sheet.getRange(rowNum, 2).setBackground('#d9d2e9').setFontColor('#351c75').setFontWeight('bold');
  } else {
    sheet.getRange(rowNum, 1, 1, numCols).setBackground(null).setFontColor('#000000');
    const rankBg = rankLabel === '★次回即決' ? '#d9ead3' : '#fff2cc';
    sheet.getRange(rowNum, 1).setBackground(rankBg).setFontWeight('bold');
    sheet.getRange(rowNum, 2).setBackground('#cfe2f3').setFontColor('#0b5394').setFontWeight('bold');
  }
}

/**
 * 〈採用報告〉および＜DS承認報告＞シートを全件走査し、医師名ごとの登録履歴マップを構築する
 */
function buildAdoptedAndDsLookupMap_(ss) {
  const map = new Map();

  const addEntry = (docKey, entry) => {
    if (!docKey) return;
    if (!map.has(docKey)) map.set(docKey, []);
    map.get(docKey).push(entry);
  };

  ss.getSheets().forEach(sh => {
    const sName = sh.getName();
    const isSaiyo = sName.includes('採用報告') && !sName.includes('不採用');
    const isDs = sName.includes('DS承認報告');
    if (!isSaiyo && !isDs) return;
    if (sh.getLastRow() < 2) return;

    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(h => String(h || '').trim());
    const findIdx = (cands) => {
      for (const c of cands) {
        const exact = headers.findIndex(h => h === c);
        if (exact !== -1) return exact;
      }
      for (const c of cands) {
        const part = headers.findIndex(h => h.includes(c));
        if (part !== -1) return part;
      }
      return -1;
    };

    const idxTs = 0;
    const idxName = findIdx(['氏名', '医師名']);
    const idxWorkDate = findIdx(['勤務日', '初回勤務日']);
    const idxClinic = findIdx(['拠点']);
    const idxAgency = findIdx(['紹介会社']);
    if (idxName === -1) return;

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const rawName = row[idxName];
      if (!rawName) continue;

      const tsRaw = row[idxTs];
      const entry = {
        type: isDs ? 'DS承認報告' : '採用報告',
        tsMs: parseTimestampMsForSokketsu_(tsRaw),
        reportDate: formatSokketsuDate_(tsRaw),
        workDate: idxWorkDate !== -1 ? formatSokketsuDate_(row[idxWorkDate]) : '',
        clinic: idxClinic !== -1 ? String(row[idxClinic] || '').trim() : '',
        agency: isDs ? '直接DS応募' : (idxAgency !== -1 ? String(row[idxAgency] || '').trim() : '')
      };

      extractCandidateDocKeys_(rawName, '').forEach(k => addEntry(k, entry));
    }
  });

  map.forEach(list => list.sort((a, b) => a.tsMs - b.tsMs));
  return map;
}

/**
 * 不採用報告のタイムスタンプ以降に〈採用報告〉または＜DS承認報告＞があるかを判定する
 */
function checkSubsequentRegistration_(doctorName, shokanRaw, fusaiyoTsRaw, regIndexMap) {
  const fusaiyoMs = parseTimestampMsForSokketsu_(fusaiyoTsRaw);
  const candidateKeys = extractCandidateDocKeys_(doctorName, shokanRaw);

  const afterHits = [];
  const beforeHits = [];

  candidateKeys.forEach(key => {
    const records = regIndexMap.get(key) || [];
    records.forEach(rec => {
      if (fusaiyoMs > 0 && rec.tsMs > 0) {
        if (rec.tsMs >= fusaiyoMs - (12 * 60 * 60 * 1000)) {
          afterHits.push(rec);
        } else {
          beforeHits.push(rec);
        }
      } else {
        afterHits.push(rec);
      }
    });
  });

  if (afterHits.length > 0) {
    afterHits.sort((a, b) => a.tsMs - b.tsMs);
    const hasSaiyo = afterHits.some(h => h.type === '採用報告');
    const hasDs = afterHits.some(h => h.type === 'DS承認報告');
    let status = '登録済（採用報告）';
    if (hasSaiyo && hasDs) status = '登録済（採用・DS両方）';
    else if (hasDs) status = '登録済（DS承認報告）';

    const latest = afterHits[afterHits.length - 1];
    const parts = [
      latest.reportDate ? `報告:${latest.reportDate}` : '',
      latest.workDate ? `勤務:${latest.workDate}` : '',
      latest.clinic,
      latest.agency
    ].filter(Boolean);

    return { isRegistered: true, status, detail: parts.join(' / ') };
  }

  if (beforeHits.length > 0) {
    const prev = beforeHits[beforeHits.length - 1];
    return {
      isRegistered: true,
      status: `以前より登録済（${prev.type}）`,
      detail: `過去報告:${prev.reportDate} ${prev.clinic ? '/ ' + prev.clinic : ''}`.trim()
    };
  }

  return { isRegistered: false, status: '未登録（声掛け対象）', detail: '' };
}

function stripChatworkQuotes_(text) {
  if (!text) return '';
  return String(text)
    .replace(/\[引用[^\]]*\][\s\S]*?\[\/引用\]/g, '')
    .replace(/\[qt\][\s\S]*?\[\/qt\]/g, '')
    .replace(/\[info\][\s\S]*?\[\/info\]/g, '')
    .trim();
}

/**
 * 異体字・旧字体の標準化ヘルパー
 */
function unifyKanjiVariantsForSokketsu_(str) {
  if (!str) return '';
  return String(str)
    .replace(/髙/g, '高')
    .replace(/[﨑嵜]/g, '崎')
    .replace(/[邊邉]/g, '辺')
    .replace(/[齋齊斎]/g, '斉')
    .replace(/廣/g, '広')
    .replace(/[濱濵]/g, '浜')
    .replace(/澤/g, '沢')
    .replace(/櫻/g, '桜')
    .replace(/德/g, '徳')
    .replace(/冨/g, '富')
    .replace(/眞/g, '真')
    .replace(/實/g, '実')
    .replace(/惠/g, '恵')
    .replace(/國/g, '国')
    .replace(/[嶋嶌]/g, '島')
    .replace(/栁/g, '柳')
    .replace(/藪/g, '薮');
}

/**
 * 氏名を名寄せ比較用の正規化キーへ変換する
 * ・括弧以降カット、括弧なしフリガナ併記カット、敬称(先生/医師/ドクター/Dr)カット、全空白除去、異体字統一
 */
function normalizeDocKeyForSokketsu_(name) {
  if (!name) return '';
  let s = String(name)
    .split(/[（(【［\/／]/)[0]
    .trim();

  // 括弧なしで「山田 太郎 ヤマダ タロウ」のように姓名の後ろにフリガナが並んでいる場合のみカナ部分を除去
  // （※「山田 エリカ」のようなカタカナ下の名前は消さない安全設計）
  const kanaTailMatch = s.match(/^([一-龠々]+[  ]+[一-龠々ぁ-んァ-ヶー]+?)[  ]+[ァ-ヶー]{2,}(?:[  ]+[ァ-ヶー]+)+$/);
  if (kanaTailMatch) {
    s = kanaTailMatch[1];
  }

  s = s
    .replace(/(?:先生|医師|ドクター|Dr\.?)$/gi, '')
    .replace(/[\s ]+/g, '')
    .trim();

  return unifyKanjiVariantsForSokketsu_(s);
}

/**
 * メイン氏名に加え、括弧内の漢字別姓（例：「矢島 理香（菅 理香）」）や所感内の「以前『植原 実紅先生』」も抽出
 */
function extractCandidateDocKeys_(rawName, shokanText) {
  const keys = new Set();
  const primary = normalizeDocKeyForSokketsu_(rawName);
  if (primary) keys.add(primary);

  const parenMatches = String(rawName || '').match(/[（(]([^）)]+)[）)]/g) || [];
  parenMatches.forEach(pm => {
    const inner = pm
      .replace(/[（()）\s ]/g, '')
      .replace(/(?:先生|医師|ドクター|Dr\.?)$/gi, '');
    if (inner && /[一-龠々]/.test(inner) && !/氏|名|旧姓/.test(inner)) {
      keys.add(unifyKanjiVariantsForSokketsu_(inner));
    } else if (inner.includes('旧姓')) {
      const cleanedMaiden = inner.replace(/旧姓[:：]?/g, '');
      if (cleanedMaiden && /[一-龠々]/.test(cleanedMaiden)) {
        keys.add(unifyKanjiVariantsForSokketsu_(cleanedMaiden));
      }
    }
  });

  if (shokanText) {
    const aliasMatch = String(shokanText).match(/「([^」]{2,10}?)先生」/);
    if (aliasMatch) {
      const alias = normalizeDocKeyForSokketsu_(aliasMatch[1]);
      if (alias) keys.add(alias);
    }
  }
  return Array.from(keys);
}

function parseTimestampMsForSokketsu_(val) {
  if (!val) return 0;
  if (val instanceof Date && !isNaN(val.getTime())) return val.getTime();
  const s = String(val).trim().replace(/[年月]/g, '/').replace(/日/g, '');
  const d = new Date(s);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function formatSokketsuDate_(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'JST', 'yyyy/MM/dd');
  }
  const s = String(val).trim();
  const m = s.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if (m) {
    return `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
  }
  return s.split(' ')[0];
}