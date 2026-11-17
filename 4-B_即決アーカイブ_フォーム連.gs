// ============================================================================
// 【4-B_即決アーカイブ_フォーム連動】（日常稼働用）
// ・大司令塔 onFormSubmit_Master(e) から呼び出されるエントリーポイント
// ・＜不採用報告＞で「今後の採用判断 ＝ 次回即決」ならアーカイブへ自動追記
// ・〈採用報告〉または＜DS承認報告＞が送信されたら、アーカイブ内の該当医師を「登録済」へ自動更新
// ============================================================================

/**
 * ★ 大司令塔.gs の onFormSubmit_Master(e) から呼び出されるリアルタイム連動関数
 */
function autoArchiveNextSokketsuOnFormSubmit(e) {
  if (!e || !e.range) return;
  const formSheet = e.range.getSheet();
  const sheetName = formSheet.getName();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nv = e.namedValues || {};
  const getVal = (keys) => {
    for (const k of keys) {
      if (nv[k] && nv[k][0]) return String(nv[k][0]).trim();
    }
    const allKeys = Object.keys(nv);
    for (const k of keys) {
      const hit = allKeys.find(ak => ak.includes(k));
      if (hit && nv[hit] && nv[hit][0]) return String(nv[hit][0]).trim();
    }
    return '';
  };

  // 1. 〈採用報告〉または＜DS承認報告＞が送信された場合 → アーカイブ側を「登録済」に自動消し込み
  if ((sheetName.includes('採用報告') && !sheetName.includes('不採用')) || sheetName.includes('DS承認報告')) {
    handleAdoptedOrDsFormSubmitForSokketsu_(ss, sheetName, getVal);
    return;
  }

  // 2. ＜不採用報告＞が送信された場合 → 「次回即決」ならアーカイブへ新規追加
  if (sheetName.includes('不採用報告')) {
    handleFusaiyoFormSubmitForSokketsu_(ss, e.range.getRow(), getVal);
  }
}

/**
 * 〈採用報告〉・＜DS承認報告＞受信時の既存アーカイブ消し込み処理
 */
function handleAdoptedOrDsFormSubmitForSokketsu_(ss, sheetName, getVal) {
  const targetName = typeof NEXT_SOKKETSU_SHEET_NAME !== 'undefined' ? NEXT_SOKKETSU_SHEET_NAME : '次回即決アーカイブ';
  const archiveSheet = ss.getSheetByName(targetName);
  if (!archiveSheet || archiveSheet.getLastRow() < 2) return;

  const rawDocName = getVal(['氏名', '医師名']);
  const candidateKeys = extractCandidateDocKeys_(rawDocName, '');
  if (candidateKeys.length === 0) return;

  const isDs = sheetName.includes('DS承認');
  const statusLabel = isDs ? '登録済（DS承認報告）' : '登録済（採用報告）';
  const tsDate = formatSokketsuDate_(getVal(['タイムスタンプ']) || new Date());
  const workDate = formatSokketsuDate_(getVal(['勤務日']));
  const clinic = getVal(['拠点']);
  const agency = isDs ? '直接DS応募' : getVal(['紹介会社']);

  const detailText = [tsDate ? `報告:${tsDate}` : '', workDate ? `勤務:${workDate}` : '', clinic, agency]
    .filter(Boolean)
    .join(' / ');

  const lastRow = archiveSheet.getLastRow();
  const data = archiveSheet.getRange(2, 1, lastRow - 1, SOKKETSU_ARCHIVE_HEADERS_.length).getValues();

  for (let i = 0; i < data.length; i++) {
    const rowKeys = extractCandidateDocKeys_(data[i][9], data[i][13]); // J列:氏名, N列:コメント抜粋
    const isHit = rowKeys.some(rk => candidateKeys.includes(rk));
    if (isHit) {
      const rowNum = i + 2;
      archiveSheet.getRange(rowNum, 2).setValue(statusLabel); // B列: その後の登録状況
      archiveSheet.getRange(rowNum, 3).setValue(detailText);  // C列: 登録詳細
      applySokketsuRowVisualStyle_(archiveSheet, rowNum, data[i][0], statusLabel);
      console.log(`🔄 [次回即決アーカイブ自動更新] 行${rowNum}（${data[i][9]}）を「${statusLabel}」に更新しました。`);
    }
  }
}

/**
 * ＜不採用報告＞受信時のアーカイブ新規追加処理
 */
function handleFusaiyoFormSubmitForSokketsu_(ss, rowNum, getVal) {
  const nextJudgment = getVal(['今後の採用判断']);
  if (nextJudgment !== '次回即決') return;

  const archiveSheet = ensureNextSokketsuSheet_(ss);
  const regIndexMap = buildAdoptedAndDsLookupMap_(ss);

  const timestampRaw = getVal(['タイムスタンプ']) || Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  const area = getVal(['エリア']);
  const clinic = getVal(['拠点']);
  const workDateRaw = getVal(['勤務日']);
  const agency = getVal(['紹介会社']);
  const doctorName = getVal(['氏名', '医師名']);
  const gender = getVal(['性別']);
  const age = getVal(['年齢']);
  const deptRaw = getVal(['応募先の診療科', '診療科']);
  const specialty = getVal(['専門科目']);
  const shokanRaw = getVal(['所感・申し送り事項', '所感']);
  const reasonRaw = getVal(['【不採用理由】', '不採用理由']);
  const resumeUrl = getVal(['履歴書']);

  const regResult = checkSubsequentRegistration_(doctorName, shokanRaw, timestampRaw, regIndexMap);

  const record = buildSokketsuArchiveRow_({
    rankLabel: '★次回即決',
    regStatus: regResult.status,
    regDetail: regResult.detail,
    timestampRaw,
    workDateRaw,
    area,
    clinic,
    deptRaw,
    agency,
    doctorName,
    age,
    gender,
    specialty,
    reasonRaw,
    shokanRaw,
    resumeUrl,
    sourceKey: `ROW_${rowNum}_${normalizeDocKeyForSokketsu_(doctorName)}_${formatSokketsuDate_(workDateRaw)}`
  });

  appendSokketsuRowWithFormat_(archiveSheet, [record]);
  console.log(`✅ [次回即決アーカイブ] ${doctorName} 先生を追加しました（登録状況: ${regResult.status}）。`);
}