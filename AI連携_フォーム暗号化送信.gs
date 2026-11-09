/**
 * ============================================================================
 * 【saiyokun 追加用】AI連携_フォーム暗号化送信.gs
 * - 〈採用報告〉、DS承認報告（DS直接応募初回勤務）、お久しぶり勤務のフォーム受信時に
 *   暗号化JSONを capsdoctor365@caps365.jp へ自動送信する
 * - 初回1回だけ既存の「チェックリスト」台帳をcaps側へ一括暗号化送信する関数も搭載
 * ============================================================================
 */

function getSaiyoBridgeConfig_() {
  return {
    RECIPIENT_EMAIL: 'capsdoctor365@caps365.jp',
    SUBJECT_FORM_EVENT: '【AI連携暗号化JSON】医師採用・初回・お久しぶり報告パック',
    SUBJECT_CHECKLIST_SYNC: '【AI連携暗号化JSON】チェックリスト医師台帳パック',
    ATTACHMENT_FILENAME: 'checklist_master_pack.enc',
    DEFAULT_SECRET_KEY: 'CAPS_MNYS_AI_AGENT_SECRET_2026_KEY_#9876'
  };
}

/**
 * ★ 大司令塔.gs の onFormSubmit_Master(e) から呼び出される関数
 */
function sendEncryptedFormEventToCaps(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // 不採用報告は対象外（採用報告、DS承認、初回勤務、お久しぶり等を対象とする）
  if (sheetName.includes('不採用')) return;
  if (!/採用報告|DS承認|初回勤務|直接応募|お久しぶり|久しぶり/.test(sheetName)) return;

  const namedValues = e.namedValues || {};
  const getVal = (candidates) => {
    for (const key of candidates) {
      if (namedValues[key] && namedValues[key][0]) {
        return String(namedValues[key][0]).trim();
      }
    }
    // 部分一致フォールバック
    const keys = Object.keys(namedValues);
    for (const cand of candidates) {
      const hit = keys.find(k => k.includes(cand));
      if (hit && namedValues[hit] && namedValues[hit][0]) {
        return String(namedValues[hit][0]).trim();
      }
    }
    return '';
  };

  const rawDoctorName = getVal(['氏名', '医師名', 'お名前', '先生名']);
  if (!rawDoctorName) return;

  const cleanDoctorName = rawDoctorName.replace(/[(（].*?[)）]/g, '').replace(/先生$/g, '').trim();
  const licenseNo = getVal(['医籍番号', '医籍登録番号']);
  const deptRaw = getVal(['診療科', '専門医資格', '専門科目']);
  const dept = deptRaw.includes('内科') ? '内科' : '小児科';
  const agency = getVal(['紹介会社', '経路']) || (sheetName.includes('DS') || sheetName.includes('直接') ? '直接DS応募' : '');
  const workDateRaw = getVal(['勤務日', '初回勤務日', '出勤日']);
  const clinic = getVal(['拠点', '勤務拠点', 'クリニック名', '拠点名']);
  const workHistory = getVal(['当法人での勤務歴', '勤務歴']);

  let normWorkDate = '';
  if (workDateRaw) {
    const dObj = new Date(workDateRaw);
    if (!isNaN(dObj.getTime())) {
      normWorkDate = Utilities.formatDate(dObj, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else {
      normWorkDate = workDateRaw;
    }
  }

  // 新規 / 復帰 の自動分類
  let category = '新規';
  if (sheetName.includes('お久しぶり') || sheetName.includes('久しぶり')) {
    category = '復帰';
  } else if (workHistory && workHistory !== 'なし' && !workHistory.includes('勤務歴なし')) {
    category = '復帰';
  }

  const payload = {
    type: 'DOCTOR_FORM_EVENT',
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    formSheetName: sheetName,
    category: category,         // "新規" または "復帰"
    doctorName: cleanDoctorName,
    licenseNo: licenseNo,
    dept: dept,
    agency: agency,
    workDate: normWorkDate,
    clinic: clinic,
    workHistory: workHistory
  };

  const cfg = getSaiyoBridgeConfig_();
  const secretKey = PropertiesService.getScriptProperties().getProperty('AI_BRIDGE_SECRET_KEY') || cfg.DEFAULT_SECRET_KEY;
  const encryptedPkg = encryptFormPayload_(JSON.stringify(payload), secretKey);

  const subject = `${cfg.SUBJECT_FORM_EVENT} (${payload.generatedAt})`;
  const body = [
    '※このメールは saiyokun から週次・月次報告AIエージェントへ自動送信された暗号化データです。',
    `フォーム種別: ${sheetName} (判定: ${category})`,
    `生成日時: ${payload.generatedAt}`,
    '',
    '---BEGIN_ENCRYPTED_AI_PAYLOAD---',
    JSON.stringify(encryptedPkg),
    '---END_ENCRYPTED_AI_PAYLOAD---'
  ].join('\n');

  MailApp.sendEmail({
    to: cfg.RECIPIENT_EMAIL,
    subject: subject,
    body: body,
    noReply: true
  });

  console.log(`✅ [AI連携] フォーム回答（${sheetName} / ${category}）を暗号化して送信しました。`);
}

/**
 * 🚀 【初回・必要時に手動実行】既存の「チェックリスト」台帳とフォームシートの履歴を
 * 一括で暗号化して capsdoctor365@caps365.jp へ送信する関数
 */
function syncExistingChecklistMasterToCaps() {
  const cfg = getSaiyoBridgeConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = [];

  // 1. Editシートから「チェックリスト」のURLを取得して読み込む
  try {
    const editSheet = ss.getSheetByName(CONFIG_FOR_CHECKLIST.EDIT_SHEET_NAME);
    if (editSheet) {
      const searchValues = editSheet.getRange(1, 1, editSheet.getLastRow(), 2).getValues();
      const hit = searchValues.find(r => r[0] === CONFIG_FOR_CHECKLIST.CHECKLIST_KEYWORD);
      if (hit && hit[1]) {
        const clSs = SpreadsheetApp.openByUrl(hit[1]);
        const clSheet = clSs.getSheetByName(CONFIG_FOR_CHECKLIST.TARGET_SHEET_NAME);
        if (clSheet && clSheet.getLastRow() >= 2) {
          const vals = clSheet.getDataRange().getDisplayValues();
          // 1〜3行目からヘッダー位置を特定
          let hRow = vals[0];
          for (let r = 0; r < Math.min(3, vals.length); r++) {
            if (vals[r].some(c => String(c).includes('医籍番号') || String(c).includes('採用可否'))) {
              hRow = vals[r];
              break;
            }
          }
          const findIdx = (kw, fallback) => {
            const idx = hRow.findIndex(h => String(h).includes(kw));
            return idx !== -1 ? idx : fallback;
          };

          const iName = findIdx('氏名', 1);      // B列
          const iStatus = findIdx('採用可否', 2); // C列
          const iAgency = findIdx('紹介会社', 4); // E列
          const iDept = findIdx('診療科', 9);     // J列
          const iHist = findIdx('勤務歴', 11);    // L列
          const iFirstDate = findIdx('勤務日', 20); // U列（初回勤務日）
          const iClinic = findIdx('拠点', 21);    // V列
          const iLic = findIdx('医籍番号', -1);

          for (let r = 1; r < vals.length; r++) {
            const row = vals[r];
            const status = String(row[iStatus] || '').trim();
            if (status.includes('不採用')) continue;
            const docName = String(row[iName] || '').replace(/[(（].*?[)）]/g, '').replace(/\s+/g, '').trim();
            if (!docName) continue;

            const lic = iLic !== -1 ? String(row[iLic] || '').trim() : '';
            const firstDate = String(row[iFirstDate] || '').trim();
            const hist = String(row[iHist] || '').trim();
            const agency = String(row[iAgency] || '').trim();
            const dept = String(row[iDept] || '').trim();
            const clinic = String(row[iClinic] || '').trim();

            if (!firstDate && !lic) continue;

            records.push({
              src: 'CHECKLIST',
              n: docName,
              id: lic,
              fd: firstDate,
              ag: agency,
              dp: dept,
              c: clinic,
              cat: (hist && hist !== 'なし' && !hist.includes('勤務歴なし')) ? '復帰' : '新規'
            });
          }
        }
      }
    }
  } catch (e) {
    console.error('チェックリスト読込エラー: ' + e.message);
  }

  // 2. saiyokun内の各フォーム回答シート（お久しぶり等）もスキャンして追加
  ss.getSheets().forEach(sh => {
    const sName = sh.getName();
    if (sName.includes('不採用')) return;
    if (!/採用報告|DS承認|初回勤務|直接応募|お久しぶり|久しぶり/.test(sName)) return;
    if (sh.getLastRow() < 2) return;

    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(h => String(h || '').trim());
    const findIdx = (cands) => {
      for (const c of cands) {
        const idx = headers.findIndex(h => h.includes(c));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const iTs = 0;
    const iName = findIdx(['氏名', '医師名', 'お名前']);
    const iLic = findIdx(['医籍番号']);
    const iDate = findIdx(['勤務日', '初回勤務日', '出勤日']);
    const iAgency = findIdx(['紹介会社', '経路']);
    const iHist = findIdx(['勤務歴']);
    const iClinic = findIdx(['拠点', 'クリニック']);
    if (iName === -1) return;

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const docName = String(row[iName] || '').replace(/[(（].*?[)）]/g, '').replace(/先生$/g, '').replace(/\s+/g, '').trim();
      if (!docName) continue;
      const lic = iLic !== -1 ? String(row[iLic] || '').trim() : '';
      const wDate = iDate !== -1 ? String(row[iDate] || '').trim() : String(row[iTs] || '').split(' ')[0];
      const hist = iHist !== -1 ? String(row[iHist] || '').trim() : '';
      const cat = (/お久しぶり|久しぶり/.test(sName) || (hist && hist !== 'なし' && !hist.includes('勤務歴なし'))) ? '復帰' : '新規';

      records.push({
        src: sName,
        n: docName,
        id: lic,
        fd: wDate,
        ag: iAgency !== -1 ? String(row[iAgency] || '').trim() : '',
        c: iClinic !== -1 ? String(row[iClinic] || '').trim() : '',
        cat: cat
      });
    }
  });

  const payload = {
    type: 'CHECKLIST_MASTER_SYNC',
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    count: records.length,
    records: records
  };

  const secretKey = PropertiesService.getScriptProperties().getProperty('AI_BRIDGE_SECRET_KEY') || cfg.DEFAULT_SECRET_KEY;
  const encryptedPkg = encryptFormPayload_(JSON.stringify(payload), secretKey);
  const blob = Utilities.newBlob(JSON.stringify(encryptedPkg), 'application/octet-stream', cfg.ATTACHMENT_FILENAME);

  GmailApp.sendEmail(
    cfg.RECIPIENT_EMAIL,
    `${cfg.SUBJECT_CHECKLIST_SYNC} (${payload.generatedAt})`,
    `チェックリスト医師台帳・フォーム履歴の一括同期パック（総件数: ${records.length}件）`,
    { attachments: [blob] }
  );

  console.log(`✅ チェックリスト＆フォーム履歴（計 ${records.length} 件）を暗号化して送信しました！`);
}

function encryptFormPayload_(plainText, secretKey) {
  const textBlob = Utilities.newBlob(plainText, 'application/json', 'payload.json');
  const gzipBytes = Utilities.gzip(textBlob).getBytes();
  const ivStr = Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  const output = new Array(gzipBytes.length);
  let keyStream = [];
  let blockIndex = 0;
  for (let i = 0; i < gzipBytes.length; i++) {
    if (i % 32 === 0) {
      keyStream = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        `${secretKey}|${ivStr}|${blockIndex}`,
        Utilities.Charset.UTF_8
      );
      blockIndex++;
    }
    output[i] = (gzipBytes[i] ^ keyStream[i % 32]) & 0xff;
    if (output[i] > 127) output[i] -= 256;
  }

  const cipherBase64 = Utilities.base64Encode(output);
  const sigBytes = Utilities.computeHmacSha256Signature(ivStr + ':' + cipherBase64, secretKey);
  return { v: 1, iv: ivStr, data: cipherBase64, sig: Utilities.base64Encode(sigBytes) };
}