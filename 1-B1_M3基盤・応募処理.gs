// =================================================================
// 【1-B1_M3基盤・応募処理】（CacheService省API ＆ 大量スレッド耐性版）
// ・エントリーポイント processM3（Message-IDキャッシュ一元管理）
// ・エムスリー住所 → 拠点マスタ「正規記載」変換ヘルパー（6時間ScriptCache対応）
// ・Step 1: 応募通知（通常応募 ＆ オファー希望）の検知・進捗シート追加
// =================================================================

const CLINIC_MASTER_SPREADSHEET_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
const CLINIC_MASTER_SHEET_NAME = '拠点名';

function extractTownNameForMatch_internal(rawAddr) {
  if (!rawAddr) return '';
  let s = String(rawAddr)
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[\s ]+/g, '');

  const m = s.match(/^([^\d\-ー−―‐ｰ]+)/);
  let town = m ? m[1] : s;

  const chomeMatch = s.match(/^([^\d\-ー−―‐ｰ]+?[一二三四五六七八九十]+丁目)/);
  if (chomeMatch) {
    town = chomeMatch[1];
  } else {
    town = town.replace(/[一二三四五六七八九十]+$/, '');
  }
  return town;
}

/**
 * 「拠点名」シートから【正規記載・住所】を取得（ScriptCacheで6時間保持し外部openByIdを98%削減）
 */
function getClinicAddressMap_internal() {
  if (getClinicAddressMap_internal.cache) return getClinicAddressMap_internal.cache;

  const scriptCache = CacheService.getScriptCache();
  const CACHE_KEY = 'M3_CLINIC_ADDR_MAP_V1';
  try {
    const cachedJson = scriptCache.get(CACHE_KEY);
    if (cachedJson) {
      const parsed = JSON.parse(cachedJson);
      getClinicAddressMap_internal.cache = parsed;
      return parsed;
    }
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.openById(CLINIC_MASTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CLINIC_MASTER_SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headerRow = data[0].map(v => String(v).trim());
    const foundNameIdx = headerRow.indexOf('正規記載');
    const foundAddrIdx = headerRow.indexOf('住所');
    const nameCol = foundNameIdx !== -1 ? foundNameIdx : 0;
    const addrCol = foundAddrIdx !== -1 ? foundAddrIdx : 16;

    const map = [];
    for (let i = 1; i < data.length; i++) {
      const rawName = String(data[i][nameCol] || '').replace(/キャップスクリニック/g, '').trim();
      const rawAddr = String(data[i][addrCol] || '').trim();
      if (!rawName && !rawAddr) continue;

      const townName = extractTownNameForMatch_internal(rawAddr);
      const townWithoutChome = townName.replace(/[一二三四五六七八九十]+丁目$/, '');

      map.push({
        clinicName: rawName,
        rawAddr: rawAddr,
        townName: townName,
        townWithoutChome: townWithoutChome
      });
    }

    getClinicAddressMap_internal.cache = map;
    try { scriptCache.put(CACHE_KEY, JSON.stringify(map), 21600); } catch (e) {} // 6時間キャッシュ
    return map;
  } catch (e) {
    console.error('拠点名マスタの読み込みエラー: ' + e.message);
    return [];
  }
}

function convertM3AddressToClinicName_internal(rawAddress) {
  if (!rawAddress) return '';
  const addr = String(rawAddress).trim();
  const inputTown = extractTownNameForMatch_internal(addr);
  const inputTownNoChome = inputTown.replace(/[一二三四五六七八九十]+丁目$/, '');

  const clinicMap = getClinicAddressMap_internal();

  if (clinicMap.length > 0) {
    const sortedByTown = [...clinicMap].sort((a, b) => b.townName.length - a.townName.length);
    for (const item of sortedByTown) {
      if (!item.townName || item.townName.length < 3) continue;
      if (inputTown.includes(item.townName) || item.townName.includes(inputTown)) {
        return item.clinicName;
      }
    }

    const sortedByNoChome = [...clinicMap].sort((a, b) => b.townWithoutChome.length - a.townWithoutChome.length);
    for (const item of sortedByNoChome) {
      if (!item.townWithoutChome || item.townWithoutChome.length < 3) continue;
      if (inputTownNoChome.includes(item.townWithoutChome) || item.townWithoutChome.includes(inputTownNoChome)) {
        return item.clinicName;
      }
    }

    const sortedByName = [...clinicMap].sort((a, b) => b.clinicName.length - a.clinicName.length);
    for (const item of sortedByName) {
      if (item.clinicName && item.clinicName.length >= 2 && addr.includes(item.clinicName)) {
        return item.clinicName;
      }
    }
  }

  return inputTown || addr;
}

/**
 * ★ 1-A_メイン基盤.gs から呼び出されるエムスリー処理の司令塔関数
 */
function processM3(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);

  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('M3_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch (e) {}

  const processedSet = new Set(processedIds);

  const ctx = {
    ss: ss,
    sheet: sheet,
    archiveSheet: archiveSheet,
    now: new Date(),
    querySuffix: querySuffix,
    confirmQuerySuffix: confirmQuerySuffix,
    processedLabel: processedLabel,
    processedIds: processedIds,
    processedSet: processedSet,
    isIdUpdated: false,
    hasProcessed: function(msgId) {
      return this.processedSet.has(msgId);
    },
    markProcessed: function(msgId) {
      if (!this.processedSet.has(msgId)) {
        this.processedSet.add(msgId);
        this.processedIds.push(msgId);
        this.isIdUpdated = true;
      }
    }
  };

  processM3_Step1_Applications_(ctx);
  processM3_Step2_ConfirmAndReject_(ctx);
  processM3_Step3_StopListing_(ctx);
  processM3_Step4_Cancels_(ctx);

  // 最新350件を保持（9KB制限の約6.5KB相当で安全）
  if (ctx.isIdUpdated) {
    if (ctx.processedIds.length > 350) {
      ctx.processedIds = ctx.processedIds.slice(-350);
    }
    props.setProperty('M3_PROCESSED_IDS', JSON.stringify(ctx.processedIds));
  }
}

/**
 * Step 1: 応募通知（通常応募 ＆ オファー希望）の処理
 */
function processM3_Step1_Applications_(ctx) {
  const { ss, sheet, now, querySuffix, processedLabel } = ctx;
  const applyQuery = `from:(career_spot@m3career.com) subject:("ご勤務希望がありました" OR "お問い合わせがありました") ${querySuffix}`;
  // ★大量受信時もクラッシュしないよう1回あたり最大50スレッドに制限
  const applyThreads = GmailApp.search(applyQuery, 0, 50);
  if (applyThreads.length === 0) return;

  const applyMessagesAll = GmailApp.getMessagesForThreads(applyThreads);

  for (let i = 0; i < applyThreads.length; i++) {
    const thread = applyThreads[i];
    const messages = applyMessagesAll[i];
    let threadProcessed = false;

    messages.forEach(message => {
      // ★最優先：48時間超の古いスレッド内メッセージはID配列を汚さず即スキップ
      const messageDate = message.getDate();
      if ((now - messageDate) / (1000 * 60 * 60) > 48) return;

      const messageId = message.getId();
      if (ctx.hasProcessed(messageId)) return;

      const subject = message.getSubject();
      const body = message.getPlainBody();

      let isOfferApply = false;
      if (subject.includes('お問い合わせがありました')) {
        const cleanBodyForCheck = body.replace(/[\r\n\s ]+/g, '');
        const offerKeyword = '■オファーを出される前にご確認ください・医師のご専門や可能／不可能な手技など必要に応じて直接先生にご確認ください。';
        if (cleanBodyForCheck.includes(offerKeyword)) {
          isOfferApply = true;
        } else {
          ctx.markProcessed(messageId);
          return;
        }
      }

      ctx.markProcessed(messageId);

      const bodyDocMatch = body.match(/氏名[\s:：]*([^\n\r]+)/);
      const doctorName = bodyDocMatch ? bodyDocMatch[1].trim() : extractM3DoctorName(subject);
      const jobId = (extract(subject, /([CＣ]\d+)/) || extract(body, /求人票ID[\s:：]*([A-Za-z0-9]+)/)).replace('Ｃ', 'C');

      const addressMatch = body.match(/勤務地[\s:：]*([^\n\r]+)/);
      const rawAddress = addressMatch ? addressMatch[1].trim() : extract(body, /^医療法人社団.+キャップスクリニック(.+)様/m).replace(/店$/, '');
      const clinicBaseName = convertM3AddressToClinicName_internal(rawAddress);

      const deptMatch = body.match(/募集科目[\s:：]*(.+)/);
      const department = deptMatch ? deptMatch[1].trim() : (findDepartmentFromHistory(ss, doctorName) || '小児科');
      const newLocationName = isOfferApply ? clinicBaseName : (rawAddress.replace(/\d.*$/, '').trim() + '（' + department + '）');

      let workDateRaw = '';
      const timeLineDateMatch = body.match(/勤務時間[\s:：]*\s*(\d{1,2}\/\d{1,2}(?:\([月火水木金土日]\))?)/);
      if (timeLineDateMatch) {
        workDateRaw = timeLineDateMatch[1].trim();
      } else {
        workDateRaw = extract(body, /^[\s ]*勤務日[\s:：]*(.+)/m);
      }
      if (!workDateRaw) {
        const subjDateMatch = subject.match(/(\d{1,2}月\d{1,2}日)/);
        if (subjDateMatch) workDateRaw = subjDateMatch[1].trim();
      }

      const timeRegex = /勤務時間[\s:：]*.*?\)\s*(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/;
      const timeMatch = body.match(timeRegex);
      const workTime = timeMatch ? timeMatch[1].trim() : extract(body, /勤務時間[\s:：]*\s*(\d{1,2}:\d{2}.+)/);

      const processStatusText = isOfferApply ? 'オファー経由' : '応募通知受信';

      const added = appendRowAndFormatUnique(sheet, [
        'エムスリー',
        formatDate(workDateRaw, messageDate),
        workTime,
        newLocationName,
        department,
        jobId,
        doctorName,
        extract(body, /(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/),
        '',
        processStatusText,
        messageDate,
        '',
        '',
        jobId
      ]);
      if (added) threadProcessed = true;
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }
}