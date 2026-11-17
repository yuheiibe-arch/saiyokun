// ============================================================================
// 【4-C_即決アーカイブ_遡及ツール】（手動メンテナンス専用）
// ・日常稼働コード（4-A / 4-B）と分離し、保守性を高めています
// ・過去データの「今後の採用判断」遡及入力 ＆ 採用・DS承認照合アーカイブ一括構築
// ============================================================================

/**
 * 🚀【初回・メンテ時に手動実行】
 * 1) 過去の「＜不採用報告＞」シートの「今後の採用判断」列(22列目)へ遡及入力
 * 2) 〈採用報告〉・＜DS承認報告＞の全履歴と突き合わせ、「その後の新規登録有無」を判定
 * 3) 『次回即決アーカイブ』シートを全行最新構造で一括生成（通知送信は一切なし）
 */
function runRetroactiveNextSokketsuMigration_once() {
  console.log('=== 🚀 不採用報告「次回即決」一括遡及 ＆ 採用/DS承認照合アーカイブ作成 開始 ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheets().find(sh => /不採用報告/.test(sh.getName()));
  if (!formSheet) throw new Error('＜不採用報告＞シートが見つかりません。');

  // 〈採用報告〉・＜DS承認報告＞から医師名ごとの登録履歴辞書を構築 [4-A]
  const regIndexMap = buildAdoptedAndDsLookupMap_(ss);

  // アーカイブシートを最新16列ヘッダーで準備 [4-A]
  const archiveSheet = resetAndEnsureNextSokketsuSheet_(ss);

  const data = formSheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const colNextEval = headers.findIndex(h => h.trim() === '今後の採用判断') + 1;
  if (colNextEval <= 0) throw new Error('「今後の採用判断」列が見つかりません。');

  // 🚫 個別除外リスト（カナ・括弧・スペース除去後の漢字氏名）
  const EXCLUDED_DOCTOR_NAMES = new Set([
    '松谷泰成', // 指定除外（行176）
    '加藤豊',   // 指定除外（行170）
    '嵐山恭志', // 71歳・年齢基準超過（行16）
    '香西光',   // 今後も採用困難記載（行129）
    '山口邦彦', // 今後も不採用・採用見送り記載（行164, 165）
    '奥野結衣'  // 2026年度時点で後期研修2年目・基準未達（行162）
  ]);

  const strongPosRegex = /即決|次回.{0,12}(採用|お願い|さ異様)|採用(できる|できそう|したい|可|して問題ない)|さ異様できる|経歴.{0,8}問題(ない|なし)|実績.{0,6}問題(ない|なし)/i;
  const explicitNgRegex = /今後のご応募につきましても(不採用|採用見送り)|今後も採用が難しい|今後のご案内を控える|医籍登録を確認できず/;

  let updatedSheetCount = 0;
  let countStar = 0;
  let countMaru = 0;
  let countRegisteredAfter = 0;
  let countUnregistered = 0;
  const rowsToInsert = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const rowNum = r + 1;

    const timestampRaw = row[0];
    const area         = row[1];
    const clinic       = row[2];
    const workDateRaw  = row[3];
    const agency       = row[6];
    const doctorName   = row[7];
    const gender       = row[8];
    const age          = row[9];
    const deptRaw      = row[10];
    const specialty    = row[15];
    const shokanRaw    = row[18];
    const reasonRaw    = row[19];
    const resumeUrl    = row[20];
    const curNextEval  = String(row[colNextEval - 1] || '').trim();

    if (!doctorName && !timestampRaw) continue;

    const normDocName = normalizeDocKeyForSokketsu_(doctorName);
    const cleanShokan = stripChatworkQuotes_(shokanRaw);
    const combinedEvalText = `${cleanShokan} ${reasonRaw}`;

    // 1. 除外対象または明確な継続不採用理由の判定
    const isHardNgReason = /基準に満たない|クレーム歴|不採用歴|クチコミ|紹介見送り/.test(reasonRaw);
    const isExcludedDoc = EXCLUDED_DOCTOR_NAMES.has(normDocName);
    const isExplicitNg = explicitNgRegex.test(combinedEvalText);

    if (isExcludedDoc || isHardNgReason || isExplicitNg) {
      if (!curNextEval) {
        formSheet.getRange(rowNum, colNextEval).setValue('今後も含め継続不採用');
        updatedSheetCount++;
      }
      continue;
    }

    // 2. 掘り起こし対象の判定
    const isEligibleReason = reasonRaw.includes('充足') || reasonRaw.includes('タイムアウト') || reasonRaw.includes('通常の枠にご応募いただければ採用できる');
    const hasStrongPos = strongPosRegex.test(combinedEvalText);

    if (!isEligibleReason && !hasStrongPos) {
      if (!curNextEval) {
        formSheet.getRange(rowNum, colNextEval).setValue('今後も含め継続不採用');
        updatedSheetCount++;
      }
      continue;
    }

    // 3. ランク分け
    const needsReCheck = /再度要選考|精査の上|再度精査|詳細未確認|確認が取れませんでした|実践の確認は必要/.test(cleanShokan);
    const rankLabel = (hasStrongPos && !needsReCheck) ? '★次回即決' : '〇充足（要確認）';

    if (!curNextEval) {
      formSheet.getRange(rowNum, colNextEval).setValue('次回即決');
      updatedSheetCount++;
    }

    // 4. 〈採用報告〉・＜DS承認報告＞との照合
    const regResult = checkSubsequentRegistration_(doctorName, shokanRaw, timestampRaw, regIndexMap);
    if (regResult.isRegistered) countRegisteredAfter++;
    else countUnregistered++;

    const sourceKey = `ROW_${rowNum}_${normDocName}_${formatSokketsuDate_(workDateRaw)}`;

    rowsToInsert.push(buildSokketsuArchiveRow_({
      rankLabel,
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
      sourceKey
    }));

    if (rankLabel === '★次回即決') countStar++;
    else countMaru++;
  }

  if (rowsToInsert.length > 0) {
    appendSokketsuRowWithFormat_(archiveSheet, rowsToInsert);
  }

  console.log(`=== 🏁 完了！ 元シート更新: ${updatedSheetCount}件 | アーカイブ総数: ${rowsToInsert.length}件 ===`);
  console.log(`  内訳(ランク) -> ★次回即決: ${countStar}件 / 〇充足(要確認): ${countMaru}件`);
  console.log(`  内訳(その後) -> 未登録(声掛け対象): ${countUnregistered}件 / その後登録済: ${countRegisteredAfter}件`);
}