// ============================================================================
// 【3-C_キャンセル_一括修復ツール】
// ・手動実行専用：過去120日分のキャンセル漏れ＆列ズレ一括クリーンアップ関数
// ・日常稼働コード（3-A / 3-B）と分離し、保守性を高めています
// ============================================================================

/**
 * 🚀 【初回・メンテ時のみ手動実行】過去120日分のキャンセル漏れ＆列ズレ一括クリーンアップ関数
 * ※ Chatwork通知は一切飛ばさず、「処理済み」シート上の過去取りこぼし分を安全に「不採用」へ修復します
 */
function runRetroactiveCancelCleanup_once() {
  console.log("=== 🧹 過去120日分 キャンセル一括遡及クリーンアップ開始（通知なし） ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('処理済み') || ss.getSheetByName('アーカイブ');
  const progressSheet = typeof SHEET_NAME_PROGRESS !== 'undefined'
    ? ss.getSheetByName(SHEET_NAME_PROGRESS)
    : ss.getSheetByName('紹介会社進捗');
  if (!archiveSheet) return;

  const colMap = getArchiveColMap_internal_(archiveSheet);
  let fixedColShiftCount = 0;

  // 1. 旧スクリプトの列ズレバグ行（I列に「キャンセル済み」またはM列に「不採用」でO列が「採用」のままの行）を修復
  const aData = archiveSheet.getDataRange().getValues();
  for (let r = 1; r < aData.length; r++) {
    const colI = String(aData[r][8] || '').trim();
    const colM = String(aData[r][12] || '').trim();
    const colO = String(aData[r][colMap.colSaiyo - 1] || '').trim();

    if (colO === '採用' && (colI === 'キャンセル済み' || colM === '不採用')) {
      const rowNum = r + 1;
      archiveSheet.getRange(rowNum, colMap.colType).setValue('お断り');
      archiveSheet.getRange(rowNum, colMap.colStatus).setValue('キャンセル済み');
      if (colM === '不採用') archiveSheet.getRange(rowNum, colMap.colPerson).setValue('自動連携');
      archiveSheet.getRange(rowNum, colMap.colSaiyo).setValue('不採用');
      archiveSheet.getRange(rowNum, 1, 1, colMap.totalCols).setBackground('#ffe599');
      console.log(`  🔧 [列ズレ修復] 行${rowNum}: ${normalizeCancelDateStr_(aData[r][1])} ${aData[r][3]} ${aData[r][6]} -> 不採用`);
      fixedColShiftCount++;
    }
  }

  // 2. 過去120日分のドクターサイト「キャンセル申請承認」メールを一括反映
  const sysEvents = [];
  const sysThreads = GmailApp.search(`subject:"お問い合わせスタッフ送信" "のキャンセル申請が承認されました" newer_than:120d`, 0, 200);
  for (let t = 0; t < sysThreads.length; t += 50) {
    const msgsList = GmailApp.getMessagesForThreads(sysThreads.slice(t, t + 50));
    msgsList.forEach(msgs => msgs.forEach(msg => {
      const body = msg.getPlainBody();
      if (body.includes('のキャンセル申請が承認されました')) {
        const parsed = parseSystemCancelApprovalBody_(body);
        if (parsed) sysEvents.push(parsed);
      }
    }));
  }
  const sysUpdatedCount = applySystemCancelEventsToSheets_(archiveSheet, progressSheet, colMap, sysEvents);

  // 3. 過去120日分の紹介会社キャンセルメールを一括反映（※M3の48時間オファー切れは除外して保護）
  const canceledJobIds = new Map();
  const stripQuotes = (txt) => String(txt || '').split(/\n(>|On .*> wrote:|.*<.*@.*>.*:|From:\s|Sent:\s|-{5,}|━{5,})/i)[0];

  const agQueries = [
    `from:career_spot@m3career.com subject:("勤務がキャンセルされました" OR "キャンセルいたしました" OR "キャンセルが申請されました" OR "辞退されました") newer_than:120d`,
    `from:career_spot@m3career.com (subject:"【お詫び】" OR "【キャンセル処理済】") newer_than:120d`,
    `(subject:"取消のご連絡" OR subject:"【キャンセル】" OR subject:"応募取消し") newer_than:120d`,
    `from:ra@mstage-corp.jp subject:"キャンセル" newer_than:120d`
  ];

  agQueries.forEach(q => {
    const ths = GmailApp.search(q, 0, 150);
    for (let t = 0; t < ths.length; t += 50) {
      const msgsList = GmailApp.getMessagesForThreads(ths.slice(t, t + 50));
      msgsList.forEach(msgs => msgs.forEach(msg => {
        const subj = msg.getSubject();
        const fullBody = msg.getPlainBody();
        const cleanTopBody = stripQuotes(fullBody);
        const from = msg.getFrom().toLowerCase();

        if (from.includes('m3career.com')) {
          if (fullBody.includes('オファーがキャンセルされました') || fullBody.includes('48時間以内に「確定」または「お断り」ボタンがクリックされなかった')) {
            return;
          }
          // 1-B3 の抽出ヘルパーを共通利用
          const targets = typeof extractM3CancelTargetsFromMail_ === 'function'
            ? extractM3CancelTargetsFromMail_(subj, fullBody)
            : [];
          targets.forEach(item => canceledJobIds.set(item.jobId, `M3キャンセル(${subj})`));
        } else if (subj.includes('取消のご連絡')) {
          const m = cleanTopBody.match(/案件番号[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
          if (m) canceledJobIds.set(m[1].trim(), '民間医局取消連絡');
        } else if (subj.includes('MRT') && (subj.includes('【キャンセル】') || subj.includes('応募取消し'))) {
          const m = cleanTopBody.match(/ID\s*[:：]\s*(\d+)/i);
          if (m) canceledJobIds.set(m[1].trim(), 'MRTキャンセル');
        } else if (from.includes('mstage-corp.jp') && subj.includes('キャンセル')) {
          const m = cleanTopBody.match(/(?:求人番号|案件ID|求人ID)[\s ]*[:：][\s ]*([A-Za-z0-9\-]+)/);
          if (m) canceledJobIds.set(m[1].trim(), 'エムステージキャンセル');
        }
      }));
    }
  });

  let agUpdatedCount = 0;
  const freshArchiveData = archiveSheet.getDataRange().getValues();
  for (let r = 1; r < freshArchiveData.length; r++) {
    const curSaiyo = String(freshArchiveData[r][colMap.colSaiyo - 1] || '').trim();
    if (curSaiyo !== '採用') continue;
    const rowJobId = String(freshArchiveData[r][colMap.colId - 1] || '').trim().replace('Ｃ', 'C');

    if (rowJobId && canceledJobIds.has(rowJobId)) {
      const rowNum = r + 1;
      archiveSheet.getRange(rowNum, colMap.colType).setValue('お断り');
      archiveSheet.getRange(rowNum, colMap.colStatus).setValue('キャンセル済み');
      archiveSheet.getRange(rowNum, colMap.colSaiyo).setValue('不採用');
      archiveSheet.getRange(rowNum, 1, 1, colMap.totalCols).setBackground('#ffe599');
      console.log(`  🔧 [紹介会社キャンセル遡及反映] 行${rowNum}: ID=${rowJobId} (${freshArchiveData[r][6]}) -> 不採用 [${canceledJobIds.get(rowJobId)}]`);
      agUpdatedCount++;
    }
  }

  console.log(`=== 🏁 クリーンアップ完了: 列ズレ修復=${fixedColShiftCount}件 | システム承認反映=${sysUpdatedCount}件 | 紹介会社キャンセル反映=${agUpdatedCount}件 ===`);
}