// =================================================================
// 【1-B3_M3キャンセル処理】
// ・Step 4: キャンセル・辞退通知（自動メール ＆ 事務局【お詫び】手動キャンセル完了メール）の自動処理
// ・48時間オファー期限切れメールによる誤「不採用」化をブロック保護
// ・A列空欄行・同一求人ID重複行の一括更新 ＆ 同一求人内の別医師保護
// =================================================================

/**
 * Step 4: エムスリーのキャンセル・辞退通知の自動処理
 */
function processM3_Step4_Cancels_(ctx) {
  const { sheet, archiveSheet, confirmQuerySuffix, processedLabel } = ctx;
  const cancelQuery = `from:(career_spot@m3career.com) (subject:("勤務がキャンセルされました" OR "キャンセルいたしました" OR "キャンセルが申請されました" OR "辞退されました") OR (subject:"【お詫び】" "キャンセル処理をさせていただきました") OR "【キャンセル処理済】") ${confirmQuerySuffix}`;
  const cancelThreads = GmailApp.search(cancelQuery);
  if (cancelThreads.length === 0) return;

  const cancelMessagesAll = GmailApp.getMessagesForThreads(cancelThreads);
  const archSaiyoColIndex = getColIndex_internal(archiveSheet, '採用可否');
  const archStatusColIndex = getColIndex_internal(archiveSheet, '処理状況');

  for (let i = 0; i < cancelThreads.length; i++) {
    const thread = cancelThreads[i];
    const messages = cancelMessagesAll[i];
    let threadProcessed = false;

    messages.forEach(message => {
      const messageId = message.getId();
      if (ctx.hasProcessed(messageId)) return;
      ctx.markProcessed(messageId);

      const fullBody = message.getPlainBody();
      const subject = message.getSubject();

      // ★重要保護：48時間オファー期限切れメールは、先に確定シフト登録済みのケースが多いため絶対に「不採用」にしない
      if (fullBody.includes('オファーがキャンセルされました') || fullBody.includes('48時間以内に「確定」または「お断り」ボタンがクリックされなかった')) {
        console.log(`[M3オファー期限切れ保護] 件名「${subject}」は48時間オファー切れのため不採用化をスキップしました。`);
        return;
      }

      const targetCancels = extractM3CancelTargetsFromMail_(subject, fullBody);
      if (targetCancels.length === 0) return;

      const archiveData = archiveSheet.getDataRange().getValues();

      targetCancels.forEach(({ jobId: mailJobId, docHint }) => {
        // ① 「処理済み」シート内の全該当行を更新（A列が空欄でもC番号一致なら対象・重複2行とも更新）
        for (let j = archiveData.length - 1; j > 0; j--) {
          const rowAgency = String(archiveData[j][0]).trim();
          const rowJobId = String(archiveData[j][5]).trim().replace('Ｃ', 'C');
          const rowDocClean = String(archiveData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');

          if ((rowAgency === 'エムスリー' || rowAgency === '') && rowJobId === mailJobId) {
            if (docHint && rowDocClean && !rowDocClean.includes(docHint) && !docHint.includes(rowDocClean)) {
              continue;
            }
            if (archSaiyoColIndex > 0) {
              archiveSheet.getRange(j + 1, archSaiyoColIndex).setValue('不採用').setBackground('#ffe599');
              archiveData[j][archSaiyoColIndex - 1] = '不採用';
            }
            if (archStatusColIndex > 0) {
              archiveSheet.getRange(j + 1, archStatusColIndex).setValue('キャンセル済み');
            }
            threadProcessed = true;
          }
        }

        // ② 「紹介会社進捗」シートにも残っていれば不採用としてアーカイブへ移動
        const sheetData = sheet.getDataRange().getValues();
        for (let j = sheetData.length - 1; j > 0; j--) {
          const rowAgency = String(sheetData[j][0]).trim();
          const rowJobId = String(sheetData[j][5]).trim().replace('Ｃ', 'C');
          const rowDocClean = String(sheetData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');

          if ((rowAgency === 'エムスリー' || rowAgency === '') && rowJobId === mailJobId) {
            if (docHint && rowDocClean && !rowDocClean.includes(docHint) && !docHint.includes(rowDocClean)) {
              continue;
            }
            const values = sheet.getRange(j + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
            values[8] = 'お断り';
            values[9] = 'キャンセル済み';
            values[11] = new Date();
            appendRowToArchive(sheet, archiveSheet, values, '不採用');
            safeDeleteRow(sheet, j + 1);
            threadProcessed = true;
          }
        }
      });
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }
}

/**
 * エムスリーのキャンセル系メール件名・本文から { jobId, docHint } の配列を抽出する内部ヘルパー
 */
function extractM3CancelTargetsFromMail_(subject, fullBody) {
  // 返信スレッド内の過去引用にあるC番号を誤爆しないよう、新規本文のみを切り出す
  const cleanTopBody = fullBody.split(/\n(>|On .*> wrote:|.*<.*@.*>.*:|From:\s|Sent:\s|-{5,}|━{5,})/i)[0];
  const targetCancels = [];
  const isAutoCancelSubj = /勤務がキャンセルされました|キャンセルいたしました|キャンセルが申請されました|辞退されました/.test(subject);

  if (isAutoCancelSubj) {
    const idMatch = subject.match(/([CＣ]\d{8,})/) || cleanTopBody.match(/([CＣ]\d{8,})/);
    if (idMatch) {
      const docM = subject.match(/(?:C\d+\s+|\d+月\d+日(?:\(.\))?\s*)([^\s 「」]+(?:\s*[^\s 「」]+)*?)先生/);
      const docHint = docM ? docM[1].replace(/[\s ]+/g, '') : '';
      targetCancels.push({ jobId: idMatch[1].replace('Ｃ', 'C'), docHint });
    }
  } else if (subject.includes('【お詫び】') && cleanTopBody.includes('キャンセル処理をさせていただきました')) {
    const docM = subject.match(/「([^」]+?)先生」/) || cleanTopBody.match(/「([^」]+?)先生」/);
    const docHint = docM ? docM[1].replace(/[\s ]+/g, '') : '';
    const ids = (subject + '\n' + cleanTopBody).match(/[CＣ]\d{8,}/g) || [];
    [...new Set(ids.map(x => x.replace('Ｃ', 'C')))].forEach(id => targetCancels.push({ jobId: id, docHint }));
  } else if (cleanTopBody.includes('【キャンセル処理済】')) {
    const secMatch = cleanTopBody.match(/【キャンセル処理済】[\s\S]*?(?=\n\s*\n|\n・|$)/);
    if (secMatch) {
      const ids = secMatch[0].match(/[CＣ]\d{8,}/g) || [];
      [...new Set(ids.map(x => x.replace('Ｃ', 'C')))].forEach(id => targetCancels.push({ jobId: id, docHint: '' }));
    }
  }

  return targetCancels;
}