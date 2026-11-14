// =================================================================
// 【1-B2_M3確定・停止処理】（省API・検索最適化版）
// ・Step 2: 確定・お断り通知の処理（不要なオファー再検索を削減・プルダウンミス自己修復含む）
// ・Step 3: 求人票掲載停止通知の処理
// =================================================================

function processM3_Step2_ConfirmAndReject_(ctx) {
  const { sheet, archiveSheet, now, confirmQuerySuffix, processedLabel } = ctx;
  const confirmQuery = `from:(career_spot@m3career.com) subject:("勤務を確定" OR "勤務が確定" OR "勤務確定" OR "をお断り") ${confirmQuerySuffix}`;
  const confirmThreads = GmailApp.search(confirmQuery, 0, 50);
  if (confirmThreads.length === 0) return;

  const confirmMessagesAll = GmailApp.getMessagesForThreads(confirmThreads);
  const archSaiyoCol = getColIndex_internal(archiveSheet, '採用可否');
  const archStatusCol = getColIndex_internal(archiveSheet, '処理状況');

  for (let i = 0; i < confirmThreads.length; i++) {
    const thread = confirmThreads[i];
    const messages = confirmMessagesAll[i];
    let threadProcessed = false;

    messages.forEach(message => {
      // ★最優先：72時間超の過去メールはIDリストを汚さず即スキップ
      const messageDate = message.getDate();
      if ((now - messageDate) / (1000 * 60 * 60) > 72) return;

      const messageId = message.getId();
      if (ctx.hasProcessed(messageId)) return;
      ctx.markProcessed(messageId);

      const subject = message.getSubject();
      const body = message.getPlainBody();

      let type = '';
      let status = '';
      if (subject.includes('勤務を確定') || subject.includes('勤務が確定') || body.includes('勤務を確定いたしました') || subject.includes('勤務確定')) {
        type = '確定';
        status = '勤務確定';
      } else if (subject.includes('をお断り') || body.includes('をお断り')) {
        type = 'お断り';
        status = 'お断り済み';
      }
      if (!type) return;

      const jobIdMatch = body.match(/求人票.*?([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
      const mailJobId = (jobIdMatch ? jobIdMatch[1].trim() : extract(body, /([CＣ]\d+)/)).replace('Ｃ', 'C');
      if (!mailJobId) return;

      const dateMatch = body.match(/求人票\s*(\d{4}年\d{1,2}月\d{1,2}日)/) || subject.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
      let mailDateStr = '';
      if (dateMatch) {
        const d = new Date(dateMatch[1].replace(/年|月/g, '/').replace('日', ''));
        mailDateStr = Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
      }

      let mailDoctorName = '';
      const nameMatch = body.match(/医師氏名:\s*([^\(（\n]+)/) || body.match(/(.+?)先生の勤務を確定/) || subject.match(/([^\s ]+?)先生/);
      if (nameMatch) {
        mailDoctorName = nameMatch[1].trim().replace(/[\s ]+/g, '');
      } else {
        mailDoctorName = extractM3DoctorName(subject).replace(/[\s ]+/g, '');
      }

      // 先に進捗シート上の該当行を探す（通常応募で情報揃い済みなら無駄なGmailオファー検索を省くため）
      let rowNum = -1;
      let progressRowStatus = '';
      const sheetData = sheet.getDataRange().getValues();
      for (let j = sheetData.length - 1; j > 0; j--) {
        const rowAgency = String(sheetData[j][0]).trim();
        if (rowAgency !== 'エムスリー' && rowAgency !== '') continue;

        const rowStatus = String(sheetData[j][9]).trim();
        if (!rowStatus.includes('応募通知受信') && !rowStatus.includes('オファー経由')) continue;

        const rowJobId = String(sheetData[j][5]).trim().replace('Ｃ', 'C');
        const rowDocClean = String(sheetData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');

        if (rowJobId === mailJobId && (!mailDoctorName || !rowDocClean || rowDocClean === mailDoctorName)) {
          rowNum = j + 1;
          progressRowStatus = rowStatus;
          break;
        }
      }

      let fetchedClinic = '';
      let fetchedWorkTime = '';
      let doctorMessage = '特になし';
      let isOffer = false;

      // ★省API化：進捗シート上で「オファー経由」の場合、または進捗シートに行が無い確定メールの場合のみオファー元メールを検索
      if (type === '確定' && (rowNum === -1 || progressRowStatus.includes('オファー経由'))) {
        const offerQuery = `from:career_spot@m3career.com subject:"お問い合わせがありました" "${mailJobId}"`;
        const offerThreads = GmailApp.search(offerQuery, 0, 1);

        if (offerThreads.length > 0) {
          isOffer = true;
          const offerBody = offerThreads[0].getMessages()[0].getPlainBody();

          const clinicMatch = offerBody.match(/勤務地[\s:：]*([^\n\r]+)/);
          if (clinicMatch) {
            fetchedClinic = convertM3AddressToClinicName_internal(clinicMatch[1].trim());
          }
          const timeMatch = offerBody.match(/勤務時間[\s:：]*[^\n\r]*?(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/);
          if (timeMatch) fetchedWorkTime = timeMatch[1].trim();

          const msgMatch = body.match(/その他連絡事項:\s*([\s\S]*?)(?=また、勤務日が近づかれましたら|※ご勤務確定後に|━━━━━━━━━━━━)/);
          if (msgMatch) doctorMessage = msgMatch[1].trim();
        }
      }

      let isNewlyProcessed = false;

      if (rowNum !== -1) {
        const values = sheetData[rowNum - 1].slice();
        if (!values[2] && fetchedWorkTime) values[2] = fetchedWorkTime;
        if (!values[3] && fetchedClinic) values[3] = fetchedClinic;
        if (!fetchedClinic && values[3]) fetchedClinic = String(values[3]);
        values[8] = type;
        values[9] = status;
        values[11] = new Date();

        appendRowToArchive(sheet, archiveSheet, values, type === '確定' ? '採用' : '不採用');
        safeDeleteRow(sheet, rowNum);

        isNewlyProcessed = true;
        threadProcessed = true;
      } else if (type === '確定') {
        const archiveData = archiveSheet.getDataRange().getValues();
        let isExistInArchive = false;

        for (let j = archiveData.length - 1; j > 0; j--) {
          const rowJobId = String(archiveData[j][5]).trim().replace('Ｃ', 'C');
          const rowDocClean = String(archiveData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');
          if (rowJobId === mailJobId && (!mailDoctorName || rowDocClean === mailDoctorName)) {
            isExistInArchive = true;
            break;
          }
        }

        if (!isExistInArchive) {
          const m3ThreadUrl = extract(body, /(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/);
          const displayDocName = nameMatch ? nameMatch[1].trim() : mailDoctorName;
          const tempValues = [
            'エムスリー', mailDateStr, fetchedWorkTime, fetchedClinic, '',
            mailJobId, displayDocName, m3ThreadUrl,
            '確定', status, new Date(), '', '', mailJobId
          ];
          appendRowToArchive(null, archiveSheet, tempValues, '採用');

          isNewlyProcessed = true;
          threadProcessed = true;
        }
      } else if (type === 'お断り' && mailDoctorName) {
        const archiveData = archiveSheet.getDataRange().getValues();
        for (let j = archiveData.length - 1; j > 0; j--) {
          const rowJobId = String(archiveData[j][5]).trim().replace('Ｃ', 'C');
          const rowDocClean = String(archiveData[j][6] || '').split(/[\/／（(]/)[0].replace(/[\s ]+/g, '');
          if (rowJobId === mailJobId && rowDocClean === mailDoctorName) {
            if (archSaiyoCol > 0) {
              archiveSheet.getRange(j + 1, archSaiyoCol).setValue('不採用').setBackground('#ffe599');
            }
            if (archStatusCol > 0) {
              archiveSheet.getRange(j + 1, archStatusCol).setValue('お断り済み');
            }
            threadProcessed = true;
          }
        }
      }

      if (isOffer && isNewlyProcessed) {
        let displayDate = mailDateStr;
        const dMatch = subject.match(/(\d{1,2}月\d{1,2}日.*?\))/);
        if (dMatch) {
          displayDate = String(messageDate.getFullYear()) + "年" + dMatch[1].replace('(', '日(');
        } else if (mailDateStr) {
          const d = new Date(mailDateStr);
          const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
          displayDate = Utilities.formatDate(d, 'JST', 'yyyy年MM月dd日') + `(${dayMap[d.getDay()]})`;
        }

        const displayDocName = nameMatch ? nameMatch[1].trim() : mailDoctorName;
        const cwMessage = `[info][title]（オファー経由）勤務確定[/title]
日時：${displayDate}
識別番号：${mailJobId}
拠点：${fetchedClinic}
医師名：${displayDocName}

医師メッセージ：【 ${doctorMessage} 】

受信時刻：${Utilities.formatDate(messageDate, 'JST', 'yyyy/MM/dd HH:mm')}
[/info]`;

        sendToChatwork('419888887', cwMessage);
        Utilities.sleep(1000);
      }
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }
}

function processM3_Step3_StopListing_(ctx) {
  const { sheet, archiveSheet, now, confirmQuerySuffix, processedLabel } = ctx;
  const stopQuery = `from:(career_spot@m3career.com) subject:("スポット求人票掲載停止のご案内") ${confirmQuerySuffix}`;
  const stopThreads = GmailApp.search(stopQuery, 0, 50);
  if (stopThreads.length === 0) return;

  const stopMessagesAll = GmailApp.getMessagesForThreads(stopThreads);

  for (let i = 0; i < stopThreads.length; i++) {
    const thread = stopThreads[i];
    const messages = stopMessagesAll[i];
    let threadProcessed = false;

    messages.forEach(message => {
      if ((now - message.getDate()) / (1000 * 60 * 60) > 72) return;

      const messageId = message.getId();
      if (ctx.hasProcessed(messageId)) return;
      ctx.markProcessed(messageId);

      const body = message.getPlainBody();
      const subject = message.getSubject();
      const jobIdMatch = body.match(/求人票.*?([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
      const mailJobId = (jobIdMatch ? jobIdMatch[1].trim() : extract(body, /([CＣ]\d+)/)).replace('Ｃ', 'C');
      if (!mailJobId) return;

      const sheetData = sheet.getDataRange().getValues();
      for (let j = sheetData.length - 1; j > 0; j--) {
        const rowAgency = String(sheetData[j][0]).trim();
        const rowJobId = String(sheetData[j][5]).trim().replace('Ｃ', 'C');
        const rowStatus = String(sheetData[j][9]).trim();

        if ((rowAgency === 'エムスリー' || rowAgency === '') && (rowStatus.includes('応募通知受信') || rowStatus.includes('オファー経由')) && rowJobId === mailJobId) {
          const values = sheetData[j].slice();
          values[8] = 'お断り';
          values[9] = '掲載停止';
          values[11] = new Date();
          appendRowToArchive(sheet, archiveSheet, values, '不採用');
          safeDeleteRow(sheet, j + 1);
          threadProcessed = true;
        }
      }
    });

    if (threadProcessed) {
      thread.addLabel(processedLabel);
    }
  }
}