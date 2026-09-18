// =================================================================
// 【1-B_エムスリー処理】
// エムスリーの応募・確定（オファー経由対応）・掲載停止・キャンセル処理
// ★修正内容：API制限対策として getMessagesForThreads による一括取得に変更
// ★修正内容：スレッド巻き込みによるすり抜けを防止するためラベル除外条件を撤廃
// =================================================================

function processM3(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const processedLabelName = processedLabel.getName();
  
  // 現在時刻（24時間前判定用）
  const now = new Date();

  // ===============================================================
  // === 1. 応募通知の処理 ===
  // ===============================================================
  // ★ ラベル除外を撤廃（既存案件の二重処理はスクリプト内のロジックで防止済み）
  const applyQuery = `from:(career_spot@m3career.com) subject:("ご勤務希望がありました") ${querySuffix}`;
  const applyThreads = GmailApp.search(applyQuery);
  
  if (applyThreads.length > 0) {
    // ★ API通信1回で全メッセージを一括取得
    const applyMessagesAll = GmailApp.getMessagesForThreads(applyThreads);
    
    for (let i = 0; i < applyThreads.length; i++) {
      const thread = applyThreads[i];
      const messages = applyMessagesAll[i];
      let threadProcessed = false;
      
      messages.forEach(message => {
        // ★ 古すぎるメールのストッパー
        const messageDate = message.getDate();
        if ((now - messageDate) / (1000 * 60 * 60) > 48) return; 

        const subject = message.getSubject();
        const body = message.getPlainBody();
        
        const doctorName = extractM3DoctorName(subject);
        const jobId = extract(subject, /(C\d+)/) || extract(body, /求人票ID：\s*([A-Za-z0-9]+)/);
        
        const addressMatch = body.match(/勤務地：([^\d\n]+)/);
        const address = addressMatch ? addressMatch[1].trim() : extract(body, /^医療法人社団.+キャップスクリニック(.+)様/m).replace(/店$/, '');
        const deptMatch = body.match(/募集科目：\s*(.+)/);
        let department = deptMatch ? deptMatch[1].trim() : (findDepartmentFromHistory(ss, doctorName) || '小児科');
        const newLocationName = address + '（' + department + '）';
        
        let workDateRaw = extract(body, /勤務日：(.+)/);
        if (!workDateRaw) {
          const dateMatch = body.match(/勤務時間：\s*(.*?)(?=\d{1,2}:\d{2})/);
          if (dateMatch) workDateRaw = dateMatch[1].trim();
        }
        const timeRegex = /勤務時間：.*?\)\s*(\d{1,2}:\d{2}\s*[〜~〜\-]\s*\d{1,2}:\d{2})/;
        const timeMatch = body.match(timeRegex);
        const workTime = timeMatch ? timeMatch[1].trim() : extract(body, /勤務時間：\s*(\d{1,2}:\d{2}.+)/);

        const added = appendRowAndFormatUnique(sheet, [
          'エムスリー', formatDate(workDateRaw, messageDate), workTime, newLocationName, department, 
          jobId, 
          doctorName, extract(body, /(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/), 
          '', '応募通知受信', messageDate, '', '', jobId 
        ]);
        if (added) threadProcessed = true;
      });
      
      if (threadProcessed) {
          thread.addLabel(processedLabel);
      }
    }
  }

  // ===============================================================
  // === 2. 確定・お断り通知の処理 ===
  // ===============================================================
  const confirmQuery = `from:(career_spot@m3career.com) subject:("勤務を確定" OR "勤務が確定" OR "勤務確定" OR "をお断り") ${confirmQuerySuffix}`;
  const confirmThreads = GmailApp.search(confirmQuery);
  
  if (confirmThreads.length > 0) {
    const confirmMessagesAll = GmailApp.getMessagesForThreads(confirmThreads);
    
    for (let i = 0; i < confirmThreads.length; i++) {
      const thread = confirmThreads[i];
      const messages = confirmMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const messageDate = message.getDate();
        if ((now - messageDate) / (1000 * 60 * 60) > 14 * 24) return; // 14日ストッパー

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
        const mailJobId = jobIdMatch ? jobIdMatch[1].trim() : extract(body, /([CＣ]\d+)/);
        
        if (!mailJobId) return; 

        const dateMatch = body.match(/求人票\s*(\d{4}年\d{1,2}月\d{1,2}日)/) || subject.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
        let mailDateStr = '';
        if (dateMatch) {
          const d = new Date(dateMatch[1].replace(/年|月/g, '/').replace('日', ''));
          mailDateStr = Utilities.formatDate(d, 'JST', 'yyyy/MM/dd');
        }
        
        let mailDoctorName = '';
        const nameMatch = body.match(/医師氏名:\s*([^\(（\n]+)/) || body.match(/(.+?)先生の勤務を確定/);
        if (nameMatch) {
          mailDoctorName = nameMatch[1].trim().replace(/\s+/g, '');
        } else {
          mailDoctorName = extractM3DoctorName(subject).replace(/\s+/g, '');
        }

        let fetchedClinic = '';
        let fetchedWorkTime = '';
        let doctorMessage = '特になし';
        let isOffer = false;

        if (type === '確定') {
            const offerQuery = `from:career_spot@m3career.com subject:"お問い合わせがありました" "${mailJobId}"`;
            const offerThreads = GmailApp.search(offerQuery, 0, 1);
            
            if (offerThreads.length > 0) {
                isOffer = true;
                const offerBody = offerThreads[0].getMessages()[0].getPlainBody();
                
                const clinicMatch = offerBody.match(/勤務地[\s:：]*([^\n\r]+)/);
                if (clinicMatch) {
                    fetchedClinic = clinicMatch[1].trim();
                    const shortClinicMatch = fetchedClinic.match(/^([^\d０-９]+)/);
                    if (shortClinicMatch) fetchedClinic = shortClinicMatch[1].trim();
                }

                const timeMatch = offerBody.match(/勤務時間[\s:：]*[^\n\r]*?(\d{1,2}:\d{2}\s*[〜~〜\-]\s*\d{1,2}:\d{2})/);
                if (timeMatch) fetchedWorkTime = timeMatch[1].trim();

                const msgMatch = body.match(/その他連絡事項:\s*([\s\S]*?)(?=また、勤務日が近づかれましたら|※ご勤務確定後に|━━━━━━━━━━━━)/);
                if (msgMatch) doctorMessage = msgMatch[1].trim();
            }
        }

        let rowNum = -1;
        const sheetData = sheet.getDataRange().getValues();
        for (let j = sheetData.length - 1; j > 0; j--) {
            const rowAgency = String(sheetData[j][0]).trim();
            if (rowAgency !== 'エムスリー') continue;
            
            const rowStatus = String(sheetData[j][9]).trim();
            if (!rowStatus.includes('応募通知受信')) continue;

            const rowJobId = String(sheetData[j][5]).trim();

            if (rowJobId === mailJobId) {
                rowNum = j + 1;
                break;
            }
        }
        
        let isNewlyProcessed = false;

        if (rowNum !== -1) {
            let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
            if (!values[2] && fetchedWorkTime) values[2] = fetchedWorkTime;
            if (!values[3] && fetchedClinic) values[3] = fetchedClinic;

            values[8] = type; values[9] = status; values[11] = new Date();
            
            appendRowToArchive(sheet, archiveSheet, values, type === '確定' ? '採用' : '不採用');
            safeDeleteRow(sheet, rowNum);
            
            isNewlyProcessed = true;
            threadProcessed = true;
        } else if (type === '確定') {
            const archiveData = archiveSheet.getDataRange().getValues();
            let isExistInArchive = false;

            for (let j = archiveData.length - 1; j > 0; j--) {
              const rowAgency = String(archiveData[j][0]).trim();
              if (rowAgency !== 'エムスリー') continue;
              const rowJobId = String(archiveData[j][5]).trim();

              if (rowJobId === mailJobId) {
                isExistInArchive = true;
                break;
              }
            }

            if (!isExistInArchive) {
              const m3ThreadUrl = extract(body, /(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/);
              const uniqueId = mailJobId;
              const displayDocName = nameMatch ? nameMatch[1].trim() : mailDoctorName;
              
              const tempValues = [
                'エムスリー', mailDateStr, fetchedWorkTime, fetchedClinic, '', 
                mailJobId, 
                displayDocName, m3ThreadUrl, 
                '確定', status, new Date(), '', '', uniqueId
              ];
              
              appendRowToArchive(null, archiveSheet, tempValues, '採用');
              
              isNewlyProcessed = true;
              threadProcessed = true;
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
        }
      });
      
      if (threadProcessed) {
          thread.addLabel(processedLabel); 
      }
    }
  }
  
  // ===============================================================
  // === 3. 掲載停止通知の処理 ===
  // ===============================================================
  const stopQuery = `from:(career_spot@m3career.com) subject:("スポット求人票掲載停止のご案内") ${confirmQuerySuffix}`;
  const stopThreads = GmailApp.search(stopQuery);

  if (stopThreads.length > 0) {
    const stopMessagesAll = GmailApp.getMessagesForThreads(stopThreads);

    for (let i = 0; i < stopThreads.length; i++) {
      const thread = stopThreads[i];
      const messages = stopMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
        const body = message.getPlainBody();
        const subject = message.getSubject();
        const jobIdMatch = body.match(/求人票.*?([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
        const mailJobId = jobIdMatch ? jobIdMatch[1].trim() : extract(body, /([CＣ]\d+)/);

        if (!mailJobId) return;

        let rowNum = -1;
        const sheetData = sheet.getDataRange().getValues();
        for (let j = sheetData.length - 1; j > 0; j--) {
            const rowAgency = String(sheetData[j][0]).trim();
            const rowJobId = String(sheetData[j][5]).trim();
            const rowStatus = String(sheetData[j][9]).trim();
            if (rowAgency === 'エムスリー' && rowStatus.includes('応募通知受信') && rowJobId === mailJobId) {
                rowNum = j + 1;
                break;
            }
        }

        if (rowNum !== -1) {
            let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
            values[8] = 'お断り'; values[9] = '掲載停止'; values[11] = new Date();
            
            appendRowToArchive(sheet, archiveSheet, values, '不採用');
            safeDeleteRow(sheet, rowNum);
            
            threadProcessed = true;
        }
      });

      if (threadProcessed) {
          thread.addLabel(processedLabel);
      }
    }
  }

  // ===============================================================
  // === 4. キャンセル・辞退通知の自動処理 ===
  // ===============================================================
  const cancelQuery = `from:(career_spot@m3career.com) subject:("キャンセルいたしました" OR "キャンセルが申請されました" OR "辞退されました") ${confirmQuerySuffix}`;
  const cancelThreads = GmailApp.search(cancelQuery);

  if (cancelThreads.length > 0) {
    const cancelMessagesAll = GmailApp.getMessagesForThreads(cancelThreads);

    for (let i = 0; i < cancelThreads.length; i++) {
      const thread = cancelThreads[i];
      const messages = cancelMessagesAll[i];
      let threadProcessed = false;

      messages.forEach(message => {
          const body = message.getPlainBody();
          const subject = message.getSubject();
          
          const jobIdMatch = body.match(/([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
          const mailJobId = jobIdMatch ? jobIdMatch[1].trim() : '';
          if (!mailJobId) return;

          const archiveData = archiveSheet.getDataRange().getValues();
          const archSaiyoColIndex = getColIndex_internal(archiveSheet, '採用可否');
          let updatedInArchive = false;

          for (let j = archiveData.length - 1; j > 0; j--) {
              const rowAgency = String(archiveData[j][0]).trim();
              const rowJobId = String(archiveData[j][5]).trim();

              if (rowAgency === 'エムスリー' && rowJobId === mailJobId) {
                  if (archSaiyoColIndex > 0) {
                      const cell = archiveSheet.getRange(j + 1, archSaiyoColIndex);
                      cell.setValue('不採用');
                      cell.setBackground('#ffe599');
                  }
                  updatedInArchive = true;
                  threadProcessed = true;
                  break;
              }
          }

          if (!updatedInArchive) {
              const sheetData = sheet.getDataRange().getValues();
              for (let j = sheetData.length - 1; j > 0; j--) {
                  const rowAgency = String(sheetData[j][0]).trim();
                  const rowJobId = String(sheetData[j][5]).trim();
                  if (rowAgency === 'エムスリー' && rowJobId === mailJobId) {
                      let values = sheet.getRange(j + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
                      values[8] = 'お断り'; values[9] = 'キャンセル済み'; values[11] = new Date();
                      appendRowToArchive(sheet, archiveSheet, values, '不採用');
                      safeDeleteRow(sheet, j + 1);
                      threadProcessed = true;
                      break;
                  }
              }
          }
      });

      if (threadProcessed) {
          thread.addLabel(processedLabel);
      }
    }
  }
}