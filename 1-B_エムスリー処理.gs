// =================================================================
// 【1-B_エムスリー処理】
// エムスリーの応募・確定（オファー経由対応）・掲載停止の各処理
// =================================================================

function processM3(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  const processedLabelName = processedLabel.getName();
  
  // === 1. 応募通知の処理 ===
  GmailApp.search(`from:(career_spot@m3career.com) subject:("ご勤務希望がありました") ${querySuffix}`).forEach(thread => {
    let threadProcessed = false;
    
    thread.getMessages().forEach(message => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const receivedDate = message.getDate();
      
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
        'エムスリー', formatDate(workDateRaw, receivedDate), workTime, newLocationName, department, 
        jobId, 
        doctorName, extract(body, /(https:\/\/career\.m3\.com\/admin\/hospital\/message_threads\/\d+)/), 
        '', '応募通知受信', receivedDate, '', '', jobId 
      ]);
      if (added) threadProcessed = true;
    });
    
    // 処理されたメールが1つでもあればラベルを付与
    if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
        thread.addLabel(processedLabel);
    }
  });

  // === 2. 確定・お断り通知の処理 ===
  const m3ConfirmQuery = `from:(career_spot@m3career.com) subject:("勤務を確定" OR "勤務が確定" OR "勤務確定" OR "をお断り") ${confirmQuerySuffix}`;
  GmailApp.search(m3ConfirmQuery).forEach(thread => {
      let threadProcessed = false;

      const messages = thread.getMessages();
      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        const receivedDate = message.getDate();
        
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

        // --- シートへのアーカイブ処理 ---
        let rowNum = -1;
        const sheetData = sheet.getDataRange().getValues();
        for (let i = sheetData.length - 1; i > 0; i--) {
            const rowAgency = String(sheetData[i][0]).trim();
            if (rowAgency !== 'エムスリー') continue;
            
            const rowStatus = String(sheetData[i][9]).trim();
            if (!rowStatus.includes('応募通知受信')) continue;

            const rowJobId = String(sheetData[i][5]).trim();

            if (rowJobId === mailJobId) {
                rowNum = i + 1;
                break;
            }
        }
        
        let isNewlyProcessed = false;

        if (rowNum !== -1) {
            // 進捗シートに存在する場合
            let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
            if (!values[2] && fetchedWorkTime) values[2] = fetchedWorkTime;
            if (!values[3] && fetchedClinic) values[3] = fetchedClinic;

            values[8] = type; values[9] = status; values[11] = new Date();
            
            // ★変更: ヘルパー関数を使って採用可否ステータスを自動付与してアーカイブ
            appendRowToArchive(sheet, archiveSheet, values, type === '確定' ? '採用' : '不採用');
            safeDeleteRow(sheet, rowNum);
            
            isNewlyProcessed = true;
            threadProcessed = true;
        } else if (type === '確定') {
            // 進捗シートに存在しない場合（オファー直確定、または処理済み）
            const archiveData = archiveSheet.getDataRange().getValues();
            let isExistInArchive = false;

            for (let i = archiveData.length - 1; i > 0; i--) {
              const rowAgency = String(archiveData[i][0]).trim();
              if (rowAgency !== 'エムスリー') continue;
              const rowJobId = String(archiveData[i][5]).trim();

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
              
              // ★変更: 進捗シートにない場合も「採用可否」を付与してアーカイブへ直接書き込み
              appendRowToArchive(null, archiveSheet, tempValues, '採用');
              
              isNewlyProcessed = true;
              threadProcessed = true;
            }
        }

        // ★★★ オファー経由の勤務確定をChatworkへ通知 ★★★
        if (isOffer && isNewlyProcessed) {
            let displayDate = mailDateStr;
            const dMatch = subject.match(/(\d{1,2}月\d{1,2}日.*?\))/);
            if (dMatch) {
                displayDate = String(receivedDate.getFullYear()) + "年" + dMatch[1].replace('(', '日('); 
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

受信時刻：${Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm')}
[/info]`;

            sendToChatwork('419888887', cwMessage);
        }
      });
      
      if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
          thread.addLabel(processedLabel); 
      }
  });
  
  // === 3. 掲載停止通知の処理 ===
  const m3StopQuery = `from:(career_spot@m3career.com) subject:("スポット求人票掲載停止のご案内") ${confirmQuerySuffix}`;
  GmailApp.search(m3StopQuery).forEach(thread => {
      let threadProcessed = false;

      const messages = thread.getMessages();
      messages.forEach(message => {
        const body = message.getPlainBody();
        const subject = message.getSubject();
        const jobIdMatch = body.match(/求人票.*?([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
        const mailJobId = jobIdMatch ? jobIdMatch[1].trim() : extract(body, /([CＣ]\d+)/);

        if (!mailJobId) return;

        let rowNum = -1;
        const sheetData = sheet.getDataRange().getValues();
        for (let i = sheetData.length - 1; i > 0; i--) {
            const rowAgency = String(sheetData[i][0]).trim();
            const rowJobId = String(sheetData[i][5]).trim();
            const rowStatus = String(sheetData[i][9]).trim();
            if (rowAgency === 'エムスリー' && rowStatus.includes('応募通知受信') && rowJobId === mailJobId) {
                rowNum = i + 1;
                break;
            }
        }

        if (rowNum !== -1) {
            let values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
            values[8] = 'お断り'; values[9] = '掲載停止'; values[11] = new Date();
            
            // ★変更: 掲載停止は「不採用」としてアーカイブへ転記
            appendRowToArchive(sheet, archiveSheet, values, '不採用');
            safeDeleteRow(sheet, rowNum);
            
            threadProcessed = true;
        }
      });

      if (threadProcessed && !thread.getLabels().some(l => l.getName() === processedLabelName)) {
          thread.addLabel(processedLabel);
      }
  });
}