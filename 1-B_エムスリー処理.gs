// =================================================================
// 【1-B_エムスリー処理】
// エムスリーの応募（通常＆オファー経由）・確定・掲載停止・キャンセル処理
// ★修正内容：API制限対策として getMessagesForThreads による一括取得に変更
// ★修正内容：スレッド巻き込みによるすり抜けを防止するためラベル除外条件を撤廃
// ★追加修正：API超節約設計（Message-IDキャッシュによる既読メールのスキップ処理を実装）
// ★修正内容：キャッシュ容量オーバー防止のため保持件数を250件に変更
// ★今回追加：「お問い合わせがありました（オファー希望）」を進捗シートへ追加
// ★今回修正：外部スプレッドシートの「拠点名」シート（A列:正規記載 × Q列:住所）と町名で照合
// =================================================================

const CLINIC_MASTER_SPREADSHEET_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
const CLINIC_MASTER_SHEET_NAME = '拠点名';

/**
 * 住所文字列から番地以降を切り落とし、「町名（漢数字の丁目を含む）」までを抽出するヘルパー関数
 * 例1: "東京都板橋区徳丸二丁目2-14..." -> "東京都板橋区徳丸二丁目"
 * 例2: "東京都品川区小山3-15-1..."     -> "東京都品川区小山"
 */
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
 * 「拠点名」シートから【正規記載（A列）・住所（Q列）】のリストを取得する（キャッシュ付き）
 */
function getClinicAddressMap_internal() {
  if (getClinicAddressMap_internal.cache) return getClinicAddressMap_internal.cache;

  try {
    const ss = SpreadsheetApp.openById(CLINIC_MASTER_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CLINIC_MASTER_SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headerRow = data[0].map(v => String(v).trim());

    // 「正規記載」列（index 0）と「住所」列（index 16）をピンポイントで取得
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
    return map;
  } catch (e) {
    console.error('拠点名マスタの読み込みエラー: ' + e.message);
    return [];
  }
}

/**
 * エムスリーの勤務地（住所）を「拠点名」シートの町名と照合して「正規記載」の拠点名に変換する関数
 */
function convertM3AddressToClinicName_internal(rawAddress) {
  if (!rawAddress) return '';
  const addr = String(rawAddress).trim();
  const inputTown = extractTownNameForMatch_internal(addr);
  const inputTownNoChome = inputTown.replace(/[一二三四五六七八九十]+丁目$/, '');

  const clinicMap = getClinicAddressMap_internal();

  if (clinicMap.length > 0) {
    // 1. 町名（漢数字の丁目を含む）で照合（文字数が長い＝より詳細な町名から優先）
    const sortedByTown = [...clinicMap].sort((a, b) => b.townName.length - a.townName.length);
    for (const item of sortedByTown) {
      if (!item.townName || item.townName.length < 3) continue;
      if (inputTown.includes(item.townName) || item.townName.includes(inputTown)) {
        return item.clinicName;
      }
    }

    // 2. 漢数字の丁目を除いた町名同士（例: シート側が「徳丸二丁目」でメール側が「徳丸2-14」の場合など）で照合
    const sortedByNoChome = [...clinicMap].sort((a, b) => b.townWithoutChome.length - a.townWithoutChome.length);
    for (const item of sortedByNoChome) {
      if (!item.townWithoutChome || item.townWithoutChome.length < 3) continue;
      if (inputTownNoChome.includes(item.townWithoutChome) || item.townWithoutChome.includes(inputTownNoChome)) {
        return item.clinicName;
      }
    }

    // 3. 住所の中にマスタの「正規記載」の拠点名そのものが含まれている場合
    const sortedByName = [...clinicMap].sort((a, b) => b.clinicName.length - a.clinicName.length);
    for (const item of sortedByName) {
      if (item.clinicName && item.clinicName.length >= 2 && addr.includes(item.clinicName)) {
        return item.clinicName;
      }
    }
  }

  // マスタに該当がない場合のフォールバック（切り出した町名を返す）
  return inputTown || addr;
}

function processM3(querySuffix, confirmQuerySuffix, processedLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PROGRESS);
  const archiveSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);
  
  // 現在時刻（48時間前判定用）
  const now = new Date();

  // 処理済みIDをプロパティから取得
  const props = PropertiesService.getScriptProperties();
  let processedIds = [];
  try {
    const stored = props.getProperty('M3_PROCESSED_IDS');
    if (stored) processedIds = JSON.parse(stored);
  } catch(e) {}
  let isIdUpdated = false;

  // ===============================================================
  // === 1. 応募通知（通常応募 ＆ オファー希望）の処理 ===
  // ===============================================================
  const applyQuery = `from:(career_spot@m3career.com) subject:("ご勤務希望がありました" OR "お問い合わせがありました") ${querySuffix}`;
  const applyThreads = GmailApp.search(applyQuery);
  
  if (applyThreads.length > 0) {
    const applyMessagesAll = GmailApp.getMessagesForThreads(applyThreads);
    
    for (let i = 0; i < applyThreads.length; i++) {
      const thread = applyThreads[i];
      const messages = applyMessagesAll[i];
      let threadProcessed = false;
      
      messages.forEach(message => {
        const messageId = message.getId();
        if (processedIds.includes(messageId)) return;

        const messageDate = message.getDate();
        if ((now - messageDate) / (1000 * 60 * 60) > 48) {
          processedIds.push(messageId);
          isIdUpdated = true;
          return; 
        }

        const subject = message.getSubject();
        const body = message.getPlainBody();

        // 「お問い合わせがありました」の場合は、オファー希望メールかどうかを判定
        let isOfferApply = false;
        if (subject.includes('お問い合わせがありました')) {
          const cleanBodyForCheck = body.replace(/[\r\n\s ]+/g, '');
          const offerKeyword = '■オファーを出される前にご確認ください・医師のご専門や可能／不可能な手技など必要に応じて直接先生にご確認ください。';
          if (cleanBodyForCheck.includes(offerKeyword)) {
            isOfferApply = true;
          } else {
            // オファー希望ではない通常の問い合わせは処理済みIDに記録してスキップ
            processedIds.push(messageId);
            isIdUpdated = true;
            return;
          }
        }

        processedIds.push(messageId);
        isIdUpdated = true;
        
        // 医師名の抽出（本文に「氏名：」があれば優先、なければ件名から抽出）
        const bodyDocMatch = body.match(/氏名[\s:：]*([^\n\r]+)/);
        const doctorName = bodyDocMatch ? bodyDocMatch[1].trim() : extractM3DoctorName(subject);

        // 求人票ID（識別番号）の抽出
        const jobId = extract(subject, /(C\d+)/) || extract(body, /求人票ID[\s:：]*([A-Za-z0-9]+)/);
        
        // 勤務地（住所）の抽出と「拠点名」シートを使った拠点名変換
        const addressMatch = body.match(/勤務地[\s:：]*([^\n\r]+)/);
        const rawAddress = addressMatch ? addressMatch[1].trim() : extract(body, /^医療法人社団.+キャップスクリニック(.+)様/m).replace(/店$/, '');
        const clinicBaseName = convertM3AddressToClinicName_internal(rawAddress);

        // 診療科の抽出
        const deptMatch = body.match(/募集科目[\s:：]*(.+)/);
        let department = deptMatch ? deptMatch[1].trim() : (findDepartmentFromHistory(ss, doctorName) || '小児科');
        
        // 拠点名の設定（オファー経由は変換後の拠点名のみ、通常応募は従来通り「拠点名（診療科）」）
        const newLocationName = isOfferApply ? clinicBaseName : (rawAddress.replace(/\d.*$/, '').trim() + '（' + department + '）');
        
        // 「最新勤務日:2026年9月30日」を誤って拾わないよう、
        // ①「勤務時間：11/15(日)...」の日付部分 → ②行頭の「勤務日：」 → ③件名の「スポット 11月15日」の順で抽出
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

        // 該当時間の抽出（例: 10:00〜13:00）
        const timeRegex = /勤務時間[\s:：]*.*?\)\s*(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/;
        const timeMatch = body.match(timeRegex);
        const workTime = timeMatch ? timeMatch[1].trim() : extract(body, /勤務時間[\s:：]*\s*(\d{1,2}:\d{2}.+)/);

        // J列（処理状況）：オファー希望は「オファー経由」、通常応募は「応募通知受信」
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
        const messageId = message.getId();
        if (processedIds.includes(messageId)) return;
        processedIds.push(messageId);
        isIdUpdated = true;

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
              fetchedClinic = convertM3AddressToClinicName_internal(clinicMatch[1].trim());
            }

            const timeMatch = offerBody.match(/勤務時間[\s:：]*[^\n\r]*?(\d{1,2}:\d{2}\s*[〜~～\-]\s*\d{1,2}:\d{2})/);
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
          // 「応募通知受信」と「オファー経由」の両方を進捗シートの検索対象にする
          if (!rowStatus.includes('応募通知受信') && !rowStatus.includes('オファー経由')) continue;

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
        const messageId = message.getId();
        if (processedIds.includes(messageId)) return;
        processedIds.push(messageId);
        isIdUpdated = true;

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
          if (rowAgency === 'エムスリー' && (rowStatus.includes('応募通知受信') || rowStatus.includes('オファー経由')) && rowJobId === mailJobId) {
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
        const messageId = message.getId();
        if (processedIds.includes(messageId)) return;
        processedIds.push(messageId);
        isIdUpdated = true;

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

  // 新しく処理したID群をプロパティに保存（最新250件のみ保持）
  if (isIdUpdated) {
    if (processedIds.length > 250) {
      processedIds = processedIds.slice(-250);
    }
    props.setProperty('M3_PROCESSED_IDS', JSON.stringify(processedIds));
  }
}