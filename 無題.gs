function debugHeaderConfiguration() {
  console.log("▼▼▼ シートのヘッダー構成 動的取得テスト 開始 ▼▼▼");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheetName = typeof SHEET_NAME_PROGRESS !== 'undefined' ? SHEET_NAME_PROGRESS : '紹介会社進捗';
  const archiveSheetName = typeof SHEET_NAME_ARCHIVE !== 'undefined' ? SHEET_NAME_ARCHIVE : '処理済み';
  
  const progressSheet = ss.getSheetByName(progressSheetName);
  const archiveSheet = ss.getSheetByName(archiveSheetName);
  
  // 今回のシステムでシステムが「動的に」探しに行く列名のリスト
  const targetColumns = ['直前応募通知', '採用可否', '医師名', '勤務希望日', '対応時間', '処理日時'];

  function checkHeaders(sheet, sheetName) {
    console.log(`\n--- 【${sheetName}】シートのヘッダー確認 ---`);
    if (!sheet) {
      console.log(`❌ シートが見つかりません。`);
      return;
    }
    
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      console.log(`❌ ヘッダー（1行目）が空です。`);
      return;
    }
    
    // 1行目（ヘッダー）をすべて取得
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    console.log(`・全列数: ${lastCol}`);
    
    // ターゲットの列を動的に探す（ハードコードなしで場所を特定する）
    for (const colName of targetColumns) {
      // 1. 完全一致で探す
      const idx = headers.findIndex(h => String(h).trim() === colName);
      if (idx !== -1) {
        // 例: 1列目ならA列、2列目ならB列と表示するための計算
        const colLetter = String.fromCharCode(65 + (idx % 26)); 
        const prefix = idx >= 26 ? String.fromCharCode(64 + Math.floor(idx / 26)) : '';
        console.log(`✅ 「${colName}」を特定: ${idx + 1}列目 (${prefix}${colLetter}列)`);
      } else {
        // 2. もし完全一致がなければ、部分一致（文字が含まれているか）で探す
        const partialIdx = headers.findIndex(h => String(h).includes(colName));
        if (partialIdx !== -1) {
          const colLetter = String.fromCharCode(65 + (partialIdx % 26));
          const prefix = partialIdx >= 26 ? String.fromCharCode(64 + Math.floor(partialIdx / 26)) : '';
          console.log(`⚠️ 類似検知: 「${colName}」を含む列名「${headers[partialIdx]}」を ${partialIdx + 1}列目 (${prefix}${colLetter}列) に発見`);
        } else {
          console.log(`❌ 「${colName}」列が見つかりません。`);
        }
      }
    }
  }
  
  checkHeaders(progressSheet, progressSheetName);
  checkHeaders(archiveSheet, archiveSheetName);
  
  console.log("\n▲▲▲ テスト完了 ▲▲▲");
  console.log("※ 『直前応募通知』が両方のシートで「✅」として特定されていれば、列を移動してもシステムは絶対に迷子になりません。");
}