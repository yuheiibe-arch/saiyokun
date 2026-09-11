function debug_processM3Cancel() {
  console.log('--- デバッグ開始：エムスリーキャンセル通知の遡及テスト ---');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('処理済み'); // 実際のシート名に合わせてください
  
  if (!archiveSheet) {
    console.error('【エラー】シート「処理済み」が見つかりません。');
    return;
  }

  // ★修正：受信日を8月20日以降に広げました
  const m3CancelQuery = `from:(career_spot@m3career.com) subject:("キャンセルいたしました" OR "キャンセルが申請されました" OR "辞退されました") after:2026/08/20`;
  const threads = GmailApp.search(m3CancelQuery);

  if (threads.length === 0) {
    console.log('【結果】条件に一致するメールが見つかりませんでした。');
    return;
  }

  console.log(`【結果】対象スレッドを ${threads.length} 件発見しました。解析を開始します。`);

  const archiveData = archiveSheet.getDataRange().getValues();
  const archSaiyoColIndex = getColIndex_internal(archiveSheet, '採用可否');

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const receivedDateStr = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd HH:mm');

      // IDの抽出
      const jobIdMatch = body.match(/([CＣ]\d+)/) || subject.match(/([CＣ]\d+)/);
      const mailJobId = jobIdMatch ? jobIdMatch[1].trim() : '';

      if (!mailJobId) {
        return;
      }

      console.log(`\n[メール検知] 受信日時: ${receivedDateStr} / 件名: ${subject}`);
      console.log(`[抽出ID] ${mailJobId}`);

      // 処理済み（アーカイブ）シートを検索
      let foundInArchive = false;
      for (let i = archiveData.length - 1; i > 0; i--) {
        const rowAgency = String(archiveData[i][0]).trim();
        const rowJobId = String(archiveData[i][5]).trim();
        const rowDoctor = String(archiveData[i][6]).trim();
        const rowCurrentStatus = archSaiyoColIndex > 0 ? String(archiveData[i][archSaiyoColIndex - 1]) : '列不明';

        if (rowAgency === 'エムスリー' && rowJobId === mailJobId) {
          foundInArchive = true;
          console.log(`★ 【シート一致】行番号: ${i + 1} / 医師名: ${rowDoctor} / 現在のステータス: ${rowCurrentStatus}`);
          console.log(`👉 本番環境では、この行（${i + 1}行目）の「採用可否」を「不採用」に変更し、背景色を黄色にします。`);
          break;
        }
      }

      if (!foundInArchive) {
        console.log(`▲ 【シート不一致】「処理済み」シートの中にID [${mailJobId}] は見つかりませんでした。`);
      }
    });
  });
  
  console.log('\n--- デバッグ完了 ---');
}