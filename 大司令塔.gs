// =============================================================
// ▼▼▼ トリガーで動作するメイン関数群（最終統合版） ▼▼▼
// =============================================================

/**
 * ★★★ 大司令塔 ★★★
 * 全てのGmail処理をこの関数で実行する
 * 【トリガー設定】：この関数を「5分おき」または「10分おき」に設定する
 */
function executeAllGmailProcesses() {
  // アカウント全体で1つのロックを取得
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) { 
    console.log('executeAllGmailProcesses: 他のプロセスが実行中のためスキップしました。');
    return; 
  }
  
  const startTime = new Date();
  console.log('=============== 【大司令塔】プロセスを開始 ===============');
  
  try {
    // --- 処理1：メール自動処理 ---
    console.log('\n--- (1/7) メール自動処理 を開始 ---'); 
    mainProcessEmails_internal();
    console.log('--- (1/7) メール自動処理 が完了 ---');
    
    Utilities.sleep(1500); 
    
    // --- 処理2：直前応募通知 ---
    console.log('\n--- (2/7) 直前応募通知 を開始 ---'); 
    executeAllProcesses_internal();
    console.log('--- (2/7) 直前応募通知 が完了 ---');

    Utilities.sleep(1500); 

    // --- 処理3：緊急キャンセル ---
    console.log('\n--- (3/7) 緊急キャンセル を開始 ---'); 
    checkAndLogEmergencyCancel_internal();
    console.log('--- (3/7) 緊急キャンセル が完了 ---');

    Utilities.sleep(1500); 

    // --- 処理4：時間外着信 ---
    console.log('\n--- (4/7) 時間外着信 を開始 ---'); 
    checkTimegaiCallWithLabel_internal();
    console.log('--- (4/7) 時間外着信 が完了 ---');

    Utilities.sleep(1500); 

    // --- 処理5：非定型キャンセル（ドタキャン） ---
    console.log('\n--- (5/7) 非定型キャンセル を開始 ---'); 
    checkInformalCancel_internal();
    console.log('--- (5/7) 非定型キャンセル が完了 ---');

    Utilities.sleep(1500); 

    // --- 処理6：エムスリーメッセージ ---
    console.log('\n--- (6/7) エムスリーメッセージ を開始 ---');
    if (typeof processM3Messages_internal === 'function') {
      processM3Messages_internal(); 
    } else {
      console.log('processM3Messages_internal関数が存在しないためスキップします。');
    }
    console.log('--- (6/7) エムスリーメッセージ が完了 ---');

    Utilities.sleep(1500); 

    // --- 処理7：個人メール検知 ---
    console.log('\n--- (7/7) 個人メール検知 を開始 ---');
    if (typeof checkPersonalEmails_internal === 'function') {
      checkPersonalEmails_internal(); 
    } else {
      console.log('checkPersonalEmails_internal関数が存在しないためスキップします。');
    }
    console.log('--- (7/7) 個人メール検知 が完了 ---');

  } catch (e) {
    console.error(`【大司令塔】プロセス実行中にエラーが発生しました: ${e.message}`);
    console.error(e.stack);
  } finally {
    lock.releaseLock();
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    console.log(`\n=============== 【大司令塔】プロセスが完了（所要時間: ${duration}秒）===============`);
  }
}

// =============================================================
// ▼▼▼ その他のトリガー ▼▼▼
// =============================================================

/**
* ★★★ 司令塔（シート編集時） ★★★
*/
function onEdit(e) {
  try {
    onEdit_processProgressSheet(e);
  } catch (err) {
    // 該当のシート以外での編集エラーなどは無視
  }

  try {
    onEdit_processUrlSheet(e);
  } catch (err) {
    // 該当のシート以外での編集エラーなどは無視
  }
}

/**
 * ★★★ 司令塔（シートを開いた時） ★★★
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('メニュー')
    .addItem('URL取得', 'displayFormUrl')
    .addItem('DS部に投稿', 'showRecipientDialog')
    .addToUi();
}

/**
* ★★★ 司令塔（フォーム送信時） ★★★
* typeofの安全装置を外し、ファイル2にある関数を直接呼び出します
*/
function onFormSubmit_Master(e) {
  // フォーム送信時に2つの処理を順番に実行
  try {
    processFormResponse(e); // 1. Chatwork投稿文の作成（こだわりの書式）
  } catch (err) {
    console.error(`processFormResponse でエラー: ${err}`);
  }
  
  try {
    autoRecordOnFormSubmit(e); // 2. チェックリストへの転記
  } catch (err) {
    console.error(`autoRecordOnFormSubmit でエラー: ${err}`);
  }
}