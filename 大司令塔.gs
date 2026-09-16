// =============================================================
// ▼▼▼ トリガーで動作するメイン関数群（エラー独立・連鎖停止防止版） ▼▼▼
// =============================================================

/**
 * ★★★ 大司令塔 ★★★
 * 全てのGmail処理をこの関数で実行する
 * 【トリガー設定】：この関数を「5分おき」または「10分おき」に設定する
 */
function executeAllGmailProcesses() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) { 
    console.log('executeAllGmailProcesses: 他のプロセスが実行中のためスキップしました。');
    return; 
  }
  
  const startTime = new Date();
  console.log('=============== 【大司令塔】プロセスを開始 ===============');
  
  // lock.releaseLock() を確実に実行するための全体ラップ
  try {
    
    // =====================================
    // (1/7) メール自動処理
    // =====================================
    try {
      console.log('\n--- (1/7) メール自動処理 を開始 ---'); 
      mainProcessEmails_internal();
      console.log('--- (1/7) メール自動処理 が完了 ---');
    } catch (e) {
      console.error(`【エラー】(1/7) メール自動処理で失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 
    
    // =====================================
    // (2/7) 直前応募通知
    // =====================================
    try {
      console.log('\n--- (2/7) 直前応募通知 を開始 ---'); 
      executeAllProcesses_internal();
      console.log('--- (2/7) 直前応募通知 が完了 ---');
    } catch (e) {
      console.error(`【エラー】(2/7) 直前応募通知で失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (3/7) 緊急キャンセル
    // =====================================
    try {
      console.log('\n--- (3/7) 緊急キャンセル を開始 ---'); 
      checkAndLogEmergencyCancel_internal();
      console.log('--- (3/7) 緊急キャンセル が完了 ---');
    } catch (e) {
      console.error(`【エラー】(3/7) 緊急キャンセルで失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (4/7) 時間外着信 (★これで絶対にすり抜けない)
    // =====================================
    try {
      console.log('\n--- (4/7) 時間外着信 を開始 ---'); 
      checkTimegaiCallWithLabel_internal();
      console.log('--- (4/7) 時間外着信 が完了 ---');
    } catch (e) {
      console.error(`【エラー】(4/7) 時間外着信で失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (5/7) 非定型キャンセル
    // =====================================
    try {
      console.log('\n--- (5/7) 非定型キャンセル を開始 ---'); 
      checkInformalCancel_internal();
      console.log('--- (5/7) 非定型キャンセル が完了 ---');
    } catch (e) {
      console.error(`【エラー】(5/7) 非定型キャンセルで失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (6/7) エムスリーメッセージ
    // =====================================
    try {
      console.log('\n--- (6/7) エムスリーメッセージ を開始 ---');
      if (typeof processM3Messages_internal === 'function') {
        processM3Messages_internal(); 
      } else {
        console.log('processM3Messages_internal関数が存在しないためスキップします。');
      }
      console.log('--- (6/7) エムスリーメッセージ が完了 ---');
    } catch (e) {
      console.error(`【エラー】(6/7) エムスリーメッセージで失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (7/8) 個人メール検知
    // =====================================
    try {
      console.log('\n--- (7/8) 個人メール検知 を開始 ---');
      if (typeof checkPersonalEmails_internal === 'function') {
        checkPersonalEmails_internal(); 
      } else {
        console.log('checkPersonalEmails_internal関数が存在しないためスキップします。');
      }
      console.log('--- (7/8) 個人メール検知 が完了 ---');
    } catch (e) {
      console.error(`【エラー】(7/8) 個人メール検知で失敗: ${e.message}`);
    }
    Utilities.sleep(1500); 

    // =====================================
    // (8/8) jinjer有給申請検知
    // =====================================
    try {
      console.log('\n--- (8/8) jinjer有給申請検知 を開始 ---');
      if (typeof processJinjerPaidLeave_internal === 'function') {
        processJinjerPaidLeave_internal(); 
      } else {
        console.log('processJinjerPaidLeave_internal関数が存在しないためスキップします。');
      }
      console.log('--- (8/8) jinjer有給申請検知 が完了 ---');
    } catch (e) {
      console.error(`【エラー】(8/8) jinjer有給申請検知で失敗: ${e.message}`);
    }

  } finally {
    // 途中のどれでエラーが起きても確実にロックを解除する
    lock.releaseLock();
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    console.log(`\n=============== 【大司令塔】プロセスが完了（所要時間: ${duration}秒）===============`);
  }
}

// =============================================================
// ▼▼▼ その他のトリガー ▼▼▼
// =============================================================

function onEdit(e) {
  try {
    onEdit_processProgressSheet(e);
  } catch (err) {}

  try {
    onEdit_processUrlSheet(e);
  } catch (err) {}

  try {
    onEdit_processPaidLeaveSheet(e);
  } catch (err) {}
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('メニュー')
    .addItem('URL取得', 'displayFormUrl')
    .addItem('DS部に投稿', 'showRecipientDialog')
    .addToUi();
}

function onFormSubmit_Master(e) {
  try {
    processFormResponse(e);
  } catch (err) {
    console.error(`processFormResponse でエラー: ${err}`);
  }
  
  try {
    autoRecordOnFormSubmit(e);
  } catch (err) {
    console.error(`autoRecordOnFormSubmit でエラー: ${err}`);
  }
}