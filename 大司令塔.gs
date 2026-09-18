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
  
  try {
    // ★ 8つの処理を配列化して一括管理（コードがスッキリし、途中停止させやすくなります）
    const tasks = [
      { name: '(1/8) メール自動処理', run: () => mainProcessEmails_internal() },
      { name: '(2/8) 直前応募通知', run: () => executeAllProcesses_internal() },
      { name: '(3/8) 緊急キャンセル', run: () => { if (typeof checkAndLogEmergencyCancel_internal === 'function') checkAndLogEmergencyCancel_internal(); } },
      { name: '(4/8) 時間外着信', run: () => { if (typeof checkTimegaiCallWithLabel_internal === 'function') checkTimegaiCallWithLabel_internal(); } },
      { name: '(5/8) 非定型キャンセル', run: () => { if (typeof checkInformalCancel_internal === 'function') checkInformalCancel_internal(); } },
      { name: '(6/8) エムスリーメッセージ', run: () => { if (typeof processM3Messages_internal === 'function') processM3Messages_internal(); } },
      { name: '(7/8) 個人メール検知', run: () => { if (typeof checkPersonalEmails_internal === 'function') checkPersonalEmails_internal(); } },
      { name: '(8/8) jinjer有給申請検知', run: () => { if (typeof processJinjerPaidLeave_internal === 'function') processJinjerPaidLeave_internal(); } }
    ];

    for (const task of tasks) {
      try {
        console.log(`\n--- ${task.name} を開始 ---`); 
        task.run();
        console.log(`--- ${task.name} が完了 ---`);
      } catch (e) {
        console.error(`【エラー】${task.name}で失敗: ${e.message}`);
        
        // ★ API制限（上限エラー）を検知した場合、Slackへ通知して後続処理をすべて強制停止する
        const errorMsg = e.message.toLowerCase();
        if (errorMsg.includes('too many times') || errorMsg.includes('quota') || errorMsg.includes('limit exceeded')) {
          sendApiErrorToSlack_internal(task.name, e.message);
          console.log('API制限エラーを検知したため、無駄な通信を避けるべく後続の全タスクをキャンセルしました。');
          break; // ループを抜け出して大司令塔を終了
        }
      }
      Utilities.sleep(1500); 
    }

  } finally {
    // 途中でストップしても確実にロックを解除する
    lock.releaseLock();
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    console.log(`\n=============== 【大司令塔】プロセスが完了（所要時間: ${duration}秒）===============`);
  }
}

/**
 * API制限エラーをSlackに直接通知するヘルパー関数
 */
function sendApiErrorToSlack_internal(stepName, errorMessage) {
  try {
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
    if (!webhookUrl) return;

    const payload = {
      channel: 'C09TRHGU64D', // 共通関数.gsで定義されているデフォルト通知先のSlackチャンネル（民間医局・直前応募などと同じチャンネル）
      text: `<!channel>\n*【🚨 システムエラー・処理停止】*\nGoogle APIの1日あたりの上限（Quota）に到達したため、大司令塔の処理を強制停止しました。\n\n*・発生箇所:* ${stepName}\n*・エラー内容:* \`${errorMessage}\`\n\n※制限がリセットされるまで、自動連携が一時的にストップします。手動での確認をお願いします。`
    };

    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });
    console.log('SlackへAPI制限エラーの警告を送信しました。');
  } catch (e) {
    console.error('Slackへの警告送信に失敗しました: ' + e.message);
  }
}

// =============================================================
// ▼▼▼ その他のトリガー ▼▼▼
// =============================================================

function onEdit(e) {
  try {
    if (typeof onEdit_processProgressSheet === 'function') onEdit_processProgressSheet(e);
  } catch (err) {}

  try {
    if (typeof onEdit_processUrlSheet === 'function') onEdit_processUrlSheet(e);
  } catch (err) {}

  try {
    if (typeof onEdit_processPaidLeaveSheet === 'function') onEdit_processPaidLeaveSheet(e);
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
    if (typeof processFormResponse === 'function') processFormResponse(e);
  } catch (err) {
    console.error(`processFormResponse でエラー: ${err}`);
  }
  
  try {
    if (typeof autoRecordOnFormSubmit === 'function') autoRecordOnFormSubmit(e);
  } catch (err) {
    console.error(`autoRecordOnFormSubmit でエラー: ${err}`);
  }
}