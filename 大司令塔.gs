// =============================================================
// ▼▼▼ トリガーで動作するメイン関数群（エラー独立・連鎖停止防止・高速化版） ▼▼▼
// =============================================================

/**
 * ★★★ 大司令塔（新バージョン） ★★★
 * 全てのGmail処理をこの関数で実行する
 * 【トリガー設定】：あなたのアカウントで、この関数（executeAllGmailProcesses_v2）を「5分おき」または「10分おき」に設定する
 */
function executeAllGmailProcesses_v2() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30000)) {
    console.log('executeAllGmailProcesses_v2: 他のプロセスが実行中のためスキップしました。');
    return;
  }

  const startTime = new Date();
  console.log('=============== 【大司令塔 v2】プロセスを開始 ===============');

  try {
    const tasks = [
      { name: '(1/9) メール自動処理', run: () => mainProcessEmails_internal() },
      { name: '(2/9) 直前応募通知', run: () => executeAllProcesses_internal() },
      { name: '(3/9) 緊急キャンセル', run: () => { if (typeof checkAndLogEmergencyCancel_internal === 'function') checkAndLogEmergencyCancel_internal(); } },
      { name: '(4/9) 紹介会社キャンセル反映', run: () => { if (typeof processAgencyCancelEmails_internal === 'function') processAgencyCancelEmails_internal(); } },
      { name: '(5/9) 時間外着信', run: () => { if (typeof checkTimegaiCallWithLabel_internal === 'function') checkTimegaiCallWithLabel_internal(); } },
      { name: '(6/9) 非定型キャンセル', run: () => { if (typeof checkInformalCancel_internal === 'function') checkInformalCancel_internal(); } },
      { name: '(7/9) エムスリーメッセージ', run: () => { if (typeof processM3Messages_internal === 'function') processM3Messages_internal(); } },
      { name: '(8/9) 個人メール検知', run: () => { if (typeof checkPersonalEmails_internal === 'function') checkPersonalEmails_internal(); } },
      { name: '(9/9) jinjer有給申請検知', run: () => { if (typeof processJinjerPaidLeave_internal === 'function') processJinjerPaidLeave_internal(); } }
    ];

    for (const task of tasks) {
      // ★ 6分タイムアウト安全装置：開始から310秒を超えていたら次回トリガーに譲って安全終了
      if ((Date.now() - startTime.getTime()) > 310 * 1000) {
        console.warn(`⏱ 実行時間が310秒を超えたため、タイムアウト防止で「${task.name}」以降を次回トリガーへ持ち越します。`);
        break;
      }

      try {
        console.log(`\n--- ${task.name} を開始 ---`);
        task.run();
        console.log(`--- ${task.name} が完了 ---`);
      } catch (e) {
        console.error(`【エラー】${task.name}で失敗: ${e.message}`);

        const errorMsg = String(e.message || '').toLowerCase();
        if (errorMsg.includes('too many times') || errorMsg.includes('quota') || errorMsg.includes('limit exceeded')) {
          sendApiErrorToSlack_internal(task.name, e.message);
          console.log('API制限エラーを検知したため、無駄な通信を避けるべく後続の全タスクをキャンセルしました。');
          break;
        }
      }
      // ★合理化：新着0件時の空待ち（旧3000ms×9=27秒）を300msに短縮（実際のChatwork送信時は各関数内で待機）
      Utilities.sleep(300);
    }

  } finally {
    lock.releaseLock();
    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000;
    console.log(`\n=============== 【大司令塔 v2】プロセスが完了（所要時間: ${duration}秒）===============`);
  }
}

/**
 * API制限エラーをSlackに直接通知するヘルパー関数
 */
function sendApiErrorToSlack_internal(stepName, errorMessage) {
  try {
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
    if (!webhookUrl) return;

    const props = PropertiesService.getScriptProperties();
    const todayStr = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
    const lastErrorDate = props.getProperty('LAST_API_ERROR_DATE');

    if (lastErrorDate === todayStr) {
      console.log('本日のAPI制限エラーは通知済みのため、Slackへの送信をスキップします。');
      return;
    }

    const currentHour = parseInt(Utilities.formatDate(new Date(), 'JST', 'H'), 10);
    const isDayTime = (currentHour >= 7 && currentHour < 19);
    const channelMention = isDayTime ? "<!channel>\n" : "";

    const payload = {
      channel: 'C09TRHGU64D',
      text: `${channelMention}*【🚨 システムエラー・処理停止】*\nGoogle APIの1日あたりの上限（Quota）に到達したため、大司令塔の処理を強制停止しました。\n\n*・発生箇所:* ${stepName}\n*・エラー内容:* \`${errorMessage}\`\n\n※制限がリセットされるまで、自動連携が一時的にストップします。手動での確認をお願いします。`
    };

    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });

    props.setProperty('LAST_API_ERROR_DATE', todayStr);
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
    .addItem('お久しぶり報告', 'openOhisashiburiFormDialog')
    .addToUi();
}

/**
 * 「お久しぶり報告」のGoogleフォームを開くダイアログを表示する
 */
function openOhisashiburiFormDialog() {
  const formUrl = 'https://docs.google.com/forms/d/e/1FAIpQLScrvlzS7GLE_Xn4APGZJuCXgxwLoVs8HUNTu9OH8sb1zSQdZQ/viewform?usp=header';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: sans-serif; text-align: center; padding: 15px;">' +
    '<p style="margin-bottom: 18px; font-size: 14px; color: #333;">下のボタンからお久しぶり報告フォームを開いてください。</p>' +
    '<a href="' + formUrl + '" target="_blank" onclick="google.script.host.close();" ' +
    'style="display: inline-block; padding: 10px 24px; background-color: #1a73e8; color: #ffffff; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px;">' +
    'お久しぶり報告フォームを開く</a>' +
    '</div>'
  )
  .setWidth(360)
  .setHeight(130);

  SpreadsheetApp.getUi().showModalDialog(html, 'お久しぶり報告');
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

  try {
    if (typeof sendEncryptedFormEventToCaps === 'function') sendEncryptedFormEventToCaps(e);
  } catch (err) {
    console.error(`sendEncryptedFormEventToCaps でエラー: ${err}`);
  }

  try {
    if (typeof autoArchiveNextSokketsuOnFormSubmit === 'function') autoArchiveNextSokketsuOnFormSubmit(e);
  } catch (err) {
    console.error(`autoArchiveNextSokketsuOnFormSubmit でエラー: ${err}`);
  }
}