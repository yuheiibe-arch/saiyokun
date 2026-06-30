// =============================================================
// 【ファイル3：メンション投稿・ダイアログ処理】
// =============================================================

/**
 * メンションタグと名前が合体した文字列（例: "[To:12345]名前"）を分割するヘルパー関数
 */
function parseMentionString(str) {
  if (!str) return null;
  const match = String(str).match(/^(\[.+?\])\s*(.*)$/);
  if (match) {
    return { tag: match[1].trim(), name: match[2].trim() || match[1] };
  }
  return { tag: String(str).trim(), name: String(str).trim() };
}

/**
 * メニューから「DS部に投稿」を押した時に呼ばれる関数（ダイアログを開く）
 */
function showRecipientDialog() {
  // ★注意★ 'dialog' の部分は、左側のメニューにあるHTMLファイルの名前と一致させてください
  const template = HtmlService.createTemplateFromFile('dialog'); 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mentionSheet = ss.getSheetByName(MENTION_SHEET_NAME); // グローバル設定を参照
  
  let mainList = [];
  let ccList = [];
  let optionalList = [];

  // メンションリストから宛先を取得
  if (mentionSheet) {
    const data = mentionSheet.getDataRange().getValues();
    
    // ヘッダー(1行目)を飛ばして2行目から読み込み
    for (let i = 1; i < data.length; i++) {
      const colA_To = data[i][0];  // A列: メイン(To)
      const colB_CC = data[i][1];  // B列: CC
      const colC_Opt = data[i][2]; // C列: 任意

      if (colA_To) {
        const parsed = parseMentionString(colA_To);
        if (parsed) mainList.push({ name: parsed.name, tag: parsed.tag, checked: true, disabled: false });
      }
      if (colB_CC) {
        const parsed = parseMentionString(colB_CC);
        if (parsed) ccList.push({ name: parsed.name, tag: parsed.tag, checked: true, disabled: false });
      }
      if (colC_Opt) {
        const parsed = parseMentionString(colC_Opt);
        if (parsed) optionalList.push({ name: parsed.name, tag: parsed.tag, checked: false, disabled: false });
      }
    }
  }

  // HTML側にデータを渡す
  template.mainList = mainList;
  template.ccList = ccList;
  template.optionalList = optionalList;

  // ダイアログを表示
  const htmlOutput = template.evaluate()
    .setWidth(450)
    .setHeight(550)
    .setTitle('Chatwork宛先選択');
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Chatwork宛先選択');
}

/**
 * HTML画面で「投稿」ボタンを押した時に呼ばれる関数（実際の送信処理）
 */
function executeActualPost(postData) {
  const mainList = postData.mainList || [];
  const ccList = postData.ccList || [];

  // メンションタグを組み立てる
  let mentionText = "";
  mainList.forEach(tag => mentionText += tag + "\n");
  ccList.forEach(tag => mentionText += tag + "\n");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urlSheet = ss.getSheetByName(URL_SHEET_NAME); // グローバル設定を参照
  if (!urlSheet) throw new Error(`シート「${URL_SHEET_NAME}」が見つかりません。`);

  // 本文(C2)と担当者名(C3)を取得
  let messageBody = urlSheet.getRange('C2').getValue();
  const picName = urlSheet.getRange('C3').getValue(); 

  if (!messageBody) throw new Error("送信するメッセージがC2セルにありません。（先にフォーム回答を処理してください）");

  // ★採用・不採用・DS承認報告の3パターンに対応して担当者を挿入
  if (picName) {
    messageBody = messageBody.replace('〈不採用報告〉', '〈不採用報告〉\n担当：' + picName);
    messageBody = messageBody.replace('〈採用報告〉', '〈採用報告〉\n担当：' + picName);
    messageBody = messageBody.replace('〈DS承認報告〉', '〈DS承認報告〉\n担当：' + picName);
  }

  // 宛先と本文を合体
  const finalMessage = mentionText + "\n" + messageBody;

  // 共通のAPIキー(sendToChatwork)を使用
  const success = sendToChatwork(SAIYO_HOKOKU_CHATWORK_ROOM_ID, finalMessage);
  
  if (!success) {
    throw new Error("Chatworkへの送信に失敗しました。APIキーを確認してください。");
  }
}