// =================================================================
// ▼▼▼ 本番用：個人メール検知・通知スクリプト ▼▼▼
// ★修正内容：API制限対策として getMessagesForThreads による一括取得に変更
// ★修正内容：検索クエリに「-label:処理済み」を追加し無駄な通信を削減
// ★追加修正：送信元名がおかしい場合、本文末尾の署名から名前を抽出して上書き
// ★追加修正：時間帯（07:00〜21:00）による [toall] の付与ロジックを追加
// ★追加修正：API超節約設計（検索範囲を過去2時間のUNIX時間に変更）
// ★追加修正：自社署名の「株式会社」による自爆誤検知を防ぐため、引用カット後に判定
// =================================================================

// 外部メンションリストの設定
const MENTION_EXT_SPREADSHEET_ID = '1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA';
const MENTION_EXT_SHEET_NAME = 'メンション先選択';

function checkPersonalEmails_internal() {
  console.log('【個人メール検知】処理を開始します。');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME_PERSONAL);
  
  // ログシートがなければ自動作成
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME_PERSONAL);
    logSheet.getRange('A1:F1').setValues([['受信日時', '送信元', '宛先', '件名', '本文プレビュー', 'Message-ID']]);
    logSheet.setFrozenRows(1);
  }

  // 既存のMessage-IDを取得（スレッド内の重複通知を完全防止）
  const lastRow = logSheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    existingIds = logSheet.getRange(2, 6, lastRow - 1, 1).getValues().flat();
  }

  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL_PERSONAL);
  if (!label) { label = GmailApp.createLabel(PROCESSED_LABEL_PERSONAL); }

  // ★★★【重要修正】バグの元凶である after:UNIX時間 を撤廃し、newer_than:1d に変更 ★★★
  const searchQuery = `(to:doctor-support@caps365.jp OR to:dr.saiyo@mnys.jp) -label:${label.getName()} newer_than:1d`;
  
  // ★ API通信上限を絶対に回避するためのストッパー: 最新の30スレッドだけをチェック
  const threads = GmailApp.search(searchQuery, 0, 30);

  if (threads.length === 0) {
    console.log('対象のメールはありませんでした。');
    return;
  }

  // --- 外部メンションリストの読み込みと異体字の標準化 ---
  let mentionList = [];
  try {
    const extSS = SpreadsheetApp.openById(MENTION_EXT_SPREADSHEET_ID);
    const mentionSheet = extSS.getSheetByName(MENTION_EXT_SHEET_NAME);
    if (mentionSheet) {
      const bColData = mentionSheet.getRange(1, 2, mentionSheet.getLastRow(), 1).getValues().flat();
      mentionList = bColData.map(val => {
        const match = val.match(/\[To:\d+\](.+)さん/);
        if (match) {
          let rawName = match[1].trim();
          // 「髙」を「高」、「﨑」を「崎」に変換（検索用の標準化）
          let normalizedName = rawName.replace(/髙/g, '高').replace(/﨑/g, '崎');
          return { tag: val, rawName: rawName, normName: normalizedName };
        }
        return null;
      }).filter(item => item !== null);
    }
  } catch (e) {
    console.error('メンションリストの読み込みに失敗しました（通知はメンションなしで続行します）:', e.message);
  }

  const newRows = [];
  const now = new Date(); // 現在時刻を取得（古いメール判定用）

  // ★ API通信1回で全メッセージを一括取得
  const allMessages = GmailApp.getMessagesForThreads(threads);

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = allMessages[i];
    let threadHasNewPersonalMail = false;

    for (const message of messages) {
      const messageId = message.getId();
      // すでに通知済みのMessage-IDはスキップ
      if (existingIds.includes(messageId)) continue;

      const receivedDate = message.getDate();
      
      // 過去メール防止のストッパー: そのメール自体の受信日時が24時間（1日）より古い場合はスキップ
      const diffHours = (now.getTime() - receivedDate.getTime()) / (1000 * 60 * 60);
      if (diffHours > 24) {
        existingIds.push(messageId); // 二度と通知判定に入らないようにIDだけ記録しておく
        continue;
      }

      const from = message.getFrom();
      const subject = message.getSubject();
      const rawBody = message.getPlainBody(); 
      
      // ★ 引用部分を先にカットする（自社の署名にある「株式会社」等の誤検知を防ぐため）
      const regexQuoteMarkers = /\n(>|On .*> wrote:|.* <.*@.*> wrote:|\d{4}[年\/]\d{1,2}[月\/]\d{1,2}日?.*(?::|のメール:)|From: .*|Sent: .*|-{3,}|={3,}|#{3,}|_+\s*$|(?:iPhone|iPad|スマートフォン)から送信)/i;
      let cleanBodyBeforeCheck = rawBody.split(regexQuoteMarkers)[0].trim();
      
      // ★ 生の本文(rawBody)ではなく、引用をカットした(cleanBodyBeforeCheck)で判定する
      if (classifyEmailForProduction(from, subject, cleanBodyBeforeCheck) === '個人候補') {
        const receivedDateStr = Utilities.formatDate(receivedDate, 'JST', 'yyyy/MM/dd HH:mm');
        let body = rawBody; 
        
        // HTMLメールのCSS（装飾コード）の残骸を削除するフィルター
        body = body.replace(/(?:[a-zA-Z0-9\*,:.#\s_-]+\s*\{[^}]+\}\s*)+/g, '').trim();
        
        // 本文の異体字も標準化（判定用）
        const normBody = body.replace(/髙/g, '高').replace(/﨑/g, '崎');
        
        // 引用部分をカットしたクリーン本文（通知用）
        let cleanBody = body.split(regexQuoteMarkers)[0].trim();
        // 引用部分をカットしたクリーン本文（判定用）
        let cleanNormBody = normBody.split(regexQuoteMarkers)[0].trim();

        // シート書き込み用のデータを準備
        let bodyPreview = cleanBody;
        if (bodyPreview.length > 300) bodyPreview = bodyPreview.substring(0, 300) + '...';
        newRows.push([receivedDateStr, from, message.getTo(), subject, bodyPreview, messageId]);
        existingIds.push(messageId);

        // 送信元(From)から名前部分だけを抽出し、クォーテーションのゴミを完全に除去
        let doctorName = from;
        const nameMatch = from.match(/^["']?(.*?)["']?(?:\s+via\s+.*?)?\s*<.+>$/i);
        if (nameMatch && nameMatch[1]) doctorName = nameMatch[1];
        doctorName = doctorName.replace(/["']/g, '').replace(/先生$/, '').trim();

        // ★★★ Vodafone対策：本文末尾（署名）から本当の名前を推測して上書き ★★★
        let realName = doctorName;
        const lines = cleanBody.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        if (lines.length > 0) {
          const lastLine = lines[lines.length - 1];
          // 英数字や記号を含まない、2〜15文字の日本語らしき文字列なら名前に採用
          if (/^[^!@#$%^&*()_+={}\[\]|\\:;"'<>,.?/0-9A-Za-z]{2,15}$/.test(lastLine)) {
            realName = lastLine;
          }
        }

        // --- 担当者メンションの判定 ---
        let targetTags = [];
        const scanHeader = cleanNormBody.substring(0, 100); // 宛名検索用（冒頭100文字）

        for (const m of mentionList) {
          const name = m.normName; // 標準化された名前で検索
          
          // パターン1（冒頭での宛名）フルネーム「岸 大貴様」や「岸大貴 様」などにも対応
          const atenaRegex = new RegExp(name + "[\\s ]*(?:[^\\s 様さん宛先生]{1,5}[\\s ]*)?(様|さん|宛|先生)", "i");
          
          // パターン2：スレッドの下の方（引用部全体）にある自分たちの署名や名乗り
          const signRegex = new RegExp(`(?:事務局|サポート部)[\\s\\n ]*` + name + `(?:[\\s\\n ]|$)`, "i");
          const nanoriRegex = new RegExp(`、` + name + `(?:です|でございます)`, "i");
          const simpleSignRegex = new RegExp(`^[\\s ]*` + name + `[\\s ]*$`, "m");

          if (atenaRegex.test(scanHeader) || signRegex.test(normBody) || nanoriRegex.test(normBody) || simpleSignRegex.test(normBody)) {
            targetTags.push(m.tag);
          }
        }
        
        // メンションタグの組み立て（重複を削除して結合）
        targetTags = [...new Set(targetTags)];
        
        // ★★★ 時間帯判定ロジックの追加 ★★★
        const currentHour = now.getHours();
        const isDayTime = (currentHour >= 7 && currentHour < 21); // 07:00 〜 20:59
        
        let mentionText = '';
        if (targetTags.length > 0) {
            mentionText = targetTags.join(' ') + '\n';
        } else if (isDayTime) {
            mentionText = '[toall]\n';
        }

        // Chatworkメッセージを作成（元のメアド名と本当の名前が、スペースの有無に関わらず一致するかチェック）
        const isSameName = (doctorName.replace(/\s/g, '') === realName.replace(/\s/g, ''));

        const chatworkMessage = `${mentionText}[info][title]医師からのメール[/title]\n医師からメールが届きました。担当者は確認してください。\n受信時刻：${receivedDateStr}\n件名：${subject}\n[info][title]${realName} 先生[/title]\n${!isSameName ? doctorName + ' 先生\n' : ''}${cleanBody}\n[/info]\n[/info]`;

        // 共通の送信関数でChatworkへ通知
        const isSuccess = sendToChatwork(PERSONAL_MAIL_CHATWORK_ROOM_ID, chatworkMessage);
        if (!isSuccess) {
            console.error(`Chatworkへの送信に失敗しました。RoomID: ${PERSONAL_MAIL_CHATWORK_ROOM_ID}`);
        }

        // ChatworkのAPI制限回避のための待機時間
        Utilities.sleep(1500);

        threadHasNewPersonalMail = true;
      }
    }

    if (threadHasNewPersonalMail) {
      thread.addLabel(label);
    }
  }

  if (newRows.length > 0) {
    logSheet.getRange(lastRow + 1, 1, newRows.length, 6).setValues(newRows);
    console.log(`${newRows.length} 件の個人メールを通知・記録しました。`);
  }
  
  console.log('【個人メール検知】処理が完了しました。');
}

/**
 * 送信元(From)や件名からカテゴリーを判定する内部関数
 */
function classifyEmailForProduction(from, subject, body) {
  const fromLower = from.toLowerCase();
  const subjectLower = subject.toLowerCase();
  const bodyLower = (body || '').toLowerCase();

  // 1. システム通知（★info@を無条件でブロック）
  if (fromLower.includes('no-reply') || 
      fromLower.includes('noreply') || 
      fromLower.includes('hospital@medrt.com') || 
      fromLower.includes('info@') ||             
      fromLower.includes('info-portal') || 
      fromLower.includes('enzine') || 
      fromLower.includes('entry@') || 
      fromLower.includes('sagpj') ||
      fromLower.includes('mfl-info@') ||         
      fromLower.includes('google.com') ||     
      fromLower.includes('nicho.co.jp') ||     
      fromLower.includes('alerts@') ||           
      fromLower.includes('mailer-daemon') ||     
      fromLower.includes('minderu.com') ||         
      subjectLower.includes('セキュリティ') ||     
      subjectLower.includes('アラート') ||         
      subjectLower.includes('ログイン') ||
      subjectLower.includes('message from') ||     
      subjectLower.includes('scanned from') ||     
       subjectLower.includes('ingage.jp') ||     
      subjectLower.includes('【要対応】') ||
      subjectLower.includes('【要確認】') ||
      subjectLower.includes('お問い合わせスタッフ送信') ||
      subjectLower.includes('capsシフト') ||
      // ★★★【重要修正】ここに残っていた「通勤経路」「交通費申請」を消し去り、自爆誤検知を解消 ★★★
      subjectLower.includes('【jinjer勤怠】') ||
      subjectLower.includes('【ジンジャー】') ||
      subjectLower.includes('招待:') || 
      subjectLower.includes('invitation:')
  ) {
    return 'システム通知';
  }

  // 2. 紹介会社（判明しているドメイン）
  if (fromLower.includes('m3career.com') || 
      fromLower.includes('medical-principle.co.jp') || 
      fromLower.includes('mstage-corp.jp') || 
      fromLower.includes('medrt.com') ||
      fromLower.includes('sogo-medical.co.jp') || 
      fromLower.includes('mynavi.jp') ||           
      fromLower.includes('mediwel.net') ||           
      fromLower.includes('prima-support-service.com') ||
     fromLower.includes('mnys.jp') ||
      fromLower.includes('jitsugenya.biz') ||
      fromLower.includes('asiantms.com') || 
      fromLower.includes('nexway.co.jp') 
  ) {
    return '紹介会社';
  }

  // 2.5 いたちごっこ防止の汎用キーワード検知 (From, 件名, 本文)
  if (fromLower.includes('株式会社') || 
      fromLower.includes('(株)') || 
      fromLower.includes('㈱') || 
      subjectLower.includes('紹介会社') || 
      subjectLower.includes('募集状況') ||
      subjectLower.includes('人材紹介') ||
      bodyLower.includes('株式会社') ||  
      bodyLower.includes('(株)') ||      
      bodyLower.includes('（株）') ||  
      bodyLower.includes('㈱')            
  ) {
    return '紹介会社・営業（キーワード検知）';
  }

  // 3. 社内・自社送信
  if (fromLower.includes('caps365.jp') || 
      fromLower.includes('mnys.jp') ||
      fromLower.includes('medicalfitness.co.jp')) {
    
    // ※重要※ メーリングリスト経由の医師からのメール ("〜 via CAPSグループ") は個人として扱う
    if (fromLower.includes('via')) {
       return '個人候補';
    }
    return '社内・自社送信';
  }

  // 4. 上記のブラックリストをすべてすり抜けたものを「個人候補（先生）」とする
  return '個人候補';
}