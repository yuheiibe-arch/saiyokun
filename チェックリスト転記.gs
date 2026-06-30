/**
 * @fileoverview 【新規】フォーム送信時に、チェックリストへ新しい記録を追記するスクリプト
 * 対応フォーム：採用報告, DS承認報告, 不採用報告
 */

// --- ▼ 設定項目 ▼ ---
// (このセクションの定義は「全体設定.gs」を参照します)
// --- ▲ 設定項目 ▲ ---

/**
 * チェックリスト転記用の、安全に値を取得するヘルパー関数
 */
function getValueForChecklist(values, key) {
  const value = values[key] ? values[key][0] : '';
  return String(value).trim();
}

/**
 * 西暦を和暦に変換するヘルパー関数
 */
function convertToJapaneseYearForChecklist(year) {
  if (!year || isNaN(year)) return '';
  if (year >= 2019) return `令和${year - 2018}年`;
  if (year >= 1989) return `平成${year - 1988}年`;
  if (year >= 1926) return `昭和${year - 1925}年`;
  return `${year}年`;
}

/**
 * チェックリスト自動転記のメイン関数
 * (★大司令塔 onFormSubmit_Master から呼び出される)
 */
function autoRecordOnFormSubmit(e) {
  try {
    const { namedValues, range } = e;
    const formSheetName = range.getSheet().getName();

    // フォームタイプの判定
    let formType = '';
    if (formSheetName.includes('不採用報告')) {
      formType = '不採用報告';
    } else if (formSheetName.includes('採用報告')) {
      formType = '採用報告';
    } else if (formSheetName.includes('DS承認報告')) {
      formType = 'DS承認報告';
    } else {
      return; // 対象外のシートなら終了
    }

    const name = getValueForChecklist(namedValues, '氏名');
    if (!name) return;

    const currentSs = SpreadsheetApp.getActiveSpreadsheet();
    const editSheet = currentSs.getSheetByName(CONFIG_FOR_CHECKLIST.EDIT_SHEET_NAME);
    const searchValues = editSheet.getRange(1, 1, editSheet.getLastRow(), 1).getValues();
    let url = '';
    for (let i = 0; i < searchValues.length; i++) {
      if (searchValues[i][0] == CONFIG_FOR_CHECKLIST.CHECKLIST_KEYWORD) {
        url = editSheet.getRange(i + 1, 2).getValue();
        break;
      }
    }
    if (!url) throw new Error('チェックリストのURLが見つかりません。');
    
    const targetSs = SpreadsheetApp.openByUrl(url);
    const targetSheet = targetSs.getSheetByName(CONFIG_FOR_CHECKLIST.TARGET_SHEET_NAME);
    if (!targetSheet) throw new Error(`シート「${CONFIG_FOR_CHECKLIST.TARGET_SHEET_NAME}」が見つかりません。`);

    const firstEmptyRow = findFirstEmptyRowForChecklist(targetSheet, 2, CONFIG_FOR_CHECKLIST.START_ROW);
    
    // ★修正: targetSheet を第4引数として渡す
    const dataToWrite = createRowDataForChecklist(formType, namedValues, currentSs, targetSheet);
    writeDataToRowForChecklist(targetSheet, firstEmptyRow, dataToWrite);

  } catch (error) {
    console.error(`自動転記スクリプトでエラー: ${error.toString()}`);
    console.error(error.stack);
  }
}

function findFirstEmptyRowForChecklist(sheet, col, startRow) {
    const values = sheet.getRange(startRow, col, sheet.getMaxRows() - startRow + 1, 1).getValues();
    for(let i = 0; i < values.length; i++) {
        if(values[i][0] === '') {
            return startRow + i;
        }
    }
    return sheet.getLastRow() + 1;
}

/**
 * フォームのタイプに応じて書き込むデータを作成する関数
 * ★修正: 「採用可否」列を動的に探し、不採用理由に応じたプルダウン連携を追加
 */
function createRowDataForChecklist(formType, values, currentSs, targetSheet) {
  let data = {};
  const today = new Date();

  const fullName = getValueForChecklist(values, '氏名');
  data['B'] = fullName.replace(/[(（].*?[)）]/g, '').trim();
  data['BB'] = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
  data['BC'] = Utilities.formatDate(today, 'JST', 'yyyy/MM');
  
  // ★追加：ターゲットシートから「採用可否」の列を動的に取得する（見つからなければデフォルトC列）
  const saiyoKahiCol = getColumnLetterByHeader(targetSheet, '採用可否') || 'C'; 
  
  // === 1. 採用報告 ===
  if (formType === '採用報告') {
    data[saiyoKahiCol] = '採用';
    data['D'] = 'スポット';
    data['E'] = getValueForChecklist(values, '紹介会社');
    const shinryoka = getValueForChecklist(values, '診療科');
    if (shinryoka.includes('小児科')) data['J'] = '小児科';
    else if (shinryoka.includes('内科')) data['J'] = '内科';
    const kinmureki = getValueForChecklist(values, '当法人での勤務歴');
    data['L'] = (kinmureki === 'なし') ? '勤務歴なし' : kinmureki;
    
    // U, V, Y列を記載する
    const kinmuDateRaw = getValueForChecklist(values, '勤務日');
    if (kinmuDateRaw) data['U'] = Utilities.formatDate(new Date(kinmuDateRaw), 'JST', 'yyyy/MM/dd');
    const kyoten = getValueForChecklist(values, '拠点');
    data['V'] = kyoten;
    if (CONFIG_FOR_CHECKLIST.KANSAI_LOCATIONS.includes(kyoten)) data['Y'] = '関西';
    else if (kyoten) data['Y'] = '関西以外';

  // === 2. 不採用報告 (修正済み) ===
  } else if (formType === '不採用報告') {
    // ★修正：フォームの「【不採用理由】」を取得し、プルダウン用の値に変換
    const reason = getValueForChecklist(values, '【不採用理由】');
    let mappedStatus = '';
    
    if (reason === '基準に満たない') {
      mappedStatus = '不採用（基準未達）';
    } else if (reason === 'クレーム歴') {
      mappedStatus = '不採用（ネガティブあり）';
    } else if (reason === '充足') {
      mappedStatus = '不採用（充足のため）';
    }

    data[saiyoKahiCol] = mappedStatus; // 動的に見つけた列に変換後のステータスをセット
    
    data['D'] = 'スポット';
    data['E'] = getValueForChecklist(values, '紹介会社');
    const shinryoka = getValueForChecklist(values, '診療科');
    if (shinryoka.includes('小児科')) data['J'] = '小児科';
    else if (shinryoka.includes('内科')) data['J'] = '内科';
    const kinmureki = getValueForChecklist(values, '当法人での勤務歴');
    data['L'] = (kinmureki === 'なし') ? '勤務歴なし' : kinmureki;

    // ★修正：勤務日(U), 拠点(V), エリア(Y) の書き込み処理を削除しました

  // === 3. DS承認報告 ===
  } else if (formType === 'DS承認報告') {
    data[saiyoKahiCol] = '採用';
    data['D'] = 'スポット';
    data['E'] = '直接DS応募';
    data['F'] = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
    data['I'] = '承認';
    const senmoni = getValueForChecklist(values, '専門医資格（任意）');
    if (senmoni.includes('小児科')) data['J'] = '小児科';
    else if (senmoni.includes('内科')) data['J'] = '内科';
    const licenseYearAD = getValueForChecklist(values, '医師免許取得年');
    data['S'] = convertToJapaneseYearForChecklist(parseInt(licenseYearAD, 10));
    const urlSheet = currentSs.getSheetByName(CONFIG_FOR_CHECKLIST.URL_SHEET_NAME);
    if (urlSheet) {
        const submitterName = urlSheet.getRange(CONFIG_FOR_CHECKLIST.SUBMITTER_CELL).getValue();
        if (submitterName) {
            const surname = submitterName.split(/[\s ]/)[0];
            data['AI'] = `${Utilities.formatDate(today, 'JST', 'M/d')}${surname}`;
        }
    }
  }
  return data;
}

function writeDataToRowForChecklist(sheet, row, data) {
  for (const colStr in data) {
    if (data[colStr] !== undefined && data[colStr] !== null) {
      sheet.getRange(`${colStr}${row}`).setValue(data[colStr]);
    }
  }
}

/**
 * ---------------------------------------------------------
 * ▼▼▼ 以下2つは動的列検索用の新規ヘルパー関数です ▼▼▼
 * ---------------------------------------------------------
 */

/**
 * シートの1〜3行目を検索し、指定したヘッダー名の列アルファベットを動的に返す
 */
function getColumnLetterByHeader(sheet, headerName) {
  if (!sheet) return null;
  const headers = sheet.getRange(1, 1, 3, sheet.getLastColumn()).getValues();
  for (let r = 0; r < headers.length; r++) {
    for (let c = 0; c < headers[r].length; c++) {
      if (headers[r][c] === headerName) {
        return convertColIndexToLetter_internal(c + 1);
      }
    }
  }
  return null;
}

/**
 * 列番号(例: 3)をアルファベット(例: C)に変換する
 */
function convertColIndexToLetter_internal(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}