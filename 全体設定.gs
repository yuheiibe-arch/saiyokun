// =================================================================
// ▼▼▼ 【全体設定】すべての設定をここにまとめる ▼▼▼
// =================================================================

// --- 1. シート名（メール自動処理 ＆ 直前応募 共通）---
const SHEET_NAME_PROGRESS = '紹介会社進捗';
const SHEET_NAME_ARCHIVE = '処理済み';
const LOG_SHEET_NAME = '通知履歴';
const URGENT_SHEET_NAME = '直前通知';
const CANCEL_LOG_SHEET_NAME = 'キャンセルログ';

// ★★★【重要】「紹介会社進捗」シートは名前を統一します
const AGENCY_SOURCE_SHEET_NAME = SHEET_NAME_PROGRESS;

// --- 2. シート名（フォーム・URL処理用）---
const URL_SHEET_NAME = 'URL';
const MENTION_SHEET_NAME = 'メンションリスト';

// --- 3. シート名（チェックリスト転記用）---
const CONFIG_FOR_CHECKLIST = {
  EDIT_SHEET_NAME: 'Edit', 
  CHECKLIST_KEYWORD: 'チェックリスト',
  TARGET_SHEET_NAME: 'チェックリスト',
  START_ROW: 1458,
  KANSAI_LOCATIONS: ['東岸和田', '阿波座', '天美', '堺鉄砲町', '長吉長原', '鶴見緑地','あべの', '豊中', '今里', '千里丘'],
  SUBMITTER_CELL: 'C3',
  URL_SHEET_NAME: URL_SHEET_NAME
};

// --- 4. Chatwork ルームID ---
const URGENT_APPLY_CHATWORK_ROOM_ID = '415381235';
const SAIYO_HOKOKU_CHATWORK_ROOM_ID = '165593914';
const CANCEL_LOG_CHATWORK_ROOM_ID = '344086172';
const TIMEGAI_CALL_CHATWORK_ROOM_ID = '415381235';

// --- 5. Gmail関連 ---
// (ファイル1：メール自動処理用)
const PROCESSED_LABEL = '紹介会社処理済み'; 
const CHECKBOX_COLUMN = 12; // ★K列(11)からL列(12)に変更

// (直前応募：part1用)
const PROCESSED_LABEL_APPLY = '処理済み-医師シフト応募';
const GMAIL_QUERY_APPLY = `subject:(("【要対応】★医師シフト応募通知★") OR ("募集シフトへの応募がありました。")) newer_than:1d`;

// (キャンセル・時間外用)
const PROCESSED_LABEL_TIMEGAI = '処理済み-時間外着信';
const PROCESSED_LABEL_CANCEL = '処理済み-キャンセル'; 
const GMAIL_QUERY_CANCEL = `subject:(("★★緊急★★＜勤務2週間以内のキャンセル申請＞") OR ("【要対応】勤務3日以内のキャンセル申請")) -label:${PROCESSED_LABEL_CANCEL}`; 
const GMAIL_QUERY_TIMEGAI = 'subject:(("【みんなにでんわ転送】時間外応答がありました") OR ("【みんなにでんわ転送】不在着信がありました"))'; 
const GMAIL_SENDER_EXCLUDE = 'no-reply-staging';

// --- 6. 紹介会社シート関連（直前応募：part2用）---
const AGENCY_STATUS_COLUMN_LETTER = 'O'; // ★N列からO列に変更
const AGENCY_STATUS_PROCESSED = '通知済み';

// --- 7. ドタキャン（非定型キャンセル）検知用 ---
const DOTAKYAN_LOG_SHEET_NAME = 'ドタキャンログ';
const DOTAKYAN_CONFIG_SHEET_NAME = 'ドタキャン';
const PROCESSED_LABEL_DOTAKYAN = '処理済み-ドタキャン';
const DOTAKYAN_CHATWORK_ROOM_ID = '415529974';

// --- 8. 個人メール通知用（新規追加） ---
const PERSONAL_MAIL_CHATWORK_ROOM_ID = '434845563'; 
const PROCESSED_LABEL_PERSONAL = '処理済み-個人連絡';
const LOG_SHEET_NAME_PERSONAL = '個人メールログ';