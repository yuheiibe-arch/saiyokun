function runRecoveryNow() {
  console.log("=============== 🚨 検知漏れ一斉回収 本番実行 =============== ");
  
  // 1. 直前応募の回収（吉冨先生の 南浦和 9/25 の件など）
  console.log("\n▶︎ 直前応募の回収を開始...");
  executeAllProcesses_internal();
  
  // 2. 個人メールの回収（吉冨先生の 通勤経路・複数拠点の件）
  console.log("\n▶︎ 個人メールの回収を開始...");
  checkPersonalEmails_internal();
  
  console.log("\n=============== 🚨 回収とChatwork通知がすべて完了しました ===============");
}