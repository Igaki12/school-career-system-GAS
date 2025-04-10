/**
 * スプレッドシートの任意のシートでA1セルが編集されたときに、
 * その内容を指定されたメールアドレスに送信するテスト関数。
 * この関数を onEdit トリガーとして設定するか、
 * 手動で実行（その場合は e オブジェクトがないため簡易的な動作）します。
 */
function testSendMailOnEdit(e) {
    // --- 設定 ---
    const TARGET_CELL = "A1"; // 監視するセル
    const RECIPIENT_EMAIL = "igatatsu1997@gmail.com"; // 送信先メールアドレス
  
    try {
      // --- トリガー経由での実行かチェック ---
      if (!e || !e.range) {
        Logger.log("この関数はスプレッドシートの編集時に自動実行されることを想定しています。(eオブジェクトがありません)");
        // 手動実行時などの簡易テスト
        const manualValue = "手動実行テスト";
        GmailApp.sendEmail(RECIPIENT_EMAIL, "GAS Gmail送信テスト (手動実行)", `これは手動実行によるテストメールです。\n値: ${manualValue}`);
        Logger.log(`テストメールを ${RECIPIENT_EMAIL} に送信しました（手動実行）。`);
        return;
      }
  
      // --- 編集情報の取得 ---
      const range = e.range;
      const sheet = range.getSheet();
      const editedCellA1Notation = range.getA1Notation();
      const editedValue = e.value; // 編集後の値
  
      // --- A1セルが編集されたかチェック ---
      if (editedCellA1Notation === TARGET_CELL) {
        Logger.log(`${TARGET_CELL} が編集されました。値: ${editedValue}`);
  
        // --- メールの準備 ---
        const subject = `GAS Gmail送信テスト (${sheet.getName()}シート)`;
        const body = `スプレッドシートの ${TARGET_CELL} セルが編集されました。\n\n`
                   + `シート名: ${sheet.getName()}\n`
                   + `セル番地: ${editedCellA1Notation}\n`
                   + `入力された値: ${editedValue || '(空欄)'}\n\n` // 値がない場合も考慮
                   + `スプレッドシートURL:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
        // --- メール送信 ---
        GmailApp.sendEmail(RECIPIENT_EMAIL, subject, body);
        Logger.log(`テストメールを ${RECIPIENT_EMAIL} に送信しました。`);
  
      } else {
        // Logger.log(`${TARGET_CELL} 以外のセル (${editedCellA1Notation}) が編集されたため、メールは送信しません。`);
      }
  
    } catch (error) {
      // --- エラー処理 ---
      Logger.log(`エラーが発生しました: ${error.message}\n${error.stack}`);
      // エラー発生時にも通知を試みる（ただし、Gmail送信自体に問題がある場合はこれも失敗する可能性あり）
      try {
        const errorSubject = "【GASエラー】Gmail送信テストコード";
        const errorBody = `Gmail送信テスト用のGASでエラーが発生しました。\n\n`
                        + `エラーメッセージ: ${error.message}\n`
                        + `スタックトレース:\n${error.stack || 'スタックトレースなし'}`;
         // エラー通知はスクリプト実行者自身に送るのが一般的
         const ownerEmail = Session.getEffectiveUser().getEmail() || RECIPIENT_EMAIL; // 実行者 or 指定アドレス
         GmailApp.sendEmail(ownerEmail, errorSubject, errorBody);
         Logger.log(`エラー通知を ${ownerEmail} に送信しようとしました。`);
      } catch (notificationError) {
         Logger.log(`エラー通知メールの送信にも失敗しました: ${notificationError.message}`);
      }
    }
  }