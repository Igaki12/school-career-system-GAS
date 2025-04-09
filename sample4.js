// --- 定数設定 ---
const SHEET_NAME = "フォームの回答 1"; // フォームの回答が記録されるシート名（実際のシート名に合わせてください）
const SETTINGS_SHEET_NAME = "設定";   // 設定シート名
const HEADERS_TO_ADD = [
  "再提出を求める",
  "担任の確認",
  "進路部の確認",
  "事務室での受領",
  "調査書作成"
];

// --- 列インデックス（フォームの項目に合わせて調整） ---
// フォームから来る列（A=1, B=2 ...）
const TIMESTAMP_COL = 1;
const STUDENT_EMAIL_COL = 2; // 「メールアドレスの入力」列
const CLASS_COL = 3;         // 「クラス」列
const STUDENT_NUMBER_COL = 4; // 「出席番号」列
const STUDENT_NAME_COL = 5;    // 「名前」列
// GASで追加する列（フォーム列の最後に続く）
// ↓↓↓ 実際のフォームの列数に合わせてインデックスを調整してください ↓↓↓
const LAST_FORM_COLUMN_INDEX = 12; // 例: フォームの最後の列がL列(12)の場合
const RETAKE_COL = LAST_FORM_COLUMN_INDEX + 1; // 再提出を求める
const TANIN_CHECK_COL = LAST_FORM_COLUMN_INDEX + 2; // 担任の確認
const SHINRO_CHECK_COL = LAST_FORM_COLUMN_INDEX + 3; // 進路部の確認
const JIMU_CHECK_COL = LAST_FORM_COLUMN_INDEX + 4; // 事務室での受領
const CREATE_DOC_COL = LAST_FORM_COLUMN_INDEX + 5; // 調査書作成
// ↑↑↑ 実際のフォームの列数に合わせてインデックスを調整してください ↑↑↑

// --- インストール可能な onEdit トリガー ---
// この関数をトリガーとして設定してください（編集 -> 現在のプロジェクトのトリガー -> トリガーを追加）
// イベントソース: スプレッドシートから / イベントの種類: 編集時
function installedOnEdit(e) {
  try {
    // 編集イベントがなければ何もしない
    if (!e || !e.range) {
      // Logger.log("編集イベントがありませんでした。");
      return;
    }

    const sheet = e.range.getSheet();
    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    const editedValue = e.value; // 編集後の値
    const oldValue = e.oldValue; // 編集前の値

    // --- 処理対象シートか確認 ---
    if (sheet.getName() !== SHEET_NAME) {
      // Logger.log(`シート名が ${SHEET_NAME} ではないため処理をスキップ: ${sheet.getName()}`);
      return;
    }

    // --- ヘッダー行の編集は無視 ---
    if (editedRow === 1) {
      // Logger.log("ヘッダー行の編集のため処理をスキップ");
      return;
    }

    // --- ヘッダーが存在するか確認し、なければ追加 ---
    initializeHeaders(sheet);

    // --- 設定シートからメールアドレスを取得 ---
    const emailAddresses = getEmailAddresses();
    if (!emailAddresses) {
      Logger.log("設定シートからメールアドレスを取得できませんでした。処理を中断します。");
      sendErrorNotification(new Error("設定シートからメールアドレスを取得できませんでした。"), emailAddresses); // 管理者に通知
      return;
    }

    // --- メール通知処理の実行 ---
    sendEmailNotifications(editedRow, editedCol, sheet, emailAddresses, editedValue, oldValue);

  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}\n${error.stack}`);
    // エラー発生時に設定シートから取得した管理者アドレスに通知
    const adminEmail = getEmailAddresses()?.admin; // 安全に管理者アドレスを取得試行
    sendErrorNotification(error, adminEmail);
  }
}

// --- ヘッダー初期化関数 ---
function initializeHeaders(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastHeader = headerRow[headerRow.length - 1];

  // 追加すべき最後のヘッダーが存在しない場合のみ追加処理を行う
  if (lastHeader !== HEADERS_TO_ADD[HEADERS_TO_ADD.length - 1]) {
     // 既に存在するフォームのヘッダーの数を取得
    const existingHeaderCount = sheet.getRange("1:1").getValues()[0].filter(String).length;

    if (existingHeaderCount >= LAST_FORM_COLUMN_INDEX) {
       const headersRange = sheet.getRange(1, existingHeaderCount + 1, 1, HEADERS_TO_ADD.length);
       // 既に何か書き込まれているかチェック（念のため）
       if (headersRange.getValues()[0].every(h => h === "")) {
          headersRange.setValues([HEADERS_TO_ADD]);
          Logger.log("ヘッダーを追加しました。");
       }
    } else {
       Logger.log(`フォームの列数が想定(${LAST_FORM_COLUMN_INDEX})より少ない(${existingHeaderCount})ため、ヘッダー追加をスキップしました。LAST_FORM_COLUMN_INDEXを確認してください。`);
       // 必要ならここでエラー通知
    }
  }
}


// --- 設定シートからメールアドレスを取得する関数 ---
function getEmailAddresses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      throw new Error(`シート名 "${SETTINGS_SHEET_NAME}" が見つかりません。`);
    }
    const data = settingsSheet.getDataRange().getValues();
    const emails = {};
    let adminEmail = null; // 管理者アドレスを別途保持

    // ヘッダー行をスキップ (i=1 から開始)
    for (let i = 1; i < data.length; i++) {
      const roleOrClass = data[i][0]; // A列
      const email = data[i][1];     // B列
      if (roleOrClass && email) {
        emails[roleOrClass] = email;
        if (roleOrClass.toLowerCase() === '管理者') { // 管理者キーを小文字で比較
          adminEmail = email;
        }
      }
    }
    // 管理者アドレスが見つからなければエラー
    if (!adminEmail && !emails['管理者']) {
        throw new Error(`設定シートに「管理者」のメールアドレスが見つかりません。`);
    }
    emails.admin = adminEmail || emails['管理者']; // admin プロパティにも設定

    // 進路部、事務室のアドレス存在チェック
    if (!emails['進路部']) Logger.log("警告: 設定シートに「進路部」のメールアドレスが見つかりません。");
    if (!emails['事務室']) Logger.log("警告: 設定シートに「事務室」のメールアドレスが見つかりません。");

    return emails;
  } catch (error) {
    Logger.log(`設定シート読み込みエラー: ${error.message}`);
    sendErrorNotification(error, null); // この時点では管理者アドレス不明の可能性あり
    return null;
  }
}

// --- メール通知処理関数 ---
function sendEmailNotifications(editedRow, editedCol, sheet, emailAddresses, editedValue, oldValue) {
  // 編集された行のデータを取得
  const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 生徒情報
  const studentClass = rowData[CLASS_COL - 1];
  const studentNumber = rowData[STUDENT_NUMBER_COL - 1];
  const studentName = rowData[STUDENT_NAME_COL - 1];
  const studentEmail = rowData[STUDENT_EMAIL_COL - 1];
  const studentInfo = `${studentClass} ${studentNumber}番 ${studentName}さん`;

  // 担任のメールアドレスを取得
  const taninEmail = emailAddresses[studentClass];
  // 進路部、事務室のメールアドレス
  const shinroEmail = emailAddresses['進路部'];
  const jimuEmail = emailAddresses['事務室'];

  // --- 条件1: 担任確認列に入力 → 進路部へメール ---
  if (editedCol === TANIN_CHECK_COL && editedValue && !oldValue) { // 空でなくなり、かつ以前は空だった場合
    if (shinroEmail) {
      const subject = `【要確認】調査書作成願 担任確認完了 (${studentInfo})`;
      const body = `以下の生徒の調査書作成願について、担任の確認が完了しました。\n\n`
                 + `生徒: ${studentInfo}\n`
                 + `担任確認内容: ${editedValue}\n\n`
                 + `進路部での確認をお願いします。\n`
                 + `該当スプレッドシート:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
      try {
        GmailApp.sendEmail(shinroEmail, subject, body);
        Logger.log(`進路部へメール送信 (${studentInfo})`);
      } catch (e) {
        Logger.log(`進路部へのメール送信失敗 (${studentInfo}): ${e.message}`);
        sendErrorNotification(new Error(`進路部へのメール送信失敗 (${studentInfo}): ${e.message}`), emailAddresses.admin);
      }
    } else {
      Logger.log(`進路部のメールアドレスが見つからないため、担任確認完了メールを送信できませんでした (${studentInfo})。`);
    }
  }

  // --- 条件2: 進路部確認列 または 事務室受領列 に入力され、両方に値が入った → 担任へメール ---
  const shinroCheckValue = rowData[SHINRO_CHECK_COL - 1];
  const jimuCheckValue = rowData[JIMU_CHECK_COL - 1];

  // 編集されたのが進路部確認列か事務室受領列であり、かつ両方の列に値が入っている場合
  if ((editedCol === SHINRO_CHECK_COL || editedCol === JIMU_CHECK_COL) && shinroCheckValue && jimuCheckValue) {
      // この編集によって初めて両方に値が入ったかをチェック（片方が元々入っていて、今回もう片方が入力されたケースのみ送信）
      let triggerSend = false;
      if (editedCol === SHINRO_CHECK_COL && editedValue && !jimuCheckValue && rowData[JIMU_CHECK_COL-1]) { // 今回進路列が入力され、事務室列は既に入力済み
          // このロジックだと「両方に値が入った」タイミングで送れない。編集前の値(oldValue)と現在の値(rowData)を比較する必要がある
          // oldValue を使って判定するほうがシンプルか？
          // 今回進路列が編集され(editedValueあり)、事務室列にも値がある(jimuCheckValueあり)。
          // かつ、編集前の進路列は空だった(!oldValue)場合に送信。
          if (editedCol === SHINRO_CHECK_COL && editedValue && jimuCheckValue && !oldValue) triggerSend = true;
          // 今回事務室列が編集され(editedValueあり)、進路部列にも値がある(shinroCheckValueあり)。
          // かつ、編集前の事務室列は空だった(!oldValue)場合に送信。
          if (editedCol === JIMU_CHECK_COL && editedValue && shinroCheckValue && !oldValue) triggerSend = true;
          // 両方とも既に値が入っている状態で片方が編集された場合は送らないようにする
          // → 上記条件で既にカバーされているはず

      }
       // 編集されたのが進路部確認列で、編集後の値があり、事務室列にも値がある。かつ編集前の進路部列は空だった。
      if (editedCol === SHINRO_CHECK_COL && editedValue && jimuCheckValue && !oldValue) triggerSend = true;
      // 編集されたのが事務室受領列で、編集後の値があり、進路部列にも値がある。かつ編集前の事務室列は空だった。
      if (editedCol === JIMU_CHECK_COL && editedValue && shinroCheckValue && !oldValue) triggerSend = true;


      if (triggerSend) {
          if (taninEmail) {
              const subject = `【確認依頼】調査書作成願 進路部・事務室確認完了 (${studentInfo})`;
              const body = `以下の生徒の調査書作成願について、進路部および事務室での確認・受領が完了しました。\n\n`
                         + `生徒: ${studentInfo}\n`
                         + `進路部確認内容: ${shinroCheckValue}\n`
                         + `事務室受領内容: ${jimuCheckValue}\n\n`
                         + `内容を確認し、調査書作成を進めてください。\n`
                         + `該当スプレッドシート:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
              try {
                GmailApp.sendEmail(taninEmail, subject, body);
                Logger.log(`担任へ進捗メール送信 (${studentInfo})`);
              } catch (e) {
                Logger.log(`担任への進捗メール送信失敗 (${studentInfo}): ${e.message}`);
                sendErrorNotification(new Error(`担任への進捗メール送信失敗 (${studentInfo}): ${e.message}`), emailAddresses.admin);
              }
          } else {
              Logger.log(`${studentClass} の担任メールアドレスが見つからないため、進捗メールを送信できませんでした (${studentInfo})。`);
          }
      }
  }


  // --- 条件3: 全ての確認列に入力 → 生徒へメール ---
  const taninCheckValue = rowData[TANIN_CHECK_COL - 1];
  // shinroCheckValue, jimuCheckValue は上で取得済み
  const createDocValue = rowData[CREATE_DOC_COL - 1];

  // 編集されたのが N, O, P, Q 列のいずれかで、かつ全ての列に値が入っている場合
  if ([TANIN_CHECK_COL, SHINRO_CHECK_COL, JIMU_CHECK_COL, CREATE_DOC_COL].includes(editedCol) &&
      taninCheckValue && shinroCheckValue && jimuCheckValue && createDocValue) {

      // この編集によって初めて全ての列に値が入ったかをチェック
      let triggerSendToStudent = false;
      // 編集前の値(oldValue)が空で、編集後の値(editedValue)があり、他の必須列にも全て値が入っている場合
      if (!oldValue && editedValue) {
          triggerSendToStudent = true;
      }

      if (triggerSendToStudent) {
          if (studentEmail) {
              const subject = `調査書作成完了のお知らせ`;
              const body = `${studentInfo}\n\n`
                         + `申請いただいた調査書が作成完了しました。\n`
                         + `担任の先生から受け取ってください。`;
              try {
                GmailApp.sendEmail(studentEmail, subject, body);
                Logger.log(`生徒へ完了メール送信 (${studentInfo})`);
              } catch (e) {
                Logger.log(`生徒への完了メール送信失敗 (${studentInfo}): ${e.message}`);
                sendErrorNotification(new Error(`生徒への完了メール送信失敗 (${studentInfo}): ${e.message}`), emailAddresses.admin);
              }
          } else {
              Logger.log(`生徒のメールアドレスが見つからないため、完了メールを送信できませんでした (${studentInfo})。`);
          }
      }
  }
}

// --- エラー通知関数 ---
function sendErrorNotification(error, adminEmail) {
  // 管理者アドレスが取得できていない場合はデフォルトアドレスを設定（または処理中断）
  const recipient = adminEmail || Session.getEffectiveUser().getEmail(); // デフォルトとしてスクリプト実行者
  // もし adminEmail が null で、Session 情報も取得できない場合は送信しない
   if (!recipient) {
       Logger.log("エラー通知の送信先が不明なため、通知をスキップします。");
       Logger.log(`エラー詳細: ${error.message}\n${error.stack}`);
       return;
   }


  try {
    const subject = "【GASエラー通知】進路調査票スプレッドシート処理";
    const body = `進路調査票スプレッドシートを処理するGASでエラーが発生しました。\n\n`
               + `エラー日時: ${new Date().toLocaleString('ja-JP')}\n`
               + `エラーメッセージ: ${error.message}\n`
               + `スタックトレース:\n${error.stack || 'スタックトレースなし'}\n\n`
               + `スプレッドシートURL:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`エラー通知を ${recipient} に送信しました。`);
  } catch (e) {
    // エラー通知メールの送信自体に失敗した場合
    Logger.log(`エラー通知メールの送信に失敗しました: ${e.message}`);
    // ここでさらに何かするなら記述（例：別の通知手段、ログレベル変更など）
  }
}