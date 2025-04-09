/**
 * @OnlyCurrentDoc
 */

// --- グローバル設定 ---
// スプレッドシートや列の設定は「設定」シートから読み込む
const SETTINGS_SHEET_NAME = '設定';

/**
 * 設定シートから設定値を取得する関数
 * @returns {object} 設定値を格納したオブジェクト
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    throw new Error(`設定シート「${SETTINGS_SHEET_NAME}」が見つかりません。`);
  }
  const settingsData = settingsSheet.getDataRange().getValues();
  const settings = {};
  // ヘッダー行(1行目)をスキップして2行目から読み込む
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] && settingsData[i][1] !== '') {
      settings[settingsData[i][0]] = settingsData[i][1];
    }
  }
  // 列番号は数値に変換
  const colKeys = [
    'タイムスタンプ列番号', 'メールアドレス列番号', 'クラス列番号', '出席番号列番号', '名前列番号',
    '再提出列番号', '担任確認列番号', '進路部確認列番号', '事務室受領列番号', '調査書作成列番号',
    '担任確認送信済列番号', '進路事務送信済列番号', '生徒通知送信済列番号'
  ];
  colKeys.forEach(key => {
    if (settings[key]) {
      settings[key] = parseInt(settings[key], 10);
      if (isNaN(settings[key])) {
         throw new Error(`設定シートの「${key}」の値が数値ではありません: ${settings[key]}`);
      }
    } else {
      throw new Error(`設定シートに「${key}」が見つかりません。`);
    }
  });

  // 必須設定のチェック
  const requiredKeys = [
    '進路部メールアドレス', '管理者メールアドレス', 'データシート名',
    'A組担任メールアドレス', // 他の組も実際には必要だが、ここでは代表してチェック
    // 列番号関連
    'メールアドレス列番号', 'クラス列番号', '出席番号列番号', '名前列番号',
    '担任確認列番号', '進路部確認列番号', '事務室受領列番号', '調査書作成列番号',
    '担任確認送信済列番号', '進路事務送信済列番号', '生徒通知送信済列番号'
  ];
   requiredKeys.forEach(key => {
    if (!settings[key]) {
      throw new Error(`設定シートに必須項目「${key}」が見つからないか、値が空です。`);
    }
  });

  // 担任メールアドレスをクラス名をキーにしたオブジェクトにまとめる
  settings.tanninEmails = {};
  for (const key in settings) {
    if (key.endsWith('組担任メールアドレス')) {
      const className = key.replace('担任メールアドレス', ''); // 例: "A組"
      settings.tanninEmails[className] = settings[key];
    }
  }
   if (Object.keys(settings.tanninEmails).length === 0) {
     throw new Error('設定シートに担任メールアドレスが見つかりません（例: A組担任メールアドレス）。');
   }

  Logger.log('設定読み込み完了: %s', JSON.stringify(settings, null, 2));
  return settings;
}


/**
 * スプレッドシートが編集されたときに実行されるトリガー関数
 * @param {Object} e イベントオブジェクト
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();

  try {
    const settings = getSettings(); // まず設定を読み込む

    // 設定シートの編集は無視、データシート以外も無視
    if (sheet.getName() === SETTINGS_SHEET_NAME || sheet.getName() !== settings['データシート名']) {
      Logger.log('編集されたシートが対象外 (%s) のため処理をスキップします。', sheet.getName());
      return;
    }

    // ヘッダー行(1行目)の編集は無視
    if (editedRow === 1) {
      Logger.log('ヘッダー行の編集のため処理をスキップします。');
      return;
    }

    // 編集された行のデータを取得 (A列から送信フラグ列まで)
    const maxCol = Math.max(
      settings['調査書作成列番号'],
      settings['担任確認送信済列番号'],
      settings['進路事務送信済列番号'],
      settings['生徒通知送信済列番号']
    );
    const rowDataRange = sheet.getRange(editedRow, 1, 1, maxCol);
    const rowDataValues = rowDataRange.getValues()[0]; // 1行分のデータを配列で取得

    // 列番号からインデックスへの変換 (0始まり)
    const cols = {
      studentEmail: settings['メールアドレス列番号'] - 1,
      className: settings['クラス列番号'] - 1,
      studentNumber: settings['出席番号列番号'] - 1,
      studentName: settings['名前列番号'] - 1,
      tanninCheck: settings['担任確認列番号'] - 1,
      shidoubuCheck: settings['進路部確認列番号'] - 1,
      jimushitsuCheck: settings['事務室受領列番号'] - 1,
      sakuseiCheck: settings['調査書作成列番号'] - 1,
      tanninMailSent: settings['担任確認送信済列番号'] - 1,
      shidouJimuMailSent: settings['進路事務送信済列番号'] - 1,
      studentMailSent: settings['生徒通知送信済列番号'] - 1
    };

    const studentInfo = `${rowDataValues[cols.className]} ${rowDataValues[cols.studentNumber]}番 ${rowDataValues[cols.studentName]}`;
    const editedValue = e.value; // 編集後の値
    const oldValue = e.oldValue; // 編集前の値

    Logger.log('編集イベント発生: Row=%s, Col=%s, Value=%s, OldValue=%s, Sheet=%s', editedRow, editedCol, editedValue, oldValue, sheet.getName());
    Logger.log('対象生徒: %s', studentInfo);


    // --- トリガー1: 担任確認メール ---
    // 担任確認列が編集され、空でなくなり、かつまだ送信していない場合
    if (editedCol === settings['担任確認列番号'] && editedValue && !rowDataValues[cols.tanninMailSent]) {
      Logger.log('担任確認メール送信条件を確認中...');
      sendMailToShidoubu(rowDataValues, cols, studentInfo, editedValue, settings);
      // 送信済みフラグを立てる
      sheet.getRange(editedRow, settings['担任確認送信済列番号']).setValue('送信済 ' + Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm"));
      Logger.log('担任確認メール送信完了、フラグ設定。');
    }

    // --- トリガー2: 進路部・事務室確認メール ---
    // 進路部確認列 または 事務室受領列 が編集され、
    // かつ両方の列に値があり、かつまだ送信していない場合
    const shidoubuCheckValue = (editedCol === settings['進路部確認列番号']) ? editedValue : rowDataValues[cols.shidoubuCheck];
    const jimushitsuCheckValue = (editedCol === settings['事務室受領列番号']) ? editedValue : rowDataValues[cols.jimushitsuCheck];

    if ((editedCol === settings['進路部確認列番号'] || editedCol === settings['事務室受領列番号']) &&
        shidoubuCheckValue && jimushitsuCheckValue && !rowDataValues[cols.shidouJimuMailSent]) {
      Logger.log('進路/事務確認メール送信条件を確認中...');
      sendMailToTannin(rowDataValues, cols, studentInfo, shidoubuCheckValue, jimushitsuCheckValue, settings);
      // 送信済みフラグを立てる
      sheet.getRange(editedRow, settings['進路事務送信済列番号']).setValue('送信済 ' + Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm"));
      Logger.log('進路/事務確認メール送信完了、フラグ設定。');
    }

    // --- トリガー3: 生徒への完了通知メール ---
    // 担任確認、進路部確認、事務室受領、調査書作成 のいずれかの列が編集され、
    // かつ全ての列に値があり、かつまだ送信していない場合
    const tanninCheckValue = (editedCol === settings['担任確認列番号']) ? editedValue : rowDataValues[cols.tanninCheck];
    const sakuseiCheckValue = (editedCol === settings['調査書作成列番号']) ? editedValue : rowDataValues[cols.sakuseiCheck];
    // 他の値は上で取得済み

    if ((editedCol === settings['担任確認列番号'] || editedCol === settings['進路部確認列番号'] ||
         editedCol === settings['事務室受領列番号'] || editedCol === settings['調査書作成列番号']) &&
        tanninCheckValue && shidoubuCheckValue && jimushitsuCheckValue && sakuseiCheckValue &&
        !rowDataValues[cols.studentMailSent]) {
       Logger.log('生徒通知メール送信条件を確認中...');
       sendMailToStudent(rowDataValues, cols, studentInfo, settings);
       // 送信済みフラグを立てる
       sheet.getRange(editedRow, settings['生徒通知送信済列番号']).setValue('送信済 ' + Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm"));
       Logger.log('生徒通知メール送信完了、フラグ設定。');
    }

  } catch (error) {
    Logger.log('エラー発生: %s', error.message);
    Logger.log('スタックトレース: %s', error.stack);
    // エラー発生時に管理者へメール通知
    try {
        const settings = getSettings(); // 設定を再取得試行
        sendErrorNotification(e, error, settings);
    } catch (settingError) {
        // 設定取得自体でエラーが起きた場合、ハードコードしたアドレスに最低限の情報を送る試み
        const adminEmail = "YOUR_FALLBACK_ADMIN_EMAIL@example.com"; // ★★★ フォールバック用管理者アドレスを設定してください ★★★
        const subject = "[緊急エラー] 進路調査票GAS エラー（設定読込失敗）";
        const body = `スプレッドシートの編集処理中にエラーが発生しましたが、設定シートの読み込みにも失敗しました。\n\n` +
                     `編集イベント情報:\n${JSON.stringify(e, null, 2)}\n\n` +
                     `内部エラー: ${error.message}\n${error.stack}\n\n` +
                     `設定読み込みエラー: ${settingError.message}\n${settingError.stack}`;
        if (adminEmail !== "YOUR_FALLBACK_ADMIN_EMAIL@example.com") {
             GmailApp.sendEmail(adminEmail, subject, body);
        }
        Logger.log('設定読み込みエラーのため、フォールバック管理者への通知試行。エラー詳細: %s', settingError.message);
    }
  }
}

/**
 * トリガー1: 担任確認 -> 進路部へメール送信
 */
function sendMailToShidoubu(rowData, cols, studentInfo, tanninCheckValue, settings) {
  const shidoubuEmail = settings['進路部メールアドレス'];
  if (!shidoubuEmail) {
    throw new Error('設定シートに進路部メールアドレスが見つかりません。');
  }

  const subject = `[担任確認完了] ${studentInfo}`;
  const body = `${studentInfo} さんの調査書について、担任確認が完了しました。\n\n` +
               `担任確認欄の入力内容:\n${tanninCheckValue}\n\n` +
               `※このメールはシステムにより自動送信されています。`;

  GmailApp.sendEmail(shidoubuEmail, subject, body);
  Logger.log('進路部へメール送信: To=%s, Subject=%s', shidoubuEmail, subject);
}

/**
 * トリガー2: 進路部確認 & 事務室受領 -> 担任へメール送信
 */
function sendMailToTannin(rowData, cols, studentInfo, shidoubuCheckValue, jimushitsuCheckValue, settings) {
  const className = rowData[cols.className]; // 例: "A組"
  const tanninEmail = settings.tanninEmails[className];

  if (!tanninEmail) {
    throw new Error(`設定シートに ${className} の担任メールアドレスが見つかりません。`);
  }

  const subject = `[進路部・事務室確認完了] ${studentInfo}`;
  const body = `${studentInfo} さんの調査書について、進路部確認と事務室での受領が完了しました。\n\n` +
               `進路部確認欄の入力内容:\n${shidoubuCheckValue}\n\n` +
               `事務室受領欄の入力内容:\n${jimushitsuCheckValue}\n\n` +
               `※このメールはシステムにより自動送信されています。`;

  GmailApp.sendEmail(tanninEmail, subject, body);
  Logger.log('担任へメール送信: To=%s, Subject=%s', tanninEmail, subject);
}

/**
 * トリガー3: 全確認完了 -> 生徒へメール送信
 */
function sendMailToStudent(rowData, cols, studentInfo, settings) {
  const studentEmail = rowData[cols.studentEmail];
  if (!studentEmail) {
    Logger.log('生徒のメールアドレスが見つからないため、生徒への通知メールは送信できません。生徒: %s', studentInfo);
    // エラーにはせず、ログだけ残して処理を続ける場合
     return;
    // エラーとして処理を止めたい場合
    // throw new Error(`生徒のメールアドレスが見つかりません: ${studentInfo}`);
  }

  const subject = `調査書作成完了のお知らせ`;
  const body = `${studentInfo} さん\n\n` +
               `依頼された調査書の作成が完了しました。\n` +
               `担任の先生から受け取ってください。\n\n` +
               `※このメールはシステムにより自動送信されています。心当たりがない場合は、お手数ですが担任または進路指導部までご連絡ください。`;

  GmailApp.sendEmail(studentEmail, subject, body);
  Logger.log('生徒へメール送信: To=%s, Subject=%s', studentEmail, subject);
}

/**
 * エラー発生時に管理者にメールで通知する関数
 * @param {Object} eventObject - トリガーイベントオブジェクト (e)
 * @param {Error} errorObject - 発生したエラーオブジェクト
 * @param {Object} settings - 読み込み済みの設定オブジェクト
 */
function sendErrorNotification(eventObject, errorObject, settings) {
  const adminEmail = settings ? settings['管理者メールアドレス'] : null;
  if (!adminEmail) {
    Logger.log('管理者メールアドレスが設定されていないため、エラー通知を送信できません。');
    // フォールバック用のアドレスを使うことも検討できる
    // const fallbackAdmin = "YOUR_FALLBACK_ADMIN_EMAIL@example.com";
    // if (fallbackAdmin !== "YOUR_FALLBACK_ADMIN_EMAIL@example.com") {
    //    GmailApp.sendEmail(fallbackAdmin, ...);
    // }
    return; // 送信先がないので終了
  }

  try {
    let subject = '[要確認] 進路調査票GAS エラー発生通知';
    let body = `進路調査票スプレッドシートの自動処理中にエラーが発生しました。\n\n` +
               `発生日時: ${new Date().toLocaleString('ja-JP')}\n`;

    if (eventObject && eventObject.range) {
      const sheet = eventObject.range.getSheet();
      const cell = eventObject.range.getA1Notation();
      body += `発生箇所: シート「${sheet.getName()}」のセル「${cell}」付近の編集時\n`;
      body += `編集後の値: ${eventObject.value}\n`;
      body += `編集前の値: ${eventObject.oldValue}\n`;
    } else {
      body += `発生箇所: 不明 (編集イベント外の可能性)\n`;
    }

    body += `\nエラーメッセージ:\n${errorObject.message}\n\n` +
            `エラー詳細 (スタックトレース):\n${errorObject.stack}\n\n` +
            `※このメールはシステムにより自動送信されています。エラーの原因を確認し、必要に応じて修正してください。`;

    GmailApp.sendEmail(adminEmail, subject, body);
    Logger.log('管理者へエラー通知メールを送信しました: To=%s', adminEmail);

  } catch (notificationError) {
    // エラー通知メールの送信自体に失敗した場合
    Logger.log('エラー通知メールの送信に失敗しました: %s', notificationError.message);
    // ここでさらにフォールバック処理を入れることも可能だが、複雑になるためログ出力に留める
  }
}