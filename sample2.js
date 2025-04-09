// --- グローバル定数 ---
// ★★★ 実際のフォーム回答シート名に合わせてください ★★★
const FORM_SHEET_NAME = 'フォームの回答 1';
const SETTINGS_SHEET_NAME = '設定';

// 各列の列番号 (A=1, B=2, ...)
const COLUMNS = {
  TIMESTAMP: 1,
  STUDENT_EMAIL: 2, // B列: 生徒のメールアドレス
  CLASS: 3,         // C列: クラス
  STUDENT_ID: 4,    // D列: 出席番号
  STUDENT_NAME: 5,  // E列: 名前
  // ... (既存のフォーム項目列)
  QUALIFICATION: 13, // M列: 受験資格
  TEACHER_CHECK: 14, // N列: 担任の確認
  CAREER_CHECK: 15,  // O列: 進路部の確認
  OFFICE_RECEIPT: 16,// P列: 事務室での受領
  DOC_CREATION: 17,  // Q列: 調査書作成
  // MAIL_FLAG は設定シートから取得
};

// --- メインの編集トリガー関数 ---
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const editedValue = e.value; // 編集後の値

  // --- 事前チェック ---
  // ヘッダー行(1行目)や対象外シートの編集は無視
  if (editedRow <= 1 || sheet.getName() !== FORM_SHEET_NAME) {
    Logger.log(`無視: シート名(${sheet.getName()}) または 行(${editedRow}) が対象外`);
    return;
  }

  // 編集された列が N, O, P, Q のいずれかでない場合は無視
  const targetColumns = [
    COLUMNS.TEACHER_CHECK,
    COLUMNS.CAREER_CHECK,
    COLUMNS.OFFICE_RECEIPT,
    COLUMNS.DOC_CREATION
  ];
  if (!targetColumns.includes(editedCol)) {
     Logger.log(`無視: 編集列(${editedCol}) が対象外`);
    return;
  }

  // 空白にされた場合は無視 (値が入力された場合のみ処理)
   if (editedValue === undefined || editedValue === null || editedValue === "") {
     Logger.log(`無視: セル (${editedRow}, ${editedCol}) が空白にされました`);
    return;
  }

  // --- 設定の読み込み ---
  const settings = getSettings();
  if (!settings) {
    Logger.log('エラー: 設定シートの読み込みに失敗しました。処理を中断します。');
    // 必要に応じてUIで通知
    // SpreadsheetApp.getUi().alert('設定シートを正しく設定してください。');
    return;
  }
  const mailFlagColIndex = getColumnIndex(settings.mailFlagColumnLetter);
  if (!mailFlagColIndex) {
      Logger.log(`エラー: 設定シートの「メール送信済みフラグ列」(${settings.mailFlagColumnLetter})が無効です。`);
      return;
  }


  // --- 編集行のデータ取得 ---
  // 必要な列 + フラグ列まで取得
  const rowDataRange = sheet.getRange(editedRow, 1, 1, mailFlagColIndex);
  const rowData = rowDataRange.getValues()[0];

  const studentEmail = rowData[COLUMNS.STUDENT_EMAIL - 1];
  const studentClass = rowData[COLUMNS.CLASS - 1];
  const studentId = rowData[COLUMNS.STUDENT_ID - 1];
  const studentName = rowData[COLUMNS.STUDENT_NAME - 1];
  const teacherCheckValue = rowData[COLUMNS.TEACHER_CHECK - 1];
  const careerCheckValue = rowData[COLUMNS.CAREER_CHECK - 1];
  const officeReceiptValue = rowData[COLUMNS.OFFICE_RECEIPT - 1];
  const docCreationValue = rowData[COLUMNS.DOC_CREATION - 1];
  let mailFlags = rowData[mailFlagColIndex - 1] ? rowData[mailFlagColIndex - 1].toString().split(',') : [];

  // メール本文などで使うプレースホルダー
  const placeholders = {
    '{クラス}': studentClass,
    '{番号}': studentId,
    '{名前}': studentName,
    '{担任確認コメント}': teacherCheckValue,
    '{進路部確認コメント}': careerCheckValue,
    '{事務室受領コメント}': officeReceiptValue
  };

  // --- メール送信処理 ---
  try {
    // 条件1: 担任確認(N列)に入力されたら進路部へメール
    if (editedCol === COLUMNS.TEACHER_CHECK && teacherCheckValue && !mailFlags.includes('N_SENT')) {
      const to = settings.shinrobuEmail;
      const subject = replacePlaceholders(settings.teacherCheckSubject, placeholders);
      const body = replacePlaceholders(settings.teacherCheckBody, placeholders);
      if (validateEmailParams(to, subject, body, "進路部")) {
        Logger.log(`条件1合致: 進路部へメール送信 (行: ${editedRow})`);
        GmailApp.sendEmail(to, subject, body);
        updateMailFlag(sheet, editedRow, mailFlagColIndex, mailFlags, 'N_SENT');
        mailFlags.push('N_SENT'); // 後続処理のために内部フラグも更新
      }
    }

    // 条件2: 進路部確認(O列) と 事務室受領(P列) の両方に入力されたら担任へメール
    // (O列またはP列が編集された時にチェック)
    if ((editedCol === COLUMNS.CAREER_CHECK || editedCol === COLUMNS.OFFICE_RECEIPT) &&
        careerCheckValue && officeReceiptValue && !mailFlags.includes('OP_SENT')) {
      const teacherEmail = settings.teacherEmails[studentClass];
      if (teacherEmail) {
        const to = teacherEmail;
        const subject = replacePlaceholders(settings.progressReportSubject, placeholders);
        const body = replacePlaceholders(settings.progressReportBody, placeholders);
         if (validateEmailParams(to, subject, body, `担任(${studentClass})`)) {
           Logger.log(`条件2合致: 担任へメール送信 (行: ${editedRow})`);
           GmailApp.sendEmail(to, subject, body);
           updateMailFlag(sheet, editedRow, mailFlagColIndex, mailFlags, 'OP_SENT');
           mailFlags.push('OP_SENT'); // 後続処理のために内部フラグも更新
         }
      } else {
        Logger.log(`警告: クラス ${studentClass} の担任メールアドレスが設定シートに見つかりません (行: ${editedRow})`);
      }
    }

    // 条件3: 担任(N), 進路部(O), 事務室(P), 作成(Q) の全てに入力されたら生徒へメール
    // (N, O, P, Q いずれかが編集された時にチェック)
     if (targetColumns.includes(editedCol) && // N,O,P,Qいずれかの編集でトリガー
        teacherCheckValue && careerCheckValue && officeReceiptValue && docCreationValue &&
        !mailFlags.includes('ALL_SENT')) {
       const to = studentEmail;
       const subject = replacePlaceholders(settings.completionSubject, placeholders);
       const body = replacePlaceholders(settings.completionBody, placeholders);
       if (validateEmailParams(to, subject, body, `生徒(${studentName})`)) {
         Logger.log(`条件3合致: 生徒へメール送信 (行: ${editedRow})`);
         GmailApp.sendEmail(to, subject, body);
         updateMailFlag(sheet, editedRow, mailFlagColIndex, mailFlags, 'ALL_SENT');
         // mailFlags.push('ALL_SENT'); // この関数内での後続処理はないので不要
       }
    }

  } catch (error) {
    Logger.log(`！！！ メール送信処理中にエラーが発生しました (行: ${editedRow}) ！！！`);
    Logger.log(`エラーメッセージ: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    // 必要に応じてUIアラートや管理者へのエラー通知メールを追加
    // SpreadsheetApp.getUi().alert(`メール送信エラーが発生しました。\n行: ${editedRow}\nエラー: ${error.message}`);
  }
}

// --- ヘルパー関数 ---

/**
 * 設定シートから設定値を読み込み、オブジェクトとして返す
 * @return {object|null} 設定値オブジェクト、またはエラー時null
 */
function getSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      Logger.log(`エラー: 設定シート "${SETTINGS_SHEET_NAME}" が見つかりません。`);
      return null;
    }
    // 設定項目を読み込む範囲をデータ量に合わせて調整 (例: A1:B20)
    const data = sheet.getRange("A1:B20").getValues();
    const settings = {
      teacherEmails: {}
    };
    let isTeacherSection = false;

    for (let i = 0; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];

      if (!key || !value) continue; // キーまたは値が空ならスキップ

      switch (key) {
        case '進路部メールアドレス':
          settings.shinrobuEmail = value;
          isTeacherSection = false;
          break;
        case 'クラス担任メールアドレス':
          isTeacherSection = true; // ここから担任アドレスセクション
          break;
        case 'メール送信済みフラグ列':
          settings.mailFlagColumnLetter = value.toUpperCase(); // 大文字に統一
           isTeacherSection = false;
          break;
        case '担任確認メール件名':
            settings.teacherCheckSubject = value;
            isTeacherSection = false;
            break;
        case '担任確認メール本文':
            settings.teacherCheckBody = value;
            isTeacherSection = false;
            break;
        case '進捗連絡メール件名':
            settings.progressReportSubject = value;
            isTeacherSection = false;
            break;
        case '進捗連絡メール本文':
            settings.progressReportBody = value;
            isTeacherSection = false;
            break;
        case '完了連絡メール件名':
            settings.completionSubject = value;
            isTeacherSection = false;
            break;
        case '完了連絡メール本文':
            settings.completionBody = value;
            isTeacherSection = false;
            break;
        default:
          // クラス担任メールアドレスセクションの場合
          if (isTeacherSection && key.endsWith('組')) {
            settings.teacherEmails[key] = value;
          }
          break;
      }
    }

     // 必須設定の簡易チェック
     if (!settings.shinrobuEmail || Object.keys(settings.teacherEmails).length === 0 || !settings.mailFlagColumnLetter ||
         !settings.teacherCheckSubject || !settings.teacherCheckBody || !settings.progressReportSubject || !settings.progressReportBody ||
         !settings.completionSubject || !settings.completionBody) {
         Logger.log('警告: 設定シートに必要な項目が不足している可能性があります。');
         // return null; // 処理を継続させる場合はコメントアウト解除しない
     }

    return settings;
  } catch (error) {
    Logger.log(`設定シートの読み込み中にエラー: ${error.message}\n${error.stack}`);
    return null;
  }
}

/**
 * メール送信フラグ列に指定のフラグを追加する
 * @param {Sheet} sheet 対象シート
 * @param {number} row 対象行番号
 * @param {number} colIndex フラグ列のインデックス
 * @param {string[]} currentFlags 現在のフラグ配列
 * @param {string} newFlag 追加するフラグ文字列
 */
function updateMailFlag(sheet, row, colIndex, currentFlags, newFlag) {
  try {
      if (!currentFlags.includes(newFlag)) {
        const updatedFlags = [...currentFlags, newFlag]; // 新しい配列を作成
        sheet.getRange(row, colIndex).setValue(updatedFlags.join(','));
        Logger.log(`フラグ更新: 行 ${row}, 列 ${colIndex}, 値 ${updatedFlags.join(',')}`);
      }
  } catch (error) {
      Logger.log(`フラグ更新エラー: 行 ${row}, 列 ${colIndex}, フラグ ${newFlag}`);
      Logger.log(`エラー: ${error.message}`);
  }
}

/**
 * 列文字(A, B, ...)を列インデックス(1, 2, ...)に変換する
 * @param {string} columnLetter 列文字 (例: "R")
 * @return {number|null} 列インデックス、または無効な場合はnull
 */
function getColumnIndex(columnLetter) {
    if (!columnLetter || typeof columnLetter !== 'string' || columnLetter.length !== 1 || !/^[A-Z]$/i.test(columnLetter)) {
        Logger.log(`無効な列文字です: ${columnLetter}`);
        return null;
    }
    return columnLetter.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

/**
 * テンプレート文字列内のプレースホルダーを実際の値に置換する
 * @param {string} template テンプレート文字列 (例: "{クラス} {番号} 様")
 * @param {object} placeholders 置換する値のオブジェクト (例: {'{クラス}': 'A組', ...})
 * @return {string} 置換後の文字列
 */
function replacePlaceholders(template, placeholders) {
    if (!template) return "";
    let result = template;
    for (const placeholder in placeholders) {
        // 正規表現でプレースホルダーを安全に置換 (gフラグで複数置換)
        const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        result = result.replace(regex, placeholders[placeholder] || ''); // 値がなければ空文字に
    }
    return result;
}

/**
 * メール送信パラメータ（宛先、件名、本文）が有効かチェックする
 * @param {string} to 宛先メールアドレス
 * @param {string} subject 件名
 * @param {string} body 本文
 * @param {string} recipientType 受信者タイプ（ログ用）
 * @return {boolean} 有効な場合はtrue、無効な場合はfalse
 */
function validateEmailParams(to, subject, body, recipientType) {
  if (!to) {
    Logger.log(`警告: ${recipientType} のメールアドレスが指定されていません。`);
    return false;
  }
  if (!subject) {
     Logger.log(`警告: ${recipientType} へのメール件名がありません。`);
    return false;
  }
   if (!body) {
     Logger.log(`警告: ${recipientType} へのメール本文がありません。`);
    return false;
  }
  // 簡単なメールアドレス形式チェック (より厳密なチェックも可能)
  if (!/.+@.+\..+/.test(to)) {
      Logger.log(`警告: ${recipientType} のメールアドレス(${to})の形式が無効です。`);
      return false;
  }
  return true;
}


// --- セットアップ用関数（初回実行または設定変更時） ---
// スクリプトエディタから手動で実行してください。
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 編集トリガーの設定
  // 既存のトリガーを削除してから再設定する（重複防止）
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('既存のonEditトリガーを削除しました。');
    }
  });
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log('新しいonEditトリガーを設定しました。');

  // 2. 設定シートの確認・作成
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    Logger.log(`シート "${SETTINGS_SHEET_NAME}" を作成しました。`);
    // 設定シートのテンプレートを書き込む (★要確認・修正)
    settingsSheet.getRange("A1:B1").setValues([["設定項目", "値"]]).setFontWeight("bold");
    settingsSheet.getRange("A2:B18").setValues([
      ["進路部メールアドレス", "shinrobu@example.com"], // ★要変更
      ["クラス担任メールアドレス", ""],
      ["A組", "tanin_a@example.com"], // ★要変更
      ["B組", "tanin_b@example.com"], // ★要変更
      ["C組", "tanin_c@example.com"], // ★要変更
      ["D組", "tanin_d@example.com"], // ★要変更
      ["E組", "tanin_e@example.com"], // ★要変更
      ["F組", "tanin_f@example.com"], // ★要変更
      ["G組", "tanin_g@example.com"], // ★要変更
      ["H組", "tanin_h@example.com"], // ★要変更
      ["メール送信済みフラグ列", "R"], // ★必要なら変更
      ["担任確認メール件名", "【要確認】調査書作成願 ({クラス} {番号} {名前})"],
      ["担任確認メール本文", "{クラス} {番号} {名前} さん\n\n担任確認欄に以下のコメントが入力されました。\n\n{担任確認コメント}"],
      ["進捗連絡メール件名", "【進捗連絡】調査書作成願 ({クラス} {番号} {名前})"],
      ["進捗連絡メール本文", "{クラス} {番号} {名前} さん\n\n調査書作成願について、以下の確認が完了しました。\n\n進路部確認: {進路部確認コメント}\n事務室受領: {事務室受領コメント}"],
      ["完了連絡メール件名", "【完了連絡】調査書作成願"],
      ["完了連絡メール本文", "{クラス} {番号} {名前} さん\n\n調査票ができましたので、担任から受け取ってください。"]
    ]).setVerticalAlignment("top"); // セルの内容が複数行になる可能性があるので上寄せ
    settingsSheet.autoResizeColumn(1);
    settingsSheet.autoResizeColumn(2);
    SpreadsheetApp.getUi().alert(`"${SETTINGS_SHEET_NAME}" シートを作成し、テンプレートを挿入しました。\nシートの内容を確認し、★印の箇所を実際の情報に修正してください。`);
  } else {
      Logger.log(`シート "${SETTINGS_SHEET_NAME}" は既に存在します。`);
  }

   // 3. フォーム回答シートに列ヘッダーを追加（存在しない場合）
   const formSheet = ss.getSheetByName(FORM_SHEET_NAME);
   if (formSheet) {
       const headersToAdd = [
           { col: COLUMNS.QUALIFICATION, name: "実際に受験資格を満たしているか" },
           { col: COLUMNS.TEACHER_CHECK, name: "担任の確認" },
           { col: COLUMNS.CAREER_CHECK, name: "進路部の確認" },
           { col: COLUMNS.OFFICE_RECEIPT, name: "事務室での受領" },
           { col: COLUMNS.DOC_CREATION, name: "調査書作成" },
       ];
       headersToAdd.forEach(headerInfo => {
           const cell = formSheet.getRange(1, headerInfo.col);
           if (!cell.getValue()) { // ヘッダーが空なら追加
                cell.setValue(headerInfo.name);
                Logger.log(`シート "${FORM_SHEET_NAME}" の ${headerInfo.col}列目にヘッダー "${headerInfo.name}" を追加しました。`);
           }
       });
       // フラグ列ヘッダーも確認・追加
       const settings = getSettings(); // 設定を再読み込み
       const flagColIndex = settings ? getColumnIndex(settings.mailFlagColumnLetter) : null;
       if (flagColIndex) {
           const flagHeaderCell = formSheet.getRange(1, flagColIndex);
           if (!flagHeaderCell.getValue()) {
                flagHeaderCell.setValue("メール送信フラグ");
                Logger.log(`シート "${FORM_SHEET_NAME}" の ${flagColIndex}列目にヘッダー "メール送信フラグ" を追加しました。`);
           }
       } else {
            Logger.log(`警告: フラグ列の列文字(${settings?.mailFlagColumnLetter})が無効なため、ヘッダーを追加できませんでした。設定シートを確認してください。`);
       }
       // ヘッダー行を固定
       if (formSheet.getFrozenRows() < 1) {
            formSheet.setFrozenRows(1);
            Logger.log(`シート "${FORM_SHEET_NAME}" のヘッダー行を固定しました。`);
       }

       SpreadsheetApp.getUi().alert('セットアップが完了しました。\n"設定"シートの内容を確認・修正し、フォーム回答シートで運用を開始してください。');

   } else {
       SpreadsheetApp.getUi().alert(`エラー: シート "${FORM_SHEET_NAME}" が見つかりません。スクリプト内の FORM_SHEET_NAME 定数を確認してください。`);
   }
}