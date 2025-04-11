/**
 * @OnlyCurrentDoc
 */

// --- 設定 ---
const FORM_RESPONSE_SHEET_NAME = "フォームの回答"; // あなたのシート名に合わせて変更してください
const SETTINGS_SHEET_NAME = "設定";
const REQUIRED_HEADERS = ["担任の確認", "進路部の確認", "事務室での受領", "調査書作成", "受付番号"];
const STUDENT_EMAIL_COL = 2; // B列
const CLASS_COL = 3; // C列
const STUDENT_NUMBER_COL = 4; // D列
const STUDENT_NAME_COL = 5; // E列
const UNIV_CODE_1_COL = 6; // G列
const UNIV_NAME_1_COL = 7; // H列
const FACULTY_1_COL = 8; // I列
const DEPT_1_COL = 9; // J列
// 必要であれば、初期メール用に大学の列を追加

// --- ヘルパー関数: 設定シートからメールアドレスを取得 ---
function getEmailAddress_(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    Logger.log(`エラー: 設定シート "${SETTINGS_SHEET_NAME}" が見つかりません。`);
    notifyAdminOnError_(`設定シート "${SETTINGS_SHEET_NAME}" が見つかりません。`, 'getEmailAddress_');
    return null;
  }
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === role) {
      return data[i][1];
    }
  }
  Logger.log(`警告: 役割 "${role}" のメールアドレスが設定に見つかりません。`);
  // 必要であれば、役割が見つからない場合に管理者に通知
  // notifyAdminOnError_(`役割 "${role}" のメールアドレスが設定に見つかりません。`, 'getEmailAddress_');
  return null; // 役割が見つからなかった
}

// --- ヘルパー関数: エラー時に管理者に通知 ---
function notifyAdminOnError_(error, functionName) {
  try {
    const adminEmail = getEmailAddress_('管理者');
    if (adminEmail) {
      const subject = `GAS スクリプトエラー: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`;
      let body = `スクリプトでエラーが発生しました。\n\n`;
      body += `スプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getName()}\n`;
      body += `シート: ${SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()}\n`; // 可能な場合はアクティブシートを取得
      body += `関数: ${functionName}\n`;
      body += `エラー: ${error.message || error}\n`;
      if (error.stack) {
        body += `スタックトレース:\n${error.stack}\n`;
      }
      // 可能であればコンテキストを追加 (例: アクティブな行/列)
      try {
        const activeRange = SpreadsheetApp.getActiveRange();
        if (activeRange) {
           body += `アクティブセル: ${activeRange.getA1Notation()}\n`;
           body += `編集された値: ${activeRange.getValue()}\n`;
        }
      } catch(e) {
         body += `アクティブセル情報の取得に失敗しました。\n`
      }

      Logger.log(`エラー通知を ${adminEmail} に送信中。 エラー: ${error}`);
      GmailApp.sendEmail(adminEmail, subject, body);
    } else {
      Logger.log(`エラー: 管理者のメールアドレスが設定に見つかりません。エラー通知を送信できません。 エラー内容: ${error}`);
    }
  } catch (e) {
    Logger.log(`重大なエラー: エラー通知の送信に失敗しました。 元のエラー: ${error}。 通知エラー: ${e}`);
  }
}


// --- ヘルパー関数: ヘッダーを確認し、なければ追加 ---
function checkAndAddHeaders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
  if (!sheet) {
    notifyAdminOnError_(`フォーム回答シート "${FORM_RESPONSE_SHEET_NAME}" が見つかりません。`, 'checkAndAddHeaders_');
    throw new Error(`シート "${FORM_RESPONSE_SHEET_NAME}" が見つかりません。`);
  }
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const headersToAdd = [];
  const headerPositions = {}; // {ヘッダー名: 列番号(1-based)}

  // 現在のヘッダー位置を取得
   headers.forEach((header, index) => {
    if (header) { // ヘッダーが空でないことを確認
        headerPositions[header.trim()] = index + 1; // 1-based index
    }
   });

  // どのヘッダーが不足しているか確認
  REQUIRED_HEADERS.forEach(requiredHeader => {
    if (!headerPositions[requiredHeader]) {
      headersToAdd.push(requiredHeader);
    }
  });

  // 不足しているヘッダーを追加
  if (headersToAdd.length > 0) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol, 1, headersToAdd.length).setValues([headersToAdd]);
    Logger.log(`不足しているヘッダーを追加しました: ${headersToAdd.join(', ')}`);
    // 追加後にヘッダー位置を更新
    headersToAdd.forEach((header, index) => {
        headerPositions[header.trim()] = nextCol + index;
    });
  }

   // 必要な全てのヘッダーの位置を再取得（追加されたものも含めて）
   REQUIRED_HEADERS.forEach(requiredHeader => {
      let found = false;
      for (const header in headerPositions) {
          if (header === requiredHeader) {
              // headerPositionsのキーが REQUIRED_HEADERS と完全に一致するように保証
              headerPositions[requiredHeader] = headerPositions[header];
              found = true;
              break;
          }
      }
       if (!found) {
           // ヘッダーが見つからない、または追加できなかった場合はエラー
           notifyAdminOnError_(`必須ヘッダー "${requiredHeader}" が見つからないか、追加できませんでした。`, 'checkAndAddHeaders_');
           throw new Error(`必須ヘッダー "${requiredHeader}" が見つからないか、追加できませんでした。`);
       }
   });


  return headerPositions;
}

// --- 新規フォーム送信を処理する関数 ---
function processNewSubmission(e) {
  const functionName = 'processNewSubmission';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);

    // フォーム送信トリガーの場合、e.rangeが存在する。それ以外の場合は最終行とみなす。
    const rowIdx = e && e.range ? e.range.getRowIndex() : sheet.getLastRow();
    if (rowIdx <= 1) return; // ヘッダー行はスキップ

    const headerPositions = checkAndAddHeaders_(); // まずヘッダーを確認
    const receptionCol = headerPositions['受付番号'];

    // 新しい行のデータを取得
    const rowData = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];

    // --- 受付番号の割り当て ---
    const currentReceptionNumber = rowData[receptionCol - 1];
    let newReceptionNumber = currentReceptionNumber;

    if (!currentReceptionNumber) { // 空の場合のみ割り当て
      let lastReceptionNumber = 0;
      if (rowIdx > 2) { // 最初のデータ行でない場合、前の行を確認
        const prevReceptionVal = sheet.getRange(rowIdx - 1, receptionCol).getValue();
        if (prevReceptionVal && !isNaN(parseInt(prevReceptionVal))) {
          lastReceptionNumber = parseInt(prevReceptionVal);
        } else { // フォールバック: 前の行が空または無効な場合、上方向に検索
             const receptionData = sheet.getRange(2, receptionCol, rowIdx - 2, 1).getValues();
             for (let i = receptionData.length - 1; i >= 0; i--) {
                 if (receptionData[i][0] && !isNaN(parseInt(receptionData[i][0]))) {
                     lastReceptionNumber = parseInt(receptionData[i][0]);
                     break;
                 }
             }
        }
      }
      newReceptionNumber = lastReceptionNumber + 1;
      sheet.getRange(rowIdx, receptionCol).setValue(newReceptionNumber);
      Logger.log(`行 ${rowIdx} に受付番号 ${newReceptionNumber} を割り当てました。`);

      // --- 生徒へ確認メールを送信 ---
      const studentEmail = rowData[STUDENT_EMAIL_COL - 1];
      const studentName = rowData[STUDENT_NAME_COL - 1];
      const studentClass = rowData[CLASS_COL - 1];
      const studentNumber = rowData[STUDENT_NUMBER_COL - 1];
      const univCode1 = rowData[UNIV_CODE_1_COL -1]; // G列
      const univName1 = rowData[UNIV_NAME_1_COL - 1]; // H列
      const faculty1 = rowData[FACULTY_1_COL - 1]; // I列
      const dept1 = rowData[DEPT_1_COL - 1]; // J列 - リクエストの文脈に基づき含める

      if (studentEmail) {
        const subject = "調査書作成願 受付完了のお知らせ";
        let body = `${studentName} さん (${studentClass} ${studentNumber}番)\n\n`;
        body += "フォームへのご提出ありがとうございました。\n";
        body += "以下の内容で調査書作成願の受付を完了しました。\n\n";
        body += `受付番号: ${newReceptionNumber}\n\n`;
        body += "--- 入力内容の控え (一部抜粋) ---\n";
        body += `大学コード1: ${univCode1 || '未入力'}\n`;
        body += `大学名1: ${univName1 || '未入力'}\n`;
        body += `学部1: ${faculty1 || '未入力'}\n`;
        body += `学科1: ${dept1 || '未入力'}\n`;
        // 必要であれば他のフィールドもここに追加し、後続の列を確認
        body += "\n------------------------------\n";
        body += "今後の手続きについては、別途連絡をお待ちください。";

        GmailApp.sendEmail(studentEmail, subject, body);
        Logger.log(`行 ${rowIdx} の ${studentEmail} に確認メールを送信しました。`);
      } else {
        Logger.log(`警告: 行 ${rowIdx} に生徒のメールアドレスが見つかりません。確認メールを送信できません。`);
      }
    } else {
        Logger.log(`行 ${rowIdx} には既に受付番号 (${currentReceptionNumber}) があります。割り当てと確認メールをスキップします。`);
    }

  } catch (error) {
    Logger.log(`${functionName} でエラーが発生しました: ${error}`);
    notifyAdminOnError_(error, functionName);
  }
}


// --- シート編集を処理する関数 ---
function onSheetEdit(e) {
  const functionName = 'onSheetEdit';
  try {
    const range = e.range; // 編集されたセル範囲
    const sheet = range.getSheet();

    // 正しいシートでの編集か、ヘッダー行でないかを確認
    if (sheet.getName() !== FORM_RESPONSE_SHEET_NAME || range.getRow() <= 1) {
      return;
    }

    const editedRow = range.getRow(); // 編集された行番号
    const editedCol = range.getColumn(); // 編集された列番号

    const headerPositions = checkAndAddHeaders_(); // ヘッダー位置を取得
    const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0]; // 編集された行の全データ

    // 共通の生徒情報を取得
    const studentClass = rowData[CLASS_COL - 1];
    const studentNumber = rowData[STUDENT_NUMBER_COL - 1];
    const studentName = rowData[STUDENT_NAME_COL - 1];
    const studentEmail = rowData[STUDENT_EMAIL_COL - 1];
    const receptionNumber = rowData[headerPositions['受付番号'] - 1];
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + sheet.getSheetId() + "&range=A" + editedRow; // 編集行へのリンク

    // --- 特定の列の編集を確認 ---
    const teacherConfirmCol = headerPositions['担任の確認'];
    const guidanceConfirmCol = headerPositions['進路部の確認'];
    const officeReceiveCol = headerPositions['事務室での受領'];
    const transcriptCreateCol = headerPositions['調査書作成'];

    // セルが実際に値で埋められたか（クリアされたのではないか）を確認
    const editedValue = range.getValue();
    if (editedValue === "" || editedValue === null || editedValue === undefined) {
        // Logger.log(`セル ${range.getA1Notation()} がクリアされたため、通知をスキップします。`);
        return; // セルがクリアされた場合は通知をトリガーしない
    }


    // 1. 担任確認 -> 進路部へ通知
    if (editedCol === teacherConfirmCol) {
      const guidanceEmail = getEmailAddress_('進路部');
      if (guidanceEmail) {
        const teacherConfirmationText = editedValue; // 入力された実際の値を使用
        const subject = `【要確認】担任確認完了: ${studentClass} ${studentNumber}番 ${studentName}`;
        let body = `担任が調査書作成願を確認しました。\n\n`;
        body += `担任入力内容: ${teacherConfirmationText}\n`;
        body += `クラス: ${studentClass}\n`;
        body += `出席番号: ${studentNumber}\n`;
        body += `氏名: ${studentName}\n`;
        body += `受付番号: ${receptionNumber}\n\n`;
        body += `詳細確認・進路部確認入力はこちら:\n${spreadsheetUrl}`;

        GmailApp.sendEmail(guidanceEmail, subject, body);
        Logger.log(`行 ${editedRow} について進路部へ通知を送信しました。`);
      } else {
        Logger.log(`警告: 進路部のメールアドレスが見つかりません。行 ${editedRow} の通知を送信できません。`);
         notifyAdminOnError_(`担任確認通知のための進路部メールアドレスが見つかりません。`, functionName);
      }
    }

    // 2. 進路部 かつ 事務室 確認済み -> 担任へ通知
    // 編集された列がこれら2つのどちらかであり、かつ両方が入力済みかを確認
    if (editedCol === guidanceConfirmCol || editedCol === officeReceiveCol) {
      const guidanceConfirmationText = rowData[guidanceConfirmCol - 1];
      const officeReceptionText = rowData[officeReceiveCol - 1];

      if (guidanceConfirmationText && officeReceptionText) { // 両方にテキストが必要
        const teacherRole = `${studentClass}担任`; // 役割名を動的に構成
        const teacherEmail = getEmailAddress_(teacherRole);
        if (teacherEmail) {
          const subject = `【進捗】進路部・事務室確認完了: ${studentClass} ${studentNumber}番 ${studentName}`;
          let body = `進路部および事務室での確認・受領が完了しました。\n\n`;
          body += `受付番号: ${receptionNumber}\n`;
          body += `進路部確認内容: ${guidanceConfirmationText}\n`;
          body += `事務室受領内容: ${officeReceptionText}\n\n`;
          body += `スプレッドシートで確認:\n${spreadsheetUrl}`;

          GmailApp.sendEmail(teacherEmail, subject, body);
          Logger.log(`行 ${editedRow} について担任 (${teacherEmail}) へ通知を送信しました。`);
        } else {
          Logger.log(`警告: 役割 "${teacherRole}" の担任メールアドレスが見つかりません。行 ${editedRow} の通知を送信できません。`);
          notifyAdminOnError_(`役割 "${teacherRole}" の担任メールアドレスが見つかりません（進路/事務確認通知）。`, functionName);
        }
      }
    }

    // 3. 全て確認済み (担任、進路部、事務室、作成) -> 生徒へ通知
    // 編集された列がこれら4つのどれかであり、かつ全てが入力済みかを確認
     if ([teacherConfirmCol, guidanceConfirmCol, officeReceiveCol, transcriptCreateCol].includes(editedCol)) {
        const teacherConfirmationText = rowData[teacherConfirmCol - 1];
        const guidanceConfirmationText = rowData[guidanceConfirmCol - 1];
        const officeReceptionText = rowData[officeReceiveCol - 1];
        const transcriptCreationText = rowData[transcriptCreateCol - 1];

        if (teacherConfirmationText && guidanceConfirmationText && officeReceptionText && transcriptCreationText) {
             if (studentEmail) {
                const subject = "調査書作成完了のお知らせ";
                let body = `${studentName} さん (${studentClass} ${studentNumber}番)\n\n`;
                body += `受付番号 ${receptionNumber} の調査書が作成できました。\n`;
                body += "担任の先生から受け取ってください。";

                GmailApp.sendEmail(studentEmail, subject, body);
                Logger.log(`行 ${editedRow} の生徒 ${studentEmail} に完了通知を送信しました。`);
            } else {
                Logger.log(`警告: 行 ${editedRow} の生徒メールアドレスが見つかりません。完了通知を送信できません。`);
                 notifyAdminOnError_(`最終完了通知送信時に、行 ${editedRow} の生徒メールアドレスが見つかりません。`, functionName);
            }
        }
     }

  } catch (error) {
    Logger.log(`${functionName} でエラーが発生しました: ${error}`);
    notifyAdminOnError_(error, functionName);
  }
}

// --- トリガー設定関数 ---
// これらのトリガーは Apps Script エディタで手動で設定する必要があります。
// 1. processNewSubmission のトリガー:
//    - イベントのソース: スプレッドシートから
//    - イベントの種類: フォーム送信時
// 2. onSheetEdit のトリガー:
//    - イベントのソース: スプレッドシートから
//    - イベントの種類: 編集時

// --- オプション: 受付番号がない既存行を手動で処理する関数 ---
function processExistingRows() {
  const functionName = 'processExistingRows';
   try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
    if (!sheet) {
        throw new Error(`シート ${FORM_RESPONSE_SHEET_NAME} が見つかりません。`);
    }

    const headerPositions = checkAndAddHeaders_();
    const receptionCol = headerPositions['受付番号'];
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()); // 2行目から開始
    const allData = dataRange.getValues();
    const receptionNumbers = sheet.getRange(2, receptionCol, lastRow -1, 1).getValues(); // 受付番号列のみ取得

    let lastAssignedNumber = 0;
     // 複数回実行した場合の重複を避けるため、最初に既存の最大番号を見つける
     for(let i = receptionNumbers.length - 1; i >= 0; i--) {
         if (receptionNumbers[i][0] && !isNaN(parseInt(receptionNumbers[i][0]))) {
            lastAssignedNumber = Math.max(lastAssignedNumber, parseInt(receptionNumbers[i][0]));
         }
     }
     Logger.log(`既存の最大の受付番号: ${lastAssignedNumber}`);


    for (let i = 0; i < allData.length; i++) {
      const currentRow = i + 2; // シートの行番号 (1-based)
      const currentReceptionNumber = receptionNumbers[i][0]; // 専用の配列を使用

      if (!currentReceptionNumber) {
        // processNewSubmission のために最小限のイベントオブジェクトをシミュレート
        const fakeEvent = {
          range: sheet.getRange(currentRow, 1) // 行の開始を示すRangeオブジェクトを提供
          // values: allData[i] // processNewSubmissionロジックでは直接使用されなくなった
        };
        Logger.log(`受付番号がない既存行 ${currentRow} を処理中。`);
        processNewSubmission(fakeEvent); // 番号割り当てとメール送信のために関数を呼び出す

        // オプション: 大量行を処理する場合、メールクォータ超過を避けるために短い遅延を追加
        // Utilities.sleep(500); // 500ミリ秒待機
      }
    }
     Logger.log(`既存行の処理が完了しました。`);
   } catch (error) {
      Logger.log(`${functionName} でエラーが発生しました: ${error}`);
      notifyAdminOnError_(error, functionName);
   }
}