// --- 設定 (支払い確認処理用) ---
const PAYMENT_CONFIRM_SHEET_NAME = "支払い確認"; // 支払い確認フォームの回答シート名
const TRANSCRIPT_REQUEST_SHEET_NAME = "フォームの回答"; // ★元の調査書作成願シート名と同じか確認
const SETTINGS_SHEET_NAME_PAYMENT = "設定"; // ★元の設定シート名と同じか確認
const PAYMENT_REQUIRED_HEADERS = ["支払い確認ID"];
const PAYMENT_CLASS_COL = 2;          // B列: クラス
const PAYMENT_STUDENT_NUMBER_COL = 3; // C列: 出席番号
const PAYMENT_STUDENT_NAME_COL = 4;   // D列: 名前
const PAYMENT_RECEPTION_NUM_COL = 5;  // E列: 受付番号の入力
const PAYMENT_NUMBER_COL = 6;         // F列: 支払い番号の入力

// --- メイン処理関数 (支払い確認フォーム送信時に実行) ---
function processPaymentConfirmation(e) {
  const functionName = 'processPaymentConfirmation';
  let mainSheet = null; // スコープを広げる
  let paymentRowIdx = -1; // スコープを広げる

  try {
    // --- 0. イベントオブジェクトとシート取得 ---
    if (!e || !e.range) {
      Logger.log("イベントオブジェクトまたは範囲がありません。手動実行では動作しません。");
      // 手動実行の場合など、e.range がない場合は処理を中断
      // 必要であれば管理者に通知
      // notifyAdminOnError_("イベント情報が不足しています。", functionName);
      return;
    }

    const paymentSheet = e.range.getSheet();
    if (paymentSheet.getName() !== PAYMENT_CONFIRM_SHEET_NAME) {
      Logger.log(`支払い確認シート (${PAYMENT_CONFIRM_SHEET_NAME}) 以外でのイベントのため処理をスキップします。`);
      return; // 関係ないシートでのイベントは無視
    }
    paymentRowIdx = e.range.getRowIndex(); // 処理対象の行番号
    if (paymentRowIdx <= 1) return; // ヘッダー行は無視

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    mainSheet = ss.getSheetByName(TRANSCRIPT_REQUEST_SHEET_NAME); // ★変数名を変更
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME_PAYMENT); // ★変数名を変更

    if (!mainSheet) throw new Error(`調査書作成願シート "${TRANSCRIPT_REQUEST_SHEET_NAME}" が見つかりません。`);
    if (!settingsSheet) throw new Error(`設定シート "${SETTINGS_SHEET_NAME_PAYMENT}" が見つかりません。`);

    // --- 1. ヘッダー確認 (支払い確認シート) ---
    const paymentHeaderPositions = checkAndAddPaymentHeaders_(paymentSheet);
    const paymentConfirmIdCol = paymentHeaderPositions['支払い確認ID'];

    // --- 2. フォーム入力データ取得 (支払い確認シート) ---
    const paymentData = paymentSheet.getRange(paymentRowIdx, 1, 1, paymentSheet.getLastColumn()).getValues()[0];
    const submittedClass = paymentData[PAYMENT_CLASS_COL - 1];
    const submittedStudentNum = paymentData[PAYMENT_STUDENT_NUMBER_COL - 1];
    const submittedStudentName = paymentData[PAYMENT_STUDENT_NAME_COL - 1];
    const submittedReceptionNum = paymentData[PAYMENT_RECEPTION_NUM_COL - 1];
    const submittedPaymentNum = paymentData[PAYMENT_NUMBER_COL - 1];

    // --- 3. バリデーション ---
    const officeEmail = getEmailAddress_('事務室'); // getEmailAddress_ は元のスクリプトから流用
    if (!officeEmail) {
         Logger.log("警告: 事務室のメールアドレスが設定シートに見つかりません。");
         // 管理者に通知しても良い
         notifyAdminOnError_("事務室のメールアドレスが設定シートに見つかりません。", functionName);
         // 処理を続けるか、中断するかは要件次第だが、ここでは中断する
         throw new Error("事務室のメールアドレスが未設定です。");
    }
    const validationErrorLink = ss.getUrl() + "#gid=" + paymentSheet.getSheetId() + "&range=A" + paymentRowIdx;

    // 3a. 支払い番号の重複チェック (支払い確認シート内)
    const paymentNumColData = paymentSheet.getRange(2, PAYMENT_NUMBER_COL, paymentSheet.getLastRow() - 1, 1).getValues();
    let duplicateFound = false;
    for (let i = 0; i < paymentNumColData.length; i++) {
        // 自分自身の行は除外してチェック
        if ((i + 2) !== paymentRowIdx && paymentNumColData[i][0] && paymentNumColData[i][0] === submittedPaymentNum) {
            duplicateFound = true;
            break;
        }
    }
    if (duplicateFound) {
      const errorMsg = `支払い番号 '${submittedPaymentNum}' は既に登録されています。(クラス: ${submittedClass}, 出席番号: ${submittedStudentNum}, 名前: ${submittedStudentName})`;
      sendValidationErrorEmail_(officeEmail, "支払い番号重複エラー", errorMsg, validationErrorLink);
      return; // 処理中断
    }

    // 3b. 受付番号の存在確認と生徒情報の一致確認 (調査書作成願シート)
    // 調査書作成願シートのデータを効率的に取得
    const mainHeaderPositions = checkAndAddHeaders_(); // 元のスクリプトから流用/確認
    const mainReceptionCol = mainHeaderPositions['受付番号'];
    const mainClassCol = CLASS_COL; // 元のスクリプトの定数を流用
    const mainStudentNumCol = STUDENT_NUMBER_COL;
    const mainStudentNameCol = STUDENT_NAME_COL;
    const mainStudentEmailCol = STUDENT_EMAIL_COL;
    const mainOfficeReceiveCol = mainHeaderPositions['事務室での受領'];
    const mainGuidanceConfirmCol = mainHeaderPositions['進路部の確認']; // 担任通知用に取得

    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    let targetRowIndex = -1; // 調査書作成願シートでの一致した行番号 (0-based index in mainData)
    let targetStudentEmail = null;
    let targetTeacherRole = null;

    for (let i = 0; i < mainData.length; i++) {
      if (mainData[i][mainReceptionCol - 1] == submittedReceptionNum) { // 型が混在する可能性があるので == で比較
        targetRowIndex = i; // 0-based index
        // 受付番号が見つかったら、クラス・出席番号・名前も一致するか確認
        if (mainData[i][mainClassCol - 1] !== submittedClass ||
            mainData[i][mainStudentNumCol - 1] != submittedStudentNum || // 番号は数値/文字列混在考慮
            mainData[i][mainStudentNameCol - 1] !== submittedStudentName)
        {
            const errorMsg = `受付番号 '${submittedReceptionNum}' は存在しますが、生徒情報（クラス, 出席番号, 名前）が一致しません。\n` +
                             `フォーム入力: ${submittedClass} ${submittedStudentNum} ${submittedStudentName}\n` +
                             `調査書シート: ${mainData[i][mainClassCol - 1]} ${mainData[i][mainStudentNumCol - 1]} ${mainData[i][mainStudentNameCol - 1]}`;
            sendValidationErrorEmail_(officeEmail, "生徒情報不一致エラー", errorMsg, validationErrorLink);
            return; // 処理中断
        }
        // 一致した場合、メールアドレスと担任ロールを取得
        targetStudentEmail = mainData[i][mainStudentEmailCol - 1];
        targetTeacherRole = `${mainData[i][mainClassCol - 1]}担任`; // 担任ロールを特定
        break; // 一致する行が見つかったのでループ終了
      }
    }

    if (targetRowIndex === -1) {
      const errorMsg = `受付番号 '${submittedReceptionNum}' が調査書作成願シートに見つかりません。(クラス: ${submittedClass}, 出席番号: ${submittedStudentNum}, 名前: ${submittedStudentName})`;
      sendValidationErrorEmail_(officeEmail, "受付番号不一致エラー", errorMsg, validationErrorLink);
      return; // 処理中断
    }

    // --- 4. 支払い確認IDの割り当て (支払い確認シート) ---
    const currentPaymentId = paymentData[paymentConfirmIdCol - 1];
    let newPaymentId = currentPaymentId;
    if (!currentPaymentId) {
        let lastPaymentIdNum = 0;
        if (paymentRowIdx > 2) {
            const prevPaymentIdVal = paymentSheet.getRange(paymentRowIdx - 1, paymentConfirmIdCol).getValue();
             if (prevPaymentIdVal && !isNaN(parseInt(prevPaymentIdVal))) {
                lastPaymentIdNum = parseInt(prevPaymentIdVal);
            } else { // Fallback
                const idData = paymentSheet.getRange(2, paymentConfirmIdCol, paymentRowIdx - 2, 1).getValues();
                for (let i = idData.length - 1; i >= 0; i--) {
                    if (idData[i][0] && !isNaN(parseInt(idData[i][0]))) {
                        lastPaymentIdNum = parseInt(idData[i][0]);
                        break;
                    }
                }
            }
        }
         newPaymentId = lastPaymentIdNum + 1;
        paymentSheet.getRange(paymentRowIdx, paymentConfirmIdCol).setValue(newPaymentId);
        Logger.log(`行 ${paymentRowIdx} に支払い確認ID ${newPaymentId} を割り当てました。`);
    } else {
         Logger.log(`行 ${paymentRowIdx} には既に支払い確認ID (${currentPaymentId}) があります。ID割り当てをスキップします。`);
         newPaymentId = currentPaymentId; // 既存のIDを使う
    }


    // --- 5. 生徒へ支払い確認メール送信 ---
    if (!targetStudentEmail) {
         Logger.log(`警告: 受付番号 ${submittedReceptionNum} に対応する生徒のメールアドレスが見つかりません。メールは送信されません。`);
         // 事務室や管理者に通知する方が親切かもしれない
         notifyAdminOnError_(`受付番号 ${submittedReceptionNum} の生徒メールアドレスが見つからず、支払い確認メールを送信できません。`, functionName);
    } else {
        const subject = "【進路】支払い確認完了のお知らせ";
        let body = `${submittedStudentName} さん (${submittedClass} ${submittedStudentNum}番)\n\n`;
        body += "調査書発行手数料の支払い確認フォームのご提出ありがとうございます。\n";
        body += "以下の内容で確認いたしました。\n\n";
        body += `受付番号: ${submittedReceptionNum}\n`;
        body += `支払い番号: ${submittedPaymentNum}\n\n`;
        body += "調査書が作成できましたら、改めてメールにてお知らせいたします。\n";
        body += "しばらくお待ちください。";

        GmailApp.sendEmail(targetStudentEmail, subject, body);
        Logger.log(`生徒 ${targetStudentEmail} へ支払い確認メールを送信しました (受付番号: ${submittedReceptionNum})。`);
    }

    // --- 6. 調査書作成願シートの「事務室での受領」列を更新 ---
    // targetRowIndex は 0-based なので、実際の行番号に +2 する
    const mainTargetRow = targetRowIndex + 2;
    mainSheet.getRange(mainTargetRow, mainOfficeReceiveCol).setValue(submittedPaymentNum); // 支払い番号を記入
    Logger.log(`調査書作成願シートの行 ${mainTargetRow} の「事務室での受領」列に支払い番号 '${submittedPaymentNum}' を記入しました。`);

    // --- 7. 担任への通知条件を確認・実行 ---
    // 更新後の値を取得して確認
    const currentGuidanceVal = mainSheet.getRange(mainTargetRow, mainGuidanceConfirmCol).getValue();
    const currentOfficeVal = mainSheet.getRange(mainTargetRow, mainOfficeReceiveCol).getValue(); // 今記入した値

    if (currentGuidanceVal && currentOfficeVal) {
        const teacherEmail = getEmailAddress_(targetTeacherRole); // 取得済みの担任ロールを使用
        if (teacherEmail) {
            const teacherNotifySubject = `【進捗】進路部・事務室確認完了: ${submittedClass} ${submittedStudentNum}番 ${submittedStudentName}`;
            const teacherNotifyLink = ss.getUrl() + "#gid=" + mainSheet.getSheetId() + "&range=A" + mainTargetRow;
            let teacherNotifyBody = `進路部および事務室での確認・受領が完了しました。\n\n`;
            teacherNotifyBody += `受付番号: ${submittedReceptionNum}\n`;
            teacherNotifyBody += `進路部確認内容: ${currentGuidanceVal}\n`; // mainDataから取得しても良い
            teacherNotifyBody += `事務室受領内容: ${currentOfficeVal}\n\n`; // 今記入した支払い番号
            teacherNotifyBody += `スプレッドシートで確認:\n${teacherNotifyLink}`;

            GmailApp.sendEmail(teacherEmail, teacherNotifySubject, teacherNotifyBody);
            Logger.log(`担任 (${teacherEmail}) へ進捗通知を送信しました (受付番号: ${submittedReceptionNum})。`);
        } else {
            Logger.log(`警告: 担任ロール "${targetTeacherRole}" のメールアドレスが見つかりません。進捗通知は送信されません。`);
            notifyAdminOnError_(`担任ロール "${targetTeacherRole}" のメールアドレスが見つからず、進捗通知を送信できません (受付番号: ${submittedReceptionNum})。`, functionName);
        }
    } else {
         Logger.log(`「進路部の確認」(${currentGuidanceVal}) または「事務室での受領」(${currentOfficeVal}) が未入力のため、担任への通知はスキップされました。`);
    }


  } catch (error) {
    Logger.log(`エラーが発生しました (${functionName}): ${error} ${error.stack}`);
    // 管理者に通知 (元のスクリプトの関数を流用)
    notifyAdminOnError_(error, functionName);
    // 失敗した場合、事務室にも通知する（オプション）
    try {
        const officeEmailOnError = getEmailAddress_('事務室');
        if(officeEmailOnError) {
            const subject = `【要確認】支払い確認処理エラー`;
            let body = `支払い確認フォームの処理中にエラーが発生しました。\n\n`;
            body += `エラー内容: ${error.message}\n`;
            if(paymentRowIdx > 0 && paymentSheet) { // paymentRowIdxが設定されていればリンクを追加
                 body += `\n関連する可能性のある支払い確認シートの行:\n`;
                 body += SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + paymentSheet.getSheetId() + "&range=A" + paymentRowIdx + "\n";
            }
             body += `\n管理者に詳細なエラーが通知されています。確認してください。`;
            GmailApp.sendEmail(officeEmailOnError, subject, body);
        }
    } catch (e) {
        Logger.log(`事務室へのエラー通知送信中にさらにエラー: ${e}`);
    }
  }
}

// --- ヘルパー関数: 支払い確認シートのヘッダーを確認・追加 ---
// 元の checkAndAddHeaders_ とほぼ同じだが、対象シートと必須ヘッダーが異なる
function checkAndAddPaymentHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const headersToAdd = [];
  const headerPositions = {}; // {HeaderName: ColumnIndex(1-based)}

  headers.forEach((header, index) => {
    if (header) headerPositions[header.trim()] = index + 1;
  });

  PAYMENT_REQUIRED_HEADERS.forEach(requiredHeader => {
    if (!headerPositions[requiredHeader]) {
      headersToAdd.push(requiredHeader);
    }
  });

  if (headersToAdd.length > 0) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol, 1, headersToAdd.length).setValues([headersToAdd]);
    Logger.log(`支払い確認シートに不足ヘッダーを追加: ${headersToAdd.join(', ')}`);
    headersToAdd.forEach((header, index) => {
        headerPositions[header.trim()] = nextCol + index;
    });
  }

   // 必須ヘッダーの位置を再確認
   PAYMENT_REQUIRED_HEADERS.forEach(requiredHeader => {
      if (!headerPositions[requiredHeader]) {
           notifyAdminOnError_(`支払い確認シートの必須ヘッダー "${requiredHeader}" が見つからないか、追加できませんでした。`, 'checkAndAddPaymentHeaders_');
           throw new Error(`支払い確認シートの必須ヘッダー "${requiredHeader}" が見つからないか、追加できませんでした。`);
       }
   });


  return headerPositions;
}


// --- ヘルパー関数: バリデーションエラーメール送信 ---
function sendValidationErrorEmail_(recipient, subjectPrefix, errorMessage, link) {
  const fullSubject = `【要確認:${subjectPrefix}】支払い確認フォーム`;
  let body = `支払い確認フォームの送信内容に問題があり、処理を中断しました。\n`;
  body += `内容を確認し、必要に応じて修正や連絡を行ってください。\n\n`;
  body += `--------------------\n`;
  body += `エラー内容:\n${errorMessage}\n\n`;
  body += `該当するフォーム回答へのリンク:\n${link}\n`;
  body += `--------------------\n`;

  try {
      GmailApp.sendEmail(recipient, fullSubject, body);
      Logger.log(`バリデーションエラーメールを ${recipient} に送信しました: ${subjectPrefix}`);
  } catch (e) {
      Logger.log(`エラー: バリデーションエラーメールの送信に失敗しました (${recipient})。エラー: ${e}`);
      // 管理者に通知
      notifyAdminOnError_(`バリデーションエラーメール(${subjectPrefix})の送信失敗: ${e}`, 'sendValidationErrorEmail_');
  }
}

// --- トリガー設定 ---
// この processPaymentConfirmation 関数に対して、
// 支払い確認フォームのスプレッドシートから「フォーム送信時」のトリガーを
// 手動で設定する必要があります。