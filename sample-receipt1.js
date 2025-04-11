/**
 * @OnlyCurrentDoc
 */

// --- 設定 (支払い確認処理用) ---
// !!注意!! これらのシート名は実際のシート名に合わせてください
const PAYMENT_CONFIRMATION_SHEET_NAME = "支払い確認"; // 支払い確認フォームの回答シート名
const TRANSCRIPT_REQUEST_SHEET_NAME = "フォームの回答"; // 調査書作成願の回答シート名
const SETTINGS_SHEET_NAME_PAY = "設定"; // 設定シート名 (前のスクリプトと同じものを使用)

// 支払い確認シートのカラム (1-based index)
const P_TIMESTAMP_COL = 1;
const P_CLASS_COL = 2;
const P_STUDENT_NUMBER_COL = 3;
const P_STUDENT_NAME_COL = 4;
const P_RECEPTION_NUMBER_INPUT_COL = 5; // 生徒が入力した受付番号
const P_PAYMENT_NUMBER_COL = 6;       // 生徒が入力した支払い番号

// 調査書作成願シートのカラム (1-based index) - 主要なもの
// (注意: これらの値は最初のスクリプトと一致させる必要があります)
const T_EMAIL_COL = 2;
const T_CLASS_COL = 3;
const T_STUDENT_NUMBER_COL = 4;
const T_STUDENT_NAME_COL = 5;
// '受付番号' と '事務室での受領' の列番号は checkAndAddHeaders_ で動的に取得します

const PAYMENT_REQUIRED_HEADERS = ["支払い確認ID"]; // 支払い確認シートに必要なヘッダー

// --- ヘルパー関数: 支払い確認シートのヘッダーを確認し追加 ---
function checkAndAddPaymentHeaders_(sheet) {
    if (!sheet) {
        throw new Error(`支払い確認シートが見つかりません。`);
    }
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    const headersToAdd = [];
    const headerPositions = {};

    // 現在の位置を取得
    headers.forEach((header, index) => {
        if (header) {
            headerPositions[header.trim()] = index + 1;
        }
    });

    // 不足しているものを確認
    PAYMENT_REQUIRED_HEADERS.forEach(requiredHeader => {
        if (!headerPositions[requiredHeader]) {
            headersToAdd.push(requiredHeader);
        }
    });

    // 不足ヘッダーを追加
    if (headersToAdd.length > 0) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol, 1, headersToAdd.length).setValues([headersToAdd]);
        Logger.log(`支払い確認シートに不足ヘッダーを追加: ${headersToAdd.join(', ')}`);
        headersToAdd.forEach((header, index) => {
            headerPositions[header.trim()] = nextCol + index;
        });
    }

    // 必須ヘッダーの位置を再取得
    PAYMENT_REQUIRED_HEADERS.forEach(requiredHeader => {
        let found = false;
        for (const header in headerPositions) {
            if (header === requiredHeader) {
                headerPositions[requiredHeader] = headerPositions[header];
                found = true;
                break;
            }
        }
        if (!found) {
            const errorMessage = `支払い確認シートの必須ヘッダー "${requiredHeader}" が見つからないか、追加できませんでした。`;
            notifyAdminOnError_(errorMessage, 'checkAndAddPaymentHeaders_'); // 管理者に通知
            throw new Error(errorMessage);
        }
    });

    return headerPositions;
}

// --- ヘルパー関数: 事務室へエラー通知 ---
function sendErrorToOffice_(subject, body) {
    try {
        const officeEmail = getEmailAddress_('事務室'); // 既存のヘルパー関数を使用
        if (officeEmail) {
            GmailApp.sendEmail(officeEmail, subject, body);
            Logger.log(`事務室 (${officeEmail}) へエラー通知を送信しました: ${subject}`);
        } else {
            Logger.log(`警告: 事務室のメールアドレスが設定に見つかりません。通知を送信できません。件名: ${subject}`);
            // 管理者にも通知した方が良いかもしれない
            notifyAdminOnError_(`事務室のメールアドレスが見つからず、エラー通知 (${subject}) を送信できませんでした。`, 'sendErrorToOffice_');
        }
    } catch (error) {
        Logger.log(`事務室へのエラー通知送信中にエラーが発生しました。件名: ${subject}, エラー: ${error}`);
        notifyAdminOnError_(`事務室へのエラー通知送信中にエラーが発生しました。件名: ${subject}, エラー: ${error}`, 'sendErrorToOffice_');
    }
}


// --- 支払い確認フォーム送信を処理する関数 ---
function processPaymentSubmission(e) {
    const functionName = 'processPaymentSubmission';
    let paymentSheet, transcriptSheet, settingsSheet; // スコープを広くする

    try {
        // --- スプレッドシートとシートを取得 ---
        // このスクリプトがどちらのスプレッドシートに紐づいているかで挙動が変わる可能性あり
        // SpreadsheetApp.getActiveSpreadsheet() はスクリプトが紐づくSSを返す
        // もし別々のSSなら、IDで開く必要があるかもしれないが、ひとまず同じSS内と仮定
        const ss = SpreadsheetApp.getActiveSpreadsheet(); // トリガー元のSSを想定
        paymentSheet = ss.getSheetByName(PAYMENT_CONFIRMATION_SHEET_NAME);
        transcriptSheet = ss.getSheetByName(TRANSCRIPT_REQUEST_SHEET_NAME);
        settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME_PAY); // getEmailAddress_で内部的に使われる

        if (!paymentSheet) throw new Error(`シート "${PAYMENT_CONFIRMATION_SHEET_NAME}" が見つかりません。`);
        if (!transcriptSheet) throw new Error(`シート "${TRANSCRIPT_REQUEST_SHEET_NAME}" が見つかりません。`);
        if (!settingsSheet) throw new Error(`シート "${SETTINGS_SHEET_NAME_PAY}" が見つかりません。`); // getEmailAddress_ が失敗する前兆

        // --- ヘッダー確認 ---
        const paymentHeaderPositions = checkAndAddPaymentHeaders_(paymentSheet);
        const paymentConfirmIdCol = paymentHeaderPositions['支払い確認ID'];
        // 調査書作成願シートのヘッダーも必要（特に受付番号と事務室受領）
        // 既存の checkAndAddHeaders_ を流用する（同じプロジェクト内にある前提）
        const transcriptHeaderPositions = checkAndAddHeaders_(); // 調査書作成願シートのヘッダー情報を取得
        const transcriptReceptionCol = transcriptHeaderPositions['受付番号'];
        const transcriptOfficeCol = transcriptHeaderPositions['事務室での受領'];


        // --- 新規行データを取得 ---
        const range = e.range; // フォーム送信イベントから編集範囲を取得
        const rowIdx = range.getRowIndex();
        if (rowIdx <= 1) return; // ヘッダー行はスキップ

        const paymentData = paymentSheet.getRange(rowIdx, 1, 1, paymentSheet.getLastColumn()).getValues()[0];

        // --- 入力データを抽出 ---
        const studentClass = paymentData[P_CLASS_COL - 1];
        const studentNumber = paymentData[P_STUDENT_NUMBER_COL - 1];
        const studentName = paymentData[P_STUDENT_NAME_COL - 1]; // メール送信用
        const enteredReceptionNumber = paymentData[P_RECEPTION_NUMBER_INPUT_COL - 1];
        const paymentNumber = paymentData[P_PAYMENT_NUMBER_COL - 1];

        // --- バリデーション１: 支払い番号の重複チェック ---
        // 自分自身を除いた支払い番号リストを取得
        const allPaymentNumbers = paymentSheet.getRange(2, P_PAYMENT_NUMBER_COL, paymentSheet.getLastRow() - 1, 1)
            .getValues()
            .flat() // 2D配列を1D配列に
            .filter((val, index) => val && (index + 2) !== rowIdx); // 空白と自分自身を除外

        if (allPaymentNumbers.includes(paymentNumber)) {
            const subject = `【要確認】支払い番号重複エラー`;
            const body = `支払い確認フォームで入力された支払い番号が、既に使用されています。\n\n`
                + `シート: ${PAYMENT_CONFIRMATION_SHEET_NAME}\n`
                + `行番号: ${rowIdx}\n`
                + `クラス: ${studentClass}\n`
                + `出席番号: ${studentNumber}\n`
                + `入力された受付番号: ${enteredReceptionNumber}\n`
                + `重複した支払い番号: ${paymentNumber}\n\n`
                + `対応する行を確認してください。`
                + `\n\n`
                + `リンク:\n`
                + `  - 支払い確認シート: ${paymentSheet.getUrl()}#gid=${paymentSheet.getSheetId()}&range=A${rowIdx}\n`;
            sendErrorToOffice_(subject, body);
            return; // 処理中断
        }

        // --- バリデーション２: 受付番号と生徒情報の一致確認 ---
        const transcriptData = transcriptSheet.getRange(2, 1, transcriptSheet.getLastRow() - 1, transcriptSheet.getLastColumn()).getValues();
        let foundTranscriptRowIndex = -1; // 調査書作成願シートでの行インデックス (0-based)
        let studentEmail = null; // 一致した場合の生徒メールアドレス

        for (let i = 0; i < transcriptData.length; i++) {
            const transcriptRow = transcriptData[i];
            const assignedReceptionNumber = transcriptRow[transcriptReceptionCol - 1];
            const transcriptClass = transcriptRow[T_CLASS_COL - 1];
            const transcriptStudentNumber = transcriptRow[T_STUDENT_NUMBER_COL - 1];

            // 受付番号が数値として一致するか、かつクラス・番号も一致するか確認
            if (assignedReceptionNumber && parseInt(assignedReceptionNumber) === parseInt(enteredReceptionNumber) &&
                transcriptClass === studentClass &&
                String(transcriptStudentNumber) === String(studentNumber)) // 文字列比較が安全
            {
                foundTranscriptRowIndex = i; // 0-based index
                studentEmail = transcriptRow[T_EMAIL_COL - 1]; // 生徒のメールアドレスを取得
                break; // 一致する行が見つかったらループ終了
            }
        }

        if (foundTranscriptRowIndex === -1) {
            const subject = `【要確認】受付番号・生徒情報不一致エラー`;
            let body = `支払い確認フォームで入力された受付番号に対応する調査書作成願が見つからないか、\n`
                + `クラス・出席番号が一致しませんでした。\n\n`
                + `支払い確認シート情報:\n`
                + `  シート: ${PAYMENT_CONFIRMATION_SHEET_NAME}\n`
                + `  行番号: ${rowIdx}\n`
                + `  クラス: ${studentClass}\n`
                + `  出席番号: ${studentNumber}\n`
                + `  入力された受付番号: ${enteredReceptionNumber}\n`
                + `  支払い番号: ${paymentNumber}\n\n`
                + `支払い確認シート (${PAYMENT_CONFIRMATION_SHEET_NAME}) の行 ${rowIdx} を確認してください。\n`
                + `もしくは、調査書作成願シート (${TRANSCRIPT_REQUEST_SHEET_NAME}) の行 ${foundTranscriptRowIndex + 2} を確認してください。\n\n`
                + `リンク:\n`
                + `  - 支払い確認シート: ${paymentSheet.getUrl()}#gid=${paymentSheet.getSheetId()}&range=A${rowIdx}\n`;
            sendErrorToOffice_(subject, body);
            return; // 処理中断
        }

        // --- バリデーション通過後の処理 ---

        // 1. 支払い確認IDを記入 (Auto Increment)
        let lastPaymentConfirmId = 0;
        if (rowIdx > 2) { // 最初のデータ行でない場合
            const paymentConfirmIds = paymentSheet.getRange(2, paymentConfirmIdCol, rowIdx - 2, 1).getValues();
            for (let i = paymentConfirmIds.length - 1; i >= 0; i--) {
                if (paymentConfirmIds[i][0] && !isNaN(parseInt(paymentConfirmIds[i][0]))) {
                    lastPaymentConfirmId = parseInt(paymentConfirmIds[i][0]);
                    break;
                }
            }
        }
        const newPaymentConfirmId = lastPaymentConfirmId + 1;
        paymentSheet.getRange(rowIdx, paymentConfirmIdCol).setValue(newPaymentConfirmId);
        Logger.log(`支払い確認ID ${newPaymentConfirmId} を ${PAYMENT_CONFIRMATION_SHEET_NAME} の行 ${rowIdx} に記入しました。`);

        // 2. 生徒へ支払い確認メールを送信
        if (studentEmail) {
            const mailSubject = "【支払い確認】調査書作成願の支払いを確認しました";
            let mailBody = `${studentName} さん (${studentClass} ${studentNumber}番)\n\n`;
            mailBody += `調査書作成願 (受付番号: ${enteredReceptionNumber}) の支払い確認フォームの送信を受け付けました。\n\n`;
            mailBody += `支払い番号: ${paymentNumber}\n\n`;
            mailBody += `内容を確認し、調査書作成を進めます。\n`;
            mailBody += `調査書が作成できましたら、改めてメールでご連絡します。\n`;
            mailBody += `しばらくお待ちください。`;
            GmailApp.sendEmail(studentEmail, mailSubject, mailBody);
            Logger.log(`生徒 (${studentEmail}) へ支払い確認メールを送信しました。受付番号: ${enteredReceptionNumber}`);
        } else {
            Logger.log(`警告: 調査書作成願シートの行 ${foundTranscriptRowIndex + 2} で生徒のメールアドレスが見つかりません。確認メールは送信されませんでした。`);
            // この場合も事務室や管理者に通知した方が良いかもしれない
            sendErrorToOffice_(`【情報】生徒メールアドレス不備`,
                `受付番号 ${enteredReceptionNumber} の支払い確認は完了しましたが、`
                + `調査書作成願シートに対応する生徒のメールアドレスがありませんでした。\n`
                + `生徒への確認メールは送信されていません。\n リンク:\n -支払い確認シート: ${paymentSheet.getUrl()}#gid=${paymentSheet.getSheetId()}&range=A${rowIdx}\n`);
        }

        // 3. 調査書作成願いシートの「事務室での受領」列に支払い番号を記入
        const transcriptTargetRow = foundTranscriptRowIndex + 2; // 1-based sheet row index
        transcriptSheet.getRange(transcriptTargetRow, transcriptOfficeCol).setValue(paymentNumber);
        Logger.log(`調査書作成願シート (${TRANSCRIPT_REQUEST_SHEET_NAME}) の行 ${transcriptTargetRow}, 列 ${transcriptOfficeCol} に支払い番号 ${paymentNumber} を記入しました。`);


    } catch (error) {
        Logger.log(`[${functionName}] でエラーが発生しました: ${error} \nStack: ${error.stack}`);
        // 調査書作成願シートの '事務室での受領' が更新される前にエラーが起きた可能性がある
        // paymentSheet が定義されているか確認
        let errorContext = `支払い確認フォーム処理中にエラーが発生しました。\nシート: ${paymentSheet ? paymentSheet.getName() : '不明'}\n`;
        try {
            const range = e.range;
            if (range) {
                errorContext += `処理中の行: ${range.getRowIndex()}\n`;
                errorContext += `入力値 (抜粋): ${paymentSheet.getRange(range.getRowIndex(), P_CLASS_COL, 1, 4).getValues()[0].join(', ')}\n`; // クラス〜受付番号あたり
            }
        } catch (e2) {
            errorContext += "詳細なコンテキスト取得に失敗しました。\n";
        }

        notifyAdminOnError_(errorContext + `エラー詳細: ${error}`, functionName);

        // 事務室にも基本的なエラー発生を通知する（詳細は管理者に）
        sendErrorToOffice_(`【重要】支払い確認処理エラー`,
            `支払い確認フォームの自動処理中にエラーが発生しました。\n`
            + `管理者へ詳細なエラーが通知されています。\n`
            + `手動での確認・対応が必要な可能性があります。\n\n`
            + `エラー概要: ${error.message} \n リンク：\n - 支払い確認シート: ${paymentSheet.getUrl()}#gid=${paymentSheet.getSheetId()}&range=A${e.range.getRowIndex()}\n`);
    }
}


// --- トリガー設定関数 ---
// この関数に対応するトリガーを Apps Script エディタで手動で設定する必要があります。
// 1. processPaymentSubmission のトリガー:
//    - イベントのソース: スプレッドシートから (支払い確認フォームが紐づくスプレッドシートを選択)
//    - イベントの種類: フォーム送信時
//    - 関数: processPaymentSubmission

// --- 既に関数がある場合の注意 ---
// 同じプロジェクトに最初のスクリプトがある場合、以下の関数は既に存在します。
// 重複して定義しないようにしてください。
/*
function getEmailAddress_(role) { ... }
function notifyAdminOnError_(error, functionName) { ... }
function checkAndAddHeaders_() { ... } // これは調査書作成願シート用なので、Payment用とは別
*/