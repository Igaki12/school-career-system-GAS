
/**
* @OnlyCurrentDoc
*/

// --- 設定項目 ---
const SHEET_NAME = "school-career-system-sample"; // 実際のシート名に合わせてください
const SHIDOBU_EMAIL = "shido@example.com"; // ★★★ 進路指導部のメールアドレスに書き換えてください ★★★
const HEADER_ROW = 1; // ヘッダー行の行番号 (通常は1)

// 列番号 (A=1, B=2, ...) ※変更する場合注意
const CLASS_COL = 3;         // クラス
const NUMBER_COL = 4;        // 出席番号
const NAME_COL = 5;          // 名前
const UNIV_NAME_COL = 7;     // 大学名
const FACULTY_COL = 8;       // 学部
const DEPARTMENT_COL = 9;    // 学科
const TANNNIN_CHECK_COL = 13; // 担任の確認 (M列)
const SHIDOBU_CHECK_COL = 14; // 進路指導部の確認 (N列)
const AI_CHECK_COL = 15;      // GeminiによるAI確認 (O列)
const JIMU_CHECK_COL = 16;    // 事務部での料金受け取り (P列)

// --- 関数 ---

/**
 * スプレッドシートを開いた時にカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('進路調査票処理')
    .addItem('ヘッダー追加 (初回のみ)', 'addHeaders')
    .addToUi();
}

/**
 * M列以降に必要なヘッダーを追加します。(初回実行用)
 */
function addHeaders() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`シート "${SHEET_NAME}" が見つかりません。`);
    }
    // ヘッダーが既に追加されているか簡易チェック
    if (sheet.getRange(HEADER_ROW, TANNNIN_CHECK_COL).getValue() === "担任の確認") {
      SpreadsheetApp.getUi().alert("ヘッダーは既に追加されているようです。");
      return;
    }

    sheet.getRange(HEADER_ROW, TANNNIN_CHECK_COL).setValue("担任の確認");
    sheet.getRange(HEADER_ROW, SHIDOBU_CHECK_COL).setValue("進路指導部の確認");
    sheet.getRange(HEADER_ROW, AI_CHECK_COL).setValue("GeminiによるAI確認");
    sheet.getRange(HEADER_ROW, JIMU_CHECK_COL).setValue("事務部での料金受け取り");
    SpreadsheetApp.getUi().alert("M列以降にヘッダーを追加しました。");
  } catch (error) {
    Logger.log(`ヘッダー追加エラー: ${error}`);
    SpreadsheetApp.getUi().alert(`ヘッダー追加中にエラーが発生しました。\n${error.message}`);
  }
}


/**
 * スプレッドシートが編集されたときに実行されるトリガー関数
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e イベントオブジェクト
 */
function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();

    // 編集されたシートが対象シートか、ヘッダー行でないかを確認
    if (sheet.getName() !== SHEET_NAME || range.getRow() <= HEADER_ROW) {
      return;
    }

    const editedCol = range.getColumn();
    const editedRow = range.getRow();
    const editedValue = e.value; // 編集後の値

    // --- 担任の確認 → 進路指導部へメール ---
    if (editedCol === TANNNIN_CHECK_COL && editedValue) { // 値が入力された場合 (空文字でない)
      sendMailToShidobu(sheet, editedRow, editedValue);
    }

    // --- 進路指導部の確認 & 事務部確認 → 担任へメール ---
    if ((editedCol === SHIDOBU_CHECK_COL || editedCol === JIMU_CHECK_COL) && editedValue) {
      const shidobuCheckValue = sheet.getRange(editedRow, SHIDOBU_CHECK_COL).getValue();
      const jimuCheckValue = sheet.getRange(editedRow, JIMU_CHECK_COL).getValue();

      // 両方の列に値が入っている場合のみメール送信
      if (shidobuCheckValue && jimuCheckValue) {
        sendMailToTannin(sheet, editedRow, shidobuCheckValue, jimuCheckValue);
      }
    }
  } catch (error) {
    Logger.log(`onEditエラー: ${error}\nスタックトレース: ${error.stack}`);
    // エラーが発生しても他の編集操作を妨げないように、ここではUIアラートは出さない方が良い場合もある
    // SpreadsheetApp.getUi().alert(`処理中にエラーが発生しました。\n${error.message}`);
  }
}

/**
 * 担任の確認内容を進路指導部へメール送信します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 対象行番号
 * @param {string} checkComment 担任の確認コメント
 */
function sendMailToShidobu(sheet, row, checkComment) {
  try {
    const studentData = sheet.getRange(row, 1, 1, DEPARTMENT_COL).getValues()[0]; // A列から学科列まで取得
    const timestamp = Utilities.formatDate(new Date(studentData[0]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const studentClass = studentData[CLASS_COL - 1];
    const studentNumber = studentData[NUMBER_COL - 1];
    const studentName = studentData[NAME_COL - 1];
    const univName = studentData[UNIV_NAME_COL - 1];
    const faculty = studentData[FACULTY_COL - 1];
    const department = studentData[DEPARTMENT_COL - 1];

    const subject = `【要確認】${studentClass} ${studentNumber}番 ${studentName} さんの調査書申請（担任確認）`;
    const body = `進路指導部 ご担当者様

${studentClass} ${studentNumber}番 ${studentName} さん (${univName} ${faculty} ${department}) の
調査書作成願について、担任による確認が行われました。

--------------------
担任確認コメント:
${checkComment}
--------------------

スプレッドシートをご確認ください。
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
`;

    GmailApp.sendEmail(SHIDOBU_EMAIL, subject, body);
    Logger.log(`進路指導部へのメール送信完了 (行: ${row})`);

  } catch (error) {
    Logger.log(`進路指導部へのメール送信エラー (行: ${row}): ${error}\nスタックトレース: ${error.stack}`);
  }
}

/**
 * 進路指導部と事務部の確認完了を担任へメール送信します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 対象行番号
 * @param {string} shidoComment 進路指導部の確認コメント
 * @param {string} jimuComment 事務部の確認コメント
 */
function sendMailToTannin(sheet, row, shidoComment, jimuComment) {
  try {
    const studentData = sheet.getRange(row, 1, 1, DEPARTMENT_COL).getValues()[0]; // A列から学科列まで取得
    const studentClass = studentData[CLASS_COL - 1];
    const studentNumber = studentData[NUMBER_COL - 1];
    const studentName = studentData[NAME_COL - 1];
    const univName = studentData[UNIV_NAME_COL - 1];
    const faculty = studentData[FACULTY_COL - 1];
    const department = studentData[DEPARTMENT_COL - 1];

    const tanninEmail = getTanninEmail(studentClass);
    if (!tanninEmail) {
      Logger.log(`担任メールアドレスが見つかりません (クラス: ${studentClass}, 行: ${row})`);
      return; // 担任アドレス不明なら送信しない
    }

    const subject = `【確認完了】${studentClass} ${studentNumber}番 ${studentName} さんの調査書申請`;
    const body = `${studentClass} 担任の先生

${studentNumber}番 ${studentName} さん (${univName} ${faculty} ${department}) の
調査書作成願について、進路指導部および事務部の確認が完了しました。

--------------------
進路指導部 確認:
${shidoComment}

事務部 確認（料金受取）:
${jimuComment}
--------------------

対応が完了しました。
`;

    GmailApp.sendEmail(tanninEmail, subject, body);
    Logger.log(`担任へのメール送信完了 (行: ${row}, クラス: ${studentClass})`);

  } catch (error) {
    Logger.log(`担任へのメール送信エラー (行: ${row}): ${error}\nスタックトレース: ${error.stack}`);
  }
}


/**
 * クラス名から担任のメールアドレスを取得します。
 * ★★★ この関数は実際の運用に合わせて修正が必要です ★★★
 * 例:
 * - 各クラス担任のメールアドレスを直接記述する
 * - 別の設定シートから参照する
 * - Google Groupsなどを利用する
 * @param {string} className クラス名 (例: "A組")
 * @return {string|null} 担任のメールアドレス or null
 */
function getTanninEmail(className) {
  // ★★★ ここを実際の担任メールアドレスに書き換えるか、別の方法で取得してください ★★★
  const mapping = {
    "A組": "tannin-a@example.com",
    "B組": "tannin-b@example.com",
    "C組": "tannin-c@example.com",
    "D組": "tannin-d@example.com",
    "E組": "tannin-e@example.com",
    "F組": "tannin-f@example.com",
    "G組": "tannin-g@example.com",
    "H組": "tannin-h@example.com"
  };
  return mapping[className] || null; // マッピングにないクラス名はnullを返す
}
// 生成に使用したプロンプト: Gemini 2.5 pro
// この進路調査票スプレッドシートにGASを使って、確認メール送信やスキャン文書の添付を行いたい。以下の条件を満たすGASコードを出力して。
// # sample.csvのデータ右側(M列以降)に、実際に受験資格を満たしているか、「担任の確認」、「進路指導部の確認」、「GeminiによるAI確認」の列をそれぞれ追加して。その右側に、「事務部での料金受け取り」列を追加して。
// # 各行の「担任の確認」欄に文字が入力されたら、その文字を進路指導部のアドレスにGmailで送信して。
// # 各行の「進路指導部の確認」欄と「事務部での料金受け取り」欄の両方に文字が入力されたら、その文字両方をわかりやすい文章で担任のアドレスにGmailで送信して。
// # GeminiなどのAPIを使って、実際に受験資格を満たしているか調べることはできる？これは日本の高校Google Workspace上で動作させることを想定しているが、技術的に可能であるかどうか？また、プライバシーの問題についてはどうか？


