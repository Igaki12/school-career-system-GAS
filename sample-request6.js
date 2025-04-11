/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const FORM_RESPONSE_SHEET_NAME = "フォームの回答"; // Change if your sheet name is different
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
// Add more university columns if needed for the initial email

// --- Helper Function: Get Email Address from Settings ---
function getEmailAddress_(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    Logger.log(`Error: Settings sheet "${SETTINGS_SHEET_NAME}" not found.`);
    notifyAdminOnError_(`Settings sheet "${SETTINGS_SHEET_NAME}" not found.`, 'getEmailAddress_');
    return null;
  }
  const data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === role) {
      return data[i][1];
    }
  }
  Logger.log(`Warning: Email address for role "${role}" not found in settings.`);
  // Optionally notify admin if a role is missing
  // notifyAdminOnError_(`Email address for role "${role}" not found in settings.`, 'getEmailAddress_');
  return null; // Role not found
}

// --- Helper Function: Notify Admin on Error ---
function notifyAdminOnError_(error, functionName) {
  try {
    const adminEmail = getEmailAddress_('管理者');
    if (adminEmail) {
      const subject = `GAS Script Error: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`;
      let body = `An error occurred in the script.\n\n`;
      body += `Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getName()}\n`;
      body += `Sheet: ${SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()}\n`;
      body += `Function: ${functionName}\n`;
      body += `Error: ${error.message || error}\n`;
      if (error.stack) {
        body += `Stack Trace:\n${error.stack}\n`;
      }
      // Add context if available (e.g., active row/column if possible)
      try {
        const activeRange = SpreadsheetApp.getActiveRange();
        if (activeRange) {
           body += `Active Cell: ${activeRange.getA1Notation()}\n`;
           body += `Edited Value: ${activeRange.getValue()}\n`;
        }
      } catch(e) {
         body += `Could not retrieve active cell info.\n`
      }

      Logger.log(`Sending error notification to ${adminEmail}. Error: ${error}`);
      GmailApp.sendEmail(adminEmail, subject, body);
    } else {
      Logger.log(`Error: Admin email not found in settings. Cannot send error notification. Error was: ${error}`);
    }
  } catch (e) {
    Logger.log(`CRITICAL ERROR: Could not send error notification. Original error: ${error}. Notification error: ${e}`);
  }
}


// --- Helper Function: Check and Add Headers ---
function checkAndAddHeaders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
  if (!sheet) {
    notifyAdminOnError_(`Form response sheet "${FORM_RESPONSE_SHEET_NAME}" not found.`, 'checkAndAddHeaders_');
    throw new Error(`Sheet "${FORM_RESPONSE_SHEET_NAME}" not found.`);
  }
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const headersToAdd = [];
  const headerPositions = {};

  // Get current positions
   headers.forEach((header, index) => {
    if (header) { // Check if header is not empty
        headerPositions[header.trim()] = index + 1; // 1-based index
    }
   });

  // Check which headers are missing
  REQUIRED_HEADERS.forEach(requiredHeader => {
    if (!headerPositions[requiredHeader]) {
      headersToAdd.push(requiredHeader);
    }
  });

  // Add missing headers
  if (headersToAdd.length > 0) {
    const nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol, 1, headersToAdd.length).setValues([headersToAdd]);
    Logger.log(`Added missing headers: ${headersToAdd.join(', ')}`);
    // Update header positions after adding
    headersToAdd.forEach((header, index) => {
        headerPositions[header.trim()] = nextCol + index;
    });
  }

   // Re-fetch all required header positions
   REQUIRED_HEADERS.forEach(requiredHeader => {
      let found = false;
      for (const header in headerPositions) {
          if (header === requiredHeader) {
              headerPositions[requiredHeader] = headerPositions[header]; // Ensure the key is exactly as in REQUIRED_HEADERS
              found = true;
              break;
          }
      }
       if (!found) {
           notifyAdminOnError_(`Required header "${requiredHeader}" could not be found or added.`, 'checkAndAddHeaders_');
           throw new Error(`Required header "${requiredHeader}" could not be found or added.`);
       }
   });


  return headerPositions;
}

// --- Function to Process New Form Submissions ---
function processNewSubmission(e) {
  const functionName = 'processNewSubmission';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);

    // If triggered by form submit, e.range exists. Otherwise, assume last row.
    const rowIdx = e && e.range ? e.range.getRowIndex() : sheet.getLastRow();
    if (rowIdx <= 1) return; // Skip header row

    const headerPositions = checkAndAddHeaders_(); // Check headers first
    const receptionCol = headerPositions['受付番号'];

    // Get data for the new row
    const rowData = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];

    // --- Assign Reception Number ---
    const currentReceptionNumber = rowData[receptionCol - 1];
    let newReceptionNumber = currentReceptionNumber;

    if (!currentReceptionNumber) { // Only assign if it's empty
      let lastReceptionNumber = 0;
      if (rowIdx > 2) { // Check previous row if not the first data row
        const prevReceptionVal = sheet.getRange(rowIdx - 1, receptionCol).getValue();
        if (prevReceptionVal && !isNaN(parseInt(prevReceptionVal))) {
          lastReceptionNumber = parseInt(prevReceptionVal);
        } else { // Fallback: Search upwards if previous row is empty or invalid
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
      Logger.log(`Assigned Reception Number ${newReceptionNumber} to row ${rowIdx}.`);

      // --- Send Confirmation Email to Student ---
      const studentEmail = rowData[STUDENT_EMAIL_COL - 1];
      const studentName = rowData[STUDENT_NAME_COL - 1];
      const studentClass = rowData[CLASS_COL - 1];
      const studentNumber = rowData[STUDENT_NUMBER_COL - 1];
      const univCode1 = rowData[UNIV_CODE_1_COL -1]; // G
      const univName1 = rowData[UNIV_NAME_1_COL - 1]; // H
      const faculty1 = rowData[FACULTY_1_COL - 1]; // I
      const dept1 = rowData[DEPT_1_COL - 1]; // J - Included based on request context

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
        // Add more fields here if needed, checking subsequent columns
        body += "\n------------------------------\n";
        body += "今後の手続きについては、別途連絡をお待ちください。";

        GmailApp.sendEmail(studentEmail, subject, body);
        Logger.log(`Sent confirmation email to ${studentEmail} for row ${rowIdx}.`);
      } else {
        Logger.log(`Warning: No student email found in row ${rowIdx}. Cannot send confirmation.`);
      }
    } else {
        Logger.log(`Row ${rowIdx} already has a Reception Number (${currentReceptionNumber}). Skipping assignment and confirmation email.`);
    }

  } catch (error) {
    Logger.log(`Error in ${functionName}: ${error}`);
    notifyAdminOnError_(error, functionName);
  }
}


// --- Function to Handle Edits on the Sheet ---
function onSheetEdit(e) {
  const functionName = 'onSheetEdit';
  try {
    const range = e.range;
    const sheet = range.getSheet();

    // Check if the edit is on the correct sheet and not in the header row
    if (sheet.getName() !== FORM_RESPONSE_SHEET_NAME || range.getRow() <= 1) {
      return;
    }

    const editedRow = range.getRow();
    const editedCol = range.getColumn();

    const headerPositions = checkAndAddHeaders_(); // Get header positions
    const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Get common student info
    const studentClass = rowData[CLASS_COL - 1];
    const studentNumber = rowData[STUDENT_NUMBER_COL - 1];
    const studentName = rowData[STUDENT_NAME_COL - 1];
    const studentEmail = rowData[STUDENT_EMAIL_COL - 1];
    const receptionNumber = rowData[headerPositions['受付番号'] - 1];
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=" + sheet.getSheetId() + "&range=A" + editedRow;

    // --- Check specific column edits ---
    const teacherConfirmCol = headerPositions['担任の確認'];
    const guidanceConfirmCol = headerPositions['進路部の確認'];
    const officeReceiveCol = headerPositions['事務室での受領'];
    const transcriptCreateCol = headerPositions['調査書作成'];

    // Check if the cell was actually filled (not cleared)
    const editedValue = range.getValue();
    if (editedValue === "" || editedValue === null || editedValue === undefined) {
        // Logger.log(`Edit cleared cell ${range.getA1Notation()}, skipping notifications.`);
        return; // Don't trigger notifications if cell is cleared
    }


    // 1. Teacher Confirmation -> Notify Guidance Dept
    if (editedCol === teacherConfirmCol) {
      const guidanceEmail = getEmailAddress_('進路部');
      if (guidanceEmail) {
        const teacherConfirmationText = editedValue; // Use the actual value entered
        const subject = `【要確認】担任確認完了: ${studentClass} ${studentNumber}番 ${studentName}`;
        let body = `担任が調査書作成願を確認しました。\n\n`;
        body += `担任入力内容: ${teacherConfirmationText}\n`;
        body += `クラス: ${studentClass}\n`;
        body += `出席番号: ${studentNumber}\n`;
        body += `氏名: ${studentName}\n`;
        body += `受付番号: ${receptionNumber}\n\n`;
        body += `詳細確認・進路部確認入力はこちら:\n${spreadsheetUrl}`;

        GmailApp.sendEmail(guidanceEmail, subject, body);
        Logger.log(`Sent notification to Guidance Dept for row ${editedRow}.`);
      } else {
        Logger.log(`Warning: Guidance Dept email not found. Cannot send notification for row ${editedRow}.`);
         notifyAdminOnError_(`Guidance Dept email not found for Teacher Confirmation notification.`, functionName);
      }
    }

    // 2. Guidance AND Office Confirmed -> Notify Teacher
    // Check if the edited column is one of the two, and then check if BOTH are filled
    if (editedCol === guidanceConfirmCol || editedCol === officeReceiveCol) {
      const guidanceConfirmationText = rowData[guidanceConfirmCol - 1];
      const officeReceptionText = rowData[officeReceiveCol - 1];

      if (guidanceConfirmationText && officeReceptionText) { // Both must have text
        const teacherRole = `${studentClass}担任`; // Construct role name dynamically
        const teacherEmail = getEmailAddress_(teacherRole);
        if (teacherEmail) {
          const subject = `【進捗】進路部・事務室確認完了: ${studentClass} ${studentNumber}番 ${studentName}`;
          let body = `進路部および事務室での確認・受領が完了しました。\n\n`;
          body += `受付番号: ${receptionNumber}\n`;
          body += `進路部確認内容: ${guidanceConfirmationText}\n`;
          body += `事務室受領内容: ${officeReceptionText}\n\n`;
          body += `スプレッドシートで確認:\n${spreadsheetUrl}`;

          GmailApp.sendEmail(teacherEmail, subject, body);
          Logger.log(`Sent notification to Teacher (${teacherEmail}) for row ${editedRow}.`);
        } else {
          Logger.log(`Warning: Teacher email for role "${teacherRole}" not found. Cannot send notification for row ${editedRow}.`);
          notifyAdminOnError_(`Teacher email not found for role "${teacherRole}" for Guidance/Office confirmation notification.`, functionName);
        }
      }
    }

    // 3. All Confirmed (Teacher, Guidance, Office, Creation) -> Notify Student
    // Check if the edited column is one of the four, then check if ALL are filled
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
                Logger.log(`Sent completion notification to student ${studentEmail} for row ${editedRow}.`);
            } else {
                Logger.log(`Warning: Student email missing for row ${editedRow}. Cannot send completion notification.`);
                 notifyAdminOnError_(`Student email missing for row ${editedRow} when sending final completion notice.`, functionName);
            }
        }
     }

  } catch (error) {
    Logger.log(`Error in ${functionName}: ${error}`);
    notifyAdminOnError_(error, functionName);
  }
}

// --- Trigger Setup Functions ---
// You need to manually set up these triggers in the Apps Script editor.
// 1. Trigger for processNewSubmission:
//    - Event Source: Spreadsheet
//    - Event Type: On form submit
// 2. Trigger for onSheetEdit:
//    - Event Source: Spreadsheet
//    - Event Type: On edit

// --- Optional: Function to run manually to process existing rows without reception numbers ---
function processExistingRows() {
  const functionName = 'processExistingRows';
   try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
    if (!sheet) {
        throw new Error(`Sheet ${FORM_RESPONSE_SHEET_NAME} not found.`);
    }

    const headerPositions = checkAndAddHeaders_();
    const receptionCol = headerPositions['受付番号'];
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()); // Start from row 2
    const allData = dataRange.getValues();
    const receptionNumbers = sheet.getRange(2, receptionCol, lastRow -1, 1).getValues();

    let lastAssignedNumber = 0;
     // Find the highest existing number first to avoid duplicates if run multiple times
     for(let i = receptionNumbers.length - 1; i >= 0; i--) {
         if (receptionNumbers[i][0] && !isNaN(parseInt(receptionNumbers[i][0]))) {
            lastAssignedNumber = Math.max(lastAssignedNumber, parseInt(receptionNumbers[i][0]));
         }
     }
     Logger.log(`Highest existing reception number found: ${lastAssignedNumber}`);


    for (let i = 0; i < allData.length; i++) {
      const currentRow = i + 2; // 1-based index for sheet rows
      const currentReceptionNumber = receptionNumbers[i][0]; // Use the dedicated array

      if (!currentReceptionNumber) {
        // Simulate the event object minimally for processNewSubmission
        const fakeEvent = {
          range: sheet.getRange(currentRow, 1) // Provide a range object pointing to the row start
          // values: allData[i] // Not directly used by processNewSubmission logic anymore
        };
        Logger.log(`Processing existing row ${currentRow} which lacks a reception number.`);
        processNewSubmission(fakeEvent); // Call the function to assign number and send email

        // Optional: Add a small delay to avoid exceeding email quotas if processing many rows
        // Utilities.sleep(500); // Sleep for 500 milliseconds
      }
    }
     Logger.log(`Finished processing existing rows.`);
   } catch (error) {
      Logger.log(`Error in ${functionName}: ${error}`);
      notifyAdminOnError_(error, functionName);
   }
}