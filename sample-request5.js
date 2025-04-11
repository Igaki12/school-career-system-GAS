/**
 * @OnlyCurrentDoc
 */

// --- Configuration ---
const CONSTANTS = {
    RESPONSE_SHEET_NAME: "【進路】調査書作成願sample", // Name of the sheet receiving Form Responses
    SETTINGS_SHEET_NAME: "設定",               // Name of the sheet with configuration
    REQUIRED_HEADERS: [
      "担任の確認",
      "進路部の確認",
      "事務室での受領",
      "調査書作成",
      "受付番号"
    ],
    // --- Column Headers (Ensure these match your sheet exactly) ---
    // Response Sheet Columns
    TIMESTAMP_COL_NAME: "タイムスタンプ",
    STUDENT_EMAIL_COL_NAME: "メールアドレスの入力",
    CLASS_COL_NAME: "クラス",
    STUDENT_NUMBER_COL_NAME: "出席番号",
    STUDENT_NAME_COL_NAME: "名前",
    UNIVERSITY_CODE_1_COL_NAME: "大学コード1(8桁か10桁で入力)", // Used in initial confirmation email
    UNIVERSITY_NAME_1_COL_NAME: "大学名",           // Used in initial confirmation email
    FACULTY_1_COL_NAME: "学部1",                 // Used in initial confirmation email
    DEPARTMENT_1_COL_NAME: "学科1",               // Used in initial confirmation email
    // Added Columns
    TEACHER_CONFIRM_COL_NAME: "担任の確認",
    GUIDANCE_CONFIRM_COL_NAME: "進路部の確認",
    OFFICE_RECEIPT_COL_NAME: "事務室での受領",
    REPORT_CREATED_COL_NAME: "調査書作成",
    RECEPTION_NUMBER_COL_NAME: "受付番号",
    // Settings Sheet Keys
    GUIDANCE_EMAIL_KEY: "進路部メール",
    OFFICE_EMAIL_KEY: "事務室メール",
    ADMIN_EMAIL_KEY: "管理者メール",
    TEACHER_EMAIL_KEY_SUFFIX: "担任メール", // e.g., "A組担任メール"
  };
  // --- End Configuration ---
  
  /**
   * Reads configuration from the Settings sheet.
   * @returns {object} Configuration object or null if sheet not found.
   */
  function getConfig() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const settingsSheet = ss.getSheetByName(CONSTANTS.SETTINGS_SHEET_NAME);
      if (!settingsSheet) {
        Logger.log(`Error: Settings sheet "${CONSTANTS.SETTINGS_SHEET_NAME}" not found.`);
        throw new Error(`Settings sheet "${CONSTANTS.SETTINGS_SHEET_NAME}" not found.`);
      }
      const data = settingsSheet.getDataRange().getValues();
      const config = {
        teacherEmails: {}
      };
      data.forEach(row => {
        const key = row[0];
        const value = row[1];
        if (key && value) {
          if (key.endsWith(CONSTANTS.TEACHER_EMAIL_KEY_SUFFIX)) {
            const className = key.replace(CONSTANTS.TEACHER_EMAIL_KEY_SUFFIX, '').trim();
            config.teacherEmails[className] = value;
          } else if (key === CONSTANTS.GUIDANCE_EMAIL_KEY) {
            config.guidanceEmail = value;
          } else if (key === CONSTANTS.OFFICE_EMAIL_KEY) {
            config.officeEmail = value;
          } else if (key === CONSTANTS.ADMIN_EMAIL_KEY) {
            config.adminEmail = value;
          }
        }
      });
      if (!config.adminEmail) {
         Logger.log("Warning: Admin email not found in settings. Error notifications will not be sent.");
         // Consider throwing an error if admin email is mandatory
         // throw new Error("Administrator email is not configured in the Settings sheet.");
      }
       if (!config.guidanceEmail) Logger.log(`Warning: "${CONSTANTS.GUIDANCE_EMAIL_KEY}" not found in settings.`);
       if (!config.officeEmail) Logger.log(`Warning: "${CONSTANTS.OFFICE_EMAIL_KEY}" not found in settings.`);
       if (Object.keys(config.teacherEmails).length === 0) Logger.log(`Warning: No teacher emails found in settings (expecting keys like "A組${CONSTANTS.TEACHER_EMAIL_KEY_SUFFIX}").`);
  
      return config;
    } catch (error) {
      Logger.log(`Error in getConfig: ${error}`);
      // Try to notify admin even if config load failed partially
      const adminEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.SETTINGS_SHEET_NAME)?.getRange("B3")?.getValue(); // Fallback attempt
      if (adminEmail) {
         MailApp.sendEmail(adminEmail, "Script Error: Failed to load configuration", `Error loading configuration from "${CONSTANTS.SETTINGS_SHEET_NAME}":\n\n${error}\n\nStack:\n${error.stack}`);
      }
      return null; // Indicate failure
    }
  }
  
  /**
   * Sends an error notification email to the administrator.
   * @param {Error} error The error object.
   * @param {string} context Additional context about where the error occurred.
   */
  function handleError(error, context) {
    Logger.log(`Error in ${context}: ${error}\nStack: ${error.stack}`);
    try {
      const config = getConfig(); // Attempt to get config again for admin email
      if (config && config.adminEmail) {
        const subject = `調査書作成願 Script Error: ${context}`;
        const body = `An error occurred in the script.\n\nContext: ${context}\nError: ${error}\n\nStack Trace:\n${error.stack}`;
        MailApp.sendEmail(config.adminEmail, subject, body);
      } else {
         Logger.log("Could not send error notification: Admin email not configured or config failed to load.");
         // Optional: Fallback - email the script owner
         // MailApp.sendEmail(Session.getEffectiveUser().getEmail(), `調査書作成願 Script Error (Admin Email Missing): ${context}`, `...`);
      }
    } catch (e) {
      Logger.log(`CRITICAL ERROR: Failed to send error notification email: ${e}`);
    }
  }
  
  
  /**
   * Finds the column number (1-indexed) for a given header name.
   * Caches results for efficiency.
   * @param {Sheet} sheet The sheet object.
   * @param {string} headerName The name of the header to find.
   * @param {object} cache An object to store cached header positions.
   * @returns {number} The 1-based column index, or -1 if not found.
   */
  function findColumn(sheet, headerName, cache) {
    if (cache && cache[headerName]) {
      return cache[headerName];
    }
  
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(headerName);
    const colNum = colIndex !== -1 ? colIndex + 1 : -1;
  
    if (cache && colNum !== -1) {
      cache[headerName] = colNum;
    } else if (colNum === -1) {
        Logger.log(`Warning: Header "${headerName}" not found in sheet "${sheet.getName()}".`);
    }
    return colNum;
  }
  
  /**
   * Checks if required headers exist and adds them if missing.
   * @param {Sheet} sheet The sheet object.
   */
  function checkAndAddHeaders(sheet) {
    try {
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const headersToAdd = [];
  
      CONSTANTS.REQUIRED_HEADERS.forEach(header => {
        if (headers.indexOf(header) === -1) {
          headersToAdd.push(header);
        }
      });
  
      if (headersToAdd.length > 0) {
        sheet.getRange(1, lastCol + 1, 1, headersToAdd.length).setValues([headersToAdd]);
        Logger.log(`Added missing headers: ${headersToAdd.join(', ')} to sheet "${sheet.getName()}"`);
        SpreadsheetApp.flush(); // Ensure header is written before proceeding
      }
    } catch (error) {
       handleError(error, `checkAndAddHeaders on sheet ${sheet.getName()}`);
    }
  }
  
  
  /**
   * Creates a custom menu when the spreadsheet is opened.
   * Also runs the header check on open.
   */
  function onOpen() {
    try {
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       const responseSheet = ss.getSheetByName(CONSTANTS.RESPONSE_SHEET_NAME) || ss.getActiveSheet(); // Fallback to active if named sheet not found
       checkAndAddHeaders(responseSheet);
  
       // Optional: Add a custom menu for manual tasks if needed later
       // SpreadsheetApp.getUi()
       //     .createMenu('調査書管理')
       //     .addItem('Manually Check Headers', 'manualHeaderCheck') // Example
       //     .addToUi();
    } catch (error) {
        handleError(error, "onOpen function");
    }
  }
  
  // --- Main Trigger Functions ---
  
  /**
   * Handles the 'on form submit' event.
   * Assigns reception number and sends confirmation to student.
   * @param {object} e The event object.
   */
  function handleFormSubmit(e) {
    const context = "handleFormSubmit";
    try {
      if (!e || !e.range) {
          Logger.log("handleFormSubmit triggered without event object or range.");
          return; // Exit if no event data
      }
  
      const range = e.range;
      const sheet = range.getSheet();
  
      // Ensure it's the correct sheet (optional but good practice)
      if (sheet.getName() !== CONSTANTS.RESPONSE_SHEET_NAME) {
          Logger.log(`Form submission event on unexpected sheet: ${sheet.getName()}. Ignoring.`);
          return;
      }
  
      const config = getConfig();
      if (!config) return; // Stop if config failed
  
      // Check headers just in case onOpen failed or script was added after headers were needed
      checkAndAddHeaders(sheet);
      const headerCache = {}; // Cache for this function execution
  
      // Get column numbers using cache
      const receptionNumberCol = findColumn(sheet, CONSTANTS.RECEPTION_NUMBER_COL_NAME, headerCache);
      const studentEmailCol = findColumn(sheet, CONSTANTS.STUDENT_EMAIL_COL_NAME, headerCache);
      const classNameCol = findColumn(sheet, CONSTANTS.CLASS_COL_NAME, headerCache);
      const studentNumCol = findColumn(sheet, CONSTANTS.STUDENT_NUMBER_COL_NAME, headerCache);
      const studentNameCol = findColumn(sheet, CONSTANTS.STUDENT_NAME_COL_NAME, headerCache);
      const uniNameCol = findColumn(sheet, CONSTANTS.UNIVERSITY_NAME_1_COL_NAME, headerCache);
      const facultyCol = findColumn(sheet, CONSTANTS.FACULTY_1_COL_NAME, headerCache);
      const departmentCol = findColumn(sheet, CONSTANTS.DEPARTMENT_1_COL_NAME, headerCache);
  
      // Ensure essential columns were found
      if ([receptionNumberCol, studentEmailCol, classNameCol, studentNumCol, studentNameCol, uniNameCol, facultyCol, departmentCol].includes(-1)) {
          throw new Error("One or more essential columns not found. Cannot process submission.");
      }
  
      const addedRow = range.getRow();
      const rowData = sheet.getRange(addedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
      // --- Assign Reception Number ---
      let nextReceptionNumber = 1;
      // Use Lock Service to prevent race conditions when getting the next number
      const lock = LockService.getScriptLock();
      lock.waitLock(15000); // Wait up to 15 seconds for lock
  
      try {
          const receptionNumberRange = sheet.getRange(2, receptionNumberCol, Math.max(sheet.getLastRow() - 1, 1)); // Range from row 2 down
          const receptionNumbers = receptionNumberRange.getValues().flat().filter(n => typeof n === 'number' && n > 0); // Filter only valid numbers
          if (receptionNumbers.length > 0) {
              nextReceptionNumber = Math.max(...receptionNumbers) + 1;
          }
          // Write the number to the newly added row
          sheet.getRange(addedRow, receptionNumberCol).setValue(nextReceptionNumber);
          SpreadsheetApp.flush(); // Ensure the number is written before sending email
      } finally {
          lock.releaseLock();
      }
      // --- End Assign Reception Number ---
  
      // --- Send Confirmation Email to Student ---
      const studentEmail = rowData[studentEmailCol - 1];
      const studentName = rowData[studentNameCol - 1];
      const className = rowData[classNameCol - 1];
      const studentNumber = rowData[studentNumCol - 1];
      const uniName = rowData[uniNameCol - 1];
      const faculty = rowData[facultyCol - 1];
      const department = rowData[departmentCol - 1];
  
      if (studentEmail) {
        const subject = "【進路】調査書作成願 受付完了のお知らせ";
        const body = `${studentName} さん (${className} ${studentNumber}番)\n\n` +
                     "フォームへのご提出ありがとうございました。\n" +
                     "以下の内容で受付を完了しました。\n\n" +
                     `受付番号: ${nextReceptionNumber}\n\n` +
                     "--- 入力内容の控え (1件目) ---\n" +
                     `大学名: ${uniName || '(未入力)'}\n` +
                     `学部: ${faculty || '(未入力)'}\n` +
                     `学科: ${department || '(未入力)'}\n` +
                     "---------------------------\n\n" +
                     "今後の手続きについては、別途連絡をお待ちください。";
        MailApp.sendEmail(studentEmail, subject, body);
        Logger.log(`Confirmation email sent to ${studentEmail} for reception number ${nextReceptionNumber}`);
      } else {
        Logger.log(`Skipping student confirmation: Email address not provided for row ${addedRow}.`);
      }
       // --- End Send Confirmation Email ---
  
    } catch (error) {
      handleError(error, context);
    }
  }
  
  
  /**
   * Handles the 'on edit' event.
   * Triggers emails based on changes in confirmation columns.
   * @param {object} e The event object.
   */
  function handleEdit(e) {
    const context = "handleEdit";
    try {
      if (!e || !e.range) {
          Logger.log("handleEdit triggered without event object or range.");
          return; // Exit if no event data
      }
  
      const range = e.range;
      const sheet = range.getSheet();
      const editedRow = range.getRow();
      const editedCol = range.getColumn();
      const newValue = e.value; // The new value entered
  
      // Ignore edits in header row or if multiple cells are edited at once
      if (editedRow === 1 || range.getNumRows() > 1 || range.getNumColumns() > 1) {
        return;
      }
  
      // Ignore if cell was cleared (newValue is undefined or empty string)
      if (newValue === undefined || newValue === null || newValue === "") {
         return;
      }
  
      // Only run on the response sheet
      if (sheet.getName() !== CONSTANTS.RESPONSE_SHEET_NAME) {
        return;
      }
  
      const config = getConfig();
      if (!config) return; // Stop if config failed
  
      // Check headers - less critical here but ensures columns exist if script added late
      // checkAndAddHeaders(sheet); // Can be commented out if onOpen is reliable
  
      // Cache header positions for this edit event
      const headerCache = {};
      const teacherConfirmCol = findColumn(sheet, CONSTANTS.TEACHER_CONFIRM_COL_NAME, headerCache);
      const guidanceConfirmCol = findColumn(sheet, CONSTANTS.GUIDANCE_CONFIRM_COL_NAME, headerCache);
      const officeReceiptCol = findColumn(sheet, CONSTANTS.OFFICE_RECEIPT_COL_NAME, headerCache);
      const reportCreatedCol = findColumn(sheet, CONSTANTS.REPORT_CREATED_COL_NAME, headerCache);
      const receptionNumberCol = findColumn(sheet, CONSTANTS.RECEPTION_NUMBER_COL_NAME, headerCache);
      const studentEmailCol = findColumn(sheet, CONSTANTS.STUDENT_EMAIL_COL_NAME, headerCache);
      const classNameCol = findColumn(sheet, CONSTANTS.CLASS_COL_NAME, headerCache);
      const studentNumCol = findColumn(sheet, CONSTANTS.STUDENT_NUMBER_COL_NAME, headerCache);
      const studentNameCol = findColumn(sheet, CONSTANTS.STUDENT_NAME_COL_NAME, headerCache);
  
      // Get the entire row's current data *after* the edit
      const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
      // --- Check 1: Teacher Confirmation -> Notify Guidance Dept ---
      if (editedCol === teacherConfirmCol && teacherConfirmCol !== -1 && config.guidanceEmail) {
        const teacherConfirmationValue = rowData[teacherConfirmCol - 1]; // Get current value
        const className = rowData[classNameCol - 1];
        const studentNumber = rowData[studentNumCol - 1];
        const studentName = rowData[studentNameCol - 1];
        const receptionNumber = rowData[receptionNumberCol - 1];
  
        if (teacherConfirmationValue) { // Ensure there's actually text entered
          const subject = `【要確認】担任確認完了:${className}_${studentNumber}_${studentName} (受付番号: ${receptionNumber})`;
          const body = `担任の確認が完了しました。\n\n` +
                       `クラス: ${className}\n` +
                       `出席番号: ${studentNumber}\n` +
                       `氏名: ${studentName}\n` +
                       `受付番号: ${receptionNumber}\n\n` +
                       `担任確認内容: ${teacherConfirmationValue}\n\n` +
                       `スプレッドシートをご確認ください。`;
          MailApp.sendEmail(config.guidanceEmail, subject, body);
          Logger.log(`Teacher confirmation email sent to Guidance Dept (${config.guidanceEmail}) for row ${editedRow}`);
        }
      }
  
      // --- Check 2: Guidance AND Office Confirmation -> Notify Teacher ---
      // This check runs if *either* guidance or office column was the one just edited
      if ((editedCol === guidanceConfirmCol || editedCol === officeReceiptCol) && guidanceConfirmCol !== -1 && officeReceiptCol !== -1) {
          const guidanceConfirmationValue = rowData[guidanceConfirmCol - 1];
          const officeReceiptValue = rowData[officeReceiptCol - 1];
          const className = rowData[classNameCol - 1];
          const receptionNumber = rowData[receptionNumberCol - 1];
  
          // Check if BOTH columns now have values
          if (guidanceConfirmationValue && officeReceiptValue) {
              const teacherEmail = config.teacherEmails[className];
              if (teacherEmail) {
                  const subject = `【進路部・事務室 受領連絡】(受付番号: ${receptionNumber})`;
                  const body = `${className}担任様\n\n`+
                               `受付番号 ${receptionNumber} の調査書作成願について、進路部および事務室での確認・受領が完了しました。\n\n` +
                               `進路部確認内容: ${guidanceConfirmationValue}\n` +
                               `事務室受領内容: ${officeReceiptValue}\n\n`+
                               `次のステップ（調査書作成）に進んでください。`;
                  MailApp.sendEmail(teacherEmail, subject, body);
                  Logger.log(`Guidance & Office confirmation email sent to Teacher (${teacherEmail}) for row ${editedRow}`);
              } else {
                  Logger.log(`Skipping Teacher notification for row ${editedRow}: Teacher email for class "${className}" not found in settings.`);
              }
          }
      }
  
      // --- Check 3: ALL Confirmations Done -> Notify Student ---
      // This check runs after any relevant edit to see if completion criteria are met
      if (teacherConfirmCol !== -1 && guidanceConfirmCol !== -1 && officeReceiptCol !== -1 && reportCreatedCol !== -1) {
          const teacherConfirmationValue = rowData[teacherConfirmCol - 1];
          const guidanceConfirmationValue = rowData[guidanceConfirmCol - 1];
          const officeReceiptValue = rowData[officeReceiptCol - 1];
          const reportCreatedValue = rowData[reportCreatedCol - 1]; // This is the final check
  
          // Check if ALL FOUR columns have values
          if (teacherConfirmationValue && guidanceConfirmationValue && officeReceiptValue && reportCreatedValue) {
              const studentEmail = rowData[studentEmailCol - 1];
              const studentName = rowData[studentNameCol - 1];
              const className = rowData[classNameCol - 1];
              const studentNumber = rowData[studentNumCol - 1];
              const receptionNumber = rowData[receptionNumberCol - 1];
  
              if (studentEmail) {
                  const subject = "【進路】調査書 準備完了のお知らせ";
                  const body = `${studentName} さん (${className} ${studentNumber}番)\n\n` +
                               `受付番号 ${receptionNumber} の調査書が作成されました。\n\n` +
                               `担任の先生から受け取ってください。`;
                  MailApp.sendEmail(studentEmail, subject, body);
                  Logger.log(`Report completion email sent to Student (${studentEmail}) for row ${editedRow}`);
                  // Optional: Mark row as complete, e.g., change background color
                  // range.getRow().setBackground("#d9ead3"); // Light green
              } else {
                   Logger.log(`Skipping student completion notification for row ${editedRow}: Student email not found.`);
              }
          }
      }
  
    } catch (error) {
      handleError(error, context);
    }
  }
  
  // Example function to allow manual header check (can be linked to menu in onOpen)
  // function manualHeaderCheck() {
  //   try {
  //      const ss = SpreadsheetApp.getActiveSpreadsheet();
  //      const responseSheet = ss.getSheetByName(CONSTANTS.RESPONSE_SHEET_NAME);
  //      if (!responseSheet) {
  //          SpreadsheetApp.getUi().alert(`Sheet "${CONSTANTS.RESPONSE_SHEET_NAME}" not found.`);
  //          return;
  //      }
  //      checkAndAddHeaders(responseSheet);
  //      SpreadsheetApp.getUi().alert("Header check complete. Missing headers (if any) were added.");
  //   } catch (error) {
  //      handleError(error, "manualHeaderCheck");
  //      SpreadsheetApp.getUi().alert(`An error occurred during header check: ${error.message}`);
  //   }
  // }