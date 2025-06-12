// Configuration Constants
const SOURCE_SPREADSHEET_ID = '16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E'; // ID of the source spreadsheet
const SOURCE_SHEET_NAME = 'Короба'; // Sheet name in the source spreadsheet, used by copyDataFromAnotherSheet and checkWildberries
const TARGET_SPREADSHEET_ID = '1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o'; // Spreadsheet ID for the target sheet in copyDataFromAnotherSheet and for the list range in checkWildberries
const TARGET_SHEET_NAME = 'Лист1'; // Sheet name in the target spreadsheet
const ARCHIVE_SHEET_PREFIX = 'Архив_'; // Prefix for archive sheet names
const LIST_RANGE_SPREADSHEET_ID = TARGET_SPREADSHEET_ID; // Named range 'СписокРЦ' is located in the target spreadsheet.
const LIST_RANGE_NAME = 'СписокРЦ'; // Name of the named range to check against
const RECIPIENT_EMAIL = 'yankoval@gmail.com'; // Email address for notifications

// Constants for checkWildberries function
const CW_COLUMN_TO_CHECK = 2; // Column number (1-indexed) in SOURCE_SHEET_NAME to check for positive values
const CW_TARGET_COLUMN = 1;   // Column number (1-indexed) in SOURCE_SHEET_NAME from which to get text if the checked column is positive

/**
 * Copies data from a specified source sheet to a target sheet and then archives the copied data.
 */
function copyDataFromAnotherSheet() {
  try {
    // var sourceSpreadsheetId = '16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E'; // Source Spreadsheet ID // Replaced by global constant SOURCE_SPREADSHEET_ID
    // var sourceSheetName = 'Короба'; // Sheet name in source spreadsheet // Replaced by global constant SOURCE_SHEET_NAME

    var targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); // Open the target spreadsheet
    var targetSheet = targetSpreadsheet.getSheetByName(TARGET_SHEET_NAME); // Get the target sheet
  
    var sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID); // Open the source spreadsheet
    var sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME); // Get the source sheet
  
    var data = sourceSheet.getDataRange().getValues(); // Get all data from the source sheet
  
    // Clear the target sheet before pasting new data
    targetSheet.clearContents();
  
    // Paste data into the target sheet
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
    // Create an archive of the copied data
    createArchive(targetSpreadsheet, data); // Pass the target spreadsheet object for archiving
  } catch (e) {
    Logger.log('Error in copyDataFromAnotherSheet: ' + e.toString());
  }
}

/**
 * Creates a new sheet in the given spreadsheet and copies the provided data into it.
 * The new sheet is named with the ARCHIVE_SHEET_PREFIX and the current timestamp.
 * @param {Spreadsheet} spreadsheet The spreadsheet where the archive sheet will be created.
 * @param {Array<Array>} data The data to be archived.
 */
function createArchive(spreadsheet, data) {
  try {
    var archiveSpreadsheet = spreadsheet; // Use the passed spreadsheet object
    // Create a unique name for the archive sheet using a prefix and timestamp
    var archiveSheetName = ARCHIVE_SHEET_PREFIX + Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd_HH-mm-ss");
    var archiveSheet = archiveSpreadsheet.insertSheet(archiveSheetName); // Insert the new sheet

    // Paste data into the archive sheet
    archiveSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  } catch (e) {
    Logger.log('Error in createArchive: ' + e.toString());
  }
}

/**
 * Checks a specified sheet for rows where a numeric value in one column is positive,
 * and a corresponding text value in another column matches entries in a named range.
 * Sends an email notification if matches are found.
 */
function checkWildberries() {
  // Configuration (local constants were here, now replaced by global constants at the top of the script)
  // const spreadsheetId = '16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E';  // Main spreadsheet ID (where "Короба" sheet is) // Replaced by global constant SOURCE_SPREADSHEET_ID
  // const listSpreadsheetId = '1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o';  // Spreadsheet ID containing "СписокРЦ" named range // Replaced by global constant LIST_RANGE_SPREADSHEET_ID
  // const sheetName = "Короба"; // Sheet name to check // Replaced by global constant SOURCE_SHEET_NAME
  // const columnToCheck = 2;  // Column 2 (1-indexed) to check for >0 // Replaced by global constant CW_COLUMN_TO_CHECK
  // const targetColumn = 1;  // Column 1 (1-indexed) to get text from // Replaced by global constant CW_TARGET_COLUMN
  // const listRangeName = "СписокРЦ";  // Name of the named range // Replaced by global constant LIST_RANGE_NAME

  try {
    // Log the start of the checking operation
    Logger.log("=== Начало проверки таблицы ==="); // "=== Starting table check ==="
    Logger.log(`ID основной таблицы: ${SOURCE_SPREADSHEET_ID}`); // Log main spreadsheet ID
    Logger.log(`ID таблицы с "СписокРЦ": ${LIST_RANGE_SPREADSHEET_ID}`); // Log spreadsheet ID with "СписокРЦ"
    Logger.log(`Проверяется лист: ${SOURCE_SHEET_NAME}`); // Log sheet being checked
    Logger.log(`Проверяемая колонка: ${CW_COLUMN_TO_CHECK} (для значений >0)`); // Log column being checked for positive values
    Logger.log(`Источная колонка: ${CW_TARGET_COLUMN} (для вывода текста)`); // Log source column for text output
    Logger.log(`Именованный диапазон для проверки: ${LIST_RANGE_NAME}`); // Log named range for checking

    // Open the main spreadsheet to check the sheet
    const mainSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    Logger.log("Основная таблица успешно открыта"); // "Main spreadsheet opened successfully"

    const sheet = mainSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Лист "${SOURCE_SHEET_NAME}" не найден`); // `Sheet "${SOURCE_SHEET_NAME}" not found`
    }
    Logger.log(`Лист "${SOURCE_SHEET_NAME}" найден`); // `Sheet "${SOURCE_SHEET_NAME}" found`

    Utilities.sleep(200); // Brief pause; purpose might need review for necessity.

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    const rowCount = data.length;
    const colCount = data[0].length;
    
    Logger.log(`Всего строк: ${rowCount}, столбцов: ${colCount}`); // "Total rows: ${rowCount}, columns: ${colCount}"
    Logger.log("Заголовки: " + data[0].slice(0, 5).join(", ") + (colCount > 5 ? "..." : "")); // "Headers: ..." (first 5 columns)

    // Open the spreadsheet containing the named range "СписокРЦ"
    const listSpreadsheet = SpreadsheetApp.openById(LIST_RANGE_SPREADSHEET_ID);
    Logger.log("Таблица с " + LIST_RANGE_NAME + " успешно открыта"); // `Spreadsheet with ${LIST_RANGE_NAME} opened successfully`

    // Get values from the named range "СписокРЦ"
    const listRange = listSpreadsheet.getRangeByName(LIST_RANGE_NAME);
    if (!listRange) {
      Logger.log(`Именованный диапазон "${LIST_RANGE_NAME}" не найден`); // `Named range "${LIST_RANGE_NAME}" not found`
      throw new Error(`Именованный диапазон "${LIST_RANGE_NAME}" не найден`); // `Named range "${LIST_RANGE_NAME}" not found`
    }
    
    // Check if the named range is empty
    const listValues = listRange.getValues();
    if (listValues.length === 0) {
      Logger.log(`Именованный диапазон "${LIST_RANGE_NAME}" пуст`); // `Named range "${LIST_RANGE_NAME}" is empty`
      throw new Error(`Именованный диапазон "${LIST_RANGE_NAME}" пуст`); // `Named range "${LIST_RANGE_NAME}" is empty`
    }
    
    // Helper function to flatten a 2D array into a 1D array and remove empty strings
    function flattenArray(arr) {
      return arr.flat().filter(item => item !== "");
    }
    
    const flatListValues = flattenArray(listValues);
    Logger.log(`Найдено ${flatListValues.length} значений в диапазоне "${LIST_RANGE_NAME}"`); // `Found ${flatListValues.length} values in range "${LIST_RANGE_NAME}"`
    
    // Iterate through rows (skip header row, hence starting from i = 1)
    const foundValues = [];
    const checkedRows = [];
    
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header row
      checkedRows.push(i);
      // Get cell value from the column to check; CW_COLUMN_TO_CHECK is 1-indexed, array access is 0-indexed.
      const cellValue = parseFloat(data[i][CW_COLUMN_TO_CHECK-1]);
      
      if (!isNaN(cellValue) && cellValue > 0) {
        // Log value found in the checked column for the current row
        Logger.log(`Строка ${i}: значение в колонке ${CW_COLUMN_TO_CHECK} = ${cellValue}`); // `Row ${i}: value in column ${CW_COLUMN_TO_CHECK} = ${cellValue}`
        
        // Get text from the target column for the current row; CW_TARGET_COLUMN is 1-indexed.
        const targetCell = data[i][CW_TARGET_COLUMN-1];
        
        // Check if the target cell's value is present in the list from the named range (case-insensitive)
        const isMatching = flatListValues.some(item => item.toString().toLowerCase() === targetCell.toString().toLowerCase());
        
        if (isMatching) {
          // Log if a match is found
          Logger.log(`Строка ${i}: значение "${targetCell}" совпадает с диапазоном "${LIST_RANGE_NAME}"`); // `Row ${i}: value "${targetCell}" matches range "${LIST_RANGE_NAME}"`
          foundValues.push(targetCell);
        } else {
          // Log if no match is found
          Logger.log(`Строка ${i}: значение "${targetCell}" НЕ совпадает с диапазоном "${LIST_RANGE_NAME}"`); // `Row ${i}: value "${targetCell}" DOES NOT match range "${LIST_RANGE_NAME}"`
        }
      }
    }

    Logger.log(`Найдено совпадений с "${LIST_RANGE_NAME}": ${foundValues.length}`); // `Found matches with "${LIST_RANGE_NAME}": ${foundValues.length}`
    
    if (foundValues.length > 0) {
      // Format the message for notification (email)
      const message = `Найдено совпадений с "${LIST_RANGE_NAME}":\n${foundValues.join('\n')}`; // `Found matches with "${LIST_RANGE_NAME}":\n${foundValues.join('\n')}`
      sendSimpleEmail(message);
      // SpreadsheetApp.getActiveSpreadsheet().toast(message, "Waildberries!!!"); // UI interaction commented out
      // Browser.msgBox(message, Browser.Buttons.OK_CANCEL); // UI interaction commented out
      Logger.log("Алерт с результатами отправлен"); // "Notification about results sent" (refers to Logger/email)
    } else {
      let errorMessage = "Значений больше 0 в указанной колонке не найдено."; // "No values greater than 0 found in the specified column."
      if (checkedRows.length > 0) {
        errorMessage += `\n\nПроверено ${checkedRows.length} строк.`; // `\n\nChecked ${checkedRows.length} rows.`
      }
      // Browser.msgBox(errorMessage, Browser.Buttons.OK_CANCEL); // UI interaction commented out
      Logger.log("Алерт с отсутствием результатов отправлен"); // "Notification about absence of results sent"
    }
    
  } catch (error) {
    Logger.log("=== Произошла ошибка ==="); // "=== An error occurred ==="
    Logger.log(`Тип ошибки: ${error.name || 'Unknown'} `); // `Error type: ${error.name || 'Unknown'} `
    Logger.log(`Описание ошибки: ${error.message}`); // `Error description: ${error.message}`
    // Browser.msgBox(`Ошибка: ${error.message}`, Browser.Buttons.OK_CANCEL); // UI interaction commented out
    Logger.log("Алерт об ошибке отправлен"); // "Error notification sent"
  }
}

/**
 * Sends a simple email with the provided text.
 * @param {string} txtRC The text content to include in the email's subject and body.
 */
function sendSimpleEmail(txtRC) {
  const recipient = RECIPIENT_EMAIL;
  const subject = "Wildberries РЦ:" + txtRC; // Subject includes the found RC details
  // Construct the email body with the matches and a standard message.
  const body = "Обнаружены следующие совпадения с указанным списком РЦ Wildberries:\n\n" + txtRC + "\n\nЭто автоматическое уведомление от Google Apps Script.";
  // "The following matches were found with the specified Wildberries RC list:\n\n" + txtRC + "\n\nThis is an automatic notification from Google Apps Script."

  MailApp.sendEmail(recipient, subject, body);
}

/**
 * Creates a time-driven trigger for the 'checkWildberries' function to run every 10 minutes.
 * Deletes any existing triggers for the project first to avoid duplicates.
 */
function createTimeDrivenTrigger() {
  // Delete existing triggers for 'checkWildberries' to avoid duplication before creating a new one.
  deleteSpecificTriggers('checkWildberries');

  // Create a new time-driven trigger to run 'checkWildberries' every 10 minutes.
  ScriptApp.newTrigger('checkWildberries')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('Триггер создан для запуска каждые 10 минут.'); // "Trigger created to run every 10 minutes."
}

/**
 * Deletes all project triggers that call the specified handler function.
 * @param {string} handlerFunctionName The name of the function whose triggers should be deleted.
 */
function deleteSpecificTriggers(handlerFunctionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerFunctionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log(`Deleted trigger for function: ${handlerFunctionName} (ID: ${triggers[i].getUniqueId()})`);
    }
  }
}
