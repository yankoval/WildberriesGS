function copyDataFromAnotherSheet() {
  var sourceSpreadsheetId = '16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E'; // ID исходной таблицы
  var sourceSheetName = 'Короба'; // Имя листа в исходной таблице
  
  var targetSpreadsheet = SpreadsheetApp.openById('1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o');//SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getSheetByName('Лист1'); // Имя листа, куда будут скопированы данные
  
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  var data = sourceSheet.getDataRange().getValues();
  
  // Очистить целевой лист
  targetSheet.clearContents();
  
  // Вставить данные
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  // Создание архива
  createArchive(data);
}

function createArchive(data) {
  var archiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var archiveSheetName = 'Архив_' + Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd_HH-mm-ss");
  var archiveSheet = archiveSpreadsheet.insertSheet(archiveSheetName);
  
  archiveSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function checkForPositiveValues() {
  // Конфигурация
  const spreadsheetId = '16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E';  // ID основной таблицы (где находится лист "Короба")
  const listSpreadsheetId = '1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o';  // ID таблицы, содержащей диапазон "СписокРЦ"
  const sheetName = "Короба";
  const columnToCheck = 2;  // Колонка 2 (индексация с 1)
  const targetColumn = 1;  // Колонка 1, из которой берем текст
  const listRangeName = "СписокРЦ";  // Имя именованного диапазона

  try {
    // Логирование начала операции
    Logger.log("=== Начало проверки таблицы ===");
    Logger.log(`ID основной таблицы: ${spreadsheetId}`);
    Logger.log(`ID таблицы с "СписокРЦ": ${listSpreadsheetId}`);
    Logger.log(`Проверяется лист: ${sheetName}`);
    Logger.log(`Проверяемая колонка: ${columnToCheck} (для значений >0)`);
    Logger.log(`Источная колонка: ${targetColumn} (для вывода текста)`);
    Logger.log(`Именованный диапазон для проверки: ${listRangeName}`);

    // Открываем главную таблицу для проверки листа "Короба"
    const mainSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Основная таблица успешно открыта");

    const sheet = mainSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Лист "${sheetName}" не найден`);
    }
    Logger.log(`Лист "${sheetName}" найден`);
Utilities.sleep(200);
    // Получаем все данные из листа
    const data = sheet.getDataRange().getValues();
    const rowCount = data.length;
    const colCount = data[0].length;
    
    Logger.log(`Всего строк: ${rowCount}, столбцов: ${colCount}`);
    Logger.log("Заголовки: " + data[0].slice(0, 5).join(", ") + (colCount > 5 ? "..." : ""));

    // Открываем таблицу с именованным диапазоном "СписокРЦ"
    const listSpreadsheet = SpreadsheetApp.openById(listSpreadsheetId);
    Logger.log("Таблица с " + listRangeName + " успешно открыта");

    // Получаем значения из именованного диапазона "СписокРЦ"
    const listRange = listSpreadsheet.getRangeByName(listRangeName);
    if (!listRange) {
      Logger.log(`Именованный диапазон "${listRangeName}" не найден`);
      throw new Error(`Именованный диапазон "${listRangeName}" не найден`);
    }
    
    // Проверяем, что диапазон не пустой
    const listValues = listRange.getValues();
    if (listValues.length === 0) {
      Logger.log(`Именованный диапазон "${listRangeName}" пуст`);
      throw new Error(`Именованный диапазон "${listRangeName}" пуст`);
    }
    
    // Функция для плоского преобразования двумерного массива в одномерный
    function flattenArray(arr) {
      return arr.flat().filter(item => item !== "");
    }
    
    const flatListValues = flattenArray(listValues);
    Logger.log(`Найдено ${flatListValues.length} значений в диапазоне "${listRangeName}"`);
    
    // Обходим строки (пропускаем заголовки - первую строку)
    const foundValues = [];
    const checkedRows = [];
    
    for (let i = 1; i < data.length; i++) {
      checkedRows.push(i);
      const cellValue = parseFloat(data[i][columnToCheck-1]); // -1 из-за 0-индексации
      
      if (!isNaN(cellValue) && cellValue > 0) {
        Logger.log(`Строка ${i}: значение в колонке ${columnToCheck} = ${cellValue}`);
        
        const targetCell = data[i][targetColumn-1]; // Текст из нужной колонки
        
        // Проверяем, содержится ли значение в списке РЦ
        const isMatching = flatListValues.some(item => item.toString().toLowerCase() === targetCell.toString().toLowerCase());
        
        if (isMatching) {
          Logger.log(`Строка ${i}: значение "${targetCell}" совпадает с диапазоном "${listRangeName}"`);
          foundValues.push(targetCell);
        } else {
          Logger.log(`Строка ${i}: значение "${targetCell}" НЕ совпадает с диапазоном "${listRangeName}"`);
        }
      }
    }

    Logger.log(`Найдено совпадений с "${listRangeName}": ${foundValues.length}`);
    
    if (foundValues.length > 0) {
      // Форматируем сообщение для вывода в алерте
      const message = `Найдено совпадений с "${listRangeName}":\n${foundValues.join('\n')}`;
      sendSimpleEmail(message);
      SpreadsheetApp.getActiveSpreadsheet().toast(message, "Waildberries!!!");
      Browser.msgBox(message, Browser.Buttons.OK_CANCEL);
      Logger.log("Алерт с результатами отправлен");
    } else {
      let errorMessage = "Значений больше 0 в указанной колонке не найдено.";
      if (checkedRows.length > 0) {
        errorMessage += `\n\nПроверено ${checkedRows.length} строк.`;
      }
      Browser.msgBox(errorMessage, Browser.Buttons.OK_CANCEL);
      Logger.log("Алерт с отсутствием результатов отправлен");
    }
    
  } catch (error) {
    Logger.log("=== Произошла ошибка ===");
    Logger.log(`Тип ошибки: ${error.name || 'Unknown'} `);
    Logger.log(`Описание ошибки: ${error.message}`);
    Browser.msgBox(`Ошибка: ${error.message}`, Browser.Buttons.OK_CANCEL);
    Logger.log("Алерт об ошибке отправлен");
  }
}

function sendSimpleEmail(txtRC) {
  const recipient = "yankoval@gmail.com";
  const subject = "Wildberries РЦ:"+txtRC;
  const body = "Привет! Это письмо отправлено через Google Apps Script.";

  MailApp.sendEmail(recipient, subject, body);
}


function createTimeDrivenTrigger() {
  // Удаляем все существующие триггеры, чтобы избежать дублирования
  deleteExistingTriggers();

  // Создаем новый триггер, который будет запускаться каждые 10 минут
  ScriptApp.newTrigger('checkWildberries')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('Триггер создан для запуска каждые 10 минут.');
}

function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}