/**
 * Google Apps Script для автоматической настройки Листа 2 в вашей Google Таблице с ответами.
 * Выполняет пункты 18 и 19 методических указаний ЛР4 (СТАУ):
 * 1. Создает Лист 2 («Проверка и оценивание»).
 * 2. Копирует ФИО участников с Листа 1.
 * 3. Расставляет формулы IF для проверки каждого из 10 вопросов.
 * 4. Считает сумму баллов =SUM(...) и выставляет оценку («Отлично», «Хорошо», «Удовл», «Неуд»).
 * 5. Форматирует шапку и автоподбор ширины колонок.
 *
 * КАК ЗАПУСТИТЬ ЗА 10 СЕКУНД:
 * 1. В вашей открытой Google Таблице нажмите в верхнем меню: «Расширения» -> «Apps Script».
 * 2. Сотрите текст и вставьте код из этого файла.
 * 3. Нажмите кнопку «Сохранить» (дискета), затем «Выполнить» (Run).
 * 4. Вернитесь во вкладку таблицы — Лист 2 уже полностью готов и заполнен!
 */

function setupGradingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var respSheet = sheets[0]; // Лист 1 с ответами
  var sheetName = respSheet.getName();
  
  var data = respSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("В таблице пока нет строк с ответами!");
    return;
  }
  
  var headers = data[0];
  var fioCol = -1;
  var qCols = [];
  
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c].toString().toLowerCase();
    if (h.indexOf("фио") !== -1) fioCol = c + 1;
    if (h.indexOf("1. ") !== -1) qCols[0] = c + 1;
    if (h.indexOf("2. ") !== -1) qCols[1] = c + 1;
    if (h.indexOf("3. ") !== -1) qCols[2] = c + 1;
    if (h.indexOf("4. ") !== -1) qCols[3] = c + 1;
    if (h.indexOf("5. ") !== -1) qCols[4] = c + 1;
    if (h.indexOf("6. ") !== -1) qCols[5] = c + 1;
    if (h.indexOf("7. ") !== -1) qCols[6] = c + 1;
    if (h.indexOf("8. ") !== -1) qCols[7] = c + 1;
    if (h.indexOf("9. ") !== -1) qCols[8] = c + 1;
    if (h.indexOf("10. ") !== -1) qCols[9] = c + 1;
  }
  
  // Если не нашли по подстроке, берем стандартные позиции
  if (fioCol === -1) fioCol = 2; // Колонка B
  
  function getColLetter(col) {
    var str = "";
    while (col > 0) {
      var mod = (col - 1) % 26;
      str = String.fromCharCode(65 + mod) + str;
      col = Math.floor((col - mod) / 26);
    }
    return str;
  }
  
  var fioLetter = getColLetter(fioCol);
  
  // Создаем или очищаем Лист 2
  var evalSheetName = "Проверка и оценивание";
  var evalSheet = ss.getSheetByName(evalSheetName);
  if (!evalSheet) {
    evalSheet = ss.insertSheet(evalSheetName);
  } else {
    evalSheet.clear();
  }
  
  // Заголовки Листа 2
  var evalHeaders = [
    "ФИО тестируемого",
    "Q1 (Джоуль)", "Q2 (1 зак. Ньютона)", "Q3 (Закон Ома)", "Q4 (Изохорный)",
    "Q5 (Оптоволокно)", "Q6 (Импульс)", "Q7 (Сила Архимеда)", "Q8 (Нейтрон)",
    "Q9 (Индукция)", "Q10 (Скорость света)",
    "ИТОГО БАЛЛОВ", "ОЦЕНКА"
  ];
  evalSheet.appendRow(evalHeaders);
  
  // Ключи правильных ответов
  var correctAnswers = [
    "Энергия и работа",
    "Равнодействующая всех сил равна нулю",
    "I = U / R",
    "Изохорный",
    "Полное внутреннее отражение",
    "импульс", // Открытый
    "Весу вытесненной жидкости",
    "Нейтральный (ноль)",
    "фарадей", // Открытый
    "300 000 км/с"
  ];
  
  var numRows = data.length; // Количество строк
  
  for (var r = 2; r <= numRows; r++) {
    var rowFormulas = [];
    
    // ФИО (='Лист 1'!C2)
    rowFormulas.push("='" + sheetName + "'!" + fioLetter + r);
    
    // Формулы IF для Q1 - Q10
    for (var q = 0; q < 10; q++) {
      var qColLetter = qCols[q] ? getColLetter(qCols[q]) : getColLetter(fioCol + 3 + q);
      var key = correctAnswers[q];
      
      if (q === 5 || q === 8) {
        // Открытые вопросы (trim и lower)
        rowFormulas.push("=IF(LOWER(TRIM('" + sheetName + "'!" + qColLetter + r + "))=\"" + key + "\", 1, 0)");
      } else {
        // Выбор из списка
        rowFormulas.push("=IF('" + sheetName + "'!" + qColLetter + r + "=\"" + key + "\", 1, 0)");
      }
    }
    
    // Итого баллов: =SUM(B_r:K_r)
    rowFormulas.push("=SUM(B" + r + ":K" + r + ")");
    
    // Итоговая оценка по шкале п. 19
    rowFormulas.push("=IF(L" + r + ">=10, \"Отлично\", IF(L" + r + ">=8, \"Хорошо\", IF(L" + r + ">=6, \"Удовлетворительно\", \"Неудовлетворительно\")))");
    
    evalSheet.appendRow(rowFormulas);
  }
  
  // Добавляем строку среднего балла
  var statRow = ["СРЕДНИЙ БАЛЛ ПО ГРУППЕ:", "", "", "", "", "", "", "", "", "", "", "=AVERAGE(L2:L" + numRows + ")", ""];
  evalSheet.appendRow(statRow);
  
  // Красивое оформление таблицы
  var headerRange = evalSheet.getRange(1, 1, 1, evalHeaders.length);
  headerRange.setBackground("#1F497D");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  
  var allDataRange = evalSheet.getRange(1, 1, numRows + 1, evalHeaders.length);
  allDataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  allDataRange.setFontFamily("Arial");
  
  // Выравнивание чисел и оценок по центру
  var scoresRange = evalSheet.getRange(2, 2, numRows, evalHeaders.length - 1);
  scoresRange.setHorizontalAlignment("center");
  
  evalSheet.setFrozenRows(1);
  for (var i = 1; i <= evalHeaders.length; i++) {
    evalSheet.autoResizeColumn(i);
  }
  
  Logger.log("Лист 2 («Проверка и оценивание») успешно настроен!");
}
