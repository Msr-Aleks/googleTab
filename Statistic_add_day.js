




/**
 * =========================================================================
 * ЧАСТЬ 1. ЕДИНЫЙ АВТОМАТИЧЕСКИЙ ТРИГГЕР (onEdit)
 * =========================================================================
 * Управляет скрытием столбцов на "Статистика" и скроллингом на "Таблица_посещений".
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const editRow = range.getRow();
  const editCol = range.getColumn();
  const isChecked = e.value === "TRUE";

  // -----------------------------------------------------------------------
  // БЛОК Б: Ваша исходная логика чекбоксов для листа "Статистика"
  // -----------------------------------------------------------------------
  if (sheetName !== "Статистика") return;
  
  const lastCol = sheet.getLastColumn();
  
  // Ищем маркер "St" в 6-й строке
  const headersRow6 = sheet.getRange(6, 1, 1, lastCol).getValues()[0];
  const stColGlobal = headersRow6.findIndex(h => h.toString().trim().toLowerCase() === "st") + 1;
  if (stColGlobal === 0) return;
  
  const dynamicColsCount = lastCol - stColGlobal;
  if (dynamicColsCount <= 0) return;
  
  const dynamicHeaders = sheet.getRange(5, stColGlobal + 1, 1, dynamicColsCount).getValues()[0];

  // --- ВАРИАНТ А: Глобальные чекбоксы выборки данных (Ячейки F2, G2, K2, L2) ---
  if (editRow === 2 && (editCol === 7 || editCol === 8 || editCol === 12 || editCol === 13)) {
    dynamicHeaders.forEach((headerText, idx) => {
      const currentGlobalColNum = stColGlobal + 1 + idx;
      let isMatch = false;
      if (editCol === 7 && headerText.startsWith("Кил")) isMatch = true;
      if (editCol === 8 && headerText.startsWith("См")) isMatch = true;
      if (editCol === 12 && headerText.startsWith("Коэф")) isMatch = true;
      if (editCol === 13 && headerText.startsWith("Проц")) isMatch = true;

      if (isMatch) {
        if (isChecked) sheet.showColumns(currentGlobalColNum); else sheet.hideColumns(currentGlobalColNum);
      }
    });
    return;
  }

  // Если кликнули левее динамической зоны дат, остальные условия игнорируем
  if (editCol <= stColGlobal) return;

  // Проверяем, что кликнули именно по столбцу Коэффициента (строка 5 начинается на "Коэф")
  const currentHeader = sheet.getRange(5, editCol).getValue().toString();
  if (!currentHeader.startsWith("Коэф")) return;

  const dateParts = currentHeader.split(" ");
  if (dateParts.length < 2) return;
  const targetDate = dateParts[1].trim(); // Извлекли чистую дату, например "19.05"

  // --- ВАРИАНТ Б: Управление локальным переключателем дня в СТРОКЕ 2 ---
  if (editRow === 2) {
    dynamicHeaders.forEach((headerText, idx) => {
      const currentGlobalColNum = stColGlobal + 1 + idx;
      
      // Реагирует на Проц, Кил и См текущей даты
      if (headerText.includes(targetDate) && (headerText.startsWith("Проц") || headerText.startsWith("Кил") || headerText.startsWith("См"))) {
        if (isChecked) {
          sheet.showColumns(currentGlobalColNum); 
        } else {
          sheet.hideColumns(currentGlobalColNum); 
        }
      }
    });
  }
}



/**
 * =========================================================================
 * ЧАСТЬ 2. ОСНОВНАЯ ФУНКЦИЯ ГЕНЕРАЦИИ ТАБЛИЦЫ СТАТИСТИКИ С НУЛЯ
 * =========================================================================
 * Очищает всю зону справа от "St" и строит полную таблицу заново.
 */
// ==========================================
// 1. КНОПКА «СФОРМИРОВАТЬ» (ПОЛНАЯ ПЕРЕЗАПИСЬ)
// ==========================================
function generateDynamicColumns(manualDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Статистика");
  if (!sheet) return;

  const tz = ss.getSpreadsheetTimeZone();
  const startDateVal = sheet.getRange("D2").getValue();
  const endDateVal = manualDate ? new Date(manualDate) : sheet.getRange("E2").getValue();
  
  const startDate = new Date(startDateVal);
  const endDate = new Date(endDateVal);
  if (isNaN(startDate) || isNaN(endDate)) {
    SpreadsheetApp.getUi().alert("Укажите корректные даты в ячейках D2 и E2!");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const dateFmt = "dd.MM.yyyy";
  const confirm = ui.alert('Подтверждение генерации', 
    `Сформировать статистику за период:\n"${Utilities.formatDate(startDate, tz, dateFmt)}" — "${Utilities.formatDate(endDate, tz, dateFmt)}"?\n\nВнимание: Старые динамические данные будут перезаписаны.`, 
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const headerRowIndex = 6; 
  let lastCol = sheet.getLastColumn(); 
  const headers = sheet.getRange(headerRowIndex, 1, 1, Math.max(lastCol, 100)).getValues()[0];
  const columnIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === "st");
  const stColumnNum = columnIndex + 1; 
  if (columnIndex === -1) {
    SpreadsheetApp.getUi().alert("Маркер 'St' в строке 6 не найден! Проверьте ячейку N6.");
    return;
  }

  // Сброс и полная очистка старых колонок
  if (lastCol > stColumnNum) {
    const totalColsToReset = sheet.getMaxColumns() - stColumnNum;
    sheet.showColumns(stColumnNum + 1, totalColsToReset);
    for (let depth = 1; depth <= 3; depth++) {
      try { sheet.getColumnGroup(stColumnNum + 1, depth).remove(); } catch(e) {}
    }
    sheet.deleteColumns(stColumnNum + 1, lastCol - stColumnNum);
  }

  // Запуск ядра в режиме полной перезаписи (insertAtCol = stColumnNum + 1)
  runDynamicGenerationCore(startDate, endDate, sheet, stColumnNum, stColumnNum + 1, true);
}

// ==========================================
// 2. КНОПКА «ДОБАВИТЬ ДЕНЬ» (ДОБОР В КОНЕЦ)
// ==========================================
function appendMissingDays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Статистика");
  if (!sheet) return;

  const tz = ss.getSpreadsheetTimeZone();
  const ui = SpreadsheetApp.getUi();

  const endDateVal = sheet.getRange("E2").getValue();
  const endDate = new Date(endDateVal);
  if (isNaN(endDate)) {
    return ui.alert("Укажите корректную конечную дату в ячейке E2!");
  }

  const headerRowIndex = 6;
  let lastCol = sheet.getLastColumn();
  const headersRow6 = sheet.getRange(headerRowIndex, 1, 1, Math.max(lastCol, 100)).getValues()[0];

  const columnIndex = headersRow6.findIndex(h => h.toString().trim().toLowerCase() === "st");
  const stColumnNum = columnIndex + 1;
  if (stColumnNum === 0) {
    return ui.alert("Маркер 'St' в строке 6 не найден!");
  }

  // Временно раскрываем, чтобы считать последнюю дату
  if (lastCol > stColumnNum) {
    sheet.showColumns(stColumnNum + 1, sheet.getMaxColumns() - stColumnNum);
  }
  
  lastCol = sheet.getLastColumn(); 
  let lastGeneratedDate = null;
  
  if (lastCol > stColumnNum) {
    const dynamicHeaders = sheet.getRange(5, stColumnNum + 1, 1, lastCol - stColumnNum).getValues()[0];
    for (let i = dynamicHeaders.length - 1; i >= 0; i--) {
      const hText = dynamicHeaders[i].toString();
      const parts = hText.split(" ");
      if (parts.length >= 2) {
        const rawDateStr = parts[1].trim();
        const dateParts = rawDateStr.split(".");
        const currentYear = new Date().getFullYear();
        lastGeneratedDate = new Date(currentYear, parseInt(dateParts[1], 10) - 1, parseInt(dateParts[0], 10));
        break;
      }
    }
  }

  let startBuildDate = new Date();
  if (lastGeneratedDate) {
    startBuildDate = new Date(lastGeneratedDate);
    startBuildDate.setDate(startBuildDate.getDate() + 1);
  } else {
    const startDateVal = sheet.getRange("D2").getValue();
    startBuildDate = new Date(startDateVal);
    if (isNaN(startBuildDate)) {
      return ui.alert("Таблица пуста. Задайте начальную дату в ячейке D2!");
    }
  }

  const compareStart = new Date(startBuildDate.getFullYear(), startBuildDate.getMonth(), startBuildDate.getDate());
  const compareEnd = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

  if (compareStart > compareEnd) {
    ss.toast("Все дни уже добавлены. Достройка не требуется.");
    if (typeof reapplyInitialHiding === "function") reapplyInitialHiding(sheet, stColumnNum);
    return;
  }

  const confirm = ui.alert('Добор статистики', 
    `Дописать недостающие дни в конец таблицы?\nПериод: "${Utilities.formatDate(compareStart, tz, "dd.MM")}" — "${Utilities.formatDate(compareEnd, tz, "dd.MM")}"`, 
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) {
    if (typeof reapplyInitialHiding === "function") reapplyInitialHiding(sheet, stColumnNum);
    return;
  }

  // ИСПРАВЛЕНИЕ 1: Если таблица пустая (lastCol равен stColumnNum), вставляем строго со следующего столбца.
  // Если на листе был пустой дефолтный хвост, сжимаем лист до stColumnNum, чтобы убрать лишние столбцы.
  if (lastCol <= stColumnNum) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols > stColumnNum) {
      sheet.deleteColumns(stColumnNum + 1, maxCols - stColumnNum);
    }
    lastCol = stColumnNum;
  }

  const targetInsertColumn = lastCol + 1;

  // Запуск ядра с точными координатами
  runDynamicGenerationCore(compareStart, compareEnd, sheet, stColumnNum, targetInsertColumn, false);

}



//-------------------------------------------------------------------
/**
 * Вспомогательное восстановление компактного вида
 */
function reapplyInitialHiding(sheet, stColGlobal) {
  const lastCol = sheet.getLastColumn();
  if (lastCol <= stColGlobal) return;
  
  // Читаем актуальное состояние глобальных чекбоксов выборки данных
  const showKills = sheet.getRange("G2").getValue();  
  const showDeaths = sheet.getRange("H2").getValue(); 
  const showKoef = sheet.getRange("L2").getValue();
  const showPercent = sheet.getRange("M2").getValue(); // Читаем глобальный чекбокс Процентов

  // Получаем плоские массивы заголовков и локальных чекбоксов строки 2
  const dynamicHeaders = sheet.getRange(5, stColGlobal + 1, 1, lastCol - stColGlobal).getValues()[0];
  const dynamicCheckboxesL2 = sheet.getRange(2, stColGlobal + 1, 1, lastCol - stColGlobal).getValues()[0];

  dynamicHeaders.forEach((headerText, idx) => {
    const currentGlobalColNum = stColGlobal + 1 + idx;
    const isL2Checked = dynamicCheckboxesL2[idx] === true || dynamicCheckboxesL2[idx] === "TRUE";
    const hText = headerText.toString().trim();

 
    
    // 2. ИСПРАВЛЕНО: Скрываем столбцы по глобальным флагам ТОЛЬКО если они равны ЛОЖЬ (галочка снята)
    if (hText.startsWith("Проц") && !showPercent) sheet.hideColumns(currentGlobalColNum);
    if (hText.startsWith("Кил") && !showKills) sheet.hideColumns(currentGlobalColNum);
    if (hText.startsWith("См") && !showDeaths) sheet.hideColumns(currentGlobalColNum);
    if (hText.startsWith("Коэф") && !showKoef) sheet.hideColumns(currentGlobalColNum);
    
    // 3. Если глобальный флаг включен (TRUE), то принудительно раскрываем столбец, чтобы он не схлопывался!
    if (hText.startsWith("Проц") && showPercent) sheet.showColumns(currentGlobalColNum);
    if (hText.startsWith("Кил") && showKills) sheet.showColumns(currentGlobalColNum);
    if (hText.startsWith("См") && showDeaths) sheet.showColumns(currentGlobalColNum);
    if (hText.startsWith("Коэф") && showKoef) sheet.showColumns(currentGlobalColNum);
  });
}



function buildGradient(range, type) {
  let rule = SpreadsheetApp.newConditionalFormatRule().setRanges([range]);
  
  if (type === "Проц") {
    // Градиент для процентов: от 0% (белый) до 30% вклада за день (сине-стальной)
    rule.setGradientMinpointWithValue("#ffffff", SpreadsheetApp.InterpolationType.NUMBER, "0")
        .setGradientMaxpointWithValue("#4a86e8", SpreadsheetApp.InterpolationType.NUMBER, "0,3");
  } else if (type === "Килы") {
    rule.setGradientMinpointWithValue("#ffffff", SpreadsheetApp.InterpolationType.NUMBER, "0")
        .setGradientMaxpointWithValue("#699ee6", SpreadsheetApp.InterpolationType.NUMBER, "80");
  } else if (type === "Смерти") {
    rule.setGradientMinpointWithValue("#ffffff", SpreadsheetApp.InterpolationType.NUMBER, "0")
        .setGradientMaxpointWithValue("#e06666", SpreadsheetApp.InterpolationType.NUMBER, "40");
  } else {
    rule.setGradientMinpointWithValue("#e06666", SpreadsheetApp.InterpolationType.NUMBER, "0")
        .setGradientMidpointWithValue("#ffffff", SpreadsheetApp.InterpolationType.NUMBER, "1")
        .setGradientMaxpointWithValue("#6aa84f", SpreadsheetApp.InterpolationType.NUMBER, "3");
  }
  
  return rule.build();
}


function columnToLetter(column) {
  let temp, letter = '';while (column > 0) {temp = (column - 1) % 26;
  letter = String.fromCharCode(temp + 65) + letter;
  column = (column - temp - 1) / 26;}return letter;
  }


