/**
 * ГЛАВНАЯ ФУНКЦИЯ
 */
function recordDailyStats() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  //ui.alert(" 1 точка ");
  const SOURCE_TABLE_NAME = "Даб_день"; 
  const DATE_CELL = "C2";


  if (ui.alert('Запись данных', 'Внести статистику в базы и очистить текущую таблицу?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;



  try {

        
    const sourceTableInfo = findTable(ss, SOURCE_TABLE_NAME);
    const sheet = ss.getSheetByName(sourceTableInfo.sheetName);
    
      // Получаем значение из C2 и принудительно форматируем в "dd.MM"
    let dateVal = sheet.getRange(DATE_CELL).getValue();
    let dateStr = (dateVal instanceof Date) 
      ? Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), "dd.MM") 
      : dateVal.toString(); 

    if (!dateStr) throw new Error("Ячейка C2 пуста!");



    const range = sourceTableInfo.range;
    const data = sheet.getRange(
      range.startRowIndex + 2, 
      range.startColumnIndex + 2, 
      range.endRowIndex - range.startRowIndex - 1, 
      4
    ).getValues(); 
    // 1. Запись основных данных (теперь valIdx всегда 2, так как функция вернет [ID, Name, Value])
    updateBaseTable(ss, "Килы", data, dateStr, 2);
    updateBaseTable(ss, "Смерти", data, dateStr, 3); 
    //ui.alert("после Записи ");
    // 2. Расчет коэффициента (теперь передает корректный массив [ID, Name, Value])
    updateRatioTable(ss, "Коэф", "Килы", "Смерти", dateStr);

    // 3. Очистка ввода
    sheet.getRange(
      range.startRowIndex + 2, 
      range.startColumnIndex + 4, 
      range.endRowIndex - range.startRowIndex - 1, 
      2 
    ).clearContent();

    ui.alert("Успех", `Данные за ${dateStr} сохранены.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("Ошибка:  " + e.message);
  }
}






/**
 * ИСПРАВЛЕННАЯ ФУНКЦИЯ ЗАПИСИ
 */
function updateBaseTable(ss, tableName, sourceData, dateStr, valIdx) {
  let tableData = findTable(ss, tableName);
  const sheet = ss.getSheetByName(tableData.sheetName);
  const r = tableData.range;

  // 1. Получаем заголовки таблицы (от начала до реального конца данных в листе)
  const lastCol = sheet.getLastColumn();
  const startCol = r.startColumnIndex + 1;
  const numCols = Math.max(lastCol - r.startColumnIndex, 1);
  
  const headerRange = sheet.getRange(r.startRowIndex + 1, startCol, 1, numCols);
  const headers = headerRange.getDisplayValues()[0];
  
  // 2. Ищем дату среди заголовков
  let colOffset = headers.map(h => h.trim()).indexOf(dateStr);
  let targetCol;

  if (colOffset === -1) {
    // Находим ПЕРВУЮ пустую ячейку в строке заголовков таблицы
    let firstEmptyRelIdx = headers.indexOf("");
    if (firstEmptyRelIdx === -1) {
      targetCol = startCol + headers.length; // Если пустых нет, берем следующую за последней
    } else {
      targetCol = startCol + firstEmptyRelIdx; // Если есть дырка, пишем в неё
    }
    
    const cell = sheet.getRange(r.startRowIndex + 1, targetCol);
    cell.setNumberFormat("@"); // Строго текстовый формат
    cell.setValue(dateStr);    // Пишем "28.03"
    cell.setFontWeight("bold").setHorizontalAlignment("center");
  } else {
    targetCol = startCol + colOffset;
  }

  // 3. Запись данных по ID (без изменений)
  const existingIds = sheet.getRange(r.startRowIndex + 2, startCol, Math.max(sheet.getLastRow() - r.startRowIndex - 1, 1), 1).getValues().flat();

  sourceData.forEach(row => {
    const id = row[0];
    const name = row[1];
    const value = row[valIdx];

    if (!id || value === "" || value === null) return;

    let rowIdx = existingIds.indexOf(id);
    if (rowIdx === -1) {
      let newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, startCol, 1, 2).setValues([[id, name]]);
      sheet.getRange(newRow, targetCol).setValue(value);
      existingIds.push(id);
    } else {
      sheet.getRange(r.startRowIndex + 2 + rowIdx, targetCol).setValue(value);
    }
  });
}

/**
 * ПОЛНОСТЬЮ ИСПРАВЛЕННЫЙ РАСЧЕТ КОЭФФИЦИЕНТА С ВЫРАВНИВАНИЕМ ПО ID
 */
function updateRatioTable(ss, rName, kName, dName, dateStr) {
  const kTab = findTable(ss, kName);
  const dTab = findTable(ss, dName);
  
  const kSheet = ss.getSheetByName(kTab.sheetName);
  const dSheet = ss.getSheetByName(dTab.sheetName);

  // 1. Получаем заголовки и ищем колонку с нужной датой
  const kHeaders = kSheet.getRange(kTab.range.startRowIndex + 1, kTab.range.startColumnIndex + 1, 1, kSheet.getLastColumn()).getDisplayValues()[0];
  const colOffset = kHeaders.map(h => h.trim()).indexOf(dateStr);
  
  if (colOffset === -1) return; 

  // 2. Получаем ВСЕ данные из таблиц Килы и Смерти (начиная со второй строки таблицы - сразу под заголовком)
  const kFullRange = kSheet.getRange(
    kTab.range.startRowIndex + 2, 
    kTab.range.startColumnIndex + 1, 
    kTab.range.endRowIndex - kTab.range.startRowIndex - 1, 
    colOffset + 1
  ).getValues();

  const dFullRange = dSheet.getRange(
    dTab.range.startRowIndex + 2, 
    dTab.range.startColumnIndex + 1, 
    dTab.range.endRowIndex - dTab.range.startRowIndex - 1, 
    colOffset + 1
  ).getValues();

  // 3. Создаем "карту" смертей для быстрого поиска по ID
  let deathsMap = {};
  dFullRange.forEach(row => {
    const id = row[0] ? row[0].toString().trim() : null;
    if (id) deathsMap[id] = parseFloat(row[colOffset]) || 0;
  });

  // 4. Считаем коэффициент, сопоставляя данные по ID
  const ratioData = kFullRange.map(row => {
    const id = row[0] ? row[0].toString().trim() : null;
    const name = row[1];
    const k = parseFloat(row[colOffset]) || 0;
    const d = deathsMap[id] || 0; // Ищем смерти этого же игрока по его ID

    // Если нет данных (0 килов и 0 смертей), возвращаем null (чтобы не писать 0)
    if (!id || (k === 0 && d === 0)) return [id, name, null];
    const ratio = (d === 0) ? k : (k / d);
    // Округляем математически, чтобы передать ЧИСЛО, а не текст
    const roundedRatio = Math.round(ratio * 100) / 100;
    return [id, name, roundedRatio];
  });

  // 5. Записываем в таблицу Коэф
  updateBaseTable(ss, rName, ratioData, dateStr, 2);
}





function findTable(ss, tableName) {
  const spreadsheet = Sheets.Spreadsheets.get(ss.getId());
  for (let s of spreadsheet.sheets) {
    if (s.tables) {
      for (let t of s.tables) {
        if ((t.displayName || t.name) === tableName) return { sheetName: s.properties.title, range: t.range || t.spec.range };
      }
    }
  }
  throw new Error(`Таблица "${tableName}" не найдена!`);
}
