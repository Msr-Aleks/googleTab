/**
 * ГЛАВНАЯ ФУНКЦИЯ (ОБНОВЛЕННАЯ С ОЧИСТКОЙ СТОЛБЦА I)
 */
function recordDailyStats_poly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const SOURCE_SHEET_NAME = "Даб_день"; 
  const DATE_CELL = "C2";

  if (ui.alert('Запись данных', 'Внести статистику в базы и очистить текущую таблицу?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  try {
    const sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sheet) throw new Error(`Лист "${SOURCE_SHEET_NAME}" не найден!`);

    let dateVal = sheet.getRange(DATE_CELL).getValue();
    let dateStr = (dateVal instanceof Date) 
      ? Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), "dd.MM") 
      : dateVal.toString(); 

    if (!dateStr) throw new Error("Ячейка C2 пуста!");

    // Собираем данные: ID(H), Имя(I), Килы(J), Смерти(K)
    const ids = sheet.getRange("H7:H38").getValues();
    const names = sheet.getRange("I7:I38").getValues(); // Читаем имена
    const kills = sheet.getRange("J7:J38").getValues();
    const deaths = sheet.getRange("K7:K38").getValues();
    
    const data = ids.map((row, i) => {
      return [
        row[0],      // ID (H)
        names[i][0], // Имя (I)
        kills[i][0], // Килы (J)
        deaths[i][0] // Смерти (K)
      ];
    });

    // Запись в базы
    updateBaseTable(ss, "Килы", data, dateStr, 2); 
    updateBaseTable(ss, "Смерти", data, dateStr, 3); 
    updateRatioTable(ss, "Коэф", "Килы", "Смерти", dateStr);

    // ОЧИСТКА: Имена (I), Килы (J) и Смерти (K)
    // Диапазон I7:K38 охватывает все три нужных столбца
    sheet.getRange("I7:K38").clearContent();

    ui.alert("Успех", `Данные за ${dateStr} сохранены.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("Ошибка: " + e.message);
  }
}
