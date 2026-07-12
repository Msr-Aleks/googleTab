// =========================================================================
// 3. ФУНКЦИЯ-ЯДРО (ГЕНЕРАЦИЯ ДАННЫХ И ИТОГОВ)
// =========================================================================
function runDynamicGenerationCore(startDate, endDate, sheet, stColumnNum, insertAtCol, isOverwriteMode) {
    
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ui = SpreadsheetApp.getUi();

  const showKills = sheet.getRange("G2").getValue();
  const showDeaths = sheet.getRange("H2").getValue();
  const showKoef = sheet.getRange("L2").getValue();
  const showProc = sheet.getRange("M2").getValue();

  let dateList = [];
  let curr = new Date(startDate);
  while (curr <= endDate) {
    dateList.push(Utilities.formatDate(curr, tz, "dd.MM"));
    curr.setDate(curr.getDate() + 1);
  }

  try {
    const spreadsheet = Sheets.Spreadsheets.get(ss.getId());
    let targetTable;
    let headersMap = {};

    spreadsheet.sheets.forEach(s => {
      if (!s.tables) return;
      s.tables.forEach(t => {
        let name = t.displayName || t.name;
        if (name === "Статистика") targetTable = t;
        if (["Килы", "Смерти", "Коэф"].includes(name)) {
          const r = t.range || t.spec.range;
          // ИСПРАВЛЕНО: Добавлен индекс [0], чтобы получить правильный плоский массив заголовков
          headersMap[name] = ss.getSheetByName(s.properties.title)
            .getRange(r.startRowIndex + 1, r.startColumnIndex + 1, 1, r.endColumnIndex - r.startColumnIndex)
            .getValues()[0].map(h => String(h).trim());
        }
      });
    });

    if (!targetTable) throw new Error('Таблица "Статистика" не найдена!');
    const rSpec = targetTable.range || targetTable.spec.range;
    const startRow = rSpec.startRowIndex + 1;
    const lastDataRow = rSpec.endRowIndex + 1;
    const totalRowIdx = lastDataRow + 1;
    const numDataRows = lastDataRow - startRow;

    let neededColumns = [];
    let checkDate = new Date(startDate);
    dateList.forEach(date => {
      // Определяем день недели (0 - воскресенье, 1 - понедельник, ..., 6 - суббота)
      const dayOfWeek = checkDate.getDay(); 

      // ВАЖНО: Добавляем Процент ПЕРЕД Килами, если Килы существуют в базе
      if (headersMap["Килы"]?.includes(date)) {
        neededColumns.push({ name: `Проц ${date}`, base: "Проц", date: date, dayOfWeek: dayOfWeek });
        neededColumns.push({ name: `Кил ${date}`, base: "Килы", date: date, dayOfWeek: dayOfWeek });
      }
      if (headersMap["Смерти"]?.includes(date)) neededColumns.push({ name: `См ${date}`, base: "Смерти", date: date, dayOfWeek: dayOfWeek });
      if (headersMap["Коэф"]?.includes(date)) neededColumns.push({ name: `Коэфф ${date}`, base: "Коэф", date: date, dayOfWeek: dayOfWeek });
      
      checkDate.setDate(checkDate.getDate() + 1);
    });



    if (neededColumns.length === 0) {
      ss.toast(isOverwriteMode ? "Нет данных в базовых таблицах за эти даты." : "Нет новых данных в базовых таблицах за эти даты.");
      if (!isOverwriteMode && typeof reapplyInitialHiding === "function") reapplyInitialHiding(sheet, stColumnNum);
      return;
    }

    const numCols = neededColumns.length;

    // Расширяем лист при необходимости (только в режиме добора)
    if (!isOverwriteMode) {
      const maxCols = sheet.getMaxColumns();
      if (insertAtCol + numCols > maxCols) {
        sheet.insertColumnsAfter(maxCols, (insertAtCol + numCols) - maxCols);
      }
    }






    let finalHeaders = [];
    let finalFormulas = Array.from({ length: numDataRows }, () => []);
    let finalFootersRow1 = [];
    let finalFootersRow2 = [];
    
    // ИСПРАВЛЕНО: Полностью сбрасываем старые динамические правила правее St во всех режимах.
    // Это исключает дублирование, так как в конце мы сгенерируем чистые правила сразу на всю обновленную ширину.
    let updatedRules = sheet.getConditionalFormatRules().filter(r => {
      return r.getRanges().every(rng => rng.getColumn() <= stColumnNum);
    });

    // ИСПРАВЛЕНИЕ: Жестко вычищаем старые данные, формулы и скрытые чекбоксы (валидацию)
    // на всем промежутке новых столбцов в строках 2 и 4 перед их генерацией
    const clearRange2 = sheet.getRange(2, insertAtCol, 1, numCols);
    const clearRange4 = sheet.getRange(4, insertAtCol, 1, numCols);
    
    clearRange2.clearContent().clearDataValidations();
    clearRange4.clearContent().clearDataValidations();

    neededColumns.forEach((col, idx) => {
      const colIdx = insertAtCol + idx;
      
      // Высчитываем смещения для формул
      // Так как добавился Процент, смещение для Коэф относительно Килов изменилось
      let koefOffset = 0;
      if (col.base === "Проц") koefOffset = 3;
      else if (col.base === "Килы") koefOffset = 2;
      else if (col.base === "Смерти") koefOffset = 1;

      const koefColLetter = columnToLetter(colIdx + koefOffset);
      finalHeaders.push(col.name);



     

      // Формулы участников
      for (let i = 0; i < numDataRows; i++) {
        const row = startRow + 1 + i;
        
        if (col.base === "Проц") {
          // Столбец Кил идет сразу после Процента, то есть на +1 столбец правее
          const killsLetter = columnToLetter(colIdx + 1);
          // Диапазон всех игроков в столбце Килов для текущего дня
          const killsDataRange = `$${killsLetter}$${startRow + 1}:$${killsLetter}$${lastDataRow}`;
          // console.log("killsDataRange - "+killsDataRange);
          // ОБНОВЛЕННАЯ ФОРМУЛА: Проверяет через COUNTIF, что участников с килами больше 10.
          // Если участников <= 10, ячейка остается пустой.
          finalFormulas[i].push(
            `=LET(` +
              `playersCount; COUNTIF(${killsDataRange}; ">0"); ` +
              `p; IFERROR(IF(ISNUMBER(VALUE(${killsLetter}${row})); VALUE(${killsLetter}${row}) / ${killsLetter}$${totalRowIdx}; ""); ""); ` +
              `IF(playersCount > 10; IF(p=0; ""; p); "")` +
            `)`
          );
        } else {
          // Базовая формула для Кил, См, Коэф
          finalFormulas[i].push(`=LET(val; IFERROR(LET(valRaw; XLOOKUP(A${row}; ${col.base}[id]; ${col.base}[${col.date}]); res; IF(valRaw="X"; "X"; IF(valRaw>0; MAX(0; valRaw - N${row}); valRaw)); IF(res=0; res; res)); ""); IF(${koefColLetter}$4=FALSE; IF(ISNUMBER(val); TEXT(val; "0"); val); val))`);
        }
      }


      const colLetter = columnToLetter(colIdx);
      const dataRangeA1 = `${colLetter}${startRow + 1}:${colLetter}${lastDataRow}`;

      // 1 СТРОКА ИТОГОВ
      if (col.base === "Килы" || col.base === "Смерти") {
        finalFootersRow1.push(`=IF(${koefColLetter}$4=FALSE; 0; IFERROR(SUM(${dataRangeA1}); 0))`);
      } else if (col.base === "Проц") {
        // Сумма всех процентов за день всегда равна 100% (или 1)
        finalFootersRow1.push(`=IFERROR(SUM(${dataRangeA1}); 0)`);
      } else if (col.base === "Коэф") {
        const killsColLetter = columnToLetter(colIdx - 2); // Смещение -2, так как перед См идет Кил
        const deathsColLetter = columnToLetter(colIdx - 1);
        finalFootersRow1.push(`=IF(${colLetter}$4=FALSE; 0; IFERROR(${killsColLetter}${totalRowIdx} / ${deathsColLetter}${totalRowIdx}; 0))`);
      }

      // 2 СТРОКА ИТОГОВ
      if (col.base === "Коэф") {
        finalFootersRow2.push(`=IF(${colLetter}$4=FALSE; 0; IFERROR(AVERAGEIF(${dataRangeA1}; ">0"); 0))`);
      } else {
        finalFootersRow2.push("");
      }

      // Настройка числовых форматов диапазона игроков и 2 строк итогов
      const fullColRange = sheet.getRange(startRow + 1, colIdx, numDataRows + 2, 1);
      if (col.base === "Проц") {
        fullColRange.setNumberFormat("0.0%");
      } else {
        fullColRange.setNumberFormat(col.base === "Коэф" ? "0.00" : "0");
      }

      if (col.base === "Коэф") {
        sheet.getRange(2, colIdx).insertCheckboxes().setValue(false).setHorizontalAlignment("center");
        sheet.getRange(4, colIdx).insertCheckboxes().setValue(true).setHorizontalAlignment("center");
      }
    });

   
    // =========================================================================
    // ШАГ 1: ФИЗИЧЕСКАЯ ЗАПИСЬ ДАННЫХ (НЕОБХОДИМА ДЛЯ ПРАВИЛЬНОГО РАСЧЕТА ДИАПАЗОНОВ)
    // =========================================================================
    const sourceMarker = sheet.getRange("K5"); 
    const rangeRow5 = sheet.getRange(5, insertAtCol, 1, numCols);
    const rangeRow6 = sheet.getRange(6, insertAtCol, 1, numCols);
    
    let bgColors = [];
    let textColors = [];
    neededColumns.forEach(col => {
      if (col.base === "Проц") {
        bgColors.push("#1f4e79"); // Темно-синий благородный цвет для контраста
        textColors.push("#ffffff"); // Белый текст
      }
      else if (col.base === "Килы") { 
        bgColors.push("#699ee6"); 
        textColors.push("#274e13"); 
      } 
      else if (col.base === "Смерти") { 
        bgColors.push("#f4cccc"); 
        textColors.push("#660000"); 
      } 
      else if (col.base === "Коэф") { 
        bgColors.push("#356854"); 
        if (col.dayOfWeek === 6) {
          textColors.push("#ff9900"); 
        } else if (col.dayOfWeek === 0) {
          textColors.push("#ff8782"); 
        } else {
          textColors.push("#ffffff"); 
        }
      } 
      else { 
        bgColors.push("#ffffff"); 
        textColors.push("#000000"); 
      }
    });

    sourceMarker.copyTo(rangeRow5, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    rangeRow5.setValues([finalHeaders]).setBackgrounds([bgColors]).setFontColors([textColors]).setFontWeight("bold").setHorizontalAlignment("center");
    rangeRow6.setValues([finalHeaders]).setFontWeight("bold").setTextRotation(0).setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    // Вставка формул участников и итог-строк
    sheet.getRange(startRow + 1, insertAtCol, numDataRows, numCols).setFormulas(finalFormulas);
    sheet.getRange(totalRowIdx, insertAtCol, 2, numCols).setFormulas([finalFootersRow1, finalFootersRow2]).setFontWeight("bold");

    // =========================================================================
    // ШАГ 2: ГЛОБАЛЬНЫЙ ПЕРЕРАСЧЕТ И СБОРКА УСЛОВНОГО ФОРМАТИРОВАНИЯ
    // =========================================================================
    const finalLastCol = sheet.getLastColumn();

    if (finalLastCol > stColumnNum) {

      // Пересканируем ВСЕ динамические столбцы от St+1 до самого правого края листа
      const allDynamicHeaders = sheet.getRange(6, stColumnNum + 1, 1, finalLastCol - stColumnNum).getValues()[0];
      
      let globalKillsRanges = [];
      let globalDeathsRanges = [];
      let globalKoefRanges = [];
      let globalPercentRanges = [];
      let globalAllRanges = []; // Для одной компактной серой заливки

      allDynamicHeaders.forEach((hText, idx) => {
        const currentGlobalColNum = stColumnNum + 1 + idx;
        const colRange = sheet.getRange(startRow + 1, currentGlobalColNum, numDataRows + 2, 1);
        globalAllRanges.push(colRange);

        // Распределяем столбцы по их фактическим именам на листе
        const hString = hText.toString();
        if (hString.startsWith("Проц")) globalPercentRanges.push(colRange); // Распознаем Процент
        if (hString.startsWith("Кил")) globalKillsRanges.push(colRange);
        if (hString.startsWith("См")) globalDeathsRanges.push(colRange);
        if (hString.startsWith("Коэфф")) globalKoefRanges.push(colRange);
      });

      // 1. Создаем ОДНО общее правило серой заливки для ВСЕХ динамических столбцов таблицы
      if (globalAllRanges.length > 0) {
        const grayRule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=R4C=ЛОЖЬ`)
          .setBackground("#efefef")
          .setFontColor("#999999")
          .setRanges(globalAllRanges)
          .build();
        updatedRules.push(grayRule);
      }

      // 2. Генерируем правила градиентов для каждого столбца
      if (typeof buildGradient === "function") {
        globalPercentRanges.forEach(r => updatedRules.push(buildGradient(r, "Проц"))); // Применяем градиент к процентам
        globalKillsRanges.forEach(r => updatedRules.push(buildGradient(r, "Килы")));
        globalDeathsRanges.forEach(r => updatedRules.push(buildGradient(r, "Смерти")));
        globalKoefRanges.forEach(r => updatedRules.push(buildGradient(r, "Коэф")));
      }

    }

    // Сохраняем чистый, обновленный массив правил условного форматирования
    sheet.setConditionalFormatRules(updatedRules);

    // Установка индивидуальной ширины столбцов

    if (numCols > 0) {
      // 1. Получаем общее количество строк на листе, чтобы правильно выровнять весь столбец
      const maxRows = sheet.getMaxRows(); 
      
      neededColumns.forEach((col, idx) => {
        const colNum = insertAtCol + idx;
        
        // 2. ИСПРАВЛЕНО: Объявляем диапазон для текущего столбца (от 5 строки до самого низа)
        const entireColumnRange = sheet.getRange(5, colNum, maxRows - 4, 1);
        
        // 3. Выставляем ширину и шрифты в зависимости от типа данных
        if (col.base === "Килы" || col.base === "Смерти") {
          sheet.setColumnWidth(colNum, 32);
          // По желанию, если нужно отцентрировать и Килы со Смертями, можно раскомментировать строку ниже:
          // entireColumnRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
          
        } else if (col.base === "Коэф") {
          sheet.setColumnWidth(colNum, 40);
          entireColumnRange.setHorizontalAlignment("center")
                            .setVerticalAlignment("middle")
                            .setFontSize(10);
                            
        } else if (col.base === "Проц") { // ИСПРАВЛЕНО: изменено на "Процент" для точного совпадения с базой
          sheet.setColumnWidth(colNum, 42);
          entireColumnRange.setHorizontalAlignment("center")
                            .setVerticalAlignment("middle")
                            .setFontSize(9);
                            
        } else {
          sheet.setColumnWidth(colNum, 38); // Дефолтный размер на всякий случай
        }
      });
    }


 
    // Автоподстановка буквы в формулу M1 со ВТОРОГО динамического столбца (Килы)
    if (isOverwriteMode) {
      const secondDynamicLetter = columnToLetter(insertAtCol + 1); // Смещение +1 дает второй столбец
      sheet.getRange("N1").setFormula(`=LEFT(${secondDynamicLetter}5; 2)`).clearFormat();
    }

    // ИСПРАВЛЕНО: Умное первичное скрытие колонок строго на основе глобальных чекбоксов
    neededColumns.forEach((col, idx) => {
      const colNum = insertAtCol + idx;
      
      // Скрываем столбец ТОЛЬКО если соответствующий глобальный флаг выключен (равен FALSE)
      if (col.base === "Проц" && !showProc) {
        sheet.hideColumns(colNum);
      }
      else if (col.base === "Килы" && !showKills) {
        sheet.hideColumns(colNum);
      }
      else if (col.base === "Смерти" && !showDeaths) {
        sheet.hideColumns(colNum);
      }
      else if (col.base === "Коэф" && !showKoef) {
        sheet.hideColumns(colNum);
      }
    });



    // Восстановление компактного вида на основе чекбоксов
    if (!isOverwriteMode && typeof reapplyInitialHiding === "function") {
      reapplyInitialHiding(sheet, stColumnNum);
    }

    // ИСПРАВЛЕНИЕ 2: Убираем пустой фантомный столбец, если он образовался в конце листа
    const finalMax = sheet.getMaxColumns();
    const expectedEnd = insertAtCol + numCols - 1;
    if (finalMax > expectedEnd && !isOverwriteMode) {
      sheet.deleteColumns(expectedEnd + 1, finalMax - expectedEnd);
    }







    ss.toast("Готово!");
  } catch (e) {
    ui.alert("Ошибка: " + e.message);
  }
}




