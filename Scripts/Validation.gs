const Validation = (function() {
  // Приватная функция инициализации настроек режима тегов
  function initializeTagModeSettings() {    
    const rule = SpreadsheetApp.newDataValidation()    
      .requireValueInList([Settings.TAG_MODES.SINGLE_COLUMN, Settings.TAG_MODES.MULTIPLE_COLUMNS])    
      .build();    
        
    sheetHelper.GetRange(    
      Settings.SHEETS.SETTINGS.NAME,     
      Settings.SHEETS.SETTINGS.CELLS.TAG_MODE    
    ).setDataValidation(rule);    
  }

  // Функция для получения уникальных непустых значений
  function uniqueNonEmpty(array) {
    return [...new Set(array.filter(v => v && v.trim() !== ''))];
  }

  // Функция поиска индекса колонки по id поля
  function findColumnIndex(fieldId) {
    return Settings.TRANSACTION_FIELDS.findIndex(field => field.id === fieldId) + 1;
  }

  // Функция установки валидации для счетов и тегов
  function setupFieldsValidation() {
    try {
      const sheet = sheetHelper.GetSheetFromSettings('EXPORT_SHEET');
      const lastRow = sheet.getLastRow();
      const maxRows = sheet.getMaxRows();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow <= 1) return; // только заголовок

      // Очищаем все данные ниже последней заполненной строки
      if (lastRow < maxRows) {
        sheet.getRange(lastRow + 1, 1, maxRows - lastRow, lastCol).clearDataValidations();      
      }

      // Загружаем словари
      const dictionaries = Dictionaries.getAllDictionaries();
      if (!dictionaries) {
        throw new Error('Не удалось загрузить справочники');
      }

      // Только реальные, уникальные значения
      const accountTitles = uniqueNonEmpty(Object.values(dictionaries.accounts)).sort((a, b) =>   
        a.localeCompare(b, 'ru')  
      );
      const tagTitles = uniqueNonEmpty(Object.values(dictionaries.tags)).sort((a, b) =>   
        a.localeCompare(b, 'ru')
      );

      // Создаём правила валидации
      const accountRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(accountTitles)
        .setAllowInvalid(true)
        .build();

      const tagRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['', ...tagTitles])
        .setAllowInvalid(true)
        .build();

      // Группируем колонки по типу валидации
      const accountColumns = [
        findColumnIndex('outcomeAccount'),
        findColumnIndex('incomeAccount')
      ].filter(idx => idx > 0);

      const tagColumns = (() => {
        const tagMode = Settings.TagMode;
        const columns = [findColumnIndex('tag')];
        
        if (tagMode === Settings.TAG_MODES.MULTIPLE_COLUMNS) {
          columns.push(
            findColumnIndex('tag1'),
            findColumnIndex('tag2')
          );
        }
        return columns.filter(idx => idx > 0);
      })();

      // Применяем валидацию пакетно для каждой группы колонок
      accountColumns.forEach(colIndex => {
        sheet.getRange(2, colIndex, lastRow - 1, 1).setDataValidation(accountRule);
      });

      tagColumns.forEach(colIndex => {
        sheet.getRange(2, colIndex, lastRow - 1, 1).setDataValidation(tagRule);
      });

    } catch(e) {
      throw new Error(`Ошибка при установке валидации: ${e.message}`);
    }
  }

  // Функция очистки валидации
  function clearAllValidation() {
    try {
      const sheet = sheetHelper.GetSheetFromSettings('EXPORT_SHEET');
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow <= 1) return; // только заголовок

      // Находим индексы столбцов, которые нужно пропустить
      const skipColumns = [
        findColumnIndex('deleted'),
        findColumnIndex('modified')
      ].filter(idx => idx > 0);

      // Если есть столбцы для пропуска, очищаем валидацию по частям
      if (skipColumns.length > 0) {
        // Сортируем индексы по возрастанию
        skipColumns.sort((a, b) => a - b);

        // Очищаем валидацию между пропускаемыми столбцами
        let startCol = 1;
        for (const skipCol of skipColumns) {
          if (skipCol > startCol) {
            // Очищаем диапазон до текущего пропускаемого столбца
            sheet.getRange(2, startCol, lastRow - 1, skipCol - startCol).clearDataValidations();
          }
          startCol = skipCol + 1;
        }

        // Очищаем оставшийся диапазон после последнего пропускаемого столбца
        if (startCol <= lastCol) {
          sheet.getRange(2, startCol, lastRow - 1, lastCol - startCol + 1).clearDataValidations();
        }
      } else {
        // Если нет столбцов для пропуска, очищаем всё
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearDataValidations();
      }

      SpreadsheetApp.getActive().toast('Валидация успешно очищена');
    } catch(e) {
      SpreadsheetApp.getActive().toast('Ошибка при очистке валидации: ' + e.message, 'Ошибка');
      console.error('Ошибка при очистке валидации:', e);
    }
  }

  // Функция для вызова из меню
  function setupValidation() {
    try {
      // Убедимся, что справочники загружены
      Dictionaries.loadDictionariesFromSheet();
      initializeTagModeSettings();
      setupFieldsValidation();
      SpreadsheetApp.getActive().toast('Валидация настроена успешно');
    } catch(e) {
      SpreadsheetApp.getActive().toast('Ошибка при настройке валидации: ' + e.message, 'Ошибка');
      console.error('Ошибка при настройке валидации:', e);
    }
  }

  // Публичный интерфейс
  return {
    setupValidation,
    clearAllValidation
  };
})();
