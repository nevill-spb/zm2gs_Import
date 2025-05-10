const Tracking = (function() {
  //═══════════════════════════════════════════════════════════════════════════
  // КОНСТАНТЫ И КОНФИГУРАЦИЯ
  // Этот блок содержит все константы и настройки модуля:
  // - PROP_KEY: ключ для хранения состояния в ScriptProperties
  // - WATCHED_FIELDS: список полей, изменения которых нужно отслеживать
  // - WATCHED_COLUMNS: преобразованные индексы колонок для отслеживаемых полей
  // - DATE_COLUMN и MODIFIED_COLUMN: специальные колонки для работы с датами и чекбоксами
  //═══════════════════════════════════════════════════════════════════════════
  const PROP_KEY = 'TRACKING_ENABLED';
  const WATCHED_FIELDS = [
    'date',
    'tag',
    'tag1',
    'tag2',
    'merchant',
    'comment',
    'outcomeAccount',
    'outcome',
    'incomeAccount',
    'income'
  ];

  const WATCHED_COLUMNS = WATCHED_FIELDS.map(fieldId => {
    const index = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === fieldId);
    if (index === -1) {
      Logger.log(`Поле ${fieldId} не найдено в TRANSACTION_FIELDS`);
      return null;
    }
    return index + 1;
  }).filter(col => col !== null);

  const WATCHED_COLUMNS_SET = new Set(WATCHED_COLUMNS);
  const DATE_COLUMN = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === 'date') + 1;
  const MODIFIED_COLUMN = Settings.TRANSACTION_FIELDS.findIndex(f => f.id === 'modified') + 1;

  //═══════════════════════════════════════════════════════════════════════════
  // СОСТОЯНИЕ МОДУЛЯ
  // Здесь хранятся переменные, отвечающие за текущее состояние модуля:
  // - cachedSheetName: кэшированное имя активного листа
  // - sheet: кэшированный объект листа
  // - trackingEnabled: флаг, включено ли отслеживание
  //═══════════════════════════════════════════════════════════════════════════
  let cachedSheetName = null;
  let sheet = null;
  let trackingEnabled = PropertiesService.getScriptProperties().getProperty(PROP_KEY) === 'true';

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  // Набор утилитарных функций для:
  // - логирования (logDebug)
  // - проверки состояния (isTrackingEnabled)
  // - проверки колонок (isWatchedColumn)
  //═══════════════════════════════════════════════════════════════════════════
  function logDebug(message, data = null) {
    Logger.log('[Tracking] ' + message);
    if (data) Logger.log(JSON.stringify(data));
  }

  function isTrackingEnabled() {
    return trackingEnabled;
  }

  function isWatchedColumn(col) {
    return WATCHED_COLUMNS_SET.has(col);
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ЛИСТАМИ
  // Функции для взаимодействия с Google Sheets:
  // - getSheet: получение и кэширование активного листа
  // - isWatchedSheet: проверка, нужно ли отслеживать изменения в листе
  //═══════════════════════════════════════════════════════════════════════════
  function getSheet() {
    try {
      const currentSheetName = sheetHelper.GetCellValue(
        Settings.SHEETS.SETTINGS.NAME, 
        Settings.SHEETS.SETTINGS.CELLS.IMPORT_SHEET
      );
      if (!sheet || currentSheetName !== cachedSheetName) {
        cachedSheetName = currentSheetName;
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentSheetName);
      }
      return sheet;
    } catch (error) {
      logDebug('Ошибка при получении листа', error);
      return null;
    }
  }

  function isWatchedSheet(checkSheet) {
    if (!checkSheet) return false;
    if (checkSheet.getName() === Settings.SHEETS.SETTINGS.NAME) return false;
    
    try {
      const importSheetName = sheetHelper.GetCellValue(
        Settings.SHEETS.SETTINGS.NAME, 
        Settings.SHEETS.SETTINGS.CELLS.IMPORT_SHEET
      );
      return checkSheet.getName() === importSheetName;
    } catch (error) {
      logDebug('Ошибка при проверке листа', error);
      return false;
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // УПРАВЛЕНИЕ ТРИГГЕРАМИ
  // Функции для управления триггерами Google Apps Script:
  // - createTriggers: создание триггеров onEdit и onChange
  // - deleteTriggers: удаление существующих триггеров
  // Триггеры необходимы для отслеживания изменений в реальном времени
  //═══════════════════════════════════════════════════════════════════════════
  function createTriggers() {
    try {
      deleteTriggers(true);

      ScriptApp.newTrigger('onEdit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();

      ScriptApp.newTrigger('onChange')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onChange()
        .create();

      trackingEnabled = true;
      PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'true');
      SpreadsheetApp.getActive().toast('Триггеры успешно установлены!', 'Успех');
      logDebug('Триггеры установлены');
    } catch (error) {
      logDebug('Ошибка при установке триггеров', error);
      SpreadsheetApp.getActive().toast('Ошибка при установке триггеров: ' + error.toString(), 'Ошибка');
    }
  }

  function deleteTriggers(silent = false) {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'onEdit' || trigger.getHandlerFunction() === 'onChange') {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      trackingEnabled = false;
      PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'false');
      
      if (!silent) {
        SpreadsheetApp.getActive().toast('Триггеры успешно отключены!', 'Успех');
      }
      logDebug('Триггеры удалены');
    } catch (error) {
      logDebug('Ошибка при удалении триггеров', error);
      if (!silent) {
        SpreadsheetApp.getActive().toast('Ошибка при удалении триггеров: ' + error.toString(), 'Ошибка');
      }
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ОБРАБОТКА ИЗМЕНЕНИЙ
  // Основной блок функций для обработки изменений в таблице:
  // - validateAndGetRange: проверка и получение диапазона изменений
  // - handleOnEdit/handleOnChange: обработчики событий редактирования
  // - handleChange: основная логика обработки изменений
  //═══════════════════════════════════════════════════════════════════════════
  function validateAndGetRange(e) {
    if (!isTrackingEnabled()) return null;
    if (!e) return null;

    let targetSheet, range;
    
    if (e.range) {
      // Событие onEdit
      targetSheet = e.range.getSheet();
      range = e.range;
    } else {
      // Событие onChange
      targetSheet = SpreadsheetApp.getActiveSheet();
      range = targetSheet.getActiveRange();
    }

    if (!isWatchedSheet(targetSheet) || !range) return null;
    
    return {
      sheet: targetSheet,
      range: range
    };
  }

  function handleOnEdit(e) {
    try {
      const validated = validateAndGetRange(e);
      if (!validated) return;

      logDebug('onEdit вызван', {
        sheet: validated.sheet.getName(),
        range: validated.range.getA1Notation()
      });

      handleChange(validated.range);
    } catch (error) {
      logDebug('Ошибка в onEdit', error);
    }
  }

  function handleOnChange(e) {
    try {
      const validated = validateAndGetRange(e);
      if (!validated) return;

      logDebug('onChange вызван', {
        sheet: validated.sheet.getName(),
        range: validated.range.getA1Notation()
      });

      handleChange(validated.range);
    } catch (error) {
      logDebug('Ошибка в onChange', error);
    }
  }

  function handleChange(range) {
    try {
      const currentSheet = getSheet();
      if (!currentSheet) return;

      const startRow = range.getRow();
      const endRow = range.getLastRow();
      const startCol = range.getColumn();
      const endCol = range.getLastColumn();

      // Проверяем, есть ли изменения в отслеживаемых колонках
      let needsModification = false;
      for (let col = startCol; col <= endCol; col++) {
        if (isWatchedColumn(col)) {
          needsModification = true;
          break;
        }
      }
      if (!needsModification) return;

      // Получаем даты для всех затронутых строк
      const dates = currentSheet.getRange(startRow, DATE_COLUMN, endRow - startRow + 1, 1).getValues();

      // Формируем списки строк и значений для обновления
      const rowsToModify = [];
      const valuesToSet = [];

      dates.forEach((dateRow, index) => {
        const date = dateRow[0];
        if (date !== "" && date !== null && date !== undefined) {
          rowsToModify.push(startRow + index);
          valuesToSet.push([true]);
        }
      });

      // Если есть строки для обновления, группируем их и обновляем
      if (rowsToModify.length > 0) {
        let currentGroup = {
          start: rowsToModify[0],
          count: 1,
          values: [valuesToSet[0]]
        };
        const groups = [currentGroup];

        // Группируем последовательные строки
        for (let i = 1; i < rowsToModify.length; i++) {
          if (rowsToModify[i] === currentGroup.start + currentGroup.count) {
            currentGroup.count++;
            currentGroup.values.push(valuesToSet[i]);
          } else {
            currentGroup = {
              start: rowsToModify[i],
              count: 1,
              values: [valuesToSet[i]]
            };
            groups.push(currentGroup);
          }
        }

        // Обновляем каждую группу строк
        groups.forEach(group => {
          currentSheet.getRange(group.start, MODIFIED_COLUMN, group.count, 1)
            .setValues(group.values);
        });
      }
    } catch (error) {
      logDebug('Ошибка в handleChange', error);
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ИНТЕРФЕЙС МЕНЮ
  // Функции для добавления пунктов управления триггерами в меню таблицы
  //═══════════════════════════════════════════════════════════════════════════
  function addMenuItems(importMenu) {
    return importMenu
      .addSeparator()
      .addItem('Setup Tracking Triggers', 'Tracking.createTriggers')
      .addItem('Clear Tracking Triggers', 'Tracking.deleteTriggers');
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЙ ИНТЕРФЕЙС
  // Экспортируемые функции модуля, доступные для внешнего использования
  //═══════════════════════════════════════════════════════════════════════════
  return {
    createTriggers,
    deleteTriggers,
    handleOnEdit,
    handleOnChange,
    addMenuItems
  };
})();
