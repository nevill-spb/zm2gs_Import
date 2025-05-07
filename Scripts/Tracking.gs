const Tracking = (function() {
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

  // Кэшируем имя и объект листа
  let cachedSheetName = null;
  let sheet = null;
  let trackingEnabled = PropertiesService.getScriptProperties().getProperty(PROP_KEY) === 'true';

  // Получить объект листа по имени из настроек
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

  // Проверить, что переданный лист — отслеживаемый и не Settings
  function isWatchedSheet(checkSheet) {
    if (!checkSheet) return false;
    if (checkSheet.getName() === Settings.SHEETS.SETTINGS.NAME) {
      return false;
    }
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

  // Проверить, что колонка отслеживается
  function isWatchedColumn(col) {
    return WATCHED_COLUMNS_SET.has(col);
  }

  // Логирование для отладки
  function logDebug(message, data = null) {
    Logger.log('[Tracking] ' + message);
    if (data) Logger.log(JSON.stringify(data));
  }

  // Проверить, включено ли отслеживание
  function isTrackingEnabled() {
    return trackingEnabled;
  }

  // Создать триггеры для отслеживания изменений
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

  // Удалить триггеры отслеживания изменений
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

  function handleOnEdit(e) {
    try {
      if (!isTrackingEnabled()) return;
      if (!e || !e.range) return;
      
      const sheet = e.range.getSheet();
      if (!isWatchedSheet(sheet)) return;

      logDebug('onEdit вызван', {
        sheet: sheet.getName(),
        range: e.range.getA1Notation()
      });

      handleChange(e.range);
    } catch (error) {
      logDebug('Ошибка в onEdit', error);
    }
  }

  function handleOnChange(e) {
    try {
      if (!isTrackingEnabled()) return;
      if (!e) return;

      const activeSheet = SpreadsheetApp.getActiveSheet();
      if (!isWatchedSheet(activeSheet)) return;

      const range = activeSheet.getActiveRange();
      if (range) {
        logDebug('onChange вызван', {
          sheet: activeSheet.getName(),
          range: range.getA1Notation()
        });

        handleChange(range);
      }
    } catch (error) {
      logDebug('Ошибка в onChange', error);
    }
  }

  // Основная функция обработки изменений
  function handleChange(range) {
    try {
      const currentSheet = getSheet();
      if (!currentSheet) return;

      const startRow = range.getRow();
      const endRow = range.getLastRow();
      const startCol = range.getColumn();
      const endCol = range.getLastColumn();

      let needsModification = false;
      for (let col = startCol; col <= endCol; col++) {
        if (isWatchedColumn(col)) {
          needsModification = true;
          break;
        }
      }
      if (!needsModification) return;

      const dates = currentSheet.getRange(startRow, DATE_COLUMN, endRow - startRow + 1, 1).getValues();

      const rowsToModify = [];
      const valuesToSet = [];

      dates.forEach((dateRow, index) => {
        const date = dateRow[0];
        if (date !== "" && date !== null && date !== undefined) {
          rowsToModify.push(startRow + index);
          valuesToSet.push([true]);
        }
      });

      if (rowsToModify.length > 0) {
        let currentGroup = {
          start: rowsToModify[0],
          count: 1,
          values: [valuesToSet[0]]
        };
        const groups = [currentGroup];

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

        groups.forEach(group => {
          currentSheet.getRange(group.start, MODIFIED_COLUMN, group.count, 1)
            .setValues(group.values);
        });
      }
    } catch (error) {
      logDebug('Ошибка в handleChange', error);
    }
  }

  // Добавить пункты меню для управления триггерами
  function addTrackingMenuItems(importMenu) {
    return importMenu
      .addSeparator()
      .addItem('Setup Tracking Triggers', 'Tracking.createTriggers')
      .addItem('Clear Tracking Triggers', 'Tracking.deleteTriggers');
  }

  return {
    createTriggers,
    deleteTriggers,
    handleOnEdit,
    handleOnChange,
    addTrackingMenuItems
  };
})();
