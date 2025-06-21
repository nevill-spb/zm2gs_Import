//═══════════════════════════════════════════════════════════════════════════
// ГЛОБАЛЬНЫЕ ТРИГГЕРЫ
// Функции, которые вызываются автоматически при определенных событиях
//═══════════════════════════════════════════════════════════════════════════
function onEdit(e) {
  try {
    if (typeof Tracking !== 'undefined') {
      Tracking.handleOnEdit(e);
    }
  } catch (error) {
    Logger.log("Ошибка в onEdit: " + error.toString());
  }
}

function onChange(e) {
  try {
    if (typeof Tracking !== 'undefined') {
      Tracking.handleOnChange(e);
    }
  } catch (error) {
    Logger.log("Ошибка в onChange: " + error.toString());
  }
}

function onOpen(e) {
  try {
    createMenu();
  } catch (error) {
    Logger.log("Ошибка в onOpen: " + error.toString());
  }
}

//═══════════════════════════════════════════════════════════════════════════
// МЕНЮ
// Функции для создания и управления меню Zen Money
//═══════════════════════════════════════════════════════════════════════════
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Дзен Мани');
  
  // Всегда добавляем основные пункты
  menu
    .addItem('Полная Синхронизация', 'doFullSync')
    .addItem('Обновить Словари', 'doUpdateDictionaries');

  // Добавляем подменю только если есть соответствующие функции
  try {
    // Подменю Экспорт
    if (typeof Export !== 'undefined') {
      menu.addSeparator().addSubMenu(
        ui.createMenu('Экспорт')
          .addItem('Полный экспорт', 'Export.doFullExport')
          .addItem('Инкрементальный экспорт', 'Export.doIncrementalExport')
          .addItem('Подготовить лист изменений', 'Export.prepareChangesSheet')
          .addItem('Применить изменения', 'Export.applyChangesToDataSheet')
      );
    }

    // Подменю Импорт
    if (typeof Import !== 'undefined') {
      const importMenu = ui.createMenu('Импорт')
        .addItem('Частичный импорт', 'Import.doUpdate');
      
      // Дополнительные пункты для Импорта
      try { if (typeof Tracking !== 'undefined') Tracking.addMenuItems?.(importMenu); } catch(e) {}
      try { if (typeof Validation !== 'undefined') Validation.addMenuItems?.(importMenu); } catch(e) {}
      
      menu.addSubMenu(importMenu);
    }

    // Универсальная функция для добавления меню настроек
    const addSettingsMenu = (title, moduleName) => {
      try {
        if (typeof eval(moduleName) !== 'undefined') {
          menu.addSubMenu(
            ui.createMenu(title)
              .addItem('Загрузить', `${moduleName}.doLoad`)
              .addItem('Сохранить', `${moduleName}.doSave`)
              .addItem('Частично', `${moduleName}.doPartial`)
              .addItem('Заменить', `${moduleName}.doReplace`)
          );
        }
      } catch(e) {}
    };

    addSettingsMenu('Настройка категорий', 'Categories');
    addSettingsMenu('Настройка счетов', 'Accounts');

  } catch(e) {
    Logger.log('Ошибка при создании подменю: ' + e.message);
  }

  // Всегда добавляем меню, даже если нет подпунктов
  menu.addToUi();
}

//═══════════════════════════════════════════════════════════════════════════
// СИНХРОНИЗАЦИЯ
// Функции для полной синхронизации и обновления справочников
//═══════════════════════════════════════════════════════════════════════════
const fullSyncHandlers = [];

function doFullSync() {
  try {
    const json = zmData.RequestData();

    doUpdateDictionaries();
    fullSyncHandlers.forEach(f => {
      try {
        f(json);
      } catch (error) {
        Logger.log(`Ошибка в обработчике полной синхронизации: ${error.toString()}`);
      }
    });
  } catch (error) {
    Logger.log("Ошибка при полной синхронизации: " + error.toString());
  }
}

function doUpdateDictionaries() {
  try {
    const requestPayload = ["account", "merchant", "instrument", "tag", "user"];
    const json = zmData.RequestForceFetch(requestPayload);
    Dictionaries.updateDictionaries(json);
    Dictionaries.saveDictionariesToSheet();
    Logger.log("Справочники обновлены");
  } catch (error) {
    Logger.log("Ошибка при обновлении справочников: " + error.toString());
  }
}

//═══════════════════════════════════════════════════════════════════════════
// SHEET HELPER
// Утилиты для работы с Google Sheets
//═══════════════════════════════════════════════════════════════════════════
const sheetHelper = (function () {
  const o = {};

  // Проверка, что активный лист - это лист настроек
  function isSettingsSheetActive() {
    const activeSheet = SpreadsheetApp.getActiveSheet();
    return activeSheet && activeSheet.getName() === Settings.SHEETS.SETTINGS.NAME;
  }

  o.Get = function (sheetName) {
    try {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet && !isSettingsSheetActive()) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(sheetName);
      }
      return sheet;
    } catch (error) {
      Logger.log(`Ошибка при получении листа ${sheetName}: ${error.toString()}`);
      return null;
    }
  };

  o.GetRange = function (sheetName, range) {
    try {
      const sheet = o.Get(sheetName);
      return sheet ? sheet.getRange(range) : null;
    } catch (error) {
      Logger.log(`Ошибка при получении диапазона ${range} на листе ${sheetName}: ${error.toString()}`);
      return null;
    }
  };

  o.GetRangeValues = function (sheetName, range) {
    try {
      const rangeObj = o.GetRange(sheetName, range);
      return rangeObj ? rangeObj.getValues() : null;
    } catch (error) {
      Logger.log(`Ошибка при получении значений диапазона ${range} на листе ${sheetName}: ${error.toString()}`);
      return null;
    }
  };

  o.GetCellValue = function (sheetName, cell) {
    try {
      const values = o.GetRangeValues(sheetName, cell);
      return values && values.length > 0 && values[0].length > 0 ? values[0][0] : null;
    } catch (error) {
      Logger.log(`Ошибка при получении значения ячейки ${cell} на листе ${sheetName}: ${error.toString()}`);
      return null;
    }
  };

  o.WriteData = function (sheetName, data) {
    try {
      const sheet = o.Get(sheetName);
      if (sheet) {
        sheet.clearContents();
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
    } catch (error) {
      Logger.log(`Ошибка при записи данных на лист ${sheetName}: ${error.toString()}`);
    }
  };

  o.GetSheetFromSettings = function (cellKey) {
    try {
      const settingsSheet = o.Get(Settings.SHEETS.SETTINGS.NAME);
      const sheetName = settingsSheet.getRange(Settings.SHEETS.SETTINGS.CELLS[cellKey]).getValue();
      if (!sheetName || !sheetName.trim()) {
        throw new Error(`Некорректное имя листа в ячейке ${Settings.SHEETS.SETTINGS.CELLS[cellKey]}`);
      }
      const sheet = o.Get(sheetName.trim());
      if (!sheet) {
        throw new Error(`Лист ${sheetName.trim()} не найден`);
      }
      return sheet;
    } catch (error) {
      Logger.log(`Ошибка при получении листа из настроек (${cellKey}): ${error.toString()}`);
      return null;
    }
  };

  return o;
})();

//═══════════════════════════════════════════════════════════════════════════
// ZM SETTINGS
// Управление настройками Zen Money (токен, timestamp)
//═══════════════════════════════════════════════════════════════════════════
const zmSettings = {
  getToken: function () {
    return sheetHelper.GetCellValue(Settings.SHEETS.SETTINGS.NAME, Settings.SHEETS.SETTINGS.CELLS.TOKEN);
  },

  getTimestamp: function () {
    return sheetHelper.GetCellValue(Settings.SHEETS.SETTINGS.NAME, Settings.SHEETS.SETTINGS.CELLS.TIMESTAMP);
  },

  setTimestamp: function (value) {
    try {
      const sheet = sheetHelper.Get(Settings.SHEETS.SETTINGS.NAME);
      if (sheet) {
        sheet.getRange(Settings.SHEETS.SETTINGS.CELLS.TIMESTAMP).setValue(value);
      }
    } catch (error) {
      Logger.log(`Ошибка при установке timestamp: ${error.toString()}`);
    }
  }
};

//═══════════════════════════════════════════════════════════════════════════
// ZM DATA
// Функции для взаимодействия с API Zen Money
//═══════════════════════════════════════════════════════════════════════════
const zmData = (function () {
  function currentTimestamp() {
    return Math.round((new Date()).getTime() / 1000);
  }

  const o = {};

  o.Request = function (data) {
    try {
      const params = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + zmSettings.getToken(),
        },
        'payload': JSON.stringify(data)
      };
      const res = UrlFetchApp.fetch("https://api.zenmoney.ru/v8/diff/", params);
      const content = res.getContentText();
      const json = JSON.parse(content);

      return json;
    } catch (err) {
      Logger.log("Ошибка при запросе к API Zen Money: " + err.toString());
      return {};
    }
  };

  o.RequestData = function () {
    try {
      const ts = currentTimestamp();
      var json = o.Request({
        'currentClientTimestamp': ts,
        'serverTimestamp': 0,
      });

      return json;
    } catch (error) {
      Logger.log("Ошибка при получении данных: " + error.toString());
      return {};
    }
  };

  o.RequestForceFetch = function (items) {
    try {
      const ts = currentTimestamp();
      var json = o.Request({
        'currentClientTimestamp': ts,
        'serverTimestamp': ts,
        'forceFetch': items,
      });

      return json;
    } catch (error) {
      Logger.log("Ошибка при принудительной загрузке данных: " + error.toString());
      return {};
    }
  };

  return o;
})();
