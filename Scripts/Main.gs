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
  try {
    const ui = SpreadsheetApp.getUi();
    const mainMenu = ui.createMenu('Zen Money')
      .addItem('Full sync', 'doFullSync')
      .addItem('Update Dictionaries', 'doUpdateDictionaries')
      .addSeparator();

    // Добавляем подменю для каждого модуля
    addSubMenu(mainMenu, 'Export', Export, [
      { name: "Full Export", func: "doFullExport" },
      { name: "Incremental Export", func: "doIncrementalExport" },
      { name: "Prepare Changes Sheet", func: "prepareChangesSheet" },
      { name: "Apply Changes to Data Sheet", func: "applyChangesToDataSheet" }
    ]);

    addSubMenu(mainMenu, 'Import', Import, [
      { name: "Partial Import", func: "doUpdate" }
    ], [Tracking, Validation]);

    addSubMenu(mainMenu, 'Setup categories', Categories, [
      { name: "Load", func: "doLoad" },
      { name: "Save", func: "doSave" },
      { name: "Partial", func: "doPartial" }
    ]);

    addSubMenu(mainMenu, 'Setup accounts', Accounts, [
      { name: "Load", func: "doLoad" },
      { name: "Save", func: "doSave" },
      { name: "Partial", func: "doPartial" }
    ]);

    mainMenu.addToUi();
  } catch (error) {
    Logger.log("Ошибка при создании меню: " + error.toString());
  }
}

function addSubMenu(mainMenu, menuName, module, items, extraModules = []) {
  if (typeof module !== 'undefined') {
    const subMenu = SpreadsheetApp.getUi().createMenu(menuName);
    items.forEach(item => subMenu.addItem(item.name, `${module.name}.${item.func}`));

    extraModules.forEach(extraModule => {
      if (typeof extraModule !== 'undefined' && typeof extraModule.addMenuItems === 'function') {
        extraModule.addMenuItems(subMenu);
      }
    });

    mainMenu.addSubMenu(subMenu);
  }
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
