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

function handleCategoryReplacement(oldCategoryId, newCategoryId) {
  try {
    if (typeof Categories === 'undefined' || 
        typeof Categories.handleCategoryReplacement !== 'function') {
      throw new Error('Модуль замены категорий не инициализирован');
    }
    return Categories.handleCategoryReplacement(oldCategoryId, newCategoryId);
  } catch (e) {
    console.error('Ошибка в handleCategoryReplacement:', e);
    return {
      error: true,
      message: e.message
    };
  }
}

//═══════════════════════════════════════════════════════════════════════════
// МЕНЮ
// Функции для создания и управления меню Дзен Мани
//═══════════════════════════════════════════════════════════════════════════

function createMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const mainMenu = ui.createMenu('Дзен Мани')
      .addItem('Полная Синхронизация', 'doFullSync')
      .addItem('Обновить Словари', 'doUpdateDictionaries')
      .addSeparator();

    if (typeof Export !== 'undefined') {
      addSubMenu(mainMenu, 'Экспорт', [
        { name: "Полный экспорт", func: "doFullExport" },
        { name: "Инкрементальный экспорт", func: "doIncrementalExport" },
        { name: "Подготовить лист изменений", func: "prepareChangesSheet" },
        { name: "Применить изменения", func: "applyChangesToDataSheet" }
      ], [], 'Export');
    }

    if (typeof Import !== 'undefined') {
      addSubMenu(mainMenu, 'Импорт', [
        { name: "Частичный импорт", func: "doUpdate" }
      ], [Tracking, Validation], 'Import');
    }

    if (typeof Categories !== 'undefined') {
      addSubMenu(mainMenu, 'Настройка категорий', [
        { name: "Загрузить", func: "doLoad" },
        { name: "Сохранить", func: "doSave" },
        { name: "Частично", func: "doPartial" },
        { name: "Заменить", func: "doReplace" }
      ], [Categories], 'Categories');
    }

    if (typeof Accounts !== 'undefined') {
      addSubMenu(mainMenu, 'Настройка счетов', [
        { name: "Загрузить", func: "doLoad" },
        { name: "Сохранить", func: "doSave" },
        { name: "Частично", func: "doPartial" },
        { name: "Заменить", func: "doReplace" }
      ], [], 'Accounts');
    }

    if (typeof Merchants !== 'undefined') {
      addSubMenu(mainMenu, 'Настройка мест', [
        { name: "Загрузить", func: "doLoad" },
        { name: "Сохранить", func: "doSave" },
        { name: "Частично", func: "doPartial" },
        { name: "Заменить", func: "doReplace" }
      ], [], 'Merchants');
    }

    mainMenu.addToUi();
  } catch (error) {
    Logger.log("Ошибка при создании меню: " + error.toString());
  }
}

function addSubMenu(mainMenu, menuName, items, extraModules = [], moduleName = null) {
  const subMenu = SpreadsheetApp.getUi().createMenu(menuName);
  const moduleNameForPath = moduleName || menuName;

  items.forEach(item => {
    const functionPath = `${moduleNameForPath}.${item.func}`;
    subMenu.addItem(item.name, functionPath);
  });

  extraModules.forEach(extraModule => {
    if (typeof extraModule !== 'undefined' && typeof extraModule.addMenuItems === 'function') {
      extraModule.addMenuItems(subMenu);
    }
  });

  mainMenu.addSubMenu(subMenu);
}

/*
// Конфигурация подменю
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Дзен Мани')
    .addItem('Полная Синхронизация', 'doFullSync')
    .addItem('Обновить Словари', 'doUpdateDictionaries')
    .addSeparator();

  // Конфигурация подменю с жестко заданными именами модулей
  const menuConfig = [
    {
      name: 'Экспорт',
      module: Export,
      items: [
        ["Полный экспорт", "Export.doFullExport"],
        ["Инкрементальный экспорт", "Export.doIncrementalExport"],
        ["Подготовить лист изменений", "Export.prepareChangesSheet"],
        ["Применить изменения", "Export.applyChangesToDataSheet"]
      ]
    },
    {
      name: 'Импорт',
      module: Import,
      items: [["Частичный импорт", "Import.doUpdate"]],
      extra: [Tracking, Validation]
    },
    {
      name: 'Настройка категорий',
      module: Categories,
      items: [
        ["Загрузить", "Categories.doLoad"],
        ["Сохранить", "Categories.doSave"],
        ["Частично", "Categories.doPartial"],
        ["Заменить", "Categories.doReplace"]
      ],
      extra: [Categories]
    },
    {
      name: 'Настройка счетов',
      module: Accounts,
      items: [
        ["Загрузить", "Accounts.doLoad"],
        ["Сохранить", "Accounts.doSave"],
        ["Частично", "Accounts.doPartial"],
        ["Заменить", "Accounts.doReplace"]
      ]
    },
    {
      name: 'Настройка мест',
      module: Merchants,
      items: [
        ["Загрузить", "Merchants.doLoad"],
        ["Сохранить", "Merchants.doSave"],
        ["Частично", "Merchants.doPartial"],
        ["Заменить", "Merchants.doReplace"]
      ]
    }
  ];

  // Динамическое создание меню
  menuConfig.forEach(config => {
    if (config.module) {
      const subMenu = ui.createMenu(config.name);
      
      // Добавляем основные пункты (имена функций уже содержат имя модуля)
      config.items.forEach(([name, funcPath]) => {
        subMenu.addItem(name, funcPath);
      });
      
      // Добавляем дополнительные пункты
      config.extra?.forEach(module => {
        if (module?.addMenuItems) module.addMenuItems(subMenu);
      });
      
      menu.addSubMenu(subMenu);
    }
  });

  menu.addToUi();
}*/

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
      if (!settingsSheet) {
        throw new Error("Лист настроек не найден");
      }
      
      const cellValue = settingsSheet.getRange(Settings.SHEETS.SETTINGS.CELLS[cellKey]).getValue();
      const sheetName = cellValue ? cellValue.toString().trim() : null;
      
      if (!sheetName) {
        throw new Error(`Некорректное имя листа в ячейке ${Settings.SHEETS.SETTINGS.CELLS[cellKey]}`);
      }
      
      const sheet = o.Get(sheetName);
      if (!sheet) {
        throw new Error(`Лист ${sheetName} не найден`);
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
