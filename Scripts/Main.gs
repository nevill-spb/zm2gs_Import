// Функции для триггеров
function onEdit(e) {  
  if (typeof Tracking !== 'undefined') {  
    Tracking.handleOnEdit(e);  
  }  
}  
  
function onChange(e) {  
  if (typeof Tracking !== 'undefined') {  
    Tracking.handleOnChange(e);  
  }  
}

function onOpen(e) {  
  createMenu();  
}

const fullSyncHandlers = [];

function createMenu() {  
  const ui = SpreadsheetApp.getUi();  
  const mainMenu = ui.createMenu('Zen Money')  
    .addItem('Full sync', 'doFullSync')  
    .addItem('Update Dictionaries', 'doUpdateDictionaries')  
    .addSeparator();  
  
  // Проверяем наличие и добавляем подменю для каждого модуля  
  if (typeof Export !== 'undefined') {  
    const exportMenu = ui.createMenu("Export")  
      .addItem("Full Export", "Export.doFullExport")  
      .addSeparator()  
      .addItem("Incremental Export", "Export.doIncrementalExport")  
      .addItem("Prepare Changes Sheet", "Export.prepareChangesSheet")  
      .addItem("Apply Changes to Data Sheet", "Export.applyChangesToDataSheet");  
    mainMenu.addSubMenu(exportMenu);  
  }  
  
  if (typeof Import !== 'undefined') {  
    const importMenu = ui.createMenu("Import")  
      .addItem("Partial Import", "Import.doUpdate");  
      
    if (typeof Tracking !== 'undefined' && Tracking.addTrackingMenuItems) {  
      Tracking.addTrackingMenuItems(importMenu);  
    }  
    if (typeof Validation !== 'undefined' && Validation.addValidationMenuItems) {  
      Validation.addValidationMenuItems(importMenu);  
    }  
    mainMenu.addSubMenu(importMenu);  
  }  
  
  if (typeof Categories !== 'undefined') {  
    const categoriesMenu = ui.createMenu("Setup categories")  
      .addItem("Load", "Categories.doLoad")  
      .addItem("Save", "Categories.doSave")  
      .addItem("Partial", "Categories.doPartial");  
    mainMenu.addSubMenu(categoriesMenu);  
  }  

  if (typeof Accounts !== 'undefined') {  
    const accountsMenu = ui.createMenu("Setup accounts")  
      .addItem("Load", "Accounts.doLoad")  
      .addItem("Save", "Accounts.doSave")  
      .addItem("Partial", "Accounts.doPartial");  
    mainMenu.addSubMenu(accountsMenu);
  }  

  mainMenu.addToUi();  
}

// Полная синхронизация
function doFullSync() {
  const json = zmData.RequestData();

  doUpdateDictionaries();
  fullSyncHandlers.forEach(f => f(json));
}

// Обновление справочников
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

const sheetHelper = (function () {  
  const o = {};  
  
  // Проверка, что активный лист - это лист настроек  
  function isSettingsSheetActive() {  
    const activeSheet = SpreadsheetApp.getActiveSheet();  
    return activeSheet && activeSheet.getName() === Settings.SHEETS.SETTINGS.NAME;  
  }  
  
  o.Get = function (sheetName) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet === null && !isSettingsSheetActive()) {
      try {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(sheetName);
      } catch (error) {
        return null;
      }
    }
    
    return sheet;
  };

  o.GetRange = function (sheetName, range) {
    return o.Get(sheetName).getRange(range);
  };

  o.GetRangeValues = function (sheetName, range) {
    return o.GetRange(sheetName, range).getValues();
  };

  o.GetCellValue = function (sheetName, cell) {
    const values = o.GetRangeValues(sheetName, cell);

    return values[0][0];
  };

  o.WriteData = function (sheetName, data) {
    const sheet = o.Get(sheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  };

  o.GetSheetFromSettings = function (cellKey) {  
    const settingsSheet = o.Get(Settings.SHEETS.SETTINGS.NAME);  
    const sheetName = settingsSheet.getRange(Settings.SHEETS.SETTINGS.CELLS[cellKey]).getValue();  
    if (!sheetName || !sheetName.trim()) {  
      throw new Error(`Некорректное имя листа в ячейке ${Settings.SHEETS.SETTINGS.CELLS[cellKey]}`);  
    }  
    return o.Get(sheetName.trim());  
  };

  return o;
})();

const zmSettings = {
  getToken: function() {
    return sheetHelper.GetCellValue(Settings.SHEETS.SETTINGS.NAME, Settings.SHEETS.SETTINGS.CELLS.TOKEN);
  },

  getTimestamp: function() {
    return sheetHelper.GetCellValue(Settings.SHEETS.SETTINGS.NAME, Settings.SHEETS.SETTINGS.CELLS.TIMESTAMP);
  },

  setTimestamp: function(value) {
    const sheet = sheetHelper.Get(Settings.SHEETS.SETTINGS.NAME);
    sheet.getRange(Settings.SHEETS.SETTINGS.CELLS.TIMESTAMP).setValue(value);
  }
};

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
      Logger.log("Error getting data");
      Logger.log(err);

      return {};
    }
  };

  o.RequestData = function () {
    const ts = currentTimestamp();
    var json = o.Request({
      'currentClientTimestamp': ts,
      'serverTimestamp': 0,
    });

    return json;
  };

  o.RequestForceFetch = function (items) {
    const ts = currentTimestamp();
    var json = o.Request({
      'currentClientTimestamp': ts,
      'serverTimestamp': ts,
      'forceFetch': items,
    });
  
    return json;
  };

  return o;
})();
