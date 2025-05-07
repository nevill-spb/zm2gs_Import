const fullSyncHandlers = [];

function onOpen(e) {  
  createMenu();  
}

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
      
    mainMenu.addSubMenu(importMenu);  
  }  
  
  if (typeof SetupCategories !== 'undefined') {  
    const categoriesMenu = ui.createMenu("Setup categories")  
      .addItem("Load", "SetupCategories.doLoad")  
      .addItem("Save", "SetupCategories.doSave")  
      .addItem("Partial", "SetupCategories.doPartial");  
    mainMenu.addSubMenu(categoriesMenu);  
  }  
  
  if (typeof Validation !== 'undefined') {  
    const validationMenu = ui.createMenu("Validation")  
      .addItem("Setup Validation", "Validation.setupValidation")  
      .addItem("Clear All Validation", "Validation.clearAllValidation");  
    mainMenu.addSubMenu(validationMenu);  
  }  
  
  mainMenu.addToUi();  
}

/*const gsMenu = SpreadsheetApp.getUi()
.createMenu((typeof paramMenuTitleMain !== 'undefined') ? paramMenuTitleMain : 'Zen Money')
.addItem((typeof paramMenuTitleFullSync !== 'undefined') ? paramMenuTitleFullSync : 'Full sync', 'doFullSync')
.addItem((typeof paramMenuTitleFullSync !== 'undefined') ? paramMenuTitleFullSync : 'Update Dictionaries', 'doUpdateDictionaries')
.addSeparator();

function onOpen() {
  gsMenu.addToUi();
}*/

// Полная синхронизация
function doFullSync() {
  const json = zmData.RequestData();

  doUpdateDictionaries()
  fullSyncHandlers.forEach(f => f(json));
}

// Обновление справочников
function doUpdateDictionaries() {
  try {
    const requestPayload = ["account", "merchant", "instrument", "tag", "user"];
    const json = zmData.RequestForceFetch(requestPayload);
    Dictionaries.updateDictionaries(json);
    Dictionaries.saveDictionariesToSheet();  // Записываем обновлённые словари на лист
    Logger.log("Справочники обновлены");
  } catch (error) {
    Logger.log("Ошибка при обновлении справочников: " + error.toString());
  }
}

const sheetHelper = (function () {
  const o = {};

  o.Get = function (sheetName) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet === null) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      sheet.setName(sheetName);
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
