const Export = (function () {
  const sheet = sheetHelper.GetSheetFromSettings('EXPORT_SHEET');
  if (!sheet) {
    Logger.log("Лист с данными не найден");
    SpreadsheetApp.getActive().toast('Лист с данными не найден', 'Ошибка');
    return;
  }

  const HEADERS = Settings.TRANSACTION_FIELDS.map(field => field.title);

  // COLUMNS генерируются из TRANSACTION_FIELDS (используем id)
  const COLUMNS = (function() {
    const columns = {};
    Settings.TRANSACTION_FIELDS.forEach((field, index) => {
      columns[field.id] = index;
    });
    return columns;
  })();

  // Вспомогательная функция для получения значения с расшифровкой по id поля
  function getDecodedValue(t, field) {  
    switch (field.id) {  
      case "tag":  
      case "tag1":  
      case "tag2":  
        if (!t.tag) return "";  
        const tagTitles = t.tag.map(tagId => {  
          const title = Dictionaries.getTagTitle(tagId) || "";  
          return title;  
        }).filter(Boolean);  
    
        const tagMode = Settings.TagMode; 
          
        if (tagMode === Settings.TAG_MODES.MULTIPLE_COLUMNS) {  
          // Режим отдельных столбцов  
          switch (field.id) {  
            case "tag": return tagTitles[0] || "";  
            case "tag1": return tagTitles[1] || "";  
            case "tag2": return tagTitles[2] || "";  
            default: return "";  
          }  
        } else {  
          // Режим одной строки  
          return field.id === "tag" ? tagTitles.join(" | ") : "";  
        }  
      case "incomeAccount":  
        return Dictionaries.getAccountTitle(t.incomeAccount) || "";  
      case "outcomeAccount":  
        return Dictionaries.getAccountTitle(t.outcomeAccount) || "";  
      case "user":  
        return Dictionaries.getUserLogin(t.user) || "";  
      case "incomeInstrument":  
        return Dictionaries.getInstrumentShortTitle(t.incomeInstrument) || "";  
      case "outcomeInstrument":  
        return Dictionaries.getInstrumentShortTitle(t.outcomeInstrument) || "";  
      case "merchant":  
        return Dictionaries.getMerchantTitle(t.merchant) || "";  
      case "opIncome":    
      case "opOutcome":  
        return t[field.id] === 0 ? "" : t[field.id];  
      default:  
        let val = t[field.id];  
        if (val === null) return "";  
        if (typeof val === "object") val = JSON.stringify(val);  
        return val !== undefined ? val : "";  
    }  
  }

  // Вспомогательная функция для форматирования транзакции под лист изменений с учётом порядка столбцов
  function formatTransactionForChangesSheet(transaction) {
    return Settings.TRANSACTION_FIELDS.map(field => {
      return getDecodedValue(transaction, field);
    });
  }

  // Обработка данных при полном экспорте
  function prepareFullData(json) {
    if (!json.transaction || !Array.isArray(json.transaction)) {
      const msg = "В JSON нет объекта 'transaction' или он некорректный";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Ошибка');
      return;
    }

    sheet.clearContents();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    let transactions = json.transaction.filter(t => !t.deleted);
    if (transactions.length === 0) {
      const msg = "Нет доступных транзакций для экспорта";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Информация');
      return;
    }

    // Сортируем по дате по убыванию (от новых к старым)  
    transactions.sort((a, b) => {  
      const dateA = new Date(a.date).getTime() || 0;  
      const dateB = new Date(b.date).getTime() || 0;  
      return dateB - dateA;
    });

    // Обновляем справочники через модуль Dictionaries
    Dictionaries.updateDictionaries(json);

    // Формируем строки с расшифровками в нужном порядке столбцов
    const data = transactions.map(t => Settings.TRANSACTION_FIELDS.map(field => getDecodedValue(t, field)));

    // Записываем транзакции, начиная со второй строки (под заголовками)
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

    sheet.getRange(2, COLUMNS.deleted + 1, data.length, 1).insertCheckboxes();  
    sheet.getRange(2, COLUMNS.modified + 1, data.length, 1).insertCheckboxes();  
    const lastRow = sheet.getMaxRows();
    if (lastRow > data.length + 1) {  
      sheet.getRange(data.length + 2, COLUMNS.deleted + 1, lastRow - data.length - 1, 1).clearContent().clearDataValidations();  
      sheet.getRange(data.length + 2, COLUMNS.modified + 1, lastRow - data.length - 1, 1).clearContent().clearDataValidations();  
    }

    Logger.log("Экспорт с расшифровками завершён");
    SpreadsheetApp.getActive().toast(`Экспорт завершен. Загружено ${data.length} операций.`, 'Экспорт завершен');
  }

  // Общая функция для запуска полного экспорта 
  function doFullExport() {
    try {
      SpreadsheetApp.getActive().toast('Начинаем полный экспорт...', 'Экспорт');
      
      const requestPayload = ["transaction", "account", "merchant", "instrument", "tag", "user"];
      const json = zmData.RequestForceFetch(requestPayload);
      Logs.logApiCall("FETCH_TRANSACTIONS", requestPayload, JSON.stringify(json));
      prepareFullData(json);
      if (json.serverTimestamp) {
        zmSettings.setTimestamp(json.serverTimestamp);
      }
    } catch (error) {
      const errorMsg = "Ошибка при экспорте: " + error.toString();
      Logger.log(errorMsg);
      Logs.logApiCall("FETCH_TRANSACTIONS_ERROR", {}, error.toString());
      SpreadsheetApp.getActive().toast(errorMsg, 'Ошибка');
    }
  }

  let newTimestamp = 0;

  function prepareChangesSheet() {
    try {
      SpreadsheetApp.getActive().toast('Получаем изменения с сервера...', 'Подготовка изменений');

      // Получаем лист для изменений
      const changesSheet = sheetHelper.Get(Settings.SHEETS.CHANGES);
      changesSheet.clearContents();
      changesSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      changesSheet.getRange(1, 1, changesSheet.getMaxRows(), changesSheet.getMaxColumns()).setBackground(null);

      const token = zmSettings.getToken().trim();
      if (!token) {
        const msg = "Ошибка: Токен не найден в Settings!B1";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Ошибка');
        return;
      }
      const lastTimestamp = zmSettings.getTimestamp() || 0;

      const requestPayload = {
        currentClientTimestamp: Math.floor(Date.now() / 1000),
        serverTimestamp: lastTimestamp
      };
      const json = zmData.Request(requestPayload);

      Logs.logApiCall("FETCH_DIFF", requestPayload, JSON.stringify(json));
      if (!json || Object.keys(json).length === 0) {
        const msg = "Ошибка при получении diff: пустой ответ";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Ошибка');
        return;
      }

      Dictionaries.loadDictionariesFromSheet()

      if (!json.transaction || !Array.isArray(json.transaction)) {
        const msg = "В JSON нет объекта 'transaction' или он некорректный";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Ошибка');
        return;
      }

      // Разделяем транзакции на удалённые и не удалённые
      const deletedTransactions = json.transaction.filter(t => t.deleted);
      const activeTransactions = json.transaction.filter(t => !t.deleted);

      if (activeTransactions.length === 0 && deletedTransactions.length === 0) {
        const msg = "Нет доступных транзакций для обновления";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Информация');
        return;
      }

      // Заголовки для листа изменений
      const changeHeaders = HEADERS;
      changesSheet.getRange(1, 1, 1, changeHeaders.length).setValues([changeHeaders]);

      // Получаем текущие данные с листа с данными
      const data = sheet.getDataRange().getValues();
      const idIndex = COLUMNS.id;
      const updatedRowMap = new Map();
      for (let i = 1; i < data.length; i++) {
        const rowId = data[i][idIndex];
        if (rowId) updatedRowMap.set(rowId, i + 1);
      }

      // Разделяем активные транзакции на новые и существующие
      const newTransactions = [];
      const existingTransactions = [];

      activeTransactions.forEach(t => {
        if (updatedRowMap.has(t.id)) {
          existingTransactions.push(t);
        } else {
          newTransactions.push(t);
        }
      });

      // Формируем данные и типы изменений
      const changesData = [];
      const changeTypes = [];

      // Добавляем удалённые транзакции
      deletedTransactions.forEach(t => {  
        changesData.push(formatTransactionForChangesSheet(t));  
        if (updatedRowMap.has(t.id)) {  
          changeTypes.push("К удалению");  // Есть в листе 
        } else {  
          changeTypes.push("Удалена");     // Уже удалена  
        }  
      });

      // Добавляем существующие (изменённые) транзакции
      existingTransactions.forEach(t => {
        changesData.push(formatTransactionForChangesSheet(t));
        changeTypes.push("Изменена");
      });

      // Добавляем новые транзакции
      newTransactions.forEach(t => {  
        changesData.push(formatTransactionForChangesSheet(t));  
        changeTypes.push("Новая");  
      });

      if (changesData.length > 0) {
        changesSheet.getRange(2, 1, changesData.length, changeHeaders.length).setValues(changesData);

        const NewColor = Settings.EXPORT.COLORS.NEW;
        const WasColor = Settings.EXPORT.COLORS.WAS_NEW; // не работает в синхронизациях
        const ModColor = Settings.EXPORT.COLORS.MODIFIED;
        const ToDoColor =Settings.EXPORT.COLORS.TO_DELETE;
        const DelColor = Settings.EXPORT.COLORS.DELETED;

        const bgColors = changeTypes.map(type => {
          switch (type) {
            case "Новая": return Array(changeHeaders.length).fill(NewColor);
            case "Бывшая новая": return Array(changeHeaders.length).fill(WasColor);
            case "Изменена": return Array(changeHeaders.length).fill(ModColor);
            case "К удалению": return Array(changeHeaders.length).fill(ToDoColor);
            case "Удалена": return Array(changeHeaders.length).fill(DelColor);
            default: return Array(changeHeaders.length).fill("#ffffff");
          }
        });

        changesSheet.getRange(2, 1, changesData.length, changeHeaders.length).setBackgrounds(bgColors);
      }

      Logger.log(`Подготовлено ${changesData.length} изменений на листе изменений`);
      SpreadsheetApp.getActive().toast(
        `Подготовлено изменений: ${changesData.length}\n` +
        `Новых: ${newTransactions.length}\n` +
        `Измененных: ${existingTransactions.length}\n` +
        `Удаленных: ${deletedTransactions.length}`,
        'Изменения подготовлены'
      );

      if (json.serverTimestamp) newTimestamp = json.serverTimestamp;

    } catch (error) {
      const errorMsg = "Ошибка при подготовке листа изменений: " + error.toString();
      Logger.log(errorMsg);
      SpreadsheetApp.getActive().toast(errorMsg, 'Ошибка');
    }
  }

  // Функция для применения изменений с листа изменений на лист с данными
  function applyChangesToDataSheet() {
    try {
      if (!sheet) return;

      const changesSheet = sheetHelper.Get(Settings.SHEETS.CHANGES);
      const changeHeaders = HEADERS;

      // Получаем все данные с листа изменений
      const changesData = changesSheet.getDataRange().getValues();
      if (changesData.length < 2) {
        const msg = "Нет данных для применения изменений";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Информация');
        return;
      }

      // Получаем текущие данные с листа с данными
      let data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        const msg = "Ошибка: нет данных или заголовков в листе данных";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Ошибка');
        return;
      }

      const idIndex = COLUMNS.id;
      if (idIndex === undefined) {
        const msg = "Ошибка: колонка 'id' не найдена в таблице с данными";
        Logger.log(msg);
        SpreadsheetApp.getActive().toast(msg, 'Ошибка');
        return;
      }

      // Создаём карту id -> номер строки в листе с данными
      const rowMap = new Map();
      for (let i = 1; i < data.length; i++) {
        const rowId = data[i][idIndex];
        if (rowId) rowMap.set(rowId.toString().trim(), i + 1);
      }

      // Массив для хранения индексов строк, которые нужно удалить
      const rowsToDelete = [];

      // Собираем индексы строк для удаления по листу изменений
      for (let i = 1; i < changesData.length; i++) {
        const row = changesData[i];
        const id = row[idIndex];
        const deleted = row[COLUMNS.deleted];
        if ((deleted === true || deleted === "TRUE") && id && rowMap.has(id.toString().trim())) {
          rowsToDelete.push(rowMap.get(id.toString().trim()));
        }
      }

      // Удаляем строки с конца, чтобы не сбивать индексы при удалении
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(rowIdx => {
        sheet.deleteRow(rowIdx);
        Logger.log(`Удалена строка ${rowIdx}`);
      });

      // После удаления заново считываем данные и строим обновлённую карту строк
      data = sheet.getDataRange().getValues();
      const updatedRowMap = new Map();
      for (let i = 1; i < data.length; i++) {
        const rowId = data[i][idIndex];
        if (rowId) updatedRowMap.set(rowId.toString().trim(), i + 1);
      }

      // Разделяем изменения на новые и существующие транзакции
      const newRows = [];
      const updatedRows = [];

      for (let i = 1; i < changesData.length; i++) {
        const row = changesData[i];
        const id = row[idIndex];
        const deleted = row[COLUMNS.deleted];
        if (deleted === true || deleted === "TRUE") {
          // Уже удалили, пропускаем
          continue;
        }
        if (id && updatedRowMap.has(id.toString().trim())) {
          updatedRows.push({ rowIndex: updatedRowMap.get(id.toString().trim()), rowData: row });
        } else {
          newRows.push(row);
        }
      }

      // Обновляем существующие строки, сравнивая и записывая только изменённые ячейки
      updatedRows.forEach(({ rowIndex, rowData }) => {
        for (let col = 0; col < rowData.length; col++) {
          if (data[rowIndex - 1][col] !== rowData[col]) {
            sheet.getRange(rowIndex, col + 1).setValue(rowData[col]);
          }
        }
      });

      // Добавляем новые строки в конец листа
      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }

      Logger.log(`Применено изменений: добавлено ${newRows.length}, обновлено ${updatedRows.length}, удалено ${rowsToDelete.length}`);
      SpreadsheetApp.getActive().toast(
        `Изменения применены:\n` +
        `Добавлено: ${newRows.length}\n` +
        `Обновлено: ${updatedRows.length}\n` +
        `Удалено: ${rowsToDelete.length}`,
        'Обновление завершено'
      );
    } catch (error) {
      const errorMsg = "Ошибка при применении изменений: " + error.toString();
      Logger.log(errorMsg);
      SpreadsheetApp.getActive().toast(errorMsg, 'Ошибка');
    }
  }

  // Инкрементальный экспорт
  function doIncrementalExport() {
    try {
      prepareChangesSheet();
      applyChangesToDataSheet();
      if (newTimestamp) zmSettings.setTimestamp(newTimestamp);
    } catch (error) {
      const errorMsg = "Ошибка при инкрементальном экспорте: " + error.toString();
      Logger.log(errorMsg);
      SpreadsheetApp.getActive().toast(errorMsg, 'Ошибка');
    }
  }

  // Добавляем обработчик полной синхронизации
  fullSyncHandlers.push(prepareFullData);

  // Регистрация функций в меню
  function createMenu() {
    const ui = SpreadsheetApp.getUi();
    const subMenu = ui.createMenu("Export")
      .addItem("Full Export", "Export.doFullExport")
      .addSeparator()
      .addItem("Incremental Export", "Export.doIncrementalExport")
      .addItem("Prepare Changes Sheet", "Export.prepareChangesSheet")
      .addItem("Apply Changes to Data Sheet", "Export.applyChangesToDataSheet")
    gsMenu.addSubMenu(subMenu);
  }

  // Вызываем создание меню при инициализации модуля
  createMenu();

  return {
    doFullExport,
    doIncrementalExport,
    prepareChangesSheet,
    applyChangesToDataSheet,
    doUpdateDictionaries
  };
})();
