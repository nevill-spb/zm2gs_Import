const Accounts = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // ИНИЦИАЛИЗАЦИЯ
  //═══════════════════════════════════════════════════════════════════════════
  const sheet = sheetHelper.GetSheetFromSettings('ACCOUNTS_SHEET');
  if (!sheet) {
    Logger.log("Лист с аккаунтами не найден");
    SpreadsheetApp.getActive().toast('Лист с аккаунтами не найден', 'Ошибка');
    return;
  }

  // Индексы полей по id для удобства доступа
  const fieldIndex = {};
  Settings.ACCOUNT_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  //═══════════════════════════════════════════════════════════════════════════
  
  // Парсеры
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  function parseNumber(value, defaultValue = 0) {
    if (value === "" || value == null) return defaultValue;
    return Number(value);
  }

  function parseDateFromString(dateStr) {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? "" : d;
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ДАННЫМИ
  //═══════════════════════════════════════════════════════════════════════════
  
  // Построение строки данных по полям из Settings
  function buildRow(account) {
    return Settings.ACCOUNT_FIELDS.map(field => {
      switch (field.id) {
        case "inBalance":
        case "private":
        case "savings":
        case "archive":
        case "enableCorrection":
        case "enableSMS":
        case "capitalization":
          return parseBool(account[field.id]) ? "TRUE" : "FALSE";
        case "delete":
        case "modify":
          return false; // чекбоксы пустые
        case "instrument":
          return Dictionaries.getInstrumentShortTitle(account.instrument) || account.instrument || "";
        case "user":
          return Dictionaries.getUserLogin(account.user) || account.user || "";
        case "startDate":
          if (typeof account.startDate === "string") return parseDateFromString(account.startDate);
          if (account.startDate != null) return new Date(account.startDate * 1000);
          return "";
        default:
          return account[field.id] != null ? account[field.id] : "";
      }
    });
  }

  // Подготовка данных аккаунтов для записи в лист
  function prepareData(json) {
    if (!('account' in json)) {
      SpreadsheetApp.getActive().toast('Нет данных об аккаунтах', 'Предупреждение');
      return;
    }

    // Сортируем аккаунты по названию + id для стабильности
    const sortedAccounts = json.account.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    const data = sortedAccounts.map(acc => buildRow(acc));
    writeDataToSheet(data);
    SpreadsheetApp.getActive().toast(`Загружено ${data.length} счетов`, 'Информация');
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ЛИСТОМ
  //═══════════════════════════════════════════════════════════════════════════

  // Запись данных в лист
  function writeDataToSheet(data) {
    // Очистка листа перед записью
    sheet.clearContents();
    sheet.clearFormats();

    // Заголовки из Settings
    const headers = Settings.ACCOUNT_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      // Список id булевых полей, для которых нужны чекбоксы
      const boolFields = [
        "inBalance", "private", "savings", "archive", 
        "enableCorrection", "enableSMS", "capitalization",
        "delete", "modify"
      ];

      // Вставляем чекбоксы по каждой колонке
      boolFields.forEach(fieldId => {
        const colIndex = fieldIndex[fieldId];
        if (colIndex === undefined) return;
        sheet.getRange(2, colIndex + 1, data.length, 1).insertCheckboxes();
      });

      // Очищаем лишние чекбоксы и валидации
      clearExtraRows(data.length, boolFields);
    }
  }

  // Очистка лишних строк (только для чекбокс-колонок)
  function clearExtraRows(dataLength, boolFields) {
    const boolCols = boolFields.map(id => fieldIndex[id]).filter(i => i !== undefined);
    if (boolCols.length > 0) {
      const minCol = Math.min(...boolCols);
      const maxCol = Math.max(...boolCols);
      const lastRow = sheet.getMaxRows();
      const startRow = dataLength + 2;
      const numRows = lastRow - startRow + 1;
      if (numRows > 0) {
        sheet.getRange(startRow, minCol + 1, numRows, maxCol - minCol + 1)
          .clearContent()
          .clearDataValidations();
      }
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ СОЗДАНИЯ И ОБНОВЛЕНИЯ АККАУНТОВ
  //═══════════════════════════════════════════════════════════════════════════

  // Создание объекта аккаунта из строки данных
  function createAccountFromRow(row, i, ts, accounts, idMap, userReverseMap, instrumentReverseMap) {
    const oldAccountId = row[fieldIndex.id];
    const title = row[fieldIndex.title];
    if (!title || typeof title !== 'string' || title.trim() === '') return null;
    const user = userReverseMap[(row[fieldIndex.user])] || Settings.DefaultUserId;
    const deleteFlag = parseBool(row[fieldIndex.delete]);

    if (deleteFlag && oldAccountId) {
      return {
        deletion: {
          id: oldAccountId,
          object: 'account',
          stamp: ts,
          user: user
        },
      };
    }

    const newAccountId = idMap[oldAccountId || `new_${i}`];
    if (!newAccountId) return null;

    let account = accounts.find(a => a.id === newAccountId) || {
      id: newAccountId,
      changed: ts
    };

    // Обновляем поля аккаунта
    account.title = title;
    account.instrument = instrumentReverseMap[row[fieldIndex.instrument]] || Settings.DefaultCurrencyId;
    account.type = row[fieldIndex.type] || "cash";
    account.balance = parseNumber(row[fieldIndex.balance]);
    account.startBalance = parseNumber(row[fieldIndex.startBalance]);
    account.inBalance = parseBool(row[fieldIndex.inBalance]);
    account.private = parseBool(row[fieldIndex.private]);
    account.savings = parseBool(row[fieldIndex.savings]);
    account.archive = parseBool(row[fieldIndex.archive]);
    account.creditLimit = parseNumber(row[fieldIndex.creditLimit]);
    account.role = row[fieldIndex.role] || null;
    account.company = row[fieldIndex.company] || null;
    account.enableCorrection = parseBool(row[fieldIndex.enableCorrection]);
    account.balanceCorrectionType = row[fieldIndex.balanceCorrectionType] || "request";
    account.startDate = null;
    const startDateValue = row[fieldIndex.startDate];
    if (startDateValue instanceof Date) {
      const yyyy = startDateValue.getFullYear();
      const mm = String(startDateValue.getMonth() + 1).padStart(2, '0');
      const dd = String(startDateValue.getDate()).padStart(2, '0');
      account.startDate = `${yyyy}-${mm}-${dd}`;
    }
    account.capitalization = parseBool(row[fieldIndex.capitalization]);
    account.percent = parseNumber(row[fieldIndex.percent]);
    account.changed = ts;
    account.syncID = row[fieldIndex.syncID] || null;
    account.enableSMS = parseBool(row[fieldIndex.enableSMS]);
    account.endDateOffset = row[fieldIndex.endDateOffset] || null;
    account.endDateOffsetInterval = row[fieldIndex.endDateOffsetInterval] || null;
    account.payoffStep = row[fieldIndex.payoffStep] || null;
    account.payoffInterval = row[fieldIndex.payoffInterval] || null;
    account.user = user;

    return { account };
  }

  // Обновление аккаунтов (полное и частичное)
  function updateAccounts(values, dictionaries, isPartial = false) {
    try {
      const json = zmData.RequestForceFetch(['account']);
      if (!('account' in json)) {
        SpreadsheetApp.getActive().toast('Нет данных об аккаунтах', 'Предупреждение');
        return;
      }
      const accounts = json['account'] || [];
      const ts = Math.floor(Date.now() / 1000);

      const newAccounts = [];
      const deletionRequests = [];
      const deletedTitles = [];
      const idMap = {};
      const newIds = Array(values.length).fill(null);

      // Первый проход: определяем новые ID
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (isPartial && !parseBool(row[fieldIndex.modify])) continue;

        const oldAccountId = row[fieldIndex.id];
        const title = row[fieldIndex.title];
        if (!title || typeof title !== 'string' || title.trim() === '') continue;

        if (!oldAccountId || !accounts.find(a => a.id === oldAccountId)) {
          const newId = Utilities.getUuid().toLowerCase();
          idMap[oldAccountId || `new_${i}`] = newId;
          newIds[i] = newId;
        } else {
          idMap[oldAccountId] = oldAccountId;
        }
      }

      // Обновляем ID в таблице
      if (newIds.some(id => id !== null)) {
        const idColumnRange = sheet.getRange(2, fieldIndex.id + 1, values.length, 1);
        const idColumnValues = idColumnRange.getValues();
        for (let i = 0; i < values.length; i++) {
          if (newIds[i]) idColumnValues[i][0] = newIds[i];
        }
        idColumnRange.setValues(idColumnValues);
      }

      // Обратные словари для преобразования названий в ID
      const userReverseMap = Object.fromEntries(
        Object.entries(dictionaries.users || {})
          .map(([id, login]) => [login, id])
      );
      
      const instrumentReverseMap = Object.fromEntries(
        Object.entries(dictionaries.instruments || {})
          .map(([id, shortTitle]) => [shortTitle, id])
      );

      // Второй проход: создаем/обновляем аккаунты
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (isPartial && !parseBool(row[fieldIndex.modify])) continue;

        const result = createAccountFromRow(row, i, ts, accounts, idMap, userReverseMap, instrumentReverseMap);
        if (!result) continue;

        if (result.deletion) {
          deletionRequests.push(result.deletion);
          // Сохраняем название удаляемого счета
          const title = row[fieldIndex.title];
          if (title) deletedTitles.push(title);
        } else if (result.account) {
          newAccounts.push(result.account);
        }
      }

      // Предупреждение об удалении счетов
      if (deletedTitles.length > 0) {
        const deletionMsg = deletedTitles.length === 1
          ? `Будет удалён счёт: ${deletedTitles[0]}`
          : `Будут удалены счета:\n${deletedTitles.slice(0, 5).join('\n')}${
              deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''
            }`;
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление счетов');
        Utilities.sleep(2000);
      }

      // Отправляем изменения на сервер
      if (newAccounts.length > 0 || deletionRequests.length > 0) {
        const data = {
          currentClientTimestamp: ts,
          serverTimestamp: ts
        };
        if (newAccounts.length > 0) data.account = newAccounts;
        if (deletionRequests.length > 0) data.deletion = deletionRequests;
        SpreadsheetApp.getActive().toast('Отправляем изменения на сервер...', 'Обновление');
        const result = zmData.Request(data);
        Logger.log(`Результат ${isPartial ? 'частичного' : 'полного'} обновления счетов: ${JSON.stringify(result)}`);
        if (typeof Logs !== 'undefined' && Logs.logApiCall) {
          Logs.logApiCall("UPDATE_ACCOUNTS", data, result);
        }
        SpreadsheetApp.getActive().toast(
          `Обновление завершено:\n` +
          `Новых/изменённых: ${newAccounts.length}\n` +
          `Удалено: ${deletionRequests.length}`,
          'Успех'
        );

        // Сброс флагов modify после частичного обновления
        if (isPartial) {
          const modifyColumn = fieldIndex.modify + 1;
          for (let i = 0; i < values.length; i++) {
            if (parseBool(values[i][fieldIndex.modify])) {
              sheet.getRange(i + 2, modifyColumn).setValue(false);
            }
          }
        }

        doLoad();
      } else {
        SpreadsheetApp.getActive().toast('Нет изменений для обработки', 'Информация');
      }
    } catch (error) {
      Logger.log("Ошибка при обновлении счетов: " + error.toString());
      SpreadsheetApp.getActive().toast("Ошибка: " + error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  // Загрузка аккаунтов
  function doLoad() {
    Dictionaries.loadDictionariesFromSheet();
    const json = zmData.RequestForceFetch(['account']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_ACCOUNTS", { account: [] }, json);
    }
    prepareData(json);
  }

  // Полное обновление аккаунтов
  function doSave() {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      const msg = "Нет данных для обновления счетов";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast('Начинаем полное обновление счетов...', 'Обновление');
    Dictionaries.loadDictionariesFromSheet();
    const values = sheet.getRange(2, 1, lastRow - 1, Settings.ACCOUNT_FIELDS.length).getValues();
    updateAccounts(values, Dictionaries, false); // Передаем данные с флагом isPartial = false
  }

  // Частичное обновление аккаунтов
  function doPartial() {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      const msg = "Нет данных для обновления счетов";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast('Начинаем частичное обновление счетов...', 'Обновление');
    Dictionaries.loadDictionariesFromSheet();
    const values = sheet.getRange(2, 1, lastRow - 1, Settings.ACCOUNT_FIELDS.length).getValues();
    updateAccounts(values, Dictionaries, true); // Передаем данные с флагом isPartial = true
  }

  // Регистрация обработчика полной синхронизации
  fullSyncHandlers.push(prepareData);

  return {
    doLoad,
    doSave,
    doPartial
  };
})();
