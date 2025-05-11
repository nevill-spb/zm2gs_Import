const Accounts = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // КОНСТАНТЫ
  //═══════════════════════════════════════════════════════════════════════════
  const UPDATE_MODES = {
    SAVE: {
      id: 'SAVE',
      description: 'полное обновление',
      logType: 'UPDATE_ACCOUNTS'
    },
    PARTIAL: {
      id: 'PARTIAL',
      description: 'частичное обновление',
      logType: 'UPDATE_ACCOUNTS'
    },
    REPLACE: {
      id: 'REPLACE',
      description: 'замену',
      logType: 'REPLACE_ACCOUNTS'
    }
  };

  // Специальный счёт "Долги" (debt) — его ID всегда сохраняется, он не удаляется и не создаётся заново
  const DEBT_ACCOUNT = {
    title: 'Долги',
    type: 'debt'
  };

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

  function writeDataToSheet(data) {
    sheet.clearContents();
    sheet.clearFormats();

    const headers = Settings.ACCOUNT_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      const boolFields = [
        "inBalance", "private", "savings", "archive",
        "enableCorrection", "enableSMS", "capitalization",
        "delete", "modify"
      ];

      boolFields.forEach(fieldId => {
        const colIndex = fieldIndex[fieldId];
        if (colIndex === undefined) return;
        sheet.getRange(2, colIndex + 1, data.length, 1).insertCheckboxes();
      });

      clearExtraRows(data.length, boolFields);
    }
  }

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
    account.startDate = row[fieldIndex.type] === 'deposit'  // обработка даты только для депозитов
      ? new Date(row[fieldIndex.startDate] || Date.now()).toISOString().split('T')[0]  
      : null;
    account.capitalization = parseBool(row[fieldIndex.capitalization]);
    account.percent = parseNumber(row[fieldIndex.percent]);
    account.changed = ts;
    account.syncID = (oldAccountId && accounts.find(a => a.id === oldAccountId))   
      ? (row[fieldIndex.syncID] || null)   
      : null; // только для существующих аккаунтов    account.enableSMS = parseBool(row[fieldIndex.enableSMS]);
    account.enableSMS = parseBool(row[fieldIndex.enableSMS]);
    account.endDateOffset = row[fieldIndex.endDateOffset] || null;
    account.endDateOffsetInterval = row[fieldIndex.endDateOffsetInterval] || null;
    account.payoffStep = row[fieldIndex.payoffStep] || null;
    account.payoffInterval = row[fieldIndex.payoffInterval] || null;
    account.user = user;

    return { account };
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ГЛАВНАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ/ЗАМЕНЫ СЧЕТОВ
  //═══════════════════════════════════════════════════════════════════════════

  function doUpdate(mode) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      const msg = "Нет данных для обновления счетов";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast(`Начинаем ${mode.description} счетов...`, 'Обновление');
    Dictionaries.loadDictionariesFromSheet();
    const values = sheet.getRange(2, 1, lastRow - 1, Settings.ACCOUNT_FIELDS.length).getValues();
    updateAccounts(values, Dictionaries, mode);
  }

  function updateAccounts(values, dictionaries, mode) {
    try {
      // Запрос данных с сервера
      const json = zmData.RequestForceFetch(['account']);
      if (!('account' in json)) {
        SpreadsheetApp.getActive().toast('Нет данных об аккаунтах', 'Предупреждение');
        return;
      }
      const accounts = json['account'] || [];
      const ts = Math.floor(Date.now() / 1000);

      // Находим существующий счёт "Долги"
      const existingDebtAccount = accounts.find(acc =>
        acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type
      );

      // Подсчёт изменений
      const changes = {
        new: 0,
        modified: 0,
        deleted: 0,
        processedIds: new Set(),
        idMap: {}
      };

      // Подготовка обратных словарей
      const userReverseMap = Object.fromEntries(
        Object.entries(dictionaries.users || {})
          .map(([id, login]) => [login, id])
      );
      const instrumentReverseMap = Object.fromEntries(
        Object.entries(dictionaries.instruments || {})
          .map(([id, shortTitle]) => [shortTitle, id])
      );

      // Первый проход: анализ изменений и формирование idMap
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (mode.id === UPDATE_MODES.PARTIAL.id && !parseBool(row[fieldIndex.modify])) continue;

        const oldAccountId = row[fieldIndex.id];
        const title = row[fieldIndex.title];
        const type = row[fieldIndex.type];
        const deleteFlag = parseBool(row[fieldIndex.delete]);

        if (!title || typeof title !== 'string' || title.trim() === '') continue;

        // Особая обработка счёта "Долги"
        if (title === DEBT_ACCOUNT.title && type === DEBT_ACCOUNT.type) {
          changes.modified++;
          // Используем ID существующего счёта "Долги" или сохраняем текущий
          const debtId = existingDebtAccount ? existingDebtAccount.id : oldAccountId;
          changes.idMap[oldAccountId || `new_${i}`] = debtId;
          changes.processedIds.add(debtId);
          continue;
        }

        if (mode.id === UPDATE_MODES.REPLACE.id) {
          // В режиме REPLACE считаем удалённые как отсутствующие старые счета
          if (!oldAccountId || !accounts.find(a => a.id === oldAccountId)) {
            changes.new++;
            changes.idMap[oldAccountId || `new_${i}`] = Utilities.getUuid().toLowerCase();
          } else {
            changes.modified++;
            changes.idMap[oldAccountId] = oldAccountId;
            changes.processedIds.add(oldAccountId);
          }
        } else {
          // В режимах SAVE/PARTIAL
          if (!oldAccountId) {
            // Счета без ID всегда считаются новыми, игнорируя флажок "Удалить"
            changes.new++;
            changes.idMap[`new_${i}`] = Utilities.getUuid().toLowerCase();
          } else {
            // Для существующих счетов учитываем флажок "Удалить"
            if (deleteFlag) {
              changes.deleted++;
              changes.processedIds.add(oldAccountId); // Чтобы не учитывать как "новые"
            } else {
              if (!accounts.find(a => a.id === oldAccountId)) {
                changes.new++;
                changes.idMap[oldAccountId] = Utilities.getUuid().toLowerCase();
              } else {
                changes.modified++;
                changes.idMap[oldAccountId] = oldAccountId;
                changes.processedIds.add(oldAccountId);
              }
            }
          }
        }
      }
      if (mode.id === UPDATE_MODES.REPLACE.id) {
        // В режиме REPLACE считаем удаляемые счета (кроме "Долги")
        changes.deleted = accounts.filter(acc =>
          !changes.processedIds.has(acc.id) &&
          !(acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type)
        ).length;
      }

      // Определяем, выбран ли счёт DEBT_ACCOUNT в зависимости от режима
      let isDebtAccountSelected = false;
      if (mode.id === UPDATE_MODES.PARTIAL.id) {
        isDebtAccountSelected = values.some(row =>
          parseBool(row[fieldIndex.modify]) && // + флажок modify
          row[fieldIndex.title] === DEBT_ACCOUNT.title &&
          row[fieldIndex.type] === DEBT_ACCOUNT.type
        );
      } else {
        isDebtAccountSelected = values.some(row =>
          row[fieldIndex.title] === DEBT_ACCOUNT.title &&
          row[fieldIndex.type] === DEBT_ACCOUNT.type
        );
      }

      // Формируем сообщение для подтверждения
      const confirmLines = [];
      if (changes.new > 0) {
        confirmLines.push(`Будет создано новых счетов: ${changes.new}`);
      }
      if (changes.modified > 0) {
        // Не считаем счёт "Долги" как модифицированный
        const modifiedCount = changes.modified - (isDebtAccountSelected ? 1 : 0);
        if (modifiedCount > 0) {
          confirmLines.push(`Будет изменено счетов: ${modifiedCount}`);
        }
      }
      if (changes.deleted > 0) {
        confirmLines.push(`Будет удалено счетов: ${changes.deleted}`);
      }

      const confirmMessage =
        confirmLines.length > 0
          ? isDebtAccountSelected
            ? `За исключением счёта "${DEBT_ACCOUNT.title}":\n${confirmLines.join('\n')}`
            : confirmLines.join('\n')
          : isDebtAccountSelected
            ? `За исключением счёта "${DEBT_ACCOUNT.title}":\nСчетов для изменения не выбрано`
            : 'Счетов для изменения не выбрано';

      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Подтверждение',
        `${confirmMessage}\n\nПродолжить?`,
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        SpreadsheetApp.getActive().toast('Операция отменена', 'Информация');
        return;
      }

      // Обновляем ID в таблице для новых счетов
      const newIds = Array(values.length).fill(null);
      for (let i = 0; i < values.length; i++) {
        const oldId = values[i][fieldIndex.id];
        const newId = changes.idMap[oldId || `new_${i}`];
        if (newId !== oldId) newIds[i] = newId;
      }
      if (newIds.some(id => id !== null)) {
        const idColumnRange = sheet.getRange(2, fieldIndex.id + 1, values.length, 1);
        const idColumnValues = idColumnRange.getValues();
        for (let i = 0; i < values.length; i++) {
          if (newIds[i]) idColumnValues[i][0] = newIds[i];
        }
        idColumnRange.setValues(idColumnValues);
      }

      // Второй проход: создание объектов для обновления
      const newAccounts = [];
      const deletionRequests = [];
      const deletedTitles = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (mode.id === UPDATE_MODES.PARTIAL.id && !parseBool(row[fieldIndex.modify])) continue;

        const result = createAccountFromRow(row, i, ts, accounts, changes.idMap, userReverseMap, instrumentReverseMap);
        if (!result) continue;

        if (result.deletion) {
          deletionRequests.push(result.deletion);
          const title = row[fieldIndex.title];
          if (title) deletedTitles.push(title);
        } else if (result.account) {
          newAccounts.push(result.account);
        }
      }

      // В режиме замены добавляем запросы на удаление для лишних счетов (кроме "Долги")
      if (mode.id === UPDATE_MODES.REPLACE.id) {
        const accountsToDelete = accounts
          .filter(acc =>
            !changes.processedIds.has(acc.id) &&
            !(acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type)
          )
          .map(acc => ({
            id: acc.id,
            object: 'account',
            stamp: ts,
            user: acc.user
          }));

        deletionRequests.push(...accountsToDelete);
        accountsToDelete.forEach(del => {
          const acc = accounts.find(a => a.id === del.id);
          if (acc && acc.title) deletedTitles.push(acc.title);
        });
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

        if (typeof Logs !== 'undefined' && Logs.logApiCall) {
          Logs.logApiCall(mode.logType, data, result);
        }

        // Сброс флагов modify после частичного обновления
        if (mode.id === UPDATE_MODES.PARTIAL.id) {
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

      // Подсчитываем количество реально изменённых счетов (без "Долги")
      const actualNewModified = newAccounts.filter(acc =>
        !(acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type)
      ).length;

      // Подсчитываем количество реально удалённых счетов (без "Долги")
      const actualDeleted = deletionRequests.filter(del => {
        const acc = accounts.find(a => a.id === del.id);
        return !(acc && acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type);
      }).length;

      SpreadsheetApp.getActive().toast(
        `${mode.id === UPDATE_MODES.REPLACE.id ? 'Замена' : 'Обновление'} завершено:\n` +
        `Новых/изменённых: ${actualNewModified}\n` +
        `Удалено: ${actualDeleted}`,
        'Успех'
      );
    } catch (error) {
      Logger.log("Ошибка при обновлении счетов: " + error.toString());
      SpreadsheetApp.getActive().toast("Ошибка: " + error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  function doLoad() {
    Dictionaries.loadDictionariesFromSheet();
    const json = zmData.RequestForceFetch(['account']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_ACCOUNTS", { account: [] }, json);
    }
    prepareData(json);
  }

  return {
    doLoad,
    doSave: () => doUpdate(UPDATE_MODES.SAVE),
    doPartial: () => doUpdate(UPDATE_MODES.PARTIAL),
    doReplace: () => doUpdate(UPDATE_MODES.REPLACE)
  };
})();
