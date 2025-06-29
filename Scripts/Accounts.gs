const Accounts = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // КОНСТАНТЫ
  //═══════════════════════════════════════════════════════════════════════════
  // Режимы обновления счетов
  const UPDATE_MODES = {  
    SAVE: { id: 'SAVE', description: 'полное обновление', logType: 'UPDATE_ACCOUNTS' },  
    PARTIAL: { id: 'PARTIAL', description: 'частичное обновление', logType: 'UPDATE_ACCOUNTS' },  
    REPLACE: { id: 'REPLACE', description: 'замену', logType: 'REPLACE_ACCOUNTS' }  
  };  
  
  // Специальный счет "Долги" (debt) — он не удаляется и не создаётся заново  
  const DEBT_ACCOUNT = { title: 'Долги', type: 'debt' };

  //═══════════════════════════════════════════════════════════════════════════
  // ИНИЦИАЛИЗАЦИЯ
  //═══════════════════════════════════════════════════════════════════════════
  // Получение листа категорий из настроек
  const sheet = sheetHelper.GetSheetFromSettings('ACCOUNTS_SHEET');
  if (!sheet) {
    Logger.log("Лист с аккаунтами не найден");
    SpreadsheetApp.getActive().toast('Лист с аккаунтами не найден', 'Ошибка');
    return;
  }

  // Определяет поля по id для удобства доступа
  const fieldIndex = {};
  Settings.ACCOUNT_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  //═══════════════════════════════════════════════════════════════════════════
  // Преобразует значение в булево
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  // Преобразует значение в число с дефолтом
  function parseNumber(value, defaultValue = 0) {
    if (value === "" || value == null) return defaultValue;
    return Number(value);
  }

  // Преобразует строку в дату или возвращает пустую строку
  function parseDateFromString(dateStr) {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? "" : d;
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ДАННЫМИ
  //═══════════════════════════════════════════════════════════════════════════
  // Генерирует карту соответствия старых и новых ID счетов, считает статистику изменений
  function generateIdMap(values, fieldIndex, mode, accountsMap, debtAccount, existingIds) {
    const idMap = {};
    const processedIds = new Set();
    const stats = { new: 0, modified: 0, deleted: 0 };

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const currentId = row[fieldIndex.id];
      const title = String(row[fieldIndex.title] || '').trim();
      const type = row[fieldIndex.type];
      if (!title) continue;

      const shouldDelete = parseBool(row[fieldIndex.delete]);
      const shouldModify = parseBool(row[fieldIndex.modify]);
      const existsOnServer = existingIds.has(currentId);
      const isDebtAccount = title === DEBT_ACCOUNT.title && type === DEBT_ACCOUNT.type;

      // Особый случай для счета "Долги"
      if (isDebtAccount) {
        idMap[currentId || `new_${i}`] = debtAccount.id;
        processedIds.add(debtAccount.id);
        continue;
      }

      // Удаление только существующих с подтверждением
      if (shouldDelete) {
        if (shouldModify && existsOnServer) {
          processedIds.add(currentId);
          stats.deleted++;
        }
        continue; // Пропускаем все остальные случаи удаления
      }

      if (!existsOnServer) {
        if (!currentId) {
          // Новые счета без ID
          idMap[`new_${i}`] = Utilities.getUuid().toLowerCase();
          stats.new++;
        } else if (mode !== UPDATE_MODES.PARTIAL || shouldModify) {
          // Новые счета с ID
          idMap[currentId] = Utilities.getUuid().toLowerCase();
          stats.new++;
        }
      } else if (shouldModify) {
        // Существующие счета с флагом modify
        idMap[currentId] = currentId;
        processedIds.add(currentId);
        stats.modified++;
      }
    }

    // В режиме REPLACE удаляем всё, чего нет в таблице (кроме счета "Долги")
    if (mode.id === UPDATE_MODES.REPLACE.id) {
      for (const [id, account] of accountsMap) {
        if (!processedIds.has(id) && 
            !(account.title === DEBT_ACCOUNT.title && account.type === DEBT_ACCOUNT.type)) {
          stats.deleted++;
        }
      }
    }

    return { ...stats, idMap, processedIds };
  }

  // Формирует строку для записи в лист из объекта счета
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
        case "role":  
          return Dictionaries.getUserLogin(account.role) || account.role || "";
        case "startDate":
          if (typeof account.startDate === "string") return parseDateFromString(account.startDate);
          if (account.startDate != null) return new Date(account.startDate * 1000);
          return "";
        default:
          return account[field.id] != null ? account[field.id] : "";
      }
    });
  }

  // Подготавливает данные для записи в лист, сортирует счета по названию и id
  function prepareData(json, showToast = true) {
    if (!('account' in json)) {
      if (showToast) SpreadsheetApp.getActive().toast('Нет данных об счетах', 'Предупреждение');  
      return;
    }

    const sortedAccounts = json.account.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    const data = sortedAccounts.map(acc => buildRow(acc));
    writeDataToSheet(data);
    if (showToast) SpreadsheetApp.getActive().toast(`Загружено ${data.length} счетов`, 'Информация');
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ЛИСТОМ
  //═══════════════════════════════════════════════════════════════════════════
  // Записывает данные в лист и форматирует строки
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

      // Упорядочивает индексы
      const checkboxColumns = boolFields
        .map(fieldId => fieldIndex[fieldId] + 1)
        .filter(col => col !== undefined)
        .sort((a, b) => a - b);

      // Группирует смежные колонки
      if (checkboxColumns.length > 0) {
        const columnGroups = checkboxColumns.reduce((groups, col) => {
          const lastGroup = groups[groups.length - 1];
          if (lastGroup && col === lastGroup.end + 1) {
            lastGroup.end = col;
          } else {
            groups.push({ start: col, end: col });
          }
          return groups;
        }, []);

        // Вставляет чекбоксы для каждой группы
        columnGroups.forEach(({ start, end }) => {
          sheet.getRange(2, start, data.length, end - start + 1)
            .insertCheckboxes();
        });
      }

      clearExtraRows(data.length, boolFields);
    }
  }

  // Очищает лишние строки после данных в чекбокс-колонках
  function clearExtraRows(dataLength, boolFields) {
    const checkboxColumns = [];

    for (const id of boolFields) {
      const col = fieldIndex[id];
      if (col !== undefined) {
        checkboxColumns.push(col);
      }
    }
    
    if (!checkboxColumns.length || sheet.getMaxRows() <= dataLength + 2) return;
    
    const range = sheet.getRange(
      dataLength + 2,
      Math.min(...checkboxColumns) + 1,
      sheet.getMaxRows() - dataLength - 1,
      Math.max(...checkboxColumns) - Math.min(...checkboxColumns) + 1
    );
    
    range.clearContent().clearDataValidations();
  }

  // Определяет строки для вставки чекбоксов
  function getRowsToModify(values, mode) {
    if (mode.id === UPDATE_MODES.SAVE.id || mode.id === UPDATE_MODES.REPLACE.id) {
      // В режимах SAVE и REPLACE отмечаем все строки с названием
      return values
        .map((row, i) => {
          const title = row[fieldIndex.title]?.toString().trim();
          return title ? i + 2 : null;
        })
        .filter(i => i !== null);
    }
    
    if (mode.id === UPDATE_MODES.PARTIAL.id) {
      // В режиме PARTIAL отмечаем только строки с флагом modify или без ID
      return values
        .map((row, i) => {
          const title = row[fieldIndex.title]?.toString().trim();
          const modify = Boolean(row[fieldIndex.modify]);
          const id = row[fieldIndex.id]?.toString().trim();
          return title && (modify || !id) ? i + 2 : null;
        })
        .filter(i => i !== null);
    }
    
    return [];
  }

  // Вставляет чекбоксы с заданным значением в указанные строки столбца листа
  function insertCheckboxesBatchWithValue(sheet, column, rowIndices, valueForRows) {
    if (rowIndices.length === 0) return;

    rowIndices.sort((a, b) => a - b);
    const minRow = rowIndices[0];
    const maxRow = rowIndices[rowIndices.length - 1];
    const totalRows = maxRow - minRow + 1;

    const rowSet = new Set(rowIndices);
    const currentValues = Array.from({length: totalRows}, (_, i) => {
      return [rowSet.has(minRow + i) ? valueForRows : null];
    });

    const range = sheet.getRange(minRow, column, totalRows, 1);
    range.insertCheckboxes();
    range.setValues(currentValues);
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ СОЗДАНИЯ И ВАЛИДАЦИИ СЧЕТОВ
  //═══════════════════════════════════════════════════════════════════════════

  // Создаёт объект счета из строки листа
  function createAccountFromRow(row, i, ts, existingIds, mode, accountsMap) {
    const currentId = row[fieldIndex.id];
    const titleRaw = row[fieldIndex.title];
    const title = (titleRaw != null) ? String(titleRaw).trim() : '';
    if (!title) return null;

    const user = Dictionaries.getUserId(row[fieldIndex.user]) || Settings.DefaultUserId;
    const shouldModify = parseBool(row[fieldIndex.modify]);
    const shouldDelete = parseBool(row[fieldIndex.delete]);
    const existsOnServer = existingIds.has(currentId);
    const isDebtAccount = title === DEBT_ACCOUNT.title && row[fieldIndex.type] === DEBT_ACCOUNT.type;
    const isDeposit = row[fieldIndex.type] === 'deposit';
    
    if (mode === UPDATE_MODES.PARTIAL && !shouldModify) return null;

    // Обработка удаления (кроме счета "Долги")
    if (shouldDelete && currentId && existsOnServer && !isDebtAccount) {
      return {
        deletion: {
          id: currentId,
          object: 'account',
          stamp: ts,
          user: user
        }
      };
    }

    if (!currentId) {
      Logger.log(`Строка ${i+1} пропущена: не определён ID`);
      return null;
    }

    let account = accountsMap.get(currentId) || { id: currentId, };

    Object.assign(account, {
      title: title,
      instrument: Dictionaries.getInstrumentId(row[fieldIndex.instrument]) || Settings.DefaultCurrencyId,
      type: row[fieldIndex.type] || "cash",
      balance: parseNumber(row[fieldIndex.balance]),
      startBalance: parseNumber(row[fieldIndex.startBalance]),
      inBalance: parseBool(row[fieldIndex.inBalance]),
      private: parseBool(row[fieldIndex.private]),
      savings: parseBool(row[fieldIndex.savings]),
      archive: parseBool(row[fieldIndex.archive]),
      creditLimit: parseNumber(row[fieldIndex.creditLimit]),
      role: Dictionaries.getUserId(row[fieldIndex.role]) || null,
      company: row[fieldIndex.company] || null,
      enableCorrection: parseBool(row[fieldIndex.enableCorrection]),
      balanceCorrectionType: row[fieldIndex.balanceCorrectionType] || "request",
      capitalization: parseBool(row[fieldIndex.capitalization]),
      percent: parseNumber(row[fieldIndex.percent]),
      syncID: existsOnServer ? (row[fieldIndex.syncID] || null) : null,
      enableSMS: parseBool(row[fieldIndex.enableSMS]),
      startDate: isDeposit ? new Date(row[fieldIndex.startDate] || Date.now()).toISOString().split('T')[0] : null,
      endDateOffset: isDeposit ? (row[fieldIndex.endDateOffset] || 1) : null,
      endDateOffsetInterval: isDeposit ? (row[fieldIndex.endDateOffsetInterval] || 'month') : null,
      payoffStep: isDeposit ? (row[fieldIndex.payoffStep] || 0) : null,
      payoffInterval: row[fieldIndex.payoffInterval] || null,
      user: user,
      changed: ts
    });

    return { account };
  }

  // Валидирует счет
  function validateAccount(account, existingAccounts = [], deletionRequests = []) {  
    // Проверка обязательных полей  
    if (!account.title || !account.type || !account.instrument) {  
      throw new Error('Не заполнены обязательные поля (название, тип, валюта)');  
    }
    
    // Проверка уникальности названия  
    const isDuplicate = existingAccounts.some(a =>  
      a.title === account.title &&  
      a.id !== account.id &&  
      !deletionRequests.some(del => del.id === a.id) // Исключаем счета, помеченные к удалению  
    );  
    if (isDuplicate) {  
      throw new Error(`счет с названием "${account.title}" уже существует`);  
    }
    
    // Валидация специального счета "Долги"  
    if (account.title === DEBT_ACCOUNT.title || account.type === DEBT_ACCOUNT.type) {  
      if (account.title !== DEBT_ACCOUNT.title || account.type !== DEBT_ACCOUNT.type) {  
        throw new Error('Нарушение правил для счета "Долги"');  
      }  
    }  
    
    // Валидация по типам счетов  
    switch (account.type) {  
      case 'deposit':  
        if (account.startDate == null) throw new Error('Для депозита требуется дата открытия');  
        if (account.percent == null) throw new Error('Для депозита требуется процентная ставка');  
        break;  
      case 'cash':  
        if (account.creditLimit > 0) throw new Error('Для наличных нельзя установить кредитный лимит');  
        break;  
    }  
    
    return true;  
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ГЛАВНАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ/ЗАМЕНЫ СЧЕТОВ
  //═══════════════════════════════════════════════════════════════════════════

  // Обновляет счета на сервере
  function updateAccounts(values, mode) {
    try {
      const errors = [];
      // Запрашивает текущие данные с сервера
      const json = zmData.RequestForceFetch(['account']);
      if (!('account' in json)) {
        SpreadsheetApp.getActive().toast('Нет данных об счетах', 'Предупреждение');
        return;
      }
      const accounts = json['account'] || [];
      const ts = Math.floor(Date.now() / 1000);

      // Создает Set и Map для быстрого поиска
      const existingIds = new Set(accounts.map(a => a.id));
      const accountsMap = new Map(accounts.map(a => [a.id, a]));
      const debtAccount = [...accountsMap.values()].find(account => 
          account.title === DEBT_ACCOUNT.title && 
          account.type === DEBT_ACCOUNT.type
        );     

      const { idMap, processedIds, new: newCount, modified: modifiedCount, deleted: deletedCount } = generateIdMap(
        values,
        fieldIndex,
        mode,
        accountsMap,
        debtAccount,
        existingIds
      );

      const confirmLines = [];
      if (newCount > 0) confirmLines.push(`Будет создано новых счетов: ${newCount}`);
      if (modifiedCount > 0) confirmLines.push(`Будет изменено счетов: ${modifiedCount}`);
      if (deletedCount > 0) confirmLines.push(`Будет удалено счетов: ${deletedCount}`);

      let confirmMessage = confirmLines.length > 0 ? confirmLines.join('\n') : 'Счетов для изменения не выбрано';

      const isDebtModified = values.some(row => 
        row[fieldIndex.title] === DEBT_ACCOUNT.title && 
        row[fieldIndex.type] === DEBT_ACCOUNT.type &&
        (mode.id !== UPDATE_MODES.PARTIAL.id || parseBool(row[fieldIndex.modify]))
      );

      if (isDebtModified) {
        confirmMessage = `За исключением счета "${DEBT_ACCOUNT.title}":\n${confirmMessage}`;
      }

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
        const newId = idMap[oldId || `new_${i}`];
        if (newId !== oldId) newIds[i] = newId;
      }
      if (newIds.some(Boolean)) {
        const updates = values.map((r, i) => [newIds[i] || r[fieldIndex.id]]);
        sheet.getRange(2, fieldIndex.id + 1, values.length, 1).setValues(updates);
        updates.forEach(([id], i) => values[i][fieldIndex.id] = id);
      }

      // Первый проход: собираем запросы на удаление
      const deletionRequests = [];
      const deletedTitles = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const result = createAccountFromRow(row, i, ts, existingIds, mode, accountsMap);
        if (!result || !result.deletion) continue;

        deletionRequests.push(result.deletion);
        deletedTitles.push(row[fieldIndex.title]);
      }

      // В режиме REPLACE добавляем запросы на удаление лишних счетов (кроме "Долги")
      if (mode.id === UPDATE_MODES.REPLACE.id) {
        for (const [id, account] of accountsMap) {
          if (!processedIds.has(id) && 
              !(account.title === DEBT_ACCOUNT.title && account.type === DEBT_ACCOUNT.type)) {
            deletionRequests.push({
              id: id,
              object: 'account',
              stamp: ts,
              user: account.user
            });
            deletedTitles.push(account.title);
          }
        }
      }

      // Второй проход: создание новых счетов с учетом удалений
      const newAccounts = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const result = createAccountFromRow(row, i, ts, existingIds, mode, accountsMap);
        if (!result || result.deletion) continue;

        try {
          validateAccount(result.account, accounts, deletionRequests);
          newAccounts.push(result.account);
        } catch (error) {
          errors.push(`Строка ${i+2}: ${error.message}`);
        }
      }

      // Уведомление об удалении
      if (deletedTitles.length > 0) {
        const deletionMsg = deletedTitles.length === 1
          ? `Будет удален счет: ${deletedTitles[0]}`
          : `Будут удалены счета:\n${deletedTitles.slice(0, 5).join(',\n')}${deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''}`;
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление счетов');
        Utilities.sleep(2000);
      }

      // Отправка изменений
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

        // Проверка ответа сервера 
        if (!result || Object.keys(result).length === 0 || 
          (Object.keys(result).length === 1 && 'serverTimestamp' in result)) {
          throw new Error('Пустой ответ сервера при обновлении счетов');
        }

        // Сброс флагов modify
        const modifyColumn = fieldIndex.modify + 1;
        const modifyRange = sheet.getRange(2, modifyColumn, values.length, 1);
        const modifyValues = modifyRange.getValues();

        values.forEach((_, i) => {
          modifyValues[i][0] = parseBool(values[i][fieldIndex.modify]) && 
                    (newAccounts.some(a => a.id === values[i][fieldIndex.id]) || 
                    deletionRequests.some(d => d.id === values[i][fieldIndex.id]))
                    ? false 
                    : modifyValues[i][0];
        });

        modifyRange.setValues(modifyValues);

        // Если все счета обработаны без ошибок, перезагружает список счетов
        if (!modifyValues.some(row => parseBool(row[0]))) {  
          doLoad(false); // не показывать toast при перезагрузке после обновления
        }

        // Подсчёт реальных изменений без учёта счёта "Долги"
        const actualChanges = newAccounts.filter(a => 
          !(a.title === DEBT_ACCOUNT.title && a.type === DEBT_ACCOUNT.type)).length;
        const actualDeletions = deletionRequests.filter(d => {
          const acc = accountsMap.get(d.id);
          return !(acc && acc.title === DEBT_ACCOUNT.title && acc.type === DEBT_ACCOUNT.type);
        }).length;

        SpreadsheetApp.getActive().toast(
          `${mode.id === UPDATE_MODES.REPLACE.id ? 'Замена' : 'Обновление'} завершено:\n` +
          `Новых/изменённых: ${actualChanges}\n` +
          `Удалено: ${actualDeletions}` +
          (errors.length > 0 ? `\nОшибок: ${errors.length}` : ''),
          'Успех'
        );

      } else {
        SpreadsheetApp.getActive().toast(
          'Нет изменений для обработки' +
          (errors.length > 0 ? `\nСтрок с ошибоками: ${errors.length}` : ''),
          'Информация'
        );
      }
      // Логирование ошибок
      if (errors.length > 0) {
        const errorSheet = sheetHelper.Get(Settings.SHEETS.ERRORS);
        errorSheet.insertRowsBefore(1, errors.length + 1);
        errorSheet.getRange(1, 1).setValue(`Ошибки ${mode.description} счетов ` + new Date().toLocaleString());
        errorSheet.getRange(2, 1, errors.length, 1).setValues(errors.map(e => [e]));
      }
    } catch (error) {
      Logger.log("Ошибка при обновлении счетов: " + error.toString());
      SpreadsheetApp.getActive().toast(error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  // Загружает счета с сервера и подготавливает данные для листа
  function doLoad(showToast = true) {
    Dictionaries.loadDictionariesFromSheet();
    
    const json = zmData.RequestForceFetch(['account']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_ACCOUNTS", { account: [] }, json);
    }
    prepareData(json, showToast);
  }

  // Основная функция обновления счетов
  function doUpdate(mode) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('Нет данных для обновления счетов');
      SpreadsheetApp.getActive().toast('Нет данных для обновления счетов', 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast(`Начинаем ${mode.description} счетов...`, 'Обновление');

    Dictionaries.loadDictionariesFromSheet();

    const dataRange = sheet.getRange(2, 1, lastRow - 1, Settings.ACCOUNT_FIELDS.length);
    let values = dataRange.getValues();
    let validRowIndices = getRowsToModify(values, mode);

    if (validRowIndices.length > 0) {
      insertCheckboxesBatchWithValue(sheet, fieldIndex.modify + 1, validRowIndices, true);
      validRowIndices.forEach(row => values[row-2][fieldIndex.modify] = true);
    }

    updateAccounts(values, mode);
  }

  // Регистрация обработчика полной синхронизации
  fullSyncHandlers.push(prepareData);

  return {
    doLoad,
    doSave: () => doUpdate(UPDATE_MODES.SAVE),
    doPartial: () => doUpdate(UPDATE_MODES.PARTIAL),
    doReplace: () => doUpdate(UPDATE_MODES.REPLACE)
  };
})();
