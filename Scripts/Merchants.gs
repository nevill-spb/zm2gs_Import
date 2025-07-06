const Merchants = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // КОНСТАНТЫ
  //═══════════════════════════════════════════════════════════════════════════
  // Режимы обновления мест
  const UPDATE_MODES = {  
    SAVE: { id: 'SAVE', description: 'полное обновление', logType: 'UPDATE_MERCHANTS' },  
    PARTIAL: { id: 'PARTIAL', description: 'частичное обновление', logType: 'UPDATE_MERCHANTS' },  
    REPLACE: { id: 'REPLACE', description: 'замену', logType: 'REPLACE_MERCHANTS' }  
  };

  //═══════════════════════════════════════════════════════════════════════════
  // ИНИЦИАЛИЗАЦИЯ
  //═══════════════════════════════════════════════════════════════════════════

  let initialized = false;
  let sheet = null;
  let fieldIndex = {};
  
  function initialize() {
    if (initialized) return true;
    
    try {
      // Получение листа мест из настроек
      sheet = sheetHelper.GetSheetFromSettings('MERCHANTS_SHEET');
      if (!sheet) {
        Logger.log("Лист со местами не найден");
        SpreadsheetApp.getActive().toast('Лист со местами не найден', 'Ошибка');
        return false;
      }

      // Определяет поля по id для удобства доступа
      Settings.MERCHANT_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);
      
      initialized = true;
      return true;
    } catch (e) {
      Logger.log("Ошибка инициализации Merchants: " + e.toString());
      return false;
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  //═══════════════════════════════════════════════════════════════════════════
  // Преобразует значение в булево
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ДАННЫМИ
  //═══════════════════════════════════════════════════════════════════════════
  // Генерирует карту соответствия старых и новых ID мест, считает статистику изменений
  function generateIdMap(values, fieldIndex, mode, merchantsMap, existingIds) {
    const idMap = {};
    const processedIds = new Set();
    const stats = { new: 0, modified: 0, deleted: 0 };
    const merchantsToDelete = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const currentId = row[fieldIndex.id];
      const title = String(row[fieldIndex.title] || '').trim();
      if (!title) continue;

      const shouldDelete = parseBool(row[fieldIndex.delete]);
      const shouldModify = parseBool(row[fieldIndex.modify]);
      const existsOnServer = existingIds.has(currentId);

      // Удаление только существующих с подтверждением
      if (shouldDelete) {
        if (shouldModify && existsOnServer) {
          processedIds.add(currentId);
          merchantsToDelete.push(currentId);
          stats.deleted++;
        }
        continue; // Пропускаем все остальные случаи удаления
      }

      if (!existsOnServer) {
        if (!currentId) {
          // Новые места без ID
          idMap[`new_${i}`] = Utilities.getUuid().toLowerCase();
          stats.new++;
        } else if (mode !== UPDATE_MODES.PARTIAL || shouldModify) {
          // Новые места с ID
          idMap[currentId] = Utilities.getUuid().toLowerCase();
          stats.new++;
        }
      } else if (shouldModify) {
        // Существующие места с флагом modify
        idMap[currentId] = currentId;
        processedIds.add(currentId);
        stats.modified++;
      }
    }

    // В режиме REPLACE удаляем всё, чего нет в таблице
    if (mode.id === UPDATE_MODES.REPLACE.id) {
      for (const [id, merchant] of merchantsMap) {
        if (!processedIds.has(id)) {
          merchantsToDelete.push(id);
          stats.deleted++;
        }
      }
    }

    return { ...stats, idMap, processedIds, merchantsToDelete };
  }

  // Формирует строку для записи в лист из объекта места
  function buildRow(merchant) {
    return Settings.MERCHANT_FIELDS.map(field => {
      switch (field.id) {
        case "delete":
        case "modify":
          return false; // чекбоксы пустые
        case "user":
          return Dictionaries.getUserLogin(merchant.user) || merchant.user || "";
        default:
          return merchant[field.id] != null ? merchant[field.id] : "";
      }
    });
  }

  // Подготавливает данные для записи в лист, сортирует места по названию и id
  function prepareData(json, showToast = true) {
    if (!initialize()) return;

    if (!('merchant' in json)) {
      if (showToast) SpreadsheetApp.getActive().toast('Нет данных о местах', 'Предупреждение');  
      return;
    }

    const sortedMerchants = json.merchant.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    const data = sortedMerchants.map(merchant => buildRow(merchant));
    writeDataToSheet(data);
    if (showToast) SpreadsheetApp.getActive().toast(`Загружено ${data.length} мест`, 'Информация');
  }

  // Проверяет транзакции, имеющие связи с местами
  function checkTransactionsWithMerchants(merchantIds) {
    try {
      const json = zmData.RequestForceFetch(['transaction']);
      if (!json.transaction || !Array.isArray(json.transaction)) {
        Logger.log("Нет данных о транзакциях");
        return null;
      }

      let totalAffectedCount = 0;
      const affectedTransactions = [];

      json.transaction.forEach(t => {
        if (!t.deleted && t.merchant && merchantIds.includes(t.merchant)) {
          totalAffectedCount++;
          if (affectedTransactions.length < 20) affectedTransactions.push(t);
        }
      });

      return {
        affcount: totalAffectedCount,
        sample: affectedTransactions.map(t => {
          return `${t.date}${
            t.income 
              ? ` | ${Dictionaries.getAccountTitle(t.incomeAccount)}: +${t.income} ${Dictionaries.getInstrumentShortTitle(t.incomeInstrument)}` 
              : t.outcome 
                ? ` | ${Dictionaries.getAccountTitle(t.outcomeAccount)}: -${t.outcome} ${Dictionaries.getInstrumentShortTitle(t.outcomeInstrument)}` : ''
          }${
            t.merchant 
              ? ` | ${Dictionaries.getMerchantTitle(t.merchant)}` 
              : t.payee ? ` | ${t.payee}` : ''
          }`;
        })
      };
    } catch (e) {
      Logger.log("Ошибка при проверке транзакций: " + e.toString());
      return null;
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ЛИСТОМ
  //═══════════════════════════════════════════════════════════════════════════
  // Записывает данные в лист и форматирует строки
  function writeDataToSheet(data) {
    sheet.clearContents();
    sheet.clearFormats();

    const headers = Settings.MERCHANT_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      const boolFields = ["delete", "modify"];

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
  // ФУНКЦИИ СОЗДАНИЯ И ВАЛИДАЦИИ МЕСТ
  //═══════════════════════════════════════════════════════════════════════════

  // Создаёт объект места из строки листа
  function createMerchantFromRow(row, i, ts, existingIds, mode, merchantsMap) {
    const currentId = row[fieldIndex.id];
    const titleRaw = row[fieldIndex.title];
    const title = (titleRaw != null) ? String(titleRaw).trim() : '';
    if (!title) return null;

    const user = Dictionaries.getUserId(row[fieldIndex.user]) || Settings.DefaultUserId;
    const shouldModify = parseBool(row[fieldIndex.modify]);
    const shouldDelete = parseBool(row[fieldIndex.delete]);
    const existsOnServer = existingIds.has(currentId);
    
    if (mode === UPDATE_MODES.PARTIAL && !shouldModify) return null;

    // Обработка удаления
    if (shouldDelete && currentId && existsOnServer) {
      return {
        deletion: {
          id: currentId,
          object: 'merchant',
          stamp: ts,
          user: user
        }
      };
    }

    if (!currentId) {
      Logger.log(`Строка ${i+1} пропущена: не определён ID`);
      return null;
    }

    let merchant = merchantsMap.get(currentId) || { id: currentId };

    Object.assign(merchant, {
      title: title,
      user: user,
      changed: ts
    });

    return { merchant };
  }

  // Валидирует место
  function validateMerchant(merchant, existingMerchants = [], deleteRequests = []) {  
    // Проверка обязательных полей  
    if (!merchant.title) {  
      throw new Error('Не заполнено обязательное поле (название)');  
    }
    
    // Проверка уникальности названия  
    const isDuplicate = existingMerchants.some(m =>  
      m.title === merchant.title &&  
      m.id !== merchant.id &&  
      !deleteRequests.some(del => del.id === m.id) // Исключаем места, помеченные к удалению  
    );  
    if (isDuplicate) {  
      throw new Error(`место с названием "${merchant.title}" уже существует`);  
    }
    
    return true;  
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ГЛАВНАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ/ЗАМЕНЫ МЕСТ
  //═══════════════════════════════════════════════════════════════════════════

  // Обновляет места на сервере
  function updateMerchants(values, mode) {
    try {
      const errors = [];
      // Запрашивает текущие данные с сервера
      const json = zmData.RequestForceFetch(['merchant']);
      if (!('merchant' in json)) {
        SpreadsheetApp.getActive().toast('Нет данных о местах', 'Предупреждение');
        return;
      }
      const merchants = json['merchant'] || [];
      const ts = Math.floor(Date.now() / 1000);

      // Создает Set и Map для быстрого поиска
      const existingIds = new Set(merchants.map(m => m.id));
      const merchantsMap = new Map(merchants.map(m => [m.id, m]));

      const { idMap, processedIds, new: newCount, modified: modifiedCount, deleted: deletedCount, merchantsToDelete } = generateIdMap(
        values,
        fieldIndex,
        mode,
        merchantsMap,
        existingIds
      );

      const confirmLines = [];
      if (newCount > 0) confirmLines.push(`Будет создано новых мест: ${newCount}`);
      if (modifiedCount > 0) confirmLines.push(`Будет изменено мест: ${modifiedCount}`);
      if (deletedCount > 0) {
        confirmLines.push(`Будет удалено мест: ${deletedCount}`);
        
        // Проверяем транзакции с удаляемыми местами
        if (merchantsToDelete.length > 0) {
          const transactionsInfo = checkTransactionsWithMerchants(merchantsToDelete);
          if (transactionsInfo?.affcount > 0) {
            const { affcount, sample } = transactionsInfo;
            confirmLines.push(
              `\nВНИМАНИЕ: Найдено ${affcount} транзакций с удаляемыми местами.`,
              `Примеры: \n${sample.slice(0, 3).join(';\n')}${sample.length > 3 ? `\n...и ещё ${affcount - 3}` : ''}\n`,
              `Поле будет очищено у ${affcount} транзакций.`
            );
          }
        }
      }

      let confirmMessage = confirmLines.length > 0 ? confirmLines.join('\n') : 'Мест для изменения не выбрано';

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

      // Обновляем ID в таблице для новых мест
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
      const deleteRequests = [];
      const deletedTitles = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const result = createMerchantFromRow(row, i, ts, existingIds, mode, merchantsMap);
        if (!result || !result.deletion) continue;

        deleteRequests.push(result.deletion);
        deletedTitles.push(row[fieldIndex.title]);
      }

      // В режиме REPLACE добавляем запросы на удаление лишних мест
      if (mode.id === UPDATE_MODES.REPLACE.id) {
        for (const [id, merchant] of merchantsMap) {
          if (!processedIds.has(id)) {
            deleteRequests.push({
              id: id,
              object: 'merchant',
              stamp: ts,
              user: merchant.user
            });
            deletedTitles.push(merchant.title);
          }
        }
      }

      // Второй проход: создание новых мест с учетом удалений
      const modifyRequests = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const result = createMerchantFromRow(row, i, ts, existingIds, mode, merchantsMap);
        if (!result || result.deletion) continue;

        try {
          validateMerchant(result.merchant, merchants, deleteRequests);
          modifyRequests.push(result.merchant);
        } catch (error) {
          errors.push(`Строка ${i+2}: ${error.message}`);
        }
      }

      // Уведомление об удалении
      if (deletedTitles.length > 0) {
        const deletionMsg = deletedTitles.length === 1
          ? `Будет удалено место: ${deletedTitles[0]}`
          : `Будут удалены места:\n${deletedTitles.slice(0, 5).join(',\n')}${deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''}`;
        Logger.log(`Cписок удаляемых мест (${deletedTitles.length}):\n${deletedTitles.join(',\n')}`);          
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление мест');
        Utilities.sleep(2000);
      }

      // Отправка изменений
      if (modifyRequests.length > 0 || deleteRequests.length > 0) {
        const data = {
          currentClientTimestamp: ts,
          serverTimestamp: ts
        };
        if (modifyRequests.length > 0) data.merchant = modifyRequests;
        if (deleteRequests.length > 0) data.deletion = deleteRequests;

        SpreadsheetApp.getActive().toast('Отправляем изменения на сервер...', 'Обновление');
        const result = zmData.Request(data);

        if (typeof Logs !== 'undefined' && Logs.logApiCall) {
          Logs.logApiCall(mode.logType, data, result);
        }

        // Проверка ответа сервера 
        if (!result || Object.keys(result).length === 0 || 
          (Object.keys(result).length === 1 && 'serverTimestamp' in result)) {
          throw new Error('Пустой ответ сервера при обновлении мест');
        }

        // Сброс флагов modify
        const modifyColumn = fieldIndex.modify + 1;
        const modifyRange = sheet.getRange(2, modifyColumn, values.length, 1);
        const modifyValues = modifyRange.getValues();

        values.forEach((_, i) => {
          modifyValues[i][0] = parseBool(values[i][fieldIndex.modify]) && 
                    (modifyRequests.some(m => m.id === values[i][fieldIndex.id]) || 
                    deleteRequests.some(d => d.id === values[i][fieldIndex.id]))
                    ? false 
                    : modifyValues[i][0];
        });

        modifyRange.setValues(modifyValues);

        // Если все места обработаны без ошибок, перезагружает список мест
        if (!modifyValues.some(row => parseBool(row[0]))) {  
          doLoad(false); // не показывать toast при перезагрузке после обновления
        }

        SpreadsheetApp.getActive().toast(
          `${mode.id === UPDATE_MODES.REPLACE.id ? 'Замена' : 'Обновление'} завершено:\n` +
          `Новых/изменённых: ${modifyRequests.length}\n` +
          `Удалено: ${deleteRequests.length}` +
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
        errorSheet.getRange(1, 1).setValue(`Ошибки ${mode.description} мест ` + new Date().toLocaleString());
        errorSheet.getRange(2, 1, errors.length, 1).setValues(errors.map(e => [e]));
      }
    } catch (error) {
      Logger.log("Ошибка при обновлении мест: " + error.toString());
      SpreadsheetApp.getActive().toast(error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  // Загружает места с сервера и подготавливает данные для листа
  function doLoad(showToast = true) {
    if (!initialize()) return;

    Dictionaries.loadDictionariesFromSheet();
    
    const json = zmData.RequestForceFetch(['merchant']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_MERCHANTS", { merchant: [] }, json);
    }
    prepareData(json, showToast);
  }

  // Основная функция обновления мест
  function doUpdate(mode) {
    if (!initialize()) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('Нет данных для обновления мест');
      SpreadsheetApp.getActive().toast('Нет данных для обновления мест', 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast(`Начинаем ${mode.description} мест...`, 'Обновление');

    Dictionaries.loadDictionariesFromSheet();

    const dataRange = sheet.getRange(2, 1, lastRow - 1, Settings.MERCHANT_FIELDS.length);
    let values = dataRange.getValues();
    let validRowIndices = getRowsToModify(values, mode);

    if (validRowIndices.length > 0) {
      insertCheckboxesBatchWithValue(sheet, fieldIndex.modify + 1, validRowIndices, true);
      validRowIndices.forEach(row => values[row-2][fieldIndex.modify] = true);
    }

    updateMerchants(values, mode);
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
