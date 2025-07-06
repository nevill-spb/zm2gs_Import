const Import = (function () {
  const sheet = sheetHelper.GetSheetFromSettings('IMPORT_SHEET');
  if (!sheet) {
    Logger.log("Лист с данными не найден");
    SpreadsheetApp.getActive().toast('Лист с данными не найден', 'Ошибка');
    return;
  }

  // Функция для преобразования номера столбца в букву
  function columnToLetter(column) {
    let letter = '';
    while (column > 0) {
      const mod = (column - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      column = Math.floor((column - mod - 1) / 26);
    }
    return letter;
  }

  function insertCheckboxesBatchWithValue(sheet, column, rowIndices, valueForRows) {
    if (rowIndices.length === 0) return;

    rowIndices.sort((a, b) => a - b);
    const minRow = rowIndices[0];
    const maxRow = rowIndices[rowIndices.length - 1];
    const totalRows = maxRow - minRow + 1;

    // Вставляем чекбоксы на весь диапазон
    const range = sheet.getRange(minRow, column, totalRows, 1);
    range.insertCheckboxes();

    // Получаем текущие значения чекбоксов в диапазоне
    const currentValues = range.getValues();

    const rowSet = new Set(rowIndices);
    for (let i = 0; i < totalRows; i++) {
      const sheetRow = minRow + i;
      if (rowSet.has(sheetRow)) {
        currentValues[i][0] = valueForRows;  // ставим нужное значение (true или false)
      }
      // для остальных оставляем текущее значение без изменений
    }

    // Записываем обновлённые значения обратно
    range.setValues(currentValues);
  }

  // COLUMNS генерируются из TRANSACTION_FIELDS по id
  const COLUMNS = (function() {
    const columns = {};
    Settings.TRANSACTION_FIELDS.forEach((field, index) => {
      columns[field.id] = index;
    });
    return columns;
  })();

  const UPDATE_COLUMNS = getUpdateColumns();

  // UPDATE_COLUMNS заполняют отсутствующие поля в листе после импорте
  function getUpdateColumns() {
    const fields = Settings.TRANSACTION_FIELDS;
    function findColumnById(id) {
      const index = fields.findIndex(field => field.id === id);
      if (index === -1) throw new Error(`Поле с id="${id}" не найдено в TRANSACTION_FIELDS`);
      return index + 1;
    }
    const result = {};
    for (const id of Settings.IMPORT.UPDATE_COLUMNS_LIST) {
      const col = findColumnById(id);
      result[id.toUpperCase()] = { name: columnToLetter(col), column: col };
    }
    return result;
  }

  // Кэш обратных справочников
  const REVERSE_DICTIONARIES = { accounts: null, instruments: null, tags: null, merchants: null, users: null };

  // Загрузка обратных справочников
  function loadDictionaries() {
    const dictionaries = Dictionaries.loadDictionariesFromSheet();
    if (!dictionaries) throw new Error("Не удалось загрузить справочники");

    // Хранит объекты обратных справочников в формате {'Название'->'id'}
    Object.entries(Dictionaries.getAllReverseDictionaries()).forEach(([key, value]) => {
      REVERSE_DICTIONARIES[key.replace('Rev', '').toLowerCase()] = new Map(Object.entries(value));
    });
  }

  // Парсеры для различных типов данных
  const parsers = {
    value: (value) => {  
      if (value === "" || value === "null" || value === undefined) return null;  
      return String(value);
    },
    number: (value) => {
      if (value === "" || value === "null" || value === undefined) return null;
      const num = Number(value);
      return isNaN(num) ? null : num;
    },
    boolean: (value, defaultValue = null) => {
      if (value === true || value === "TRUE") return true;
      if (value === false || value === "FALSE") return false;
      return defaultValue;
    },
    tags: (value) => {
      if (!value || value === "null" || value === "") return null;
      try {
        const parsed = JSON.parse(value);
        return Array.isArray(parsed) ? parsed : [parsed];
      } catch (e) {
        return value.split(' | ')
          .map(t => t.trim())
          .filter(t => t.length > 0);
      }
    },
    date: (dateValue) => {
      if (!dateValue) return null;
      const date = new Date(dateValue);
      return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
    }
  };

  // Проверка наличия данных в строке
  function hasData(row) {
    if (!row) return false;
    const dateFilled = row[COLUMNS.date] !== "" && row[COLUMNS.date] != null;
    const outcomeValid = (row[COLUMNS.outcomeAccount] !== "" && row[COLUMNS.outcomeAccount] != null) && (row[COLUMNS.outcome] > 0);
    const incomeValid = (row[COLUMNS.incomeAccount] !== "" && row[COLUMNS.incomeAccount] != null) && (row[COLUMNS.income] > 0);
    return dateFilled && (outcomeValid || incomeValid);
  }

  // Безопасное получение ID из справочников
  function getAccountIdSafe(accountName) {
    if (!accountName || accountName === "null") return null;
    const id = REVERSE_DICTIONARIES.accounts.get(accountName);
    if (!id) {
      throw new Error(`Не найден ID для счета "${accountName}". Проверьте наличие счета в справочнике.`);
    }
    return id;
  }

  function getInstrumentIdSafe(instrumentName) {
    if (!instrumentName || instrumentName === "null") return null;
    const id = REVERSE_DICTIONARIES.instruments.get(instrumentName);
    if (!id && instrumentName) {
      throw new Error(`Не найдена валюта "${instrumentName}". Проверьте справочник валют.`);
    }
    return id ? Number(id) : null;
  }

  function getUserIdSafe(userName) {
    if (!userName || userName === "null") return null;
    const id = REVERSE_DICTIONARIES.users.get(userName);
    if (!id && userName) {
      throw new Error(`Не найден пользователь "${userName}". Проверьте справочник пользователей.`);
    }
    return id ? Number(id) : null;
  }

  function getMerchantIdSafe(merchantName) {
    if (!merchantName || merchantName === "null") return null;
    const id = REVERSE_DICTIONARIES.merchants.get(merchantName);
    if (!id) {
      throw new Error(`Не найден ID для места "${merchantName}". Проверьте справочник мест.`);
    }
    return id;
  }

  function processTags(row) {  
    const tagMode = Settings.TagMode; 
      
    let tags;  
    if (tagMode === Settings.TAG_MODES.MULTIPLE_COLUMNS) {  
      // Режим отдельных столбцов  
      tags = [  
        String(row[COLUMNS.tag]),   // Основной тег  
        String(row[COLUMNS.tag1]),  // Дополнительный тег 1  
        String(row[COLUMNS.tag2])   // Дополнительный тег 2  
      ]  
      .filter(tag => tag && tag !== "null" && tag.trim() !== "")  
      .map(tag => tag.trim());  
    } else {  
      // Режим одной строки  
      const tagString = row[COLUMNS.tag];  
      if (!tagString || tagString === "null" || tagString.trim() === "") return null;  
      tags = tagString.split(' | ')  
        .map(t => t.trim())  
        .filter(t => t.length > 0);  
    }  
    
    if (tags.length === 0) return null;  
    
    // Преобразуем названия тегов в ID  
    return tags.map(tag => {  
      const id = REVERSE_DICTIONARIES.tags.get(tag);  
      if (!id) {  
        throw new Error(`Не найден ID для категории "${tag}". Проверьте справочник категорий.`);  
      }  
      return id;  
    });  
  }

  // Создание объекта транзакции из строки данных
  function createTransaction(row, ts) {
    const isNew = !row[COLUMNS.id];
    const id = isNew ? Utilities.getUuid().toLowerCase() : row[COLUMNS.id];
    const deleted = isNew ? false : parsers.boolean(row[COLUMNS.deleted], false);
    
    try {
      const transaction = {
        id: id,
        date: parsers.date(row[COLUMNS.date]),
        tag: processTags(row), // поправлено для универсальности кода
        merchant: getMerchantIdSafe(row[COLUMNS.merchant]),
        comment: parsers.value(row[COLUMNS.comment]),
        outcomeAccount: getAccountIdSafe(row[COLUMNS.outcomeAccount]),
        outcome: parsers.number(row[COLUMNS.outcome]) || 0,
        outcomeInstrument: getInstrumentIdSafe(row[COLUMNS.outcomeInstrument]),
        incomeAccount: getAccountIdSafe(row[COLUMNS.incomeAccount]),
        income: parsers.number(row[COLUMNS.income]) || 0,
        incomeInstrument: getInstrumentIdSafe(row[COLUMNS.incomeInstrument]),
        created: parsers.number(row[COLUMNS.created]) || ts,
        changed: ts,
        user: getUserIdSafe(row[COLUMNS.user]) || Settings.DefaultUserId,
        deleted: deleted,
        viewed: parsers.boolean(row[COLUMNS.viewed], false),
        hold: parsers.boolean(row[COLUMNS.hold]),
        payee: parsers.value(row[COLUMNS.payee]),
        originalPayee: parsers.value(row[COLUMNS.originalPayee]),
        qrCode: parsers.value(row[COLUMNS.qrCode]),
        source: parsers.value(row[COLUMNS.source]),
        opIncome: parsers.number(row[COLUMNS.opIncome]) || 0,
        opOutcome: parsers.number(row[COLUMNS.opOutcome]) || 0,
        opIncomeInstrument: parsers.value(row[COLUMNS.opIncomeInstrument]),
        opOutcomeInstrument: parsers.value(row[COLUMNS.opOutcomeInstrument]),
        incomeBankID: parsers.value(row[COLUMNS.incomeBankID]),
        outcomeBankID: parsers.value(row[COLUMNS.outcomeBankID]),
        latitude: parsers.number(row[COLUMNS.latitude]),
        longitude: parsers.number(row[COLUMNS.longitude]),
        reminderMarker: parsers.value(row[COLUMNS.reminderMarker])
      };

      // Если не указана валюта для основных сумм, используем по умолчанию
      if (!transaction.outcomeInstrument) transaction.outcomeInstrument = Settings.DefaultCurrencyId;
      if (!transaction.incomeInstrument) transaction.incomeInstrument = Settings.DefaultCurrencyId;

      return transaction;
    } catch (error) {
      throw new Error(`Ошибка создания транзакции: ${error.message}`);
    }
  }

  // Валидация транзакции
  function validateTransaction(transaction) {
    const rules = [
      { condition: !transaction.date, message: 'Отсутствует дата' },
      { condition: !transaction.incomeAccount && !transaction.outcomeAccount, message: 'Должен быть указан хотя бы один счет' },
      { condition: transaction.income > 0 && !transaction.incomeAccount, message: 'Указан доход, но не указан счет дохода' },
      { condition: transaction.outcome > 0 && !transaction.outcomeAccount, message: 'Указан расход, но не указан счет расхода' },
      { condition: transaction.incomeAccount && transaction.outcomeAccount && transaction.incomeAccount === transaction.outcomeAccount && !(transaction.income === 0 || transaction.outcome === 0),
        message: 'Указан одновременный расход и доход по счету'},
      { condition: transaction.incomeAccount && transaction.outcomeAccount && transaction.incomeAccount !== transaction.outcomeAccount && !(transaction.income > 0 && transaction.outcome > 0),
        message: 'Перевод с нулевой суммой'}
    ];

    const errors = rules.filter(r => r.condition).map(r => r.message);
    if (errors.length > 0) throw new Error(errors.join('\n'));
    return true;
  }

  // Отправка пакета транзакций на сервер
  function sendBatch(transactions, ts) {
    const requestPayload = {
      currentClientTimestamp: ts,
      serverTimestamp: ts,
      transaction: transactions
    };
    const result = zmData.Request(requestPayload);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {  
      Logs.logApiCall("IMPORT_TRANSACTIONS", requestPayload, result);
    }
    if (result.error) {
      throw new Error(`API Error: ${result.error}`);
    }
    const keys = Object.keys(result);
    if (  
      keys.length === 0 ||   
      (keys.length === 1 && keys[0] === "serverTimestamp")  
    ) {  
      throw new Error("Неудачный импорт. Пустой ответ или только serverTimestamp");  
    }
    return result;
  }

  // Пакетное обновление ячеек
  function applyUpdates(sheet, updates) {
    for (const [key, update] of Object.entries(updates)) {
      const rows = Object.keys(update.values).map(Number).sort((a, b) => a - b);
      if (rows.length === 0) continue;
      // Группируем последовательные строки для пакетной записи
      let batchStart = null, batchValues = [];
      for (let i = 0; i < rows.length; i++) {
        if (batchStart === null) batchStart = rows[i];
        batchValues.push([update.values[rows[i]]]);
        const isLast = i === rows.length - 1;
        const nextIsNotSequential = !isLast && rows[i + 1] !== rows[i] + 1;
        if (isLast || nextIsNotSequential) {
          if (batchValues.length === 1) {
            sheet.getRange(batchStart, update.column).setValue(batchValues[0][0]);
          } else {
            sheet.getRange(batchStart, update.column, batchValues.length, 1).setValues(batchValues);
          }
          batchStart = null;
          batchValues = [];
        }
      }
      update.values = {};
    }
  }

  // Универсальная обработка пакета (batch)
  function processBatch(currentBatch, data, ts, updates, sheet, processedCount, errorCount, errors, totalCount) {
    if (currentBatch.length === 0) return { processedCount, errorCount };
    try {
      const result = sendBatch(currentBatch.map(item => item.transaction), ts);

      // Сброс флага Modified для успешно отправленных
      currentBatch.forEach(({ rowIndex }) => {
        const row = data[rowIndex];
        const isModified = row[COLUMNS.modified] === true || row[COLUMNS.modified] === "TRUE";
        if (isModified) {
          updates[UPDATE_COLUMNS.MODIFIED.name].values[rowIndex + 1] = false;
        }
      });

      applyUpdates(sheet, updates);

      processedCount += currentBatch.length;

      // Показываем прогресс только если это не последний пакет
      if (processedCount < totalCount && processedCount % Settings.IMPORT.PROGRESS_INTERVAL === 0) {  
        showProgress(processedCount, totalCount);  
      }

      if (result && result.serverTimestamp) {
        zmSettings.setTimestamp(result.serverTimestamp);
      }
    } catch (error) {
      errorCount += currentBatch.length;
      errors.push(`Ошибка отправки пакета: ${error.message}`);
    }
    return { processedCount, errorCount };
  }

  // Отображение прогресса
  function showProgress(current, total) {
    SpreadsheetApp.getActive().toast(
      `Обработано ${current} из ${total} операций...`,
      'Прогресс',
      3
    );
  }

  // Основная функция импорта
  function doUpdate() {
    try {
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) throw new Error("Нет данных для импорта");

      // Находим строки для импорта и устанавливаем чекбоксы
      const modifiedRowIndices = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[COLUMNS.id] && hasData(row)) {
          modifiedRowIndices.push(i + 1);
        }
      }

      insertCheckboxesBatchWithValue(sheet, UPDATE_COLUMNS.MODIFIED.column, modifiedRowIndices, true);
      insertCheckboxesBatchWithValue(sheet, UPDATE_COLUMNS.DELETED.column, modifiedRowIndices, false);

      loadDictionaries();

      // Создаём и отправляем новые места перед импортом транзакций
      prepareMerchantsBeforeImport(data);

      const ts = Math.round(Date.now() / 1000);

      // Формируем массив импортируемых строк
      const rowsToProcess = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id = row[COLUMNS.id];
        const isModified = row[COLUMNS.modified] === true || row[COLUMNS.modified] === "TRUE";
        if (hasData(row) && (!id || isModified)) {
          rowsToProcess.push({ row, rowIndex: i });
        }
      }

      if (rowsToProcess.length === 0) {
        SpreadsheetApp.getActive().toast('Нет данных для импорта');
        return;
      }

      // Запрос подтверждения
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Подтверждение',
        `Будет импортировано ${rowsToProcess.length} операций. Продолжить?`,
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return;

      // Инициализация переменных для пакетной обработки
      let processedCount = 0;
      let errorCount = 0;
      const errors = [];
      let currentBatch = [];
      const totalCount = rowsToProcess.length;
      
      // Показываем начальный прогресс
      showProgress(0, totalCount);

      // Структура для группировки обновлений по столбцам
      const updates = Object.values(UPDATE_COLUMNS).reduce((acc, { name, column }) => {
        acc[name] = { column, values: {} };
        return acc;
      }, {});

      // Обработка по пакетам
      for (let idx = 0; idx < rowsToProcess.length; idx++) {
        const { row, rowIndex } = rowsToProcess[idx];
        if (!row) {
          Logger.log(`Пустая строка в rowsToProcess на индексе ${rowIndex}`);
          continue;
        }

        try {
          const transaction = createTransaction(row, ts);
          validateTransaction(transaction);
          currentBatch.push({ transaction, rowIndex });

          // После импорта. Если строка новая, дописываем поля
          if (!row[COLUMNS.id]) {
            Settings.IMPORT.UPDATE_COLUMNS_LIST.forEach(key => {
              const upperKey = key.toUpperCase();
              updates[UPDATE_COLUMNS[upperKey].name].values[rowIndex + 1] =
                key === 'id' ? transaction.id :
                key === 'deleted' || key === 'modified' ? false :
                key === 'created' ? transaction.created :
                key === 'changed' ? transaction.changed :
                Dictionaries.getUserLogin(transaction.user);
            });
          }

          // Отправляем пакет, когда он достиг нужного размера
          if (currentBatch.length >= Settings.IMPORT.BATCH_SIZE) {
            const res = processBatch(currentBatch, data, ts, updates, sheet, processedCount, errorCount, errors, totalCount);
            processedCount = res.processedCount;
            errorCount = res.errorCount;
            currentBatch = []; // Очищаем текущий пакет
          }

        } catch (error) {
          errorCount++;
          errors.push(`Строка ${rowIndex + 1}: ${error.message}`);
          Logger.log(`Ошибка в строке ${rowIndex + 1}: ${error.message}`);
        }
      }

      // Отправляем оставшиеся транзакции
      if (currentBatch.length > 0) {
        const res = processBatch(currentBatch, data, ts, updates, sheet, processedCount, errorCount, errors, totalCount);
        processedCount = res.processedCount;
        errorCount = res.errorCount;
      }

      // Итоговое сообщение
      const successMessage = `Успешно импортировано: ${processedCount}`;
      const errorMessage = errorCount > 0 ? `\nОшибок: ${errorCount}` : '';
      SpreadsheetApp.getActive().toast(successMessage + errorMessage, 'Результат импорта');

      // Логируем ошибки в отдельный лист
      if (errors.length > 0) {
        const errorSheet = sheetHelper.Get(Settings.SHEETS.ERRORS);
        const header = "Ошибки импорта " + new Date().toLocaleString();
        const numNewErrors = errors.length;
        errorSheet.insertRowsBefore(1, numNewErrors + 1);
        errorSheet.getRange(1, 1).setValue(header);
        errorSheet.getRange(2, 1, numNewErrors, 1).setValues(errors.map(e => [e]));
      }

    } catch (error) {
      Logger.log("Ошибка при импорте: " + error.toString());
      if (typeof Logs !== 'undefined' && Logs.logApiCall) {  
        Logs.logApiCall("IMPORT_ERROR", {}, error.toString());
      }
      SpreadsheetApp.getActive().toast('Ошибка: ' + error.toString(), 'Ошибка импорта');
    }
  }

  // Подготовка новых мест перед импортом
  function prepareMerchantsBeforeImport(data) {
    const merchantNamesSet = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];  
      const isModified = row[COLUMNS.modified] === true || row[COLUMNS.modified] === "TRUE";  
      const isNew = !row[COLUMNS.id];
      if (!isModified && !isNew) continue;  // добавляем только для строк, подлежащих импорту

      const merchantName = data[i][COLUMNS.merchant];
      if (merchantName && merchantName !== "null" && merchantName.trim() !== "") {
        merchantNamesSet.add(merchantName.trim());
      }
    }

    const newMerchants = [];

    merchantNamesSet.forEach(name => {
      if (!REVERSE_DICTIONARIES.merchants.has(name)) {
        const newId = Utilities.getUuid().toLowerCase();
        REVERSE_DICTIONARIES.merchants.set(name, newId);
        newMerchants.push({ id: newId, title: name });
      }
    });

    if (newMerchants.length > 0) {
      createMerchantsOnServer(newMerchants);
      if (Dictionaries && Dictionaries.getAllDictionaries) {
        const allDicts = Dictionaries.getAllDictionaries();
        newMerchants.forEach(({id, title}) => {
          allDicts.merchants[id] = title;
        });
      }
      if (Dictionaries && Dictionaries.saveDictionariesToSheet) {
        Dictionaries.saveDictionariesToSheet();
      }
    }
  }

  // Импорт новых мест на сервер
  function createMerchantsOnServer(newMerchants) {
    if (newMerchants.length === 0) return;

    const ts = Math.floor(Date.now() / 1000);
    const merchantObjects = newMerchants.map(({id, title}) => ({
      id,
      title,
      changed: ts,
      user: Settings.DefaultUserId
    }));

    const data = {
      currentClientTimestamp: ts,
      serverTimestamp: ts,
      merchant: merchantObjects
    };

    const result = zmData.Request(data);
    if (result.error) {
      throw new Error(`Ошибка создания мест: ${result.error}`);
    }
    Logger.log(`Создано новых мест: ${newMerchants.length}`);
  }

  return {
    doUpdate
  };
})();	
