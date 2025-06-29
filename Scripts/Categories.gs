const Categories = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // КОНСТАНТЫ
  //═══════════════════════════════════════════════════════════════════════════
  // Режимы обновления категорий
  const UPDATE_MODES = {
    SAVE: { id: 'SAVE', description: 'полное обновление', logType: 'UPDATE_TAGS' },
    PARTIAL: { id: 'PARTIAL', description: 'частичное обновление', logType: 'UPDATE_TAGS' },
    REPLACE: { id: 'REPLACE', description: 'замену', logType: 'REPLACE_TAGS' }
  };

  //═══════════════════════════════════════════════════════════════════════════
  // ИНИЦИАЛИЗАЦИЯ
  //═══════════════════════════════════════════════════════════════════════════
  // Получение листа категорий из настроек
  const sheet = sheetHelper.GetSheetFromSettings('TAGS_SHEET');
  if (!sheet) {
    Logger.log("Лист с категориями не найден");
    SpreadsheetApp.getActive().toast('Лист с категориями не найден', 'Ошибка');
    return;
  }

  // Определяет поля по id для удобства доступа
  const fieldIndex = {};
  Settings.CATEGORY_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  //═══════════════════════════════════════════════════════════════════════════
  // Преобразует значение в булево
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  // Преобразует hex-строку цвета в числовой формат ZenMoney
  function hexToZenColor(hex) {
    if (!hex) return null;
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    if (!result) return null;
    const r = parseInt(result[1], 16);
    const g = parseInt(result[2], 16);
    const b = parseInt(result[3], 16);
    return (r << 16) + (g << 8) + b;
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ДАННЫМИ
  //═══════════════════════════════════════════════════════════════════════════
  // Генерирует карту соответствия старых и новых ID категорий, считает статистику изменений
  function generateIdMap(values, tags, mode, fieldIndex) {
    const idMap = {};
    const processedIds = new Set();
    const stats = { new: 0, modified: 0, deleted: 0 };
    const tagsToDelete = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const currentId = row[fieldIndex.id];
      const title = String(row[fieldIndex.title] || '').trim();
      if (!title) continue;

      const shouldDelete = parseBool(row[fieldIndex.delete]);
      const shouldModify = parseBool(row[fieldIndex.modify]);
      const existsOnServer = currentId && tags.some(t => t.id === currentId);

      // Удаление только существующих с подтверждением
      if (shouldDelete) {
        if (shouldModify && existsOnServer) {
          processedIds.add(currentId);
          tagsToDelete.push(currentId);
          stats.deleted++;
        }
        continue; // Пропускаем все остальные случаи удаления
      }

      if (!existsOnServer) {
          if (!currentId) {
              // Новые категории без ID
              idMap[`new_${i}`] = Utilities.getUuid().toLowerCase();
              stats.new++;
          } else if (mode !== UPDATE_MODES.PARTIAL || shouldModify) {
              // Новые категории с ID
              idMap[currentId] = Utilities.getUuid().toLowerCase();
              stats.new++;
          }
      } else if (shouldModify) {
          // Существующие категорий с флагом modify
          idMap[currentId] = currentId;
          processedIds.add(currentId);
          stats.modified++;
      }
    }

    // В режиме REPLACE удаляем всё, чего нет в таблице
    if (mode.id === UPDATE_MODES.REPLACE.id) {
      tags.forEach(tag => {
        if (!processedIds.has(tag.id)) {
          tagsToDelete.push(tag.id);
          stats.deleted++;
        }
      });
    }

    return { ...stats, idMap, processedIds, tagsToDelete };
  }

  // Формирует строку для записи в лист из объекта категории, учитывая типы полей
  function buildRow(tag) {
    return Settings.CATEGORY_FIELDS.map(field => {
      switch (field.id) {
        case "showIncome":
        case "showOutcome":
        case "budgetIncome":
        case "budgetOutcome":
        case "required":
          return parseBool(tag[field.id]) ? "TRUE" : "FALSE";
        case "delete":
        case "modify":
          return false;
        case "color":
          return tag.color != null ? tag.color : "";
        case "user":
          return Dictionaries.getUserLogin(tag.user) || tag.user || "";
        case "changed":
          return tag[field.id] !== undefined ? tag[field.id] : "";
        default:
          return tag[field.id] || "";
      }
    });
  }

  // Подготавливает данные для записи в лист, группирует категории по родителям и детям
  function prepareData(json, showToast = true) {
    if (!('tag' in json)) {
      if (showToast) SpreadsheetApp.getActive().toast('Нет данных о категориях', 'Предупреждение');
      return;
    }

    const sortedTags = json.tag.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    const childrenMap = {};
    sortedTags.forEach(tag => {
      if (tag.parent) {
        if (!childrenMap[tag.parent]) childrenMap[tag.parent] = [];
        childrenMap[tag.parent].push(tag);
      }
    });

    const data = [];
    sortedTags.forEach(tag => {
      if (!tag.parent) {
        data.push(buildRow(tag));
        const children = childrenMap[tag.id] || [];
        children.forEach(child => data.push(buildRow(child)));
      }
    });

    writeDataToSheet(data);
    if (showToast) SpreadsheetApp.getActive().toast(`Загружено ${data.length} категорий`, 'Информация');
  }

  // Проверяет транзакции, имеющие связи с категориями
  function checkTransactionsWithTags(tagIds) {
    try {
      const json = zmData.RequestForceFetch(['transaction']);
      if (!json.transaction || !Array.isArray(json.transaction)) {
        Logger.log("Нет данных о транзакциях");
        return null;
      }

      const affectedTransactions = json.transaction.filter(t => 
        !t.deleted && t.tag && t.tag.some(tagId => tagIds.includes(tagId))
      ).slice(0, 20); // Ограничиваем выборку для производительности

      return affectedTransactions.length > 0 ? {
        count: affectedTransactions.length,
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
      } : null;
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

    const headers = Settings.CATEGORY_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      const boolFields = ["showIncome", "showOutcome", "budgetIncome", "budgetOutcome", "required", "delete", "modify"];

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
            .insertCheckboxes()
        });
      }

      clearExtraRows(data.length, boolFields);
      setColorBackgrounds(data);
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

  // Устанавливает цвета фона ячеек
  function setColorBackgrounds(data) {
    for (let i = 0; i < data.length; i++) {
      const colorNum = data[i][fieldIndex.color];
      if (colorNum && colorNum > 0) {
        let num = colorNum >>> 0;
        const b = num & 0xFF;
        const g = (num & 0xFF00) >>> 8;
        const r = (num & 0xFF0000) >>> 16;
        const hexColor = "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
        const cell = sheet.getRange(i + 2, fieldIndex.color + 1);
        cell.setBackground(hexColor);
        cell.setValue("");
      } else {
        sheet.getRange(i + 2, fieldIndex.color + 1).setBackground(null);
      }
    }
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
  // ФУНКЦИИ СОЗДАНИЯ И ОБНОВЛЕНИЯ КАТЕГОРИЙ
  //═══════════════════════════════════════════════════════════════════════════
  // Создаёт объект категории из строки листа
  function createTagFromRow(row, i, ts, tags, mode, idMap) {
    const currentId = row[fieldIndex.id];
    const oldParentId = row[fieldIndex.parent] || null;
    const titleRaw = row[fieldIndex.title];
    const title = (titleRaw != null) ? String(titleRaw).trim() : '';
    if (!title) return null;

    const colorHex = sheet.getRange(i + 2, fieldIndex.color + 1).getBackground();
    const user = Dictionaries.getUserId(row[fieldIndex.user]) || Settings.DefaultUserId;
    const shouldModify = parseBool(row[fieldIndex.modify]);
    const shouldDelete = parseBool(row[fieldIndex.delete]);
    const existsOnServer = currentId && tags.some(t => t.id === currentId);    
    
    if (mode === UPDATE_MODES.PARTIAL && !shouldModify) return null;

    // Обработка удаления
    if (shouldDelete && currentId && existsOnServer) {
      return {
        deletion: {
          id: currentId,
          object: 'tag',
          stamp: ts,
          user: user
        }
      };
    }

    if (!currentId) {
      Logger.log(`Строка ${i+1} пропущена: не определён ID`);
      return null;
    }

    const newParentId = (oldParentId && idMap[oldParentId]) ? idMap[oldParentId] : null;

    let tag = tags.find(t => t.id === currentId) || { id: currentId, };

    Object.assign(tag, {
      parent: newParentId,
      title: title,
      color: hexToZenColor(colorHex),
      icon: row[fieldIndex.icon],
      showIncome: parseBool(row[fieldIndex.showIncome]),
      showOutcome: parseBool(row[fieldIndex.showOutcome]),
      budgetIncome: parseBool(row[fieldIndex.budgetIncome]),
      budgetOutcome: parseBool(row[fieldIndex.budgetOutcome]),
      required: parseBool(row[fieldIndex.required]),
      user: user,
      changed: ts
    });

    return { tag };
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ГЛАВНАЯ ФУНКЦИЯ ОБНОВЛЕНИЯ/ЗАМЕНЫ СЧЕТОВ
  //═══════════════════════════════════════════════════════════════════════════

  // Обновляет счета на сервере
  function updateTags(values, mode) {
    try {
      const errors = [];
      // Запрашивает текущие данные с сервера
      const json = zmData.RequestForceFetch(['tag']);
      if (!('tag' in json)) {
        SpreadsheetApp.getActive().toast('Нет данных о категориях', 'Предупреждение');
      }
      const tags = json['tag'] || [];
      const ts = Math.floor(Date.now() / 1000);

      // Создает Set и Map для быстрого поиска
      const existingIds = new Set(tags.map(t => t.id));
      const tagsMap = new Map(tags.map(t => [t.id, t]));

      const { idMap, processedIds, new: newCount, modified: modifiedCount, deleted: deletedCount, tagsToDelete } = generateIdMap(
        values,
        tags,
        mode,
        fieldIndex
      );

      const confirmLines = [];
      if (newCount > 0) confirmLines.push(`Будет создано новых категорий: ${newCount}`);
      if (modifiedCount > 0) confirmLines.push(`Будет изменено категорий: ${modifiedCount}`);
      if (deletedCount > 0) {
        confirmLines.push(`Будет удалено категорий: ${deletedCount}`);
        
        // Проверяем транзакции с удаляемыми категориями
        if (tagsToDelete.length > 0) {
          const transactionsInfo = checkTransactionsWithTags(tagsToDelete);
          if (transactionsInfo) {
            confirmLines.push(
              `\nВНИМАНИЕ: Найдено ${transactionsInfo.count} транзакций с удаляемыми категориями.`,
              `Примеры: \n${transactionsInfo.sample.slice(0, 3).join(';\n')}${
                transactionsInfo.sample.length > 3 
                  ? `\n...и ещё ${transactionsInfo.count - 3}` 
                  : ''
              }`,
              `\nКатегории этих транзакций будут стёрты.`
            );
          }
        }
      }

      const confirmMessage = confirmLines.length > 0 ? confirmLines.join('\n') : 'Категорий для изменения не выбрано';

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

      // Обновляем ID в таблице для новых категорий
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

      const deleteRequests = [];
      const modifyRequests = [];
      const deletedTitles = [];

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const result = createTagFromRow(row, i, ts, tags, mode, idMap);
        if (!result) continue;

        if (result.deletion) {
          deleteRequests.push(result.deletion);
          deletedTitles.push(row[fieldIndex.title]);
        } else if (result.tag) {
          modifyRequests.push(result.tag);
        }
      }

      // В режиме REPLACE добавляем запросы на удаление лишних категорий
      if (mode.id === UPDATE_MODES.REPLACE.id) {
        tags
          .filter(tag => !processedIds.has(tag.id))
          .forEach(tag => {
            deleteRequests.push({
              id: tag.id,
              object: 'tag',
              stamp: ts,
              user: tag.user
            });
            deletedTitles.push(tag.title);
          });
      }

      // Уведомление об удалении
      if (deletedTitles.length > 0) {
        const deletionMsg = deletedTitles.length === 1
          ? `Будет удалена категория: ${deletedTitles[0]}`
          : `Будут удалены категории:\n${deletedTitles.slice(0, 5).join(',\n')}${deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''}`;
        Logger.log(`Cписок удаляемых категорий (${deletedTitles.length}):\n${deletedTitles.join(',\n')}`);          
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление категорий');
        Utilities.sleep(2000);
      }

      // Отправка изменений
      if (modifyRequests.length > 0 || deleteRequests.length > 0) {
        const data = {
          currentClientTimestamp: ts,
          serverTimestamp: ts
        };
        if (modifyRequests.length > 0) data.tag = modifyRequests;
        if (deleteRequests.length > 0) data.deletion = deleteRequests;

        SpreadsheetApp.getActive().toast('Отправляем изменения на сервер...', 'Обновление');
        const result = zmData.Request(data);

        if (typeof Logs !== 'undefined' && Logs.logApiCall) {
          Logs.logApiCall(mode.logType, data, result);
        }

        // Проверка ответа сервера 
        if (!result || Object.keys(result).length === 0 || 
          (Object.keys(result).length === 1 && 'serverTimestamp' in result)) {
          throw new Error('Пустой ответ сервера при обновлении категорий');
        }

        // Сброс флагов modify
        const modifyColumn = fieldIndex.modify + 1;
        const modifyRange = sheet.getRange(2, modifyColumn, values.length, 1);
        const modifyValues = modifyRange.getValues();

        values.forEach((_, i) => {
          modifyValues[i][0] = parseBool(values[i][fieldIndex.modify]) && 
                    (modifyRequests.some(a => a.id === values[i][fieldIndex.id]) || 
                    deleteRequests.some(d => d.id === values[i][fieldIndex.id]))
                    ? false 
                    : modifyValues[i][0];
        });

        modifyRange.setValues(modifyValues);

        // Если все счета обработаны без ошибок, перезагружает список счетов
        if (!modifyValues.some(row => parseBool(row[0]))) {  
          doLoad(false); // не показывать toast при перезагрузке после обновления
        }

        // Подсчёт реальных изменений
        SpreadsheetApp.getActive().toast(
        `${mode.id === UPDATE_MODES.REPLACE.id ? 'Замена' : 'Обновление'} завершено:\n` +
        `Новых/изменённых: ${modifyRequests.length}\n` +
        `Удалено: ${deleteRequests.length}` +
        (errors.length > 0 ? `\nОшибок: ${errors.length}` : ''),
        'Успех'
      );
      } else {
        SpreadsheetApp.getActive().toast('Нет изменений для обработки', 'Информация');
      }
    } catch (error) {
      Logger.log("Ошибка при обновлении категорий: " + error.toString());
      SpreadsheetApp.getActive().toast(error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  // Загружает категории с сервера и подготавливает данные для листа
  function doLoad(showToast = true) {
    Dictionaries.loadDictionariesFromSheet();

    const json = zmData.RequestForceFetch(['tag']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_TAGS", { tag: [] }, json);
    }
    prepareData(json, showToast);
  }

  // Основная функция обновления категорий
  function doUpdate(mode) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('Нет данных для обновления категорий');
      SpreadsheetApp.getActive().toast('Нет данных для обновления категорий', 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast(`Начинаем ${mode.description} категорий...`, 'Обновление');

    Dictionaries.loadDictionariesFromSheet();

    const dataRange = sheet.getRange(2, 1, lastRow - 1, Settings.CATEGORY_FIELDS.length);
    let values = dataRange.getValues();
    let validRowIndices = getRowsToModify(values, mode);

    if (validRowIndices.length > 0) {
      insertCheckboxesBatchWithValue(sheet, fieldIndex.modify + 1, validRowIndices, true);
      validRowIndices.forEach(row => values[row-2][fieldIndex.modify] = true);
    }

    updateTags(values, mode);
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
