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

  let initialized = false;
  let sheet = null;
  let fieldIndex = {};

  function initialize() {
    if (initialized) return true;
    
    try {
      // Получение листа категорий из настроек
      sheet = sheetHelper.GetSheetFromSettings('TAGS_SHEET');
      if (!sheet) {
        Logger.log("Лист с категориями не найден");
        SpreadsheetApp.getActive().toast('Лист с категориями не найден', 'Ошибка');
        return false;
      }

      // Определяет поля по id для удобства доступа
      Settings.CATEGORY_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);
      
      initialized = true;
      return true;
    } catch (e) {
      Logger.log("Ошибка инициализации Categories: " + e.toString());
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

  function highlightDeletedTags() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const deletedTags = JSON.parse(scriptProperties.getProperty('deletedTags') || '[]');
    if (deletedTags.length === 0) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const idColumn = fieldIndex.id + 1;
    const idRange = sheet.getRange(2, idColumn, lastRow - 1, 1);
    const ids = idRange.getValues();
    
    const textStyles = ids.map(row => {
      const id = row[0];
      return [deletedTags.includes(id) ? 'red' : null];
    });

    // Применяем стили текста ко всем столбцам
    const columnsCount = Settings.CATEGORY_FIELDS.length;
    for (let col = 1; col <= columnsCount; col++) {
      sheet.getRange(2, col, lastRow - 1, 1).setFontColors(textStyles);
    }
  }

  function cleanDeletedTags(serverTags) {
    const scriptProperties = PropertiesService.getScriptProperties();
    const deletedTags = JSON.parse(scriptProperties.getProperty('deletedTags') || '[]');
    if (deletedTags.length === 0) return;

    // Создаем Set из актуальных ID с сервера
    const serverTagIds = new Set(serverTags.map(tag => tag.id));

    // Фильтруем deletedTags, оставляя только те, что есть на сервере
    const updatedDeletedTags = deletedTags.filter(id => serverTagIds.has(id));

    if (updatedDeletedTags.length === 0) {
      scriptProperties.deleteProperty('deletedTags');
    } else if (updatedDeletedTags.length !== deletedTags.length) {
      scriptProperties.setProperty('deletedTags', JSON.stringify(updatedDeletedTags));
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ РАБОТЫ С ДАННЫМИ
  //═══════════════════════════════════════════════════════════════════════════
  // Генерирует карту соответствия старых и новых ID категорий, считает статистику изменений
  function generateIdMap(values, fieldIndex, mode, tagsMap, existingIds) {
    const idMap = {};
    const processedIds = new Set();
    const stats = { new: 0, modified: 0, deleted: 0 };
    const tagsToDelete = [];
    const childTagsToDelete = []; // Отдельный массив для дочерних категорий

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
      } else {
          // Составляем полную карту существующих категорий
          idMap[currentId] = currentId;
          // Подсчитываем категории с флагом modify
          if (shouldModify) {
            processedIds.add(currentId);
            stats.modified++;}
      }
    }

    // В режиме REPLACE удаляем всё, чего нет в таблице
    if (mode.id === UPDATE_MODES.REPLACE.id) {
      tagsMap.forEach((tag, id) => {
        if (!processedIds.has(id)) {
          tagsToDelete.push(id);
          stats.deleted++;
        }
      });
    }

    // Дети родительских категорий также будут удалены
    tagsToDelete.forEach(parentId => {
      Array.from(tagsMap.values()).filter(t => t.parent === parentId)
        .filter(child => !tagsToDelete.includes(child.id))
        .forEach(child => {
          const childRow = values.find(r => r[fieldIndex.id] === child.id);
          if (!childRow) return;
          
          const shouldModify = parseBool(childRow[fieldIndex.modify]);
          const sameParent = child.parent === childRow[fieldIndex.parent];
          const parentMarkedForDelete = tagsToDelete.includes(child.parent);
          
          if ((shouldModify && sameParent && parentMarkedForDelete) || !shouldModify) {
            childTagsToDelete.push(child.id);
            stats.deleted++;
            if (shouldModify && sameParent) stats.modified--;
          }
        });
    });

    tagsToDelete.push(...childTagsToDelete);

    if (tagsToDelete.length > 0) {
      const scriptProperties = PropertiesService.getScriptProperties();
      const existingDeletedTags = JSON.parse(scriptProperties.getProperty('deletedTags') || '[]');
      const updatedDeletedTags = [...new Set([...existingDeletedTags, ...tagsToDelete])];
      scriptProperties.setProperty('deletedTags', JSON.stringify(updatedDeletedTags));
    }

    return { ...stats, idMap, processedIds, tagsToDelete, deletedChildTagsCount: childTagsToDelete.length };
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
    if (!initialize()) return;

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

      let affectedCount = 0;
      const affectedTransactions = [];

      json.transaction.forEach(t => {
        if (!t.deleted && t.tag && t.tag.some(tagId => tagIds.includes(tagId))) {
          affectedCount++;
          if (affectedTransactions.length < 20) {
            affectedTransactions.push(t);
          }
        }
      });

      return affectedTransactions.length > 0 ? {
        count: affectedCount,
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
      //sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      const batchSize = 500;
      for (let i = 0, batchNum = 1; i < data.length; i += batchSize, batchNum++) {
        const batch = data.slice(i, i + batchSize);
        sheet.getRange(i + 2, 1, batch.length, data[0].length)
          .setValues(batch);
        Utilities.sleep(100);
      }

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
    // Быстро собираем колонки с чекбоксами
    const checkboxColumns = boolFields
      .map(id => fieldIndex[id])
      .filter(col => col !== undefined);
    
    if (!checkboxColumns.length || sheet.getMaxRows() <= dataLength + 2) return;
    
    const minCol = Math.min(...checkboxColumns) + 1;
    const maxCol = Math.max(...checkboxColumns) + 1;
    const startRow = dataLength + 2;
    const numRows = sheet.getMaxRows() - startRow + 1;
    const numCols = maxCol - minCol + 1;
    
    sheet.getRange(startRow, minCol, numRows, numCols)
      .clearContent()
      .clearDataValidations();
  }

  // Устанавливает цвета фона ячеек
  function setColorBackgrounds(data) {
    const backgrounds = [];

    for (let i = 0; i < data.length; i++) {
      const colorNum = data[i][fieldIndex.color];
      if (colorNum && colorNum > 0) {
        let num = colorNum >>> 0;
        const b = num & 0xFF;
        const g = (num & 0xFF00) >>> 8;
        const r = (num & 0xFF0000) >>> 16;
        const hexColor = "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
        backgrounds.push([hexColor]);
      } else {
        backgrounds.push([null]);
      }
    }
    const colorRange = sheet.getRange(2, fieldIndex.color + 1, data.length, 1);
    colorRange.setBackgrounds(backgrounds);
    colorRange.setValue("");
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
  function createTagFromRow(row, i, ts, existingIds, mode, idMap, tagsMap, colorHex) {
    const currentId = row[fieldIndex.id];
    const oldParentId = row[fieldIndex.parent] || null;
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

    let tag = tagsMap.get(currentId) || { id: currentId };

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

      const { idMap, processedIds, new: newCount, modified: modifiedCount, deleted: deletedCount, tagsToDelete, deletedChildTagsCount } = generateIdMap(
        values,
        fieldIndex,
        mode,
        tagsMap,
        existingIds
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
              `\nПримеры: \n${transactionsInfo.sample.slice(0, 3).join(';\n')}${
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

      const colorHexes = sheet.getRange(2, fieldIndex.color + 1, values.length, 1).getBackgrounds();
      const modifyFlags = values.map(row => parseBool(row[fieldIndex.modify]));
      const deleteFlags = values.map(row => parseBool(row[fieldIndex.delete]));

      const deleteRequests = [];
      const modifyRequests = [];
      const deletedTitles = [];

      for (let i = 0; i < values.length; i++) {

        if (!modifyFlags[i]) continue;

        const row = values[i];
        const result = createTagFromRow(row, i, ts, existingIds, mode, idMap, tagsMap, colorHexes[i][0]);
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
        let deletionMsg = deletedTitles.length === 1 
          ? `Будет удалена категория: ${deletedTitles[0]}`
          : `Будут удалены категории:\n${deletedTitles.slice(0, 5).join(',\n')}${
              deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''
            }`;
        
        if (deletedChildTagsCount > 0) {
          deletionMsg += `, а также ${deletedChildTagsCount} дочерних`;
        }
        
        Logger.log(`Удаление ${deletedTitles.length} категорий:\n${deletedTitles.join('\n')}${deletedChildTagsCount > 0 ? ` + ${deletedChildTagsCount} дочерних` : ''}`);
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление категорий');
        Utilities.sleep(2000);
      }

      // Отправка изменений
      if (modifyRequests.length > 0 || deleteRequests.length > 0) {
        const data = {
          currentClientTimestamp: ts,
          serverTimestamp: ts
        };

        if (modifyRequests.length > 0) {
          data.tag = modifyRequests;
          SpreadsheetApp.getActive().toast('Отправляем изменения категорий...', 'Обновление');
          const modifyResult = zmData.Request(data);
          
          if (typeof Logs !== 'undefined' && Logs.logApiCall) {
            Logs.logApiCall(mode.logType, data, modifyResult);
          }

          if (!modifyResult || Object.keys(modifyResult).length === 0 || 
            (Object.keys(modifyResult).length === 1 && 'serverTimestamp' in modifyResult)) {
            throw new Error('Пустой ответ сервера при обновлении категорий');
          }
        }

        if (deleteRequests.length > 0) {
          data.deletion = deleteRequests;
          SpreadsheetApp.getActive().toast('Отправляем удаление категорий...', 'Обновление');
          const deleteResult = zmData.Request(data);
          
          if (typeof Logs !== 'undefined' && Logs.logApiCall) {
            Logs.logApiCall(mode.logType, data, deleteResult);
          }

          if (!deleteResult || Object.keys(deleteResult).length === 0 || 
            (Object.keys(deleteResult).length === 1 && 'serverTimestamp' in deleteResult)) {
            throw new Error('Пустой ответ сервера при удалении категорий');
          }
        }

        // Сброс флагов modify - оптимизированная версия
        const modifyColumn = fieldIndex.modify + 1;
        const modifyRange = sheet.getRange(2, modifyColumn, values.length, 1);

        // Создаем Set для быстрого поиска
        const modifiedIds = new Set([
          ...modifyRequests.map(a => a.id),
          ...deleteRequests.map(d => d.id)
        ]);

        // Готовим новые значения одним проходом
        const newModifyValues = values.map(row => [
          parseBool(row[fieldIndex.modify]) && modifiedIds.has(row[fieldIndex.id]) ? false : row[fieldIndex.modify]
        ]);

        // Применяем изменения одним вызовом
        modifyRange.setValues(newModifyValues);

        // Проверяем нужно ли перезагружать
        if (newModifyValues.every(row => !parseBool(row[0]))) {
          doLoad(false);
        }

        // Подсчёт реальных изменений
        SpreadsheetApp.getActive().toast(
        `${mode.id === UPDATE_MODES.REPLACE.id ? 'Замена' : 'Обновление'} завершено:\n` +
        `Новых/изменённых: ${modifyRequests.length}\n` +
        `Удалено: ${deleteRequests.length + deletedChildTagsCount}` +
        (errors.length > 0 ? `\nОшибок: ${errors.length}` : ''),
        'Успех'
      );
      } else {
        SpreadsheetApp.getActive().toast('Нет изменений для обработки', 'Информация');
      }
      // Логирование ошибок
      if (errors.length > 0) {
        const errorSheet = sheetHelper.Get(Settings.SHEETS.ERRORS);
        errorSheet.insertRowsBefore(1, errors.length + 1);
        errorSheet.getRange(1, 1).setValue(`Ошибки ${mode.description} категорий ` + new Date().toLocaleString());
        errorSheet.getRange(2, 1, errors.length, 1).setValues(errors.map(e => [e]));
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
    if (!initialize()) return;

    Dictionaries.loadDictionariesFromSheet();

    const json = zmData.RequestForceFetch(['tag']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_TAGS", { tag: [] }, json);
    }
    prepareData(json, showToast);
    cleanDeletedTags(json.tag || []);
    highlightDeletedTags();
  }

  // Основная функция обновления категорий
  function doUpdate(mode) {
    if (!initialize()) return;

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

  //═══════════════════════════════════════════════════════════════════════════
  // ЗАМЕНА КАТЕГОРИЙ В ОПЕРАЦИЯХ
  //═══════════════════════════════════════════════════════════════════════════
  function showCategoryReplacementDialog() {
    if (!initialize()) return;

    const categories = getSortedCategories();
    if (!categories.length) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось загрузить категории', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    try {
      const html = HtmlService.createTemplateFromFile('CategoryReplacementDialog');
      html.categories = categories;
      SpreadsheetApp.getUi().showModalDialog(html.evaluate()
        .setWidth(500)
        .setHeight(350), 'Заменить категорию в операциях');
    } catch (error) {
      Logger.log('Не удалось открыть диалоговое окно: ' + error.message);
      SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось открыть диалоговое окно', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }

  function getSortedCategories() {
    const { tag: tags } = zmData.RequestForceFetch(['tag']) || {};
    if (!tags) {
      Logger.log('Не удалось получить список категорий');
      return [];
    }
    
    const tagsMap = new Map(tags.map(tag => [tag.id, tag]));
    
    return tags.map(tag => {
      const parentTitle = tag.parent ? tagsMap.get(tag.parent)?.title || null : null;
      
      return {
        id: tag.id,
        title: tag.title,
        parentTitle: parentTitle,
        fullTitle: parentTitle ? `${parentTitle} / ${tag.title}` : tag.title
      };
    }).sort((a, b) => a.fullTitle.localeCompare(b.fullTitle));
  }

  function handleCategoryReplacement(oldId, newId) {
    if (!initialize()) return;
    try {
      if (!oldId || !newId || oldId === newId) {
        throw new Error(oldId === newId ? 'Нельзя заменять одинаковые категории' : 'Не выбраны категории');
      }

      const { transaction: transactions } = zmData.RequestForceFetch(['transaction']) || {};
      if (!transactions) throw new Error('Не удалось загрузить транзакции');

      const toUpdate = transactions
        .filter(t => !t.deleted && t.tag?.includes(oldId))
        .map(t => ({ ...t, tag: t.tag.map(id => id === oldId ? newId : id), changed: Math.floor(Date.now()/1000) }));

      if (!toUpdate.length) {
        Logger.log('Нет операций с выбранной категорией');
        return { success: true, count: 0, message: 'Нет операций с выбранной категорией' };
      }

      Logger.log(`Начало замены категории ${oldId} → ${newId} (${toUpdate.length} операций)`);
      
      const result = zmData.Request({
        currentClientTimestamp: Math.floor(Date.now()/1000),
        serverTimestamp: Math.floor(Date.now()/1000),
        transaction: toUpdate
      });

      if (!result.serverTimestamp) throw new Error('Сервер не подтвердил изменения');
      
      Logger.log(`Успешно заменено категорий: ${toUpdate.length}`);
      return { success: true, count: toUpdate.length, message: `Заменено в ${toUpdate.length} операциях` };
    } catch (error) {
      Logger.log('Ошибка замены категорий: ' + error.message);
      return { error: true, message: error.message };
    }
  }

  function addMenuItems(menu) {
    if (!initialize()) return;

    try {
      HtmlService.createTemplateFromFile('CategoryReplacementDialog');
      menu.addSeparator().addItem('Заменить в операциях', 'Categories.showCategoryReplacementDialog');
    } catch(e) {
      Logger.log('HTML-диалог замены недоступен, пункт меню не добавлен: ' + e.message);
    }
  }

  return {
    doLoad,
    doSave: () => doUpdate(UPDATE_MODES.SAVE),
    doPartial: () => doUpdate(UPDATE_MODES.PARTIAL),
    doReplace: () => doUpdate(UPDATE_MODES.REPLACE),
    showCategoryReplacementDialog,
    handleCategoryReplacement,
    addMenuItems
  };
})();
