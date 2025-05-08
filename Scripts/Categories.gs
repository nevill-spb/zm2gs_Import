const Categories = (function () {
  //═══════════════════════════════════════════════════════════════════════════
  // ИНИЦИАЛИЗАЦИЯ
  //═══════════════════════════════════════════════════════════════════════════
  const sheet = sheetHelper.GetSheetFromSettings('TAGS_SHEET');
  if (!sheet) {
    Logger.log("Лист с категориями не найден");
    SpreadsheetApp.getActive().toast('Лист с категориями не найден', 'Ошибка');
    return;
  }

  // Индексы полей по id для удобства доступа
  const fieldIndex = {};
  Settings.CATEGORY_FIELDS.forEach((f, i) => fieldIndex[f.id] = i);

  //═══════════════════════════════════════════════════════════════════════════
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  //═══════════════════════════════════════════════════════════════════════════
  
  // Парсеры для булевых значений
  function parseBool(value) {
    return value === true || value === "TRUE";
  }

  function parseBoolRequired(value) {
    return value || value === null ? "TRUE" : "FALSE";
  }

  // Преобразует hex-цвет в формат ZenMoney (число)
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
  
  // Построение строки данных по полям из Settings
  function buildRow(tag) {
    return Settings.CATEGORY_FIELDS.map(field => {
      switch (field.id) {
        case "showIncome":
        case "showOutcome":
        case "budgetIncome":
        case "budgetOutcome":
          return parseBool(tag[field.id]) ? "TRUE" : "FALSE";
        case "required":
          return parseBoolRequired(tag[field.id]);
        case "delete":
        case "modify":
          return false; // чекбоксы пустые
        case "color":
          return tag.color != null ? tag.color : "";
        case "user":
        case "changed":
          return tag[field.id] !== undefined ? tag[field.id] : "";
        default:
          return tag[field.id] || "";
      }
    });
  }

  // Подготовка данных категорий для записи в лист
  function prepareData(json) {
    if (!('tag' in json)) {
      SpreadsheetApp.getActive().toast('Нет данных о категориях', 'Предупреждение');
      return;
    }

    // Сортируем категории по названию и id для стабильности
    const sortedTags = json.tag.slice().sort((a, b) => {
      if (a.title === b.title) return a.id.localeCompare(b.id);
      return a.title.localeCompare(b.title);
    });

    // Создаем словарь для быстрого поиска дочерних категорий по parent id
    const childrenMap = {};
    sortedTags.forEach(tag => {
      if (tag.parent) {
        if (!childrenMap[tag.parent]) childrenMap[tag.parent] = [];
        childrenMap[tag.parent].push(tag);
      }
    });

    // Формируем массив данных с группировкой: родительские + дочерние
    const data = [];
    sortedTags.forEach(tag => {
      if (!tag.parent) {
        data.push(buildRow(tag));
        const children = childrenMap[tag.id] || [];
        children.forEach(child => data.push(buildRow(child)));
      }
    });

    writeDataToSheet(data);
    SpreadsheetApp.getActive().toast(`Загружено ${data.length} категорий`, 'Информация');
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
    const headers = Settings.CATEGORY_FIELDS.map(f => f.title);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

      // Список id булевых полей, для которых нужны чекбоксы
      const boolFields = ["showIncome", "showOutcome", "budgetIncome", "budgetOutcome", "required", "delete", "modify"];

      // Вставляем чекбоксы по каждой колонке
      boolFields.forEach(fieldId => {
        const colIndex = fieldIndex[fieldId];
        if (colIndex === undefined) return;
        sheet.getRange(2, colIndex + 1, data.length, 1).insertCheckboxes();
      });

      // Очищаем лишние чекбоксы и валидации
      clearExtraRows(data.length, boolFields);

      // Устанавливаем цвета фона
      setColorBackgrounds(data);
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

  // Установка цветов фона
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

  //═══════════════════════════════════════════════════════════════════════════
  // ФУНКЦИИ СОЗДАНИЯ И ОБНОВЛЕНИЯ КАТЕГОРИЙ
  //═══════════════════════════════════════════════════════════════════════════

  // Создание объекта категории из строки данных
  function createTagFromRow(row, i, ts, tags, idMap) {
    const oldTagId = row[fieldIndex.id];
    const oldParentId = row[fieldIndex.parent] || null;
    const title = row[fieldIndex.title];
    if (!title || typeof title !== 'string' || title.trim() === '') return null;
    const colorHex = sheet.getRange(i + 2, fieldIndex.color + 1).getBackground();
    const user = row[fieldIndex.user] || Number(Settings.DefaultUserId);
    const deleteFlag = row[fieldIndex.delete] === true;

    if (deleteFlag && oldTagId) {
      return {
        deletion: {
          id: oldTagId,
          object: 'tag',
          stamp: ts,
          user: user
        }
      };
    }

    const newTagId = idMap[oldTagId || `new_${i}`];
    if (!newTagId) return null;

    const newParentId = (oldParentId && idMap[oldParentId]) ? idMap[oldParentId] : null;

    let tag = tags.find(t => t.id === newTagId) || {
      id: newTagId,
      user: user,
      changed: ts
    };

    // Обновляем поля тега
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

  // Обновление категорий
  function updateTags(values, isPartial = false) {
    try {
      const json = zmData.RequestForceFetch(['tag']);
      const tags = json['tag'] || [];
      const ts = Math.floor(Date.now() / 1000);

      const newTags = [];
      const deletionRequests = [];
      const deletedTitles = [];
      const idMap = {};
      const newIds = Array(values.length).fill(null);

      // Первый проход: определяем новые ID
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (isPartial && !row[fieldIndex.modify]) continue;

        const oldTagId = row[fieldIndex.id];
        const title = row[fieldIndex.title];
        if (!title || typeof title !== 'string' || title.trim() === '') continue;

        if (!oldTagId || !tags.find(t => t.id === oldTagId)) {
          const newId = Utilities.getUuid().toLowerCase();
          idMap[oldTagId || `new_${i}`] = newId;
          newIds[i] = newId;
        } else {
          idMap[oldTagId] = oldTagId;
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

      // Второй проход: создаем/обновляем категории
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (isPartial && !row[fieldIndex.modify]) continue;

        const result = createTagFromRow(row, i, ts, tags, idMap);
        if (!result) continue;

        if (result.deletion) {
          deletionRequests.push(result.deletion);
          // Сохраняем название удаляемой категории
          const title = row[fieldIndex.title];
          if (title) deletedTitles.push(title);
        } else if (result.tag) {
          newTags.push(result.tag);
        }
      }

      // Предупреждение об удалении категорий
      if (deletedTitles.length > 0) {
        const deletionMsg = deletedTitles.length === 1
          ? `Будет удалена категория: ${deletedTitles[0]}`
          : `Будут удалены категории:\n${deletedTitles.slice(0, 5).join('\n')}${
              deletedTitles.length > 5 ? `\n...и еще ${deletedTitles.length - 5}` : ''
          }`;
        SpreadsheetApp.getActive().toast(deletionMsg, 'Удаление категорий');
        // Небольшая пауза, чтобы пользователь успел увидеть сообщение
        Utilities.sleep(2000);
      }

      // Отправляем изменения на сервер
      if (newTags.length > 0 || deletionRequests.length > 0) {
        const data = {
          currentClientTimestamp: ts,
          serverTimestamp: ts
        };

        if (newTags.length > 0) data.tag = newTags;
        if (deletionRequests.length > 0) data.deletion = deletionRequests;
        SpreadsheetApp.getActive().toast('Отправляем изменения на сервер...', 'Обновление');
        const result = zmData.Request(data);
        Logger.log(`Результат ${isPartial ? 'частичного' : 'полного'} обновления категорий: ${JSON.stringify(result)}`);
        if (typeof Logs !== 'undefined' && Logs.logApiCall) {
          Logs.logApiCall("UPDATE_TAGS", data, result);
        }
        // Показываем итоговое сообщение
        SpreadsheetApp.getActive().toast(
          `Обновление завершено:\n` +
          `Новых/измененных: ${newTags.length}\n` +
          `Удалено: ${deletionRequests.length}`,
          'Успех'
        );

        // Сбрасываем флаги modify после успешного обновления
        if (isPartial) {
          const modifyColumn = fieldIndex.modify + 1;
          for (let i = 0; i < values.length; i++) {
            if (values[i][fieldIndex.modify] === true) {
              sheet.getRange(i + 2, modifyColumn).setValue(false);
            }
          }
        }

        doLoad();
      } else {
        SpreadsheetApp.getActive().toast('Нет изменений для обработки', 'Информация');
      }
    } catch (error) {
      Logger.log("Ошибка при обновлении категорий: " + error.toString());
      SpreadsheetApp.getActive().toast("Ошибка: " + error.toString(), "Ошибка");
    }
  }

  //═══════════════════════════════════════════════════════════════════════════
  // ПУБЛИЧНЫЕ МЕТОДЫ
  //═══════════════════════════════════════════════════════════════════════════

  // Загрузка категорий
  function doLoad() {
    Dictionaries.loadDictionariesFromSheet();
    const json = zmData.RequestForceFetch(['tag']);
    if (typeof Logs !== 'undefined' && Logs.logApiCall) {
      Logs.logApiCall("FETCH_TAGS", { tag: [] }, json);
    }
    prepareData(json);
  }

  // Полное обновление категорий
  function doSave() {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      const msg = "Нет данных для обновления категорий";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast('Начинаем полное обновление категорий...', 'Обновление');
    Dictionaries.loadDictionariesFromSheet();
    const values = sheet.getRange(2, 1, lastRow - 1, Settings.CATEGORY_FIELDS.length).getValues();
    updateTags(values, false); // Передаем данные с флагом isPartial = false
  }

  // Частичное обновление категорий
  function doPartial() {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      const msg = "Нет данных для обновления категорий";
      Logger.log(msg);
      SpreadsheetApp.getActive().toast(msg, 'Предупреждение');
      return;
    }

    SpreadsheetApp.getActive().toast('Начинаем частичное обновление категорий...', 'Обновление');
    Dictionaries.loadDictionariesFromSheet();
    const values = sheet.getRange(2, 1, lastRow - 1, Settings.CATEGORY_FIELDS.length).getValues();
    updateTags(values, true); // Передаем данные с флагом isPartial = true
  }

  // Регистрация обработчика полной синхронизации
  fullSyncHandlers.push(prepareData);

  return {
    doLoad,
    doSave,
    doPartial
  };
})();
